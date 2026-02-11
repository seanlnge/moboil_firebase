import { onDocumentCreated, FirestoreEvent } from "firebase-functions/v2/firestore";
import { onCall, HttpsError } from "firebase-functions/v2/https";
import { defineSecret } from "firebase-functions/params";
import * as admin from "firebase-admin";
import sgMail from "@sendgrid/mail";
import { QueryDocumentSnapshot } from "firebase-admin/firestore";

admin.initializeApp();

const SENDGRID_API_KEY = defineSecret("SENDGRID_API_KEY");
const ADMIN_EMAIL = defineSecret("ADMIN_EMAIL");
const SENDGRID_FROM = defineSecret("SENDGRID_FROM");

interface BookingData {
  email?: string;
  firstName?: string;
  lastName?: string;
  dateTime?: number;
  service?: string;
  address?: string;
  notes?: string;
  phone?: string;
  carYear?: string;
  carMake?: string;
  carModel?: string;
  status?: string;
  emailStatus?: {
    customer?: string;
    admin?: string;
  };
  [key: string]: any;
}

// Available time slots as { hour, minute } pairs
const TIME_SLOTS = [
  { hour: 9, minute: 30 },   // 9:30 AM
  { hour: 11, minute: 0 },   // 11:00 AM
  { hour: 13, minute: 0 },   // 1:00 PM
  { hour: 14, minute: 30 },  // 2:30 PM
  { hour: 16, minute: 0 },   // 4:00 PM
  { hour: 17, minute: 30 },  // 5:30 PM
];

/**
 * Returns the UTC offset in hours for America/New_York at a given date.
 * Accounts for EST (-5) vs EDT (-4) automatically.
 */
function getEasternOffsetHours(date: Date): number {
  // Build a formatter that tells us the UTC offset for Eastern time
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/New_York",
    timeZoneName: "shortOffset",
  }).formatToParts(date);

  const tzPart = parts.find((p) => p.type === "timeZoneName");
  // tzPart.value will be something like "GMT-5" or "GMT-4"
  if (tzPart) {
    const match = tzPart.value.match(/GMT([+-]\d+)/);
    if (match) return parseInt(match[1], 10);
  }
  return -5; // Fallback to EST
}

/**
 * Callable function that returns available booking slots.
 * Generates all Mon-Sat slots for the next 3 weeks (after Feb 20, 2026),
 * then removes any slots already taken by existing bookings.
 * All times are generated in America/New_York timezone.
 */
export const getAvailableSlots = onCall(
  { region: "us-east1" },
  async () => {
    const now = Date.now();

    // Build a helper that creates a UTC timestamp for a given
    // Eastern-time calendar date + time-of-day.
    function toEasternTimestamp(year: number, month: number, day: number, hour: number, minute: number): number {
      // Create a date in UTC that *looks like* the Eastern date
      const utc = Date.UTC(year, month, day, hour, minute, 0, 0);
      // Determine the Eastern offset for that moment
      const offset = getEasternOffsetHours(new Date(utc));
      // Shift by the offset so the resulting UTC instant corresponds
      // to the desired Eastern wall-clock time
      return utc - offset * 60 * 60 * 1000;
    }

    // Determine start date in Eastern time
    const nowEastern = new Date(
      new Date().toLocaleString("en-US", { timeZone: "America/New_York" })
    );
    const launchDate = new Date(2026, 1, 21); // Feb 21, 2026
    const startDate = nowEastern > launchDate ? new Date(nowEastern) : new Date(launchDate);

    if (nowEastern > launchDate) {
      startDate.setDate(startDate.getDate() + 1);
    }
    startDate.setHours(0, 0, 0, 0);

    // Generate all possible slots for the next 21 days (Mon-Sat)
    const allSlots: number[] = [];
    const current = new Date(startDate);

    for (let i = 0; i < 21; i++) {
      const dayOfWeek = current.getDay();
      // Skip Sundays (0)
      if (dayOfWeek !== 0) {
        const y = current.getFullYear();
        const m = current.getMonth();
        const d = current.getDate();

        for (const slot of TIME_SLOTS) {
          const ts = toEasternTimestamp(y, m, d, slot.hour, slot.minute);
          // Only include future slots
          if (ts > now) {
            allSlots.push(ts);
          }
        }
      }
      current.setDate(current.getDate() + 1);
    }

    // Get the range boundaries for querying
    const minTime = allSlots.length > 0 ? allSlots[0] : 0;
    const maxTime = allSlots.length > 0 ? allSlots[allSlots.length - 1] : 0;

    if (allSlots.length === 0) {
      return { slots: [] };
    }

    // Query existing bookings that fall within our slot range
    const bookingsSnap = await admin
      .firestore()
      .collection("bookings")
      .where("dateTime", ">=", minTime)
      .where("dateTime", "<=", maxTime)
      .where("status", "in", ["pending", "confirmed"])
      .get();

    const takenTimes = new Set<number>();
    bookingsSnap.forEach((doc) => {
      const data = doc.data() as BookingData;
      if (data.dateTime) {
        takenTimes.add(data.dateTime);
      }
    });

    // Filter out taken slots
    const availableSlots = allSlots.filter((ts) => !takenTimes.has(ts));

    return { slots: availableSlots };
  }
);

export const onBookingCreated = onDocumentCreated(
  {
    document: "bookings/{bookingId}",
    secrets: [SENDGRID_API_KEY, ADMIN_EMAIL, SENDGRID_FROM],
    region: "us-east1",
  },
  async (event: FirestoreEvent<QueryDocumentSnapshot | undefined>) => {
    const snap = event.data;
    if (!snap) return;

    const bookingId = event.params.bookingId;
    const bookingRef = snap.ref;
    const booking = snap.data() as BookingData;

    const customerEmail = booking.email || booking.userEmail;
    if (!customerEmail) {
      // No customer email; still notify admin if you want
      console.log("No customer email found for booking", bookingId);
      return;
    }

    // 1) Transaction: claim the right to send emails (idempotency)
    // If function retries, it won't send twice.
    const claim = await admin.firestore().runTransaction(async (tx) => {
      const fresh = await tx.get(bookingRef);
      const data = (fresh.data() || {}) as BookingData;

      const status = data.emailStatus || {};
      const alreadyCustomer = status.customer === "sent";
      const alreadyAdmin = status.admin === "sent";

      // If both already sent, exit early.
      if (alreadyCustomer && alreadyAdmin) return { shouldSend: false, alreadyCustomer, alreadyAdmin };

      // Mark as "sending" so concurrent retries don't double-send.
      tx.set(
        bookingRef,
        {
          emailStatus: {
            customer: alreadyCustomer ? "sent" : "sending",
            admin: alreadyAdmin ? "sent" : "sending",
          },
          emailUpdatedAt: admin.firestore.FieldValue.serverTimestamp(),
        },
        { merge: true }
      );

      return { shouldSend: true, alreadyCustomer, alreadyAdmin };
    });

    if (!claim.shouldSend) return;

    // Initialize tracking variables
    let customerSent = claim.alreadyCustomer;
    let adminSent = claim.alreadyAdmin;

    // 2) Configure SendGrid
    sgMail.setApiKey(SENDGRID_API_KEY.value());

    const from = SENDGRID_FROM.value();
    const adminEmail = ADMIN_EMAIL.value();

    // --- 4) Build template data (matches your HTML handlebars vars) ---
    const year = new Date().getFullYear();

    // Decipher date and time from dateTime timestamp
    let bookingDate = "";
    let bookingTime = "";
    
    if (booking.dateTime) {
      const dateObj = new Date(booking.dateTime);
      // Format: "Feb 20, 2026"
      bookingDate = dateObj.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
        year: "numeric",
        timeZone: "America/New_York", // Or your preferred timezone
      });
      // Format: "2:00 PM"
      bookingTime = dateObj.toLocaleTimeString("en-US", {
        hour: "numeric",
        minute: "2-digit",
        timeZone: "America/New_York",
      });
    } else {
       // Fallback for old data
       bookingDate = (booking.bookingDate || booking.date || "").toString(); 
       bookingTime = (booking.bookingTime || booking.time || "").toString();
    }

    // Construct name if not present
    const customerName = booking.firstName && booking.lastName ? `${booking.firstName} ${booking.lastName}` : booking.firstName || "there";

    // Car fields (use whatever your form collects)
    const carYear = booking.carYear ?? "";
    const carMake = booking.carMake ?? "";
    const carModel = booking.carModel ?? "";
    const carTrim = booking.carTrim ?? "";
    const carColor = booking.carColor ?? "";
    const carPlate = booking.carPlate ?? "";

    const carSummaryParts = [
      [carYear, carMake, carModel].filter(Boolean).join(" ").trim(),
      carTrim ? `(${carTrim})` : "",
      carColor ? `• ${carColor}` : "",
      carPlate ? `• Plate: ${carPlate}` : "",
    ].filter(Boolean);

    const carSummary = carSummaryParts.join(" ").replace(/\s+/g, " ").trim();

    // Location fields
    const addressLine1 = booking.addressLine1 || booking.address || "";
    const addressLine2 = booking.addressLine2 ?? "";
    const city = booking.city ?? "";
    const state = booking.state ?? "";
    const zip = booking.zip ?? "";
    const parkingNotes = booking.parkingNotes ?? "";

    const locationLine = (
      booking.locationLine ||
      [addressLine1, city && state ? `${city}, ${state}` : city || state, zip].filter(Boolean).join(" ")
    ).toString();

    const specialInstructions = (booking.specialInstructions || booking.notes || "").toString();

    const hasCarOrLocationOrInstructions = Boolean(
      carSummary ||
        carMake ||
        carModel ||
        addressLine1 ||
        locationLine ||
        parkingNotes ||
        specialInstructions
    );

    const dynamicTemplateData = {
      brandName: booking.brandName || "Moboil",
      year,
      name: customerName,
      service: booking.service || "Oil Change",
      bookingDate,
      bookingTime,

      // buttons
      manageUrl: booking.manageUrl || "", // optional
      cancelUrl: `https://moboil.org/booking?cancel=${bookingId}`,

      // car
      carSummary,
      carYear,
      carMake,
      carModel,
      carTrim,
      carColor,
      carPlate,

      // location
      locationLine,
      addressLine1,
      addressLine2,
      city,
      state,
      zip,
      parkingNotes,

      // notes
      specialInstructions,
      hasCarOrLocationOrInstructions,
    };

    // --- 5) Customer email (dynamic template) ---
    const customerMsg = {
      to: customerEmail,
      from,
      templateId: "d-a8eb9e326aa54075b829b7e1fc20458d",
      dynamicTemplateData,
    };

    // --- 6) Admin notification email (simple text) ---
    const adminMsg = {
      to: adminEmail,
      from,
      subject: "New booking received",
      text:
        `New booking\n\n` +
        `Name: ${customerName}\n` +
        `Email: ${customerEmail}\n` +
        `Phone: ${booking.phone || "N/A"}\n` +
        `When: ${bookingDate} ${bookingTime}\n` +
        `Service: ${booking.service || "Oil Change"}\n` +
        `Car: ${carSummary || [carYear, carMake, carModel].filter(Boolean).join(" ")}\n` +
        `Location: ${locationLine}\n` +
        `Address: ${addressLine1}${addressLine2 ? ", " + addressLine2 : ""}, ${city} ${state} ${zip}\n` +
        `Parking: ${parkingNotes}\n` +
        `Special instructions: ${specialInstructions}\n\n` +
        `Booking ID: ${bookingRef.id}\n`,
    };

    try {
      if (!customerSent) {
        await sgMail.send(customerMsg);
        customerSent = true;
      }
      if (!adminSent) {
        await sgMail.send(adminMsg);
        adminSent = true;
      }

      await bookingRef.set(
        {
          emailStatus: {
            customer: customerSent ? "sent" : "error",
            admin: adminSent ? "sent" : "error",
          },
          emailSentAt: admin.firestore.FieldValue.serverTimestamp(),
        },
        { merge: true }
      );
    } catch (err: any) {
      // Mark error (don’t hide failures)
      console.error("Error sending email:", err);
      await bookingRef.set(
        {
          emailStatus: {
            customer: customerSent ? "sent" : "error",
            admin: adminSent ? "sent" : "error",
          },
          emailError: String(err?.message || err),
          emailErrorAt: admin.firestore.FieldValue.serverTimestamp(),
        },
        { merge: true }
      );
      throw err; // let Firebase retry
    }
  }
);

interface CancelBookingRequest {
  bookingId: string;
}

/**
 * Callable function to cancel a booking.
 * Verifies the caller owns the booking, then sets status to "cancelled".
 */
export const cancelBooking = onCall(
  { region: "us-east1" },
  async (request) => {
    // Require authentication
    if (!request.auth) {
      throw new HttpsError("unauthenticated", "You must be signed in to cancel a booking.");
    }

    const { bookingId } = request.data as CancelBookingRequest;

    if (!bookingId || typeof bookingId !== "string") {
      throw new HttpsError("invalid-argument", "A valid bookingId is required.");
    }

    const bookingRef = admin.firestore().collection("bookings").doc(bookingId);
    const bookingSnap = await bookingRef.get();

    if (!bookingSnap.exists) {
      throw new HttpsError("not-found", "Booking not found.");
    }

    const bookingData = bookingSnap.data() as BookingData;

    // Verify ownership
    if (bookingData.userId !== request.auth.uid) {
      throw new HttpsError("permission-denied", "You can only cancel your own bookings.");
    }

    // Check if already cancelled
    if (bookingData.status === "cancelled") {
      return { success: true, message: "Booking was already cancelled." };
    }

    // Set status to cancelled
    await bookingRef.update({
      status: "cancelled",
      cancelledAt: admin.firestore.FieldValue.serverTimestamp(),
    });

    return { success: true, message: "Booking cancelled successfully." };
  }
);
