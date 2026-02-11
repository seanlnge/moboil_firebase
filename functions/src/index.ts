import { onDocumentCreated, FirestoreEvent } from "firebase-functions/v2/firestore";
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
  emailStatus?: {
    customer?: string;
    admin?: string;
  };
  [key: string]: any;
}

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
      bookingId,

      // buttons
      manageUrl: booking.manageUrl || "", // optional
      cancelUrl: booking.cancelUrl || "", // required in your template

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
