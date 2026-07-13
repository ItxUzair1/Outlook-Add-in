const nodemailer = require("nodemailer");

let transporter = null;

function getTransporter() {
  if (transporter) return transporter;

  const host = process.env.SMTP_HOST;
  const port = process.env.SMTP_PORT || 587;
  const user = process.env.SMTP_EMAIL;
  const pass = process.env.SMTP_PASSWORD;

  if (!host || !user || !pass) {
    console.warn("SMTP credentials not fully provided. Email notifications will be disabled.");
    return null;
  }

  transporter = nodemailer.createTransport({
    host: host,
    port: port,
    secure: port == 465, // true for 465, false for other ports
    auth: {
      user: user,
      pass: pass,
    },
  });

  return transporter;
}

async function sendApprovalNotification(userEmail, projectName) {
  const mailer = getTransporter();
  if (!mailer) {
    console.log(`[EmailService] Skipping notification to ${userEmail} for project '${projectName}' because SMTP is not configured.`);
    return;
  }

  try {
    const info = await mailer.sendMail({
      from: `"Koyomail Admin" <${process.env.SMTP_EMAIL}>`,
      to: userEmail,
      subject: "Your Project Indexing Request is Complete",
      text: `Good news! The project you requested indexing for (${projectName}) has been indexed and is now available in search.`,
      html: `
        <div style="font-family: sans-serif; padding: 20px;">
          <h2>Project Indexing Complete</h2>
          <p>Good news! The project you requested indexing for (<strong>${projectName}</strong>) has been indexed.</p>
          <p>You can now search for it within Koyomail.</p>
        </div>
      `,
    });
    console.log(`[EmailService] Notification sent: ${info.messageId}`);
  } catch (err) {
    console.error("[EmailService] Error sending email notification:", err);
  }
}

async function sendRejectionNotification(userEmail, projectName, rejectionMessage) {
  const mailer = getTransporter();
  if (!mailer) {
    console.log(`[EmailService] Skipping notification to ${userEmail} for project '${projectName}' because SMTP is not configured.`);
    return;
  }

  try {
    const info = await mailer.sendMail({
      from: `"Koyomail Admin" <${process.env.SMTP_EMAIL}>`,
      to: userEmail,
      subject: "Update on Your Project Indexing Request",
      text: `Your request to index the project (${projectName}) could not be completed.\n\nReason from Admin: ${rejectionMessage}`,
      html: `
        <div style="font-family: sans-serif; padding: 20px;">
          <h2>Project Indexing Request Update</h2>
          <p>Your request to index the project (<strong>${projectName}</strong>) could not be completed at this time.</p>
          <div style="background-color: #fff1f2; border-left: 4px solid #ef4444; padding: 12px 16px; margin: 16px 0; border-radius: 4px;">
            <p style="margin: 0; color: #991b1b;"><strong>Reason from Admin:</strong><br/><br/>${rejectionMessage}</p>
          </div>
          <p>Please contact your administrator if you have further questions.</p>
        </div>
      `,
    });
    console.log(`[EmailService] Rejection notification sent: ${info.messageId}`);
  } catch (err) {
    console.error("[EmailService] Error sending rejection notification:", err);
  }
}

module.exports = {
  sendApprovalNotification,
  sendRejectionNotification
};
