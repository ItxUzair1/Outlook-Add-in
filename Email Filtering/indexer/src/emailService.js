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

module.exports = {
  sendApprovalNotification
};
