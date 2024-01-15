require("dotenv").config();
const fs = require("fs");
const mailgun = require("mailgun-js");

async function sendEmail() {
  try {
    const mg = mailgun({
      apiKey: process.env.MAILGUN_API_KEY,
      domain: process.env.MAILGUN_DOMAIN,
    });

    const data = {
      from: "Anaphora <noreply@" + process.env.MAILGUN_DOMAIN + ">",
      to: process.env.REPORT_RECIPIENT_EMAIL,
      subject: process.env.REPORT_EMAIL_SUBJECT,
      text: `Hi ,
             Please find the Daily report attached.
             Thank you,
             Trinity`,
      attachment:  new mg.Attachment({
        data: fs.readFileSync(process.env.PDF_NAME + ".pdf"),
        filename: process.env.PDF_NAME + ".pdf",
      }),
      cc: process.env.REPORT_CC_EMAIL,
    };
    // console.log("data=== ", data);
    await mg.messages().send(data, function (error, body) {
      if (error) {
        console.error("Error sending email:", error);
      } else {
        console.log("Email sent:", body);
        fs.unlinkSync(process.env.PDF_NAME + ".pdf");
      }
    });
  } catch (err) {
    console.log(err);
    throw err;
  }
}
// sendEmail()

module.exports = { sendEmail };