const fs = require("fs");
const axios = require("axios");
require("dotenv").config();

async function convertSheetToPDF(auth) {
  const destPath = process.env.PDF_NAME + ".pdf"; // Path to save the PDF
  console.log("__dirname: ", destPath);
  await auth.authorize();
  // Fetch the access token
  const accessToken = await auth.getAccessToken();
  const exportUrl = `https://docs.google.com/spreadsheets/d/${process.env.SPREADSHEET_ID}/export?format=pdf&gid=${process.env.TIMESHEET_GID}`;

  const response = await axios({
    method: "get",
    url: exportUrl,
    headers: {
      Authorization: `Bearer ${accessToken.token}`,
      ResponseType: "stream",
    },
    responseType: "stream",
  });

  // Upload pdf locally
  return new Promise((resolve, reject) => {
    const dest = fs.createWriteStream(destPath);
    response.data
      .pipe(dest)
      .on("finish", () => {
        console.log("Upload pdf locally");
        resolve("done");
      })
      .on("error", error => {
        console.error("Error downloading file:", error);
        reject(error);
      });
  });
}

module.exports = { convertSheetToPDF };