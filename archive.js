const { google } = require("googleapis");
const sheets = google.sheets("v4");
require("dotenv").config();

function archiveSheets() {
  const spreadsheetId = process.env.SPREADSHEET_ID;

  async function copyInfo() {
    const ss = sheets.spreadsheets;

    const copySheets = [
      { name: "Job", range: "AV" },
      { name: "Timesheet", range: "E" },
      { name: "Route", range: "H" },
      { name: "Morning Routes", range: "H" },
    ];
    const pasteSheets = {
      Timesheet: ["Archived timesheet", "Weekly timesheet"],
      Job: ["Archived Data"],
      Route: ["Archived Route"],
      "Morning Routes": ["Route"],
    };

    const authClient = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      null,
      process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY,

      [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/calendar",
      ]
    );

    await authClient.authorize();

    try {
      for (copySheet of copySheets) {
        console.log(
          `Trying to get data of a copy sheet with name ${copySheet.name}`
        );
        const sourceData = await ss.values.get({
          auth: authClient,
          spreadsheetId,
          range: `${copySheet.name}!A2:${copySheet.range}`,
        });
        console.log(
          `The length data of the copy sheet: ${
            sourceData?.data?.values?.length || 0
          }`
        );
        const copyValues = sourceData.data?.values;

        if (!copyValues?.length) {
          console.log(`The ${copySheet.name} has no values`);
          if (copySheet.name == "Route") {
            await copyFormula(ss, authClient, spreadsheetId);
          }
          console.log(
            "---------------------------------------------------------------"
          );

          continue;
        }

        for (let pasteSheet of pasteSheets[copySheet.name]) {
          console.log(
            `Trying to get data of paste sheet with name ${pasteSheet}`
          );
          const pasteData = await ss.values.get({
            auth: authClient,
            spreadsheetId,
            range: `${pasteSheet}!A2:${copySheet.range}`,
          });
          const lastRow = pasteData.data?.values?.length
            ? pasteData.data.values.length
            : 1 + 1;
          console.log(`${pasteSheet} contains ${lastRow} rows`);

          if (lastRow > 2)
            copyValues.unshift(
              pasteData.data.values[pasteData.data.values.length - 1]
            );
          const body = {
            values: copyValues,
            range: `${pasteSheet}!A${lastRow}:${copySheet.range}`,
          };

          console.log(
            `Trying to update ${pasteSheet} with updated rows: ${sourceData?.data?.values?.length}`
          );
          await ss.values
            .update({
              auth: authClient,
              spreadsheetId,
              range: body.range,
              valueInputOption: "USER_ENTERED",
              requestBody: body,
            })
            .then(data => {
              console.log("updatedRows: ", data.data.updatedRows);
              console.log("updatedRange: ", data.data.updatedRange);
            })
            .catch(err => {
              console.log("Error:  ", err);
            });

          if (copySheet.name == "Route") {
            await copyFormula(ss, authClient, spreadsheetId);
          }
        }

        console.log(`Trying to clear all rows in the ${copySheet.name} sheet.`);
        // Clear the source sheet
        await ss.values.clear({
          auth: authClient,
          spreadsheetId,
          range: `${copySheet.name}!A2:${copySheet.range}`,
        });
        console.log(
          "---------------------------------------------------------------"
        );
      }
    } catch (error) {
      console.error("The API returned an error: " + error);
    }
  }

  copyInfo();
}

async function copyFormula(ss, authClient, spreadsheetId) {
  console.log("Trying to copy formulas to job");

  // copy formula from formulas to job
  const formulaData = await ss.values.get({
    auth: authClient,
    spreadsheetId,
    valueRenderOption: "FORMULA",
    range: `Formulas!A2:AV`,
  });
  // Trying to update job sheet with formula value
  console.log(`Copy the formula to the Job sheet`);
  console.log(`Formula length: ${formulaData.data.values.length}`);
  await ss.values.update({
    auth: authClient,
    spreadsheetId,
    range: `Job!A2:AV`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: formulaData?.data?.values,
      range: `Job!A2:AV`,
    },
  });
}

module.exports = { archiveSheets };
