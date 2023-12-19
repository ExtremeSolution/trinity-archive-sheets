const { google } = require("googleapis");
const sheets = google.sheets("v4");
require("dotenv").config();

function archiveSheets() {
  const spreadsheetId = process.env.SPREADSHEET_ID;

  async function copyInfo() {
    const spreadsheets = sheets.spreadsheets;

    const copySheets = [
      { name: "Job", range: "AV" },
      { name: "Timesheet", range: "E" },
      { name: "Route", range: "H" },
      { name: "Morning route entry", range: "H" },
      { name: "Morning Routes", range: "H" },
    ];
    const pasteSheets = {
      Job: ["Archived Data"],
      Timesheet: ["Archived timesheet", "Weekly timesheet"],
      Route: ["Archived Route"],
      "Morning route entry": ["Morning Routes"],
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
        const sourceData = await spreadsheets.values.get({
          auth: authClient,
          spreadsheetId,
          range: `${copySheet.name}!A2:${copySheet.range}`,
        });
        console.log(
          `The length data of the copy sheet: ${
            sourceData?.data?.values?.length || 0
          }`
        );
        let copyValues = sourceData.data?.values;
        if (copySheet.name === "Morning route entry") {
          if (copyValues?.length) {
            for (let i = 0; i < copyValues.length; i++) {
              copyValues[i].splice(
                2,
                0,
                `=VLOOKUP(D${i + 2}, Driver!$B$2:$D$300,3, FALSE)`
              );
            }
          }
        }
        if (!copyValues?.length) {
          console.log(`The ${copySheet.name} has no values`);
          if (copySheet.name == "Job") {
            await copyFormula(spreadsheets, authClient, spreadsheetId);
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
          const pasteData = await spreadsheets.values.get({
            auth: authClient,
            spreadsheetId,
            range: `${pasteSheet}!A2:${copySheet.range}`,
          });
          const lastRow = pasteData.data?.values?.length
            ? pasteData.data.values.length
            : 1 + 1;
          console.log(`${pasteSheet} contains ${lastRow} rows`);
          const body = {
            values: copyValues,
            range: pasteSheet,
          };

          console.log(
            `Trying to update ${pasteSheet} with updated rows: ${sourceData?.data?.values?.length}`
          );
          if (copySheet.name === "Morning route entry") {
            await spreadsheets.values.batchClear({
              auth: authClient,
              spreadsheetId,
              requestBody: {
                ranges: [`${pasteSheet}!A2:${copySheet.range}`],
              },
            });
          }
          await spreadsheets.values.append({
            auth: authClient,
            spreadsheetId,
            range: body.range,
            valueInputOption: "USER_ENTERED",
            requestBody: body,
          });

          if (copySheet.name == "Job") {
            await copyFormula(spreadsheets, authClient, spreadsheetId);
          }
        }

        console.log(`Trying to clear all rows in the ${copySheet.name} sheet.`);
        // Clear the source sheet
        await spreadsheets.values.clear({
          auth: authClient,
          spreadsheetId,
          range: `${copySheet.name}!A2:${copySheet.range}`,
        });

        console.log(
          "---------------------------------------------------------------"
        );
      }
      // update formula for morning route and Route sheet
      for (let sheet of ["Morning Routes", "Route"]) {
        let response = await spreadsheets.values.get({
          auth: authClient,
          spreadsheetId,
          range: `${sheet}!A2:H`,
        });
        let { formulas, lastIndex } = getFormulasList(response.data?.range);
        await spreadsheets.values
          .update({
            auth: authClient,
            spreadsheetId,
            valueInputOption: "USER_ENTERED",
            range: `${sheet}!C2:C${lastIndex}`,

            requestBody: {
              values: formulas,
              range: `${sheet}!C2:C${lastIndex}`,
            },
          })
          .then(data => {
            console.log();
            console.log(
              `Add formulas...................`,
              data.data.updatedRange
            );
          })
          .catch(err => console.log(`Cant add formulas`, err));
      }
    } catch (error) {
      console.error("The API returned an error: " + error);
    }
  }

  copyInfo();
}

function getFormulasList(range) {
  // get the last row index
  let lastIndex = extractNumberAfterColon(range);
  // generate array of formulas
  let formulas = [];
  for (let i = 2; i <= lastIndex; i++) {
    formulas.push([`=VLOOKUP(D${i}, Driver!$B$2:$D$300,3, FALSE)`]);
  }
  return { formulas, lastIndex };
}

function extractNumberAfterColon(rangeString) {
  // Regular expression to find one or more letters followed by one or more digits
  const regex = /:([A-Z]+)(\d+)/i;

  // Apply the regex to the range string
  const match = rangeString.match(regex);

  if (match && match[2]) {
    // Extract and return the number part
    return parseInt(match[2], 10);
  } else {
    // Return null or appropriate response if no match found
    return null;
  }
}

async function copyFormula(spreadsheets, authClient, spreadsheetId) {
  console.log("Trying to copy formulas to job");

  // copy formula from formulas to job
  const formulaData = await spreadsheets.values.get({
    auth: authClient,
    spreadsheetId,
    valueRenderOption: "FORMULA",
    range: `Formulas!A2:AV`,
  });
  // Trying to update job sheet with formula value
  console.log(`Copy the formula to the Job sheet`);
  console.log(`Formula length: ${formulaData.data.values.length}`);
  await spreadsheets.values.update({
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
