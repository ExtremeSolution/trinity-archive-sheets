const { google } = require("googleapis");
const sheets = google.sheets("v4");
// let fs = require("fs");
require("dotenv").config();

async function copyFromRouteEntryToRouteSheet() {
  const spreadsheetId = process.env.SPREADSHEET_ID;

  const spreadsheets = sheets.spreadsheets;
  /*
    copy from sheets to other sheets
    in case of updating morning sheet, we will delete all data in the morning sheet & update the new values with email formula

*/

  const copySheets = [{ name: "Route Entry", range: "H" }];

  const authClient = new google.auth.JWT(
    process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    null,
    process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY.split(String.raw`\n`).join(
      "\n"
    ),
    [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ]
  );

  await authClient.authorize();

  try {
    for (copySheet of copySheets) {
      console.log(`Trying to get data of ${copySheet.name}`);

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
      let lastIndex = 0;
      for (
        let i = copyValues?.length ? copyValues?.length - 1 : 0;
        i >= 0;
        i--
      ) {
        if (copyValues[i]?.length == 8) {
          lastIndex = i + 1;
          console.log("the last index is " + lastIndex);
          break;
        }
      }
      const newData = await spreadsheets.values.get({
        auth: authClient,
        spreadsheetId,
        range: `${copySheet.name}!A${lastIndex ? lastIndex + 2 : 2}:${
          copySheet.range
        }`,
      });

      let routeData = await spreadsheets.values.get({
        auth: authClient,
        spreadsheetId,
        range: `Route!A2:H`,
      });

      let lastRoutIndex = routeData.data?.values?.length + 2;
      console.log("routeData length: ", lastRoutIndex);

      let newValues = newData.data?.values;
      let newValuesAddedForCurrentDay = [];
      console.log(`The length of the new data: ${newValues?.length || 0}`);
      if (newValues?.length) {
        for (let i = 0; i < newValues.length; i++) {
          if (
            new Date(newValues[i][3]).toDateString() ==
            new Date().toDateString()
          ) {
            console.log("new row === ", newValues[i]);
            newValues[i].splice(
              2,
              0,
              `=VLOOKUP(D${lastRoutIndex}, Driver!$B$2:$D$300,3, FALSE)`
            );
            newValuesAddedForCurrentDay.push(newValues[i]);
            lastRoutIndex = lastRoutIndex + 1;
          }
        }

        const body = {
          values: newValuesAddedForCurrentDay,
          range: "Route",
        };

        console.log("append data to Route");

        await spreadsheets.values
          .append({
            auth: authClient,
            spreadsheetId,
            range: body.range,
            valueInputOption: "USER_ENTERED",
            requestBody: body,
          })
          .catch(err => {
            console.log("append error========== ", err);
          });

        console.log("Mark the last row in Route entry");
        await spreadsheets.values
          .update({
            auth: authClient,
            spreadsheetId,
            range: `Route entry!H${lastIndex + 1 + newValues.length}`,
            valueInputOption: "USER_ENTERED",
            requestBody: {
              values: [["the last row!!!!!!!!"]],
              range: `Route entry!H${lastIndex + 1 + newValues.length}`,
            },
          })
          .then(data => {
            //   console.log(data);
          })
          .catch(err => {
            console.log("error############## ", err);
          });
      }
    }
  } catch (error) {
    console.error("The API returned an error: " + error);
  }
}

module.exports = { copyFromRouteEntryToRouteSheet };
