require("dotenv").config();
const spreadsheetId = process.env.SPREADSHEET_ID;

async function handleRouteEntry(spreadsheets, authClient, data) {
  let { currentData, futureData } = getArchivedDataOfRouteEntry(data);

  // clear all sheet
  console.log("clear all sheet");
  await spreadsheets.values.clear({
    auth: authClient,
    spreadsheetId,
    range: `Route entry!A2:H`,
  });
  // update currentData with formula
  if (currentData?.length) {
    console.log("update currentData with formula");

    for (let i = 0; i < currentData.length; i++) {
      currentData[i].splice(
        2,
        0,
        `=VLOOKUP(D${i + 2}, Driver!$B$2:$D$300,3, FALSE)`
      );
    }
    const body = {
      values: currentData,
      range: "Route!A2:H",
    };

    console.log("add currentData to Route sheet");

    // add currentData to Route sheet

    await spreadsheets.values.append({
      auth: authClient,
      spreadsheetId,
      range: body.range,
      valueInputOption: "USER_ENTERED",
      requestBody: body,
    });
  }
  console.log("add futureData to sheet");

  if (futureData?.length) {
    // add futureData to sheet
    const body = {
      values: futureData,
      range: "Route entry",
    };
    await spreadsheets.values.append({
      auth: authClient,
      spreadsheetId,
      range: body.range,
      valueInputOption: "USER_ENTERED",
      requestBody: body,
    });
  }

  //   await updateRouteFormula(spreadsheets, authClient);
}

function getArchivedDataOfRouteEntry(data) {
  let currentData = [];
  let previousData = [];
  let futureData = [];
  let rangeStartIndex = 0;
  let currentState;
  let previousState;
  const currentDateStr = new Date().toDateString();
  const currentTime = new Date().getTime();

  data.forEach((row, index) => {
    const itemDate = new Date(row[3]);
    const itemDateStr = itemDate.toDateString();
    const itemTime = itemDate.getTime();

    console.log(
      `Index: ${index}, Item Date: ${itemDateStr}, Current Date: ${currentDateStr}`
    );

    if (itemDateStr === currentDateStr) {
      currentState = "currentData";
    } else if (itemTime > currentTime) {
      currentState = "futureData";
    } else {
      currentState = "previousData";
    }

    if (currentState !== previousState && previousState !== null) {
      const rangeEndIndex = index === 0 ? 0 : index - 1;
      addRange(previousState, rangeStartIndex, rangeEndIndex);
      rangeStartIndex = index;
    }

    if (index === data.length - 1) {
      addRange(currentState, rangeStartIndex, index);
    }

    previousState = currentState;
  });

  function addRange(state, start, end) {
    console.log(`Adding range: ${state}, Start: ${start}, End: ${end}`);
    if (state === "previousData") {
      previousData.push([start, end]);
    } else if (state === "currentData") {
      currentData.push(...data.slice(start, end + 1));
    } else if (state === "futureData") {
      futureData.push(...data.slice(start, end + 1));
    }
  }

  return {
    currentData,
    futureData,
    previousData,
  };
}

module.exports = { handleRouteEntry };
