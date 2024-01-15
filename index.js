const { archiveSheets } = require("./archive");

// test local
const {
  copyFromRouteEntryToRouteSheet,
} = require("./copy-from-route-entry-to-route");

exports.main = async (req, res) => {
  await archiveSheets();
  res.status(200).send("Archive Sheets successfully");
};

// in case using it locale uncomment these code
async function bootstrap() {
  await archiveSheets();
  // await copyFromRouteEntryToRouteSheet();
}

bootstrap();

// test on cloud function

/*
archive time sheet last row: 2796
weekly archive sheet last row: 547
*/

// const functions = require('@google-cloud/functions-framework');
// const { archiveSheets } = require("./archive");
// const {
//   copyFromRouteEntryToRouteSheet,
// } = require("./copy-from-route-entry-to-route");

// functions.http('trinity',async (req, res) => {
//   if(req.body.action == 'archive sheets')
//     await archiveSheets();
//   else if( req.body.action == 'update route sheet')
//     await copyFromRouteEntryToRouteSheet();

//   res.send(`Hello ${process.env.PDF_NAME}!`);
// });
