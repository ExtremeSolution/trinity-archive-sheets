const { archiveSheets } = require("./archive");

exports.main = async (req, res) => {
  await archiveSheets();
  res.status(200).send("Archive Sheets successfully");
};

// in case using it locale uncomment these code
async function bootstrap() {
  await archiveSheets();
}

bootstrap();
