# trinity-archive-sheets

This project is responsible for archive sheets

## Installation

```bash
$ npm i
```

### Project Dependencies

- NodeJS v16 or higher.
- Yarn v1.22.17 or higher.
- googleapis ^95.0.0
- dotenv ^16.0.2

### Environment Variables

Copy `.env-sample` to a new file `.env` and add the values that match with each key and description. All variables must be loaded in `.env` to be able to run the project.

## Running the app

To run the project locally, please uncomment the following lines

```bash
async function bootstrap() {
  await archiveSheets();
}
bootstrap();
```
