# Blue Pyramid - WhatsApp to Excel Converter

Converts WhatsApp chat exports from the BPR production floor into structured Excel workbooks.

## Quick Start

```bash
npm install
npm run dev
```

Then open http://localhost:5173 in your browser.

## Parse Modes

- `Baling Production`: dedicated baling parser + baling workbook writer.
- `Cutting Data`: existing cutting parser and cutting workbook writer.

The UI lets you choose baling vs cutting before parsing.

## CLI Parse Command

```bash
npm run parse -- --type=cutting --format=old --input=./sample-data/WhatsApp_Chat_with_Blue_pyramid_recycling-_Cut_Tyres.txt --output=./tmp/cutting.json
npm run parse -- --type=baling --input=./sample-data/WhatsApp_Chat_with_Blue_pyramid_recycling-_Baled_Finished_Goods.txt --output=./tmp/baling.json
```

Arguments:
- `--type=cutting|baling`
- `--format=old|new` (cutting only)
- `--input=<chat.txt>`
- `--output=<json-file>` (optional)

## Project Structure

- `src/parsing/cuttingParser.js`
- `src/parsing/balingParser.js`
- `src/parsing/commonParsingUtils.js`
- `src/excel/cuttingExcelWriter.js`
- `src/excel/balingExcelWriter.js`
- `src/excel/createBalingWorkbook.js`
- `src/excel/excelCommon.js`
- `src/config/cuttingSchemas.js`
- `src/config/balingSchemas.js`

## Baling Workbook (Dedicated)

Default filename in baling mode: `BPR_Baling_Data.xlsx`

Sheets:
- `Bales_Production`
- `Failed_Bales`
- `Scrap_Sidewalls`
- `CR_CA_Tests`
- `Daily_Summaries`
- `Validation_Log`

## Sample Data

- `sample-data/WhatsApp_Chat_with_Blue_pyramid_recycling-_Baled_Finished_Goods.txt`
- `sample-data/WhatsApp_Chat_with_Blue_pyramid_recycling-_Cut_Tyres.txt`
