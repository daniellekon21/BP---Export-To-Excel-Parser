# Blue Pyramid - WhatsApp to Excel Converter

Converts WhatsApp chat exports from the BPR production floor into structured CSV/Excel spreadsheets.

## Quick Start

```bash
npm install
npm run dev
```

Then open http://localhost:5173 in your browser.

## How to Use

1. Pick the report type: **Baling Production** or **Cutting Data**
2. Drag-and-drop (or paste) the exported WhatsApp `.txt` file
3. Click **Parse Chat** to extract the data
4. Review the preview table
5. Click **Download CSV** — opens directly in Excel

## Sample Data

The `sample-data/` folder contains real WhatsApp exports you can test with:
- `WhatsApp_Chat_with_Blue_pyramid_recycling-_Baled_Finished_Goods.txt` → use with **Baling Production**
- `WhatsApp_Chat_with_Blue_pyramid_recycling-_Cut_Tyres.txt` → use with **Cutting Data**

## Debugging with Claude Code

Open this folder in VS Code with the Claude Code extension, then ask Claude to help debug or extend the parsers. The parsing logic is all in `src/App.jsx`:
- `parseBalingMessages()` — handles bale reports
- `parseCuttingMessages()` — handles hourly cutting data
- `generateBalingCSV()` / `generateCuttingCSV()` — output formatters
