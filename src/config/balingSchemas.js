export const BALING_SHEET_NAMES = {
  production: "Bales_Production",
  failed: "Failed_Bales",
  scrap: "Scrap_Sidewalls",
  crca: "CR_CA_Tests",
  summaries: "Daily_Summaries",
  validation: "Validation_Log",
};

export const BALING_PRODUCTION_HEADERS = [
  "Chat Date Parsed",
  "Baling Machine Number",
  "Bale Number",
  "Record Type",
  "Production Type",
  "Baler Operator",
  "Baler Assistant Operator",
  "Start Time",
  "Finish Time",
  "Total Time to Bale",
  "Passenger",
  "4 x 4",
  "LC",
  "Total Number of Tyres",
  "Bale Weight KG",
  "Bale Weight TONS",
  "Notes / Flags",
  "Raw Message",
];

export const BALING_FAILED_HEADERS = [
  "Chat Date Parsed",
  "Baling Machine Number",
  "Bale Number",
  "Failure Type",
  "Failure Reason",
  "Baler Operator",
  "Baler Assistant Operator",
  "Start Time",
  "Finish Time",
  "Passenger Qty",
  "4x4 Qty",
  "LC Qty",
  "Motorcycle Qty",
  "SR Qty",
  "Agri Qty",
  "Tread Qty",
  "Side Wall Qty",
  "Total Qty",
  "Weight Kg",
  "Notes / Flags",
  "Raw Message",
];

export const BALING_SCRAP_HEADERS = [
  "Chat Date Parsed",
  "Baling Machine Number",
  "Production Label",
  "Bale Number",
  "Scrap Type",
  "Scrap Qty",
  "Weight Kg",
  "Baler Operator",
  "Baler Assistant Operator",
  "Start Time",
  "Finish Time",
  "Notes / Flags",
  "Raw Message",
];

export const BALING_CRCA_HEADERS = [
  "Chat Date Parsed",
  "Baling Machine Number",
  "Bale/Test Code",
  "Test Type",
  "Record Type",
  "Baler Operator",
  "Baler Assistant Operator",
  "Start Time",
  "Finish Time",
  "Total Time to Bale",
  "Tread Qty",
  "Side Wall Qty",
  "Passenger Qty",
  "4x4 Qty",
  "LC Qty",
  "Motorcycle Qty",
  "SR Qty",
  "Agri Qty",
  "Weight Kg",
  "Notes / Flags",
  "Raw Message",
];

export const BALING_SUMMARY_HEADERS = [
  "Chat Date Parsed",
  "Summary Type",
  "Baling Machine Number",
  "Bale Count",
  "Weight Kg",
  "Tons",
  "Passenger Qty",
  "4x4 Qty",
  "LC Qty",
  "Motorcycle Qty",
  "SR Qty",
  "Tread Qty",
  "Side Wall Qty",
  "Agri Qty",
  "Total Tyres",
  "Machine 1 Start Hour",
  "Machine 1 Finish Hour",
  "Machine 2 Start Hour",
  "Machine 2 Finish Hour",
  "Notes / Flags",
  "Raw Message",
];

// Bale code prefix definitions (for reference):
//   CA    = Cut Agricultural
//   PCR   = Passenger Car Radial (legacy code: PB → PCR)
//   CRC   = Cut Radial (will shorten to CR in future)
//   CRS   = Cut Radial Scrap
//   TB    = Tubes (will relabel to "Tubes")
//   CN    = Cut Nylon
//   PShrB = Passenger Shred Bulkbag
//   CR    = Cut Radial (old SR/Scrap Radial merged into CR from Nov 2024)
//
// Tyre component abbreviations used in bale messages:
//   T      = Tread
//   SW     = Side Wall
//   HC     = Heavy Commercial (HC(Cut) for CRC bales; HC(PS) for CR/CRS bales)
//   MC     = Motorcycle
//   LC     = Light Commercial
//   P / Pass. = Passenger
//   4x4    = 4x4
//   Tube   = Tubes (TB bales)
//
// IMPORTANT — alias order matters: the FIRST match wins.
//   sr must come before hc so HC(PS) maps to sr (pre-shredded/scrap radial),
//   while HC(Cut) and bare HC map to hc (cut heavy commercial).

export const BALING_CATEGORY_ALIASES = [
  { key: "passenger", pattern: /\bpassengers?\b|\bpass\b|\bpcr\b|\bPass\./i },
  { key: "fourx4", pattern: /\b4\s*x\s*4\b|\b4x4\b/i },
  { key: "lc", pattern: /\blight\s*commercials?\b|\blight\s*comm\b|\blc\b/i },
  { key: "motorcycle", pattern: /\bmotor\s*cycle\b|\bmotorcycle\b|\bmc\b/i },
  // HC(PS) = pre-shredded heavy commercial (same category as old Scrap Radial / SR)
  { key: "sr", pattern: /\bsr\b|\bside\s*wall\s*radial\b|\bradial\s*side\s*wall\b|\bHC\s*\(\s*ps\s*\)/i },
  // HC(Cut) and bare HC = cut heavy commercial (CRC bales)
  { key: "hc", pattern: /\bheavy\s*commercials?\b|\bHC\s*\(\s*cut\s*\)|\bHC\b/i },
  { key: "agri", pattern: /\bagri\b|\bagricultural\b/i },
  // \bT\b matches standalone "T" used for treads in CA/CN bale messages
  { key: "tread", pattern: /\btyre\s*treads?\b|\btire\s*treads?\b|\btreads?\b|\blct\b|\bhct\b|\bhc\s*\(\s*treads?\s*\)\b|\bT\b/i },
  { key: "sideWall", pattern: /\bside\s*walls?\b|\bsidewall\b|\bsw\b/i },
  { key: "tube", pattern: /\btubes?\b/i },
];
