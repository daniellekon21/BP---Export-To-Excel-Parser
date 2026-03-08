import { workbookToBrowserDownload } from "./excelCommon.js";
import { createBalingWorkbook } from "./createBalingWorkbook.js";

export async function downloadBalingWorkbook(data, filename = "BPR_Baling_Data.xlsx") {
  const wb = await createBalingWorkbook(data);
  await workbookToBrowserDownload(wb, filename || "BPR_Baling_Data.xlsx");
}
