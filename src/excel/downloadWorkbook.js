import { downloadCuttingWorkbook } from "./cuttingExcelWriter.js";
import { downloadBalingWorkbook } from "./balingExcelWriter.js";

export async function downloadWorkbook({ chatType, filename, cutting, baling }) {
  if (chatType === "baling") {
    return downloadBalingWorkbook(baling, filename);
  }
  return downloadCuttingWorkbook(cutting.records || [], filename, {
    summaryRecords: cutting.summaryRecords || [],
    validationLog: cutting.validationLog || [],
  });
}
