import { downloadCuttingWorkbook } from "./cuttingExcelWriter.js";
import { downloadBalingWorkbook } from "./balingExcelWriter.js";
import { downloadDeliveriesWorkbook } from "./deliveriesExcelWriter.js";

export async function downloadWorkbook({ chatType, filename, cutting, baling }) {
  if (chatType === "baling") {
    return downloadBalingWorkbook(baling, filename);
  }
  if (chatType === "deliveries") {
    return downloadDeliveriesWorkbook(cutting.records || [], filename, {
      validationLog: cutting.validationLog || [],
    });
  }
  return downloadCuttingWorkbook(cutting.records || [], filename, {
    summaryRecords: cutting.summaryRecords || [],
    validationLog: cutting.validationLog || [],
  });
}
