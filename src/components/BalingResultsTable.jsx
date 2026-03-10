import { formatTime, dateToStr } from "../helpers.js";

const mutedText = "#64748B";
const borderColor = "#E2E8F0";

const TH = ({ children }) => (
  <th style={{ padding: "10px 8px", textAlign: "left", borderBottom: `1px solid ${borderColor}`, fontWeight: 600, whiteSpace: "nowrap", color: mutedText, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5 }}>
    {children}
  </th>
);
const TD = ({ children, bold }) => (
  <td style={{ padding: "8px", fontWeight: bold ? 600 : undefined }}>{children ?? "—"}</td>
);

export default function BalingResultsTable({ records }) {
  return (
    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
      <thead>
        <tr style={{ background: "#F9FAFB" }}>
          {["Date", "Type", "Machine", "Bale", "Status", "Start", "Finish", "Passenger", "4x4", "LC", "MC", "SR", "Agri", "Weight", "Notes"].map((h) => (
            <TH key={h}>{h}</TH>
          ))}
        </tr>
      </thead>
      <tbody>
        {records.slice(0, 100).map((r, i) => (
          <tr key={i} style={{ borderBottom: "1px solid #F3F4F6" }}>
            <TD>{r.date ? dateToStr(r.date) : null}</TD>
            <TD>{r.recordType}</TD>
            <TD>{r.machine}</TD>
            <TD>{r.baleNumber || r.baleTestCode}</TD>
            <TD>{r.status}</TD>
            <TD>{r.startTime ? formatTime(r.startTime) : null}</TD>
            <TD>{r.finishTime ? formatTime(r.finishTime) : null}</TD>
            <TD>{r.passengerQty}</TD>
            <TD>{r.fourx4Qty}</TD>
            <TD>{r.lcQty}</TD>
            <TD>{r.motorcycleQty}</TD>
            <TD>{r.srQty}</TD>
            <TD>{r.agriQty}</TD>
            <TD bold>{r.weightKg}</TD>
            <TD>{r.notesFlags}</TD>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
