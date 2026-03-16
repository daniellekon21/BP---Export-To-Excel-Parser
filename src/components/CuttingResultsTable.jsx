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

export default function CuttingResultsTable({ records }) {
  return (
    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
      <thead>
        <tr style={{ background: "#F9FAFB" }}>
          {["Date", "Machine", "Time", "Radials LC", "Radials HC", "Radials Agri", "Radials Agri Treads", "Nylons LC"].map((h) => (
            <TH key={h}>{h}</TH>
          ))}
        </tr>
      </thead>
      <tbody>
        {records.slice(0, 100).map((r, i) => (
          <tr key={i} style={{ borderBottom: "1px solid #F3F4F6" }}>
            <TD>{dateToStr(r.date)}</TD>
            <TD>{r.cmNumber}</TD>
            <TD>{r.startTime && r.finishTime ? `${formatTime(r.startTime)}-${formatTime(r.finishTime)}` : null}</TD>
            <TD bold>{r.radialsLC}</TD>
            <TD bold>{r.radialsHC}</TD>
            <TD bold>{r.radialsAgri}</TD>
            <TD>{r.radialsAgriTreads}</TD>
            <TD>{r.nylonsLC}</TD>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
