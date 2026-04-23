import { dateToStr } from "../helpers.js";

const mutedText = "#64748B";
const borderColor = "#E2E8F0";

const TH = ({ children }) => (
  <th style={{ padding: "10px 8px", textAlign: "left", borderBottom: `1px solid ${borderColor}`, fontWeight: 600, whiteSpace: "nowrap", color: mutedText, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5 }}>
    {children}
  </th>
);
const TD = ({ children, bold }) => (
  <td style={{ padding: "8px", fontWeight: bold ? 600 : undefined, whiteSpace: "nowrap" }}>{children ?? "—"}</td>
);

export default function DeliveriesResultsTable({ records }) {
  return (
    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
      <thead>
        <tr style={{ background: "#F9FAFB" }}>
          {["Date", "Truck", "Depot", "Transporter", "Passenger", "4x4", "Motorcycle", "LC", "HC", "HC SW", "HC T", "Agri"].map((h) => (
            <TH key={h}>{h}</TH>
          ))}
        </tr>
      </thead>
      <tbody>
        {records.slice(0, 100).map((r, i) => (
          <tr key={i} style={{ borderBottom: "1px solid #F3F4F6" }}>
            <TD>{r.date ? dateToStr(r.date) : null}</TD>
            <TD>{r.delTruckNo}</TD>
            <TD>{r.delDepot}</TD>
            <TD>{r.delTransporter}</TD>
            <TD bold>{r.delPassenger}</TD>
            <TD bold>{r.delFourByFour}</TD>
            <TD bold>{r.delMotorcycle}</TD>
            <TD bold>{r.delLightCommercial}</TD>
            <TD bold>{r.delHeavyCommercial}</TD>
            <TD bold>{r.delHeavyCommercialSW}</TD>
            <TD bold>{r.delHeavyCommercialT}</TD>
            <TD bold>{r.delAgricultural}</TD>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
