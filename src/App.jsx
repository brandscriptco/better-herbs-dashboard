import { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// âââ CONSTANTS âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
const PRODUCTS = ["All Products", "Lactify", "Brainfy Powder", "Brainfy Drops", "Mamafy", "Flow Joy Drops"];

const PRODUCT_COLORS = {
  "All Products": "#c8a2f8",
  "Lactify": "#60d394",
  "Brainfy Powder": "#5ca4f7",
  "Brainfy Drops": "#f7a75c",
  "Mamafy": "#f76b8a",
  "Flow Joy Drops": "#5cf7e0",
  "Other": "#888",
};

// âââ DEMO DATA ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
const DEMO_ADS = [
  { product: "Lactify", name: "Lactify_Dr. Nayana", spend: 29920.97, sales: 107868, purchases: 124, roas: 3.61, type: "Doctor" },
  { product: "Lactify", name: "Lactify - 2nd Doctor Compilation", spend: 12600.59, sales: 18336, purchases: 22, roas: 1.48, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Customer Review", spend: 11997.83, sales: 24681, purchases: 28, roas: 2.10, type: "UGC" },
  { product: "Lactify", name: "Lactify - Dr. Rohit Bhardwaj", spend: 6449.40, sales: 4989, purchases: 7, roas: 0.80, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Dr. Mayur", spend: 5501.84, sales: 9517, purchases: 12, roas: 1.73, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Dr. Priyanka D", spend: 4600.99, sales: 5523, purchases: 7, roas: 3.10, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Dr. Garima", spend: 3586.59, sales: 3305, purchases: 5, roas: 0.92, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Static 27", spend: 3572.62, sales: 6071, purchases: 8, roas: 1.70, type: "Static" },
  { product: "Lactify", name: "Lactify - Dr. Naaz", spend: 3550.11, sales: 8165, purchases: 10, roas: 2.30, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Dr. Srimukhi (Tamil)", spend: 2689.65, sales: 3605, purchases: 5, roas: 4.71, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Dr. Sandeep G", spend: 1645.26, sales: 2537, purchases: 3, roas: 1.54, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Dr. Priyanka (2nd)", spend: 1587.13, sales: 7689, purchases: 11, roas: 4.93, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Dr. Anuradha", spend: 960.49, sales: 1298, purchases: 2, roas: 1.35, type: "Doctor" },
  { product: "Lactify", name: "Lactify - Mother Compilation", spend: 929.65, sales: 2446, purchases: 3, roas: 2.63, type: "UGC" },
  { product: "Brainfy Drops", name: "Drops - Maninder Kaur (Blessings) Video", spend: 76725.88, sales: 247220.45, purchases: 320, roas: 3.22, type: "UGC" },
  { product: "Brainfy Drops", name: "Drops - Dr. Vinod Video", spend: 71320.64, sales: 204505.20, purchases: 276, roas: 2.87, type: "Doctor" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Dr. Sajid", spend: 42538.74, sales: 125683, purchases: 169, roas: 2.95, type: "Doctor" },
  { product: "Brainfy Drops", name: "Drops - Static 7", spend: 31810.78, sales: 94990.60, purchases: 119, roas: 2.99, type: "Static" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Dr. Sajid (2nd)", spend: 31295.22, sales: 76436, purchases: 106, roas: 2.44, type: "Doctor" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Rajmani", spend: 16589.64, sales: 38827, purchases: 56, roas: 2.34, type: "UGC" },
  { product: "Brainfy Drops", name: "Drops - Dr. Maninder Kaur (Copycat Brand)", spend: 11199.54, sales: 27172.20, purchases: 32, roas: 2.43, type: "Doctor" },
  { product: "Brainfy Drops", name: "Mothers Day - Drops", spend: 10533.15, sales: 40082, purchases: 43, roas: 3.80, type: "Static" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Static 11", spend: 9856.68, sales: 20245, purchases: 24, roas: 2.05, type: "Static" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Static 12", spend: 9766.42, sales: 32583, purchases: 43, roas: 3.34, type: "Static" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Static 3", spend: 8924.17, sales: 22710, purchases: 26, roas: 2.54, type: "Static" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Rajmani Patel Edited", spend: 7933.11, sales: 17268, purchases: 19, roas: 2.18, type: "UGC" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Dr. Neha", spend: 5104.66, sales: 12911, purchases: 15, roas: 2.53, type: "Doctor" },
  { product: "Brainfy Drops", name: "Brainfy Drops - Static 8", spend: 5534.29, sales: 12605, purchases: 15, roas: 2.28, type: "Static" },
  { product: "Flow Joy Drops", name: "FlowJoy Drops - Static 2", spend: 27785.57, sales: 54363.40, purchases: 60, roas: 1.96, type: "Static" },
  { product: "Flow Joy Drops", name: "Flow Drop - Dr. Ankit", spend: 22315.49, sales: 43853, purchases: 52, roas: 1.97, type: "Doctor" },
  { product: "Flow Joy Drops", name: "Flow Drop - Dispatch Video", spend: 19097.73, sales: 34730, purchases: 46, roas: 1.82, type: "UGC" },
  { product: "Flow Joy Drops", name: "Flow Drops - Dr. Sushma Mogri", spend: 9063.66, sales: 15289, purchases: 21, roas: 1.69, type: "Doctor" },
  { product: "Flow Joy Drops", name: "FlowJoy Drops - Dr. Garima", spend: 5889.86, sales: 23955, purchases: 36, roas: 4.07, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Divya Bajpai", spend: 62595.97, sales: 211753.82, purchases: 244, roas: 3.38, type: "UGC" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Static 13 (Ingredients)", spend: 48015.05, sales: 139503.40, purchases: 167, roas: 2.91, type: "Static" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Ankit Jha Video", spend: 36147.68, sales: 134105.96, purchases: 167, roas: 3.71, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Vinod 2", spend: 34148.19, sales: 121652, purchases: 160, roas: 3.56, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Vinod Post ID", spend: 21521.84, sales: 52808, purchases: 67, roas: 2.45, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Pushpendra 2nd", spend: 16791.22, sales: 37857, purchases: 50, roas: 2.25, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Prachi Mahajan", spend: 16488.92, sales: 46782, purchases: 49, roas: 2.84, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Doctor Compilation 2", spend: 12692.45, sales: 43337, purchases: 55, roas: 3.41, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Static 11", spend: 7369.96, sales: 13779, purchases: 20, roas: 1.87, type: "Static" },
  { product: "Brainfy Powder", name: "Brainfy Powder - USP Static", spend: 7243.52, sales: 28133.20, purchases: 34, roas: 3.88, type: "Static" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Ankit Jha (2nd)", spend: 6841.49, sales: 17435, purchases: 24, roas: 2.55, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Static 12", spend: 6838.35, sales: 28006, purchases: 34, roas: 4.10, type: "Static" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Maninder Kaur (Blessings)", spend: 4125.93, sales: 8876, purchases: 13, roas: 2.15, type: "UGC" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Riya", spend: 3519.23, sales: 7343, purchases: 7, roas: 2.09, type: "Doctor" },
  { product: "Brainfy Powder", name: "Brainfy Powder - Dr. Suryakamal", spend: 3290.11, sales: 4374, purchases: 6, roas: 1.33, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Dr. Garima", spend: 28464.37, sales: 79689, purchases: 101, roas: 2.80, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Static 13", spend: 23968.92, sales: 59926, purchases: 75, roas: 2.50, type: "Static" },
  { product: "Mamafy", name: "Mamafy - Dr. Shaifali", spend: 23065.44, sales: 68987, purchases: 88, roas: 2.99, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Static 4 (Do Not Buy)", spend: 20628.66, sales: 49872, purchases: 61, roas: 2.42, type: "Static" },
  { product: "Mamafy", name: "Mamafy - Static 2", spend: 12758.97, sales: 33348, purchases: 42, roas: 2.61, type: "Static" },
  { product: "Mamafy", name: "Mamafy - Dr. Smriti Edited", spend: 11809.94, sales: 21520, purchases: 28, roas: 1.82, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Static 6 USP", spend: 10232.50, sales: 29735, purchases: 33, roas: 2.91, type: "Static" },
  { product: "Mamafy", name: "Mamafy - Dr. Samra Edit 2 (Direct)", spend: 9360.38, sales: 21614, purchases: 25, roas: 2.31, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Doctor Compilation Video", spend: 9199.31, sales: 27753, purchases: 37, roas: 3.02, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Doctor Compilation March", spend: 6473.92, sales: 15929, purchases: 20, roas: 2.46, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Dispatch Video March", spend: 5663.82, sales: 12123, purchases: 17, roas: 2.14, type: "UGC" },
  { product: "Mamafy", name: "Mamafy - Dr. Samra Edit 4 (Nutrition)", spend: 5478.25, sales: 12313, purchases: 15, roas: 2.25, type: "Doctor" },
  { product: "Mamafy", name: "Mothers Day - Mamafy", spend: 4392.87, sales: 14140, purchases: 15, roas: 3.22, type: "Static" },
  { product: "Mamafy", name: "Mamafy - Static 14", spend: 4112.96, sales: 12953, purchases: 15, roas: 3.15, type: "Static" },
  { product: "Mamafy", name: "Mamafy - Founder's Dispatch Video May", spend: 4103.58, sales: 16258, purchases: 21, roas: 3.96, type: "UGC" },
  { product: "Mamafy", name: "Mamafy - Dr. Priya Soni", spend: 3697.88, sales: 3934, purchases: 6, roas: 1.06, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Dr. Soniya Gupta", spend: 3383.37, sales: 14412, purchases: 17, roas: 4.26, type: "Doctor" },
  { product: "Mamafy", name: "Mamafy - Dispatch Video April", spend: 2729.07, sales: 3585, purchases: 5, roas: 1.31, type: "UGC" },
];

const DEMO_GOOGLE = { spend: 47604.54, sales: 505685.75, purchases: 471, roas: 8.82 };

// âââ PARSING HELPERS ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function extractProduct(adName, campaignName) {
  for (const src of [(adName || "").toLowerCase(), (campaignName || "").toLowerCase()]) {
    if (src.includes("lactify")) return "Lactify";
    if (src.includes("brainfy powder") || (src.includes("brainfy") && src.includes("powder"))) return "Brainfy Powder";
    if (src.includes("brainfy drops") || (src.includes("brainfy") && src.includes("drop"))) return "Brainfy Drops";
    if (src.includes("mamafy")) return "Mamafy";
    if (src.includes("flow joy") || src.includes("flowjoy") || src.includes("flow drop") || src.includes("flow drops")) return "Flow Joy Drops";
    if ((src.includes("drops") || src.includes("drop")) && !src.includes("mamafy")) return "Brainfy Drops";
  }
  return "Other";
}

function extractAdType(adName) {
  const n = (adName || "").toLowerCase();
  if (n.includes("dr.") || n.includes("dr ") || n.includes("doctor") || n.includes("compilation")) return "Doctor";
  if (n.includes("static") || n.includes("usp") || n.includes("ingredient")) return "Static";
  return "UGC";
}

function findColIdx(headers, candidates) {
  for (const c of candidates) {
    const idx = headers.findIndex((h) => h.toLowerCase().includes(c.toLowerCase()));
    if (idx !== -1) return idx;
  }
  return -1;
}

function cleanNum(val) {
  if (val === null || val === undefined || val === "") return 0;
  const s = String(val).replace(/[â¹,\s]/g, "").replace(/[^0-9.-]/g, "");
  return parseFloat(s) || 0;
}

function parseMetaRows(headers, rows) {
  const nameIdx = findColIdx(headers, ["ad name", "name"]);
  const spendIdx = findColIdx(headers, ["amount spent", "spend", "cost"]);
  const salesIdx = findColIdx(headers, ["purchase conversion value", "conversion value", "revenue", "value"]);
  const purchasesIdx = findColIdx(headers, ["website purchases", "purchases", "conversions"]);
  const campaignIdx = findColIdx(headers, ["campaign name", "campaign"]);

  const ads = [];
  for (const row of rows) {
    const name = String(row[nameIdx] || "").trim();
    const spend = cleanNum(row[spendIdx]);
    const sales = cleanNum(row[salesIdx]);
    const purchases = Math.round(cleanNum(row[purchasesIdx]));
    const campaign = String(row[campaignIdx] || "").trim();
    if (!name || spend <= 0) continue;
    const product = extractProduct(name, campaign);
    const type = extractAdType(name);
    const roas = spend > 0 && sales > 0 ? sales / spend : 0;
    ads.push({ name, spend, sales, purchases, roas, product, type });
  }
  return ads;
}

function parseGoogleRows(headers, rows) {
  const costIdx = findColIdx(headers, ["cost", "spend", "amount"]);
  const convsIdx = findColIdx(headers, ["conversions", "conv.", "purchases"]);
  const valueIdx = findColIdx(headers, ["conversion value", "conv. value", "all conv. value", "revenue", "value"]);
  let spend = 0, sales = 0, purchases = 0;
  for (const row of rows) {
    const rowCost = cleanNum(row[costIdx]);
    if (rowCost > 0) {
      spend += rowCost;
      sales += cleanNum(row[valueIdx]);
      purchases += Math.round(cleanNum(row[convsIdx]));
    }
  }
  return { spend, sales, purchases, roas: spend > 0 && sales > 0 ? sales / spend : 0 };
}

function sheetToHeadersRows(sheet) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!raw || raw.length < 2) return { headers: [], rows: [] };
  let headerIdx = 0;
  for (let i = 0; i < Math.min(raw.length, 10); i++) {
    if (raw[i].filter((c) => c !== "").length >= 3) { headerIdx = i; break; }
  }
  const headers = raw[headerIdx].map((h) => String(h || "").trim());
  const rows = raw.slice(headerIdx + 1).filter((r) => r.some((c) => c !== ""));
  return { headers, rows };
}

// âââ AGGREGATION ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function computeProductTotals(ads) {
  const totals = {};
  for (const ad of ads) {
    if (!totals[ad.product]) totals[ad.product] = { spend: 0, sales: 0, purchases: 0 };
    totals[ad.product].spend += ad.spend;
    totals[ad.product].sales += ad.sales;
    totals[ad.product].purchases += ad.purchases;
  }
  const result = {};
  for (const [p, d] of Object.entries(totals)) {
    result[p] = { ...d, roas: d.spend > 0 && d.sales > 0 ? d.sales / d.spend : 0 };
  }
  const allSpend = Object.values(result).reduce((s, d) => s + d.spend, 0);
  const allSales = Object.values(result).reduce((s, d) => s + d.sales, 0);
  const allPurchases = Object.values(result).reduce((s, d) => s + d.purchases, 0);
  result["All Products"] = {
    spend: allSpend, sales: allSales, purchases: allPurchases,
    roas: allSpend > 0 && allSales > 0 ? allSales / allSpend : 0,
  };
  return result;
}

// âââ FORMATTERS âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
const fmt = (n) => {
  if (n >= 100000) return `â¹${(n / 100000).toFixed(2)}L`;
  if (n >= 1000) return `â¹${(n / 1000).toFixed(1)}K`;
  return `â¹${n.toFixed(0)}`;
};
const fmtFull = (n) => `â¹${n.toLocaleString("en-IN", { maximumFractionDigits: 0 })}`;

// âââ UI: ROAS BADGE ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function RoasBadge({ value }) {
  const color = value >= 3 ? "#34d399" : value >= 2 ? "#fbbf24" : "#f87171";
  const bg = value >= 3 ? "rgba(52,211,153,0.12)" : value >= 2 ? "rgba(251,191,36,0.12)" : "rgba(248,113,113,0.12)";
  return (
    <span style={{ background: bg, color, padding: "3px 10px", borderRadius: 6, fontWeight: 700, fontSize: 13, fontFamily: "'JetBrains Mono', monospace" }}>
      {value.toFixed(2)}x
    </span>
  );
}

// âââ UI: MINI BAR âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function MiniBar({ value, max }) {
  const pct = max > 0 ? Math.min((value / max) * 100, 100) : 0;
  return (
    <div style={{ width: "100%", height: 6, background: "rgba(255,255,255,0.06)", borderRadius: 3 }}>
      <div style={{ width: `${pct}%`, height: "100%", background: "linear-gradient(90deg, #5ca4f7, #c8a2f8)", borderRadius: 3, transition: "width 0.5s ease" }} />
    </div>
  );
}

// âââ UI: UPLOAD ZONE ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function UploadZone({ onData, onDemo }) {
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const inputRef = useRef(null);

  const processFile = useCallback(async (file) => {
    setLoading(true);
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const workbook = XLSX.read(buf, { type: "array" });
      if (!workbook.SheetNames.length) throw new Error("No sheets found in the file.");

      const { headers: metaHeaders, rows: metaRows } = sheetToHeadersRows(workbook.Sheets[workbook.SheetNames[0]]);
      if (!metaHeaders.length) throw new Error("Sheet 1 (Meta Ads) appears to be empty.");

      const metaAds = parseMetaRows(metaHeaders, metaRows);
      if (!metaAds.length) throw new Error("No valid ad rows found. Check columns: Ad Name, Amount Spent, Purchase Conversion Value, Website Purchases.");

      let googleData = null;
      if (workbook.SheetNames.length >= 2) {
        const { headers: gH, rows: gR } = sheetToHeadersRows(workbook.Sheets[workbook.SheetNames[1]]);
        if (gH.length && gR.length) googleData = parseGoogleRows(gH, gR);
      }

      onData({ metaAds, googleData, fileName: file.name });
    } catch (e) {
      setError(e.message || "Failed to read file.");
    } finally {
      setLoading(false);
    }
  }, [onData]);

  return (
    <div style={{ marginBottom: 28 }}>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) processFile(f); }}
        onClick={() => inputRef.current?.click()}
        style={{
          border: `2px dashed ${dragging ? "#c8a2f8" : "rgba(255,255,255,0.12)"}`,
          borderRadius: 14, padding: "36px 24px", textAlign: "center", cursor: "pointer",
          background: dragging ? "rgba(200,162,248,0.06)" : "rgba(255,255,255,0.02)",
          transition: "all 0.2s ease",
        }}
      >
        <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={(e) => { const f = e.target.files[0]; if (f) processFile(f); }} style={{ display: "none" }} />
        {loading ? (
          <div style={{ color: "#c8a2f8", fontSize: 14, fontWeight: 600 }}>â³ Reading fileâ¦</div>
        ) : (
          <>
            <div style={{ fontSize: 36, marginBottom: 12 }}>ð</div>
            <div style={{ fontSize: 15, fontWeight: 700, color: "#fff", marginBottom: 8 }}>
              Drop your Excel file here, or click to browse
            </div>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,0.35)", lineHeight: 1.8 }}>
              Accepts <strong style={{ color: "rgba(255,255,255,0.6)" }}>.xlsx</strong> files
              <br />
              <strong style={{ color: "rgba(255,255,255,0.5)" }}>Sheet 1</strong> â Meta Ads data &nbsp;Â·&nbsp;
              <strong style={{ color: "rgba(255,255,255,0.5)" }}>Sheet 2</strong> â Google Ads data (optional)
              <br />
              Required columns: <em>Ad Name Â· Amount Spent Â· Purchase Conversion Value Â· Website Purchases</em>
            </div>
          </>
        )}
      </div>

      {error && (
        <div style={{ marginTop: 10, padding: "10px 14px", borderRadius: 8, background: "rgba(248,113,113,0.1)", border: "1px solid rgba(248,113,113,0.25)", color: "#f87171", fontSize: 13 }}>
          â ï¸ {error}
        </div>
      )}

      <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
        <div style={{ flex: 1, height: 1, background: "rgba(255,255,255,0.06)" }} />
        <span style={{ fontSize: 11, color: "rgba(255,255,255,0.25)" }}>or</span>
        <div style={{ flex: 1, height: 1, background: "rgba(255,255,255,0.06)" }} />
      </div>
      <div style={{ textAlign: "center", marginTop: 14 }}>
        <button
          onClick={onDemo}
          style={{
            padding: "10px 24px", borderRadius: 8, border: "1px solid rgba(200,162,248,0.25)",
            background: "rgba(200,162,248,0.08)", color: "#c8a2f8", fontSize: 13, fontWeight: 600,
            cursor: "pointer", fontFamily: "inherit",
          }}
        >
          Load Demo Data (May 1â19, 2026)
        </button>
      </div>
    </div>
  );
}

// âââ MAIN DASHBOARD âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
export default function Dashboard() {
  const [metaAds, setMetaAds] = useState(null);
  const [googleData, setGoogleData] = useState(null);
  const [isDemo, setIsDemo] = useState(false);
  const [fileName, setFileName] = useState("");
  const [selectedProduct, setSelectedProduct] = useState("All Products");
  const [searchQuery, setSearchQuery] = useState("");
  const [sortField, setSortField] = useState("spend");
  const [sortDir, setSortDir] = useState("desc");
  const [adTypeFilter, setAdTypeFilter] = useState("All");

  const hasData = metaAds !== null;

  const loadDemo = useCallback(() => {
    setMetaAds(DEMO_ADS);
    setGoogleData(DEMO_GOOGLE);
    setIsDemo(true);
    setFileName("");
    setSelectedProduct("All Products");
    setSearchQuery("");
    setAdTypeFilter("All");
  }, []);

  const handleFileData = useCallback(({ metaAds: ads, googleData: gd, fileName: fn }) => {
    setMetaAds(ads);
    setGoogleData(gd);
    setIsDemo(false);
    setFileName(fn);
    setSelectedProduct("All Products");
    setSearchQuery("");
    setAdTypeFilter("All");
  }, []);

  const clearData = useCallback(() => {
    setMetaAds(null);
    setGoogleData(null);
    setIsDemo(false);
    setFileName("");
  }, []);

  const productTotals = useMemo(() => metaAds ? computeProductTotals(metaAds) : {}, [metaAds]);

  const presentProducts = useMemo(() => PRODUCTS.filter((p) => p === "All Products" || productTotals[p]), [productTotals]);

  const productData = productTotals[selectedProduct] || { spend: 0, sales: 0, purchases: 0, roas: 0 };

  const filteredAds = useMemo(() => {
    if (!metaAds) return [];
    let ads = selectedProduct !== "All Products" ? metaAds.filter((a) => a.product === selectedProduct) : metaAds;
    if (searchQuery) { const q = searchQuery.toLowerCase(); ads = ads.filter((a) => a.name.toLowerCase().includes(q) || a.product.toLowerCase().includes(q)); }
    if (adTypeFilter !== "All") ads = ads.filter((a) => a.type === adTypeFilter);
    return [...ads].sort((a, b) => sortDir === "desc" ? b[sortField] - a[sortField] : a[sortField] - b[sortField]);
  }, [metaAds, selectedProduct, searchQuery, sortField, sortDir, adTypeFilter]);

  const maxSpend = useMemo(() => filteredAds.reduce((m, a) => Math.max(m, a.spend), 0), [filteredAds]);

  const handleSort = (field) => {
    if (sortField === field) setSortDir(sortDir === "desc" ? "asc" : "desc");
    else { setSortField(field); setSortDir("desc"); }
  };

  const SortIcon = ({ field }) => {
    if (sortField !== field) return <span style={{ opacity: 0.3, fontSize: 10 }}>â</span>;
    return <span style={{ fontSize: 10 }}>{sortDir === "desc" ? "â" : "â"}</span>;
  };

  const cpa = productData.purchases > 0 ? productData.spend / productData.purchases : 0;
  const gd = googleData || { spend: 0, sales: 0, purchases: 0, roas: 0 };

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", color: "#e8e6f0", fontFamily: "'DM Sans', sans-serif", padding: "28px 24px" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600;700&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: rgba(255,255,255,0.03); }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 3px; }
        input::placeholder { color: rgba(255,255,255,0.3); }
        button { transition: opacity 0.15s ease; }
        button:hover { opacity: 0.8; }
      `}</style>

      {/* ââ HEADER ââ */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 8, height: 8, borderRadius: "50%", background: hasData ? "#34d399" : "#444", boxShadow: hasData ? "0 0 10px #34d399" : "none", transition: "all 0.3s" }} />
            <h1 style={{ fontSize: 22, fontWeight: 700, letterSpacing: "-0.5px", color: "#fff" }}>
              Better Herbs â Performance Dashboard
            </h1>
          </div>
          <p style={{ fontSize: 12, color: "rgba(255,255,255,0.35)", marginTop: 4, marginLeft: 18 }}>
            {isDemo
              ? "Demo data Â· May 1â19, 2026 Â· Meta + Google Ads"
              : hasData
              ? `Live data Â· ${fileName} Â· Meta${googleData ? " + Google" : " only"}`
              : "Upload your Excel file below to populate the dashboard"}
          </p>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {hasData && (
            <button onClick={clearData} style={{ padding: "6px 14px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(255,255,255,0.1)", background: "rgba(255,255,255,0.04)", color: "rgba(255,255,255,0.45)" }}>
              â Change File
            </button>
          )}
          <div style={{
            padding: "6px 14px", borderRadius: 8,
            background: isDemo ? "rgba(251,191,36,0.1)" : hasData ? "rgba(52,211,153,0.1)" : "rgba(255,255,255,0.04)",
            border: `1px solid ${isDemo ? "rgba(251,191,36,0.25)" : hasData ? "rgba(52,211,153,0.25)" : "rgba(255,255,255,0.08)"}`,
            fontSize: 12, fontWeight: 600,
            color: isDemo ? "#fbbf24" : hasData ? "#34d399" : "rgba(255,255,255,0.25)",
          }}>
            {isDemo ? "DEMO MODE" : hasData ? "LIVE DATA" : "NO DATA"}
          </div>
        </div>
      </div>

      {/* ââ UPLOAD ZONE (no data state) ââ */}
      {!hasData && <UploadZone onData={handleFileData} onDemo={loadDemo} />}

      {/* ââ DASHBOARD ââ */}
      {hasData && (
        <>
          {/* Product Tabs */}
          <div style={{ display: "flex", gap: 6, marginBottom: 24, overflowX: "auto", paddingBottom: 4 }}>
            {presentProducts.map((p) => (
              <button key={p} onClick={() => setSelectedProduct(p)} style={{
                padding: "8px 18px", borderRadius: 8,
                border: `1px solid ${selectedProduct === p ? PRODUCT_COLORS[p] : "rgba(255,255,255,0.08)"}`,
                background: selectedProduct === p ? `${PRODUCT_COLORS[p]}18` : "rgba(255,255,255,0.03)",
                color: selectedProduct === p ? PRODUCT_COLORS[p] : "rgba(255,255,255,0.5)",
                fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap",
              }}>
                {p}
              </button>
            ))}
          </div>

          {/* Metric Cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 14, marginBottom: 28 }}>
            {[
              { label: "Meta Spend", value: fmtFull(productData.spend), sub: selectedProduct },
              { label: "Meta Revenue", value: fmtFull(productData.sales), sub: "Reported" },
              { label: "Meta ROAS", value: `${productData.roas.toFixed(2)}x`, sub: productData.roas >= 2.5 ? "â Healthy" : "â  Needs Work", highlight: true },
              { label: "Purchases", value: productData.purchases.toLocaleString("en-IN"), sub: cpa > 0 ? `CPA: ${fmtFull(cpa)}` : "â" },
              { label: "Google Ads", value: gd.spend > 0 ? `${gd.roas.toFixed(2)}x` : "â", sub: gd.spend > 0 ? `${fmt(gd.spend)} spend` : "No Google data" },
            ].map((m, i) => (
              <div key={i} style={{
                background: m.highlight
                  ? "linear-gradient(135deg, rgba(200,162,248,0.08), rgba(92,164,247,0.08))"
                  : "rgba(255,255,255,0.025)",
                border: `1px solid ${m.highlight ? "rgba(200,162,248,0.2)" : "rgba(255,255,255,0.06)"}`,
                borderRadius: 14, padding: "18px 20px",
              }}>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.4)", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 8 }}>{m.label}</div>
                <div style={{ fontSize: 24, fontWeight: 700, fontFamily: "'JetBrains Mono', monospace", color: "#fff", marginBottom: 4 }}>{m.value}</div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.35)" }}>{m.sub}</div>
              </div>
            ))}
          </div>

          {/* Product Comparison Bar */}
          {selectedProduct === "All Products" && (
            <div style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 14, padding: 20, marginBottom: 28 }}>
              <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 16, color: "#fff" }}>Product-wise Spend & ROAS</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {PRODUCTS.filter((p) => p !== "All Products" && productTotals[p]).map((p) => {
                  const d = productTotals[p];
                  const pct = (d.spend / (productTotals["All Products"]?.spend || 1)) * 100;
                  return (
                    <div key={p} style={{ display: "flex", alignItems: "center", gap: 14 }}>
                      <div style={{ width: 120, fontSize: 12, fontWeight: 600, color: PRODUCT_COLORS[p], flexShrink: 0 }}>{p}</div>
                      <div style={{ flex: 1, height: 26, background: "rgba(255,255,255,0.04)", borderRadius: 6, position: "relative", overflow: "hidden" }}>
                        <div style={{ width: `${pct}%`, height: "100%", borderRadius: 6, background: `linear-gradient(90deg, ${PRODUCT_COLORS[p]}40, ${PRODUCT_COLORS[p]}80)`, transition: "width 0.6s ease" }} />
                        <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", fontSize: 11, fontWeight: 600, color: "rgba(255,255,255,0.6)", fontFamily: "'JetBrains Mono', monospace" }}>
                          {fmt(d.spend)}
                        </span>
                      </div>
                      <RoasBadge value={d.roas} />
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Ad Table */}
          <div style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 14, padding: 20 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 12 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: "#fff" }}>
                Ad-Level Performance
                <span style={{ fontSize: 12, fontWeight: 400, color: "rgba(255,255,255,0.35)", marginLeft: 8 }}>{filteredAds.length} ads</span>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {["All", "Doctor", "Static", "UGC"].map((t) => (
                  <button key={t} onClick={() => setAdTypeFilter(t)} style={{
                    padding: "5px 14px", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer",
                    border: `1px solid ${adTypeFilter === t ? "#c8a2f8" : "rgba(255,255,255,0.08)"}`,
                    background: adTypeFilter === t ? "rgba(200,162,248,0.12)" : "transparent",
                    color: adTypeFilter === t ? "#c8a2f8" : "rgba(255,255,255,0.45)",
                  }}>{t}</button>
                ))}
              </div>
            </div>

            <div style={{ position: "relative", marginBottom: 16 }}>
              <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", color: "rgba(255,255,255,0.3)", fontSize: 14 }}>ð</span>
              <input
                type="text" value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search by creator, ad name, or productâ¦"
                style={{ width: "100%", padding: "12px 16px 12px 40px", borderRadius: 10, border: "1px solid rgba(255,255,255,0.08)", background: "rgba(255,255,255,0.04)", color: "#e8e6f0", fontSize: 14, outline: "none" }}
              />
            </div>

            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 2px" }}>
                <thead>
                  <tr>
                    {[
                      { key: "name", label: "Ad Name", align: "left", sortable: false },
                      { key: "product", label: "Product", align: "left", sortable: false },
                      { key: "type", label: "Type", align: "center", sortable: false },
                      { key: "spend", label: "Spend", align: "right", sortable: true },
                      { key: "sales", label: "Revenue", align: "right", sortable: true },
                      { key: "purchases", label: "Orders", align: "right", sortable: true },
                      { key: "roas", label: "ROAS", align: "right", sortable: true },
                    ].map((col) => (
                      <th key={col.key} onClick={() => col.sortable && handleSort(col.key)} style={{
                        padding: "10px 12px", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px",
                        color: "rgba(255,255,255,0.35)", textAlign: col.align,
                        cursor: col.sortable ? "pointer" : "default",
                        borderBottom: "1px solid rgba(255,255,255,0.06)", whiteSpace: "nowrap", userSelect: "none",
                      }}>
                        {col.label} {col.sortable && <SortIcon field={col.key} />}
                      </th>
                    ))}
                    <th style={{ padding: "10px 12px", fontSize: 11, fontWeight: 600, color: "rgba(255,255,255,0.35)", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>SPEND SHARE</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredAds.map((ad, i) => (
                    <tr key={i} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.015)" : "transparent" }}>
                      <td style={{ padding: "11px 12px", fontSize: 13, fontWeight: 500, maxWidth: 280, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: "#fff" }}>{ad.name}</td>
                      <td style={{ padding: "11px 12px" }}>
                        <span style={{ fontSize: 11, fontWeight: 600, color: PRODUCT_COLORS[ad.product] || "#888", background: `${PRODUCT_COLORS[ad.product] || "#888"}15`, padding: "3px 8px", borderRadius: 4 }}>{ad.product}</span>
                      </td>
                      <td style={{ padding: "11px 12px", textAlign: "center" }}>
                        <span style={{
                          fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px",
                          color: ad.type === "Doctor" ? "#60d394" : ad.type === "Static" ? "#5ca4f7" : "#fbbf24",
                          background: ad.type === "Doctor" ? "rgba(96,211,148,0.1)" : ad.type === "Static" ? "rgba(92,164,247,0.1)" : "rgba(251,191,36,0.1)",
                          padding: "3px 8px", borderRadius: 4,
                        }}>{ad.type}</span>
                      </td>
                      <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmt(ad.spend)}</td>
                      <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmt(ad.sales)}</td>
                      <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600, color: "#fff" }}>{ad.purchases}</td>
                      <td style={{ padding: "11px 12px", textAlign: "right" }}><RoasBadge value={ad.roas} /></td>
                      <td style={{ padding: "11px 12px", width: 120 }}><MiniBar value={ad.spend} max={maxSpend} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filteredAds.length === 0 && (
                <div style={{ textAlign: "center", padding: 40, color: "rgba(255,255,255,0.3)", fontSize: 14 }}>No ads match your filters</div>
              )}
            </div>
          </div>

          <div style={{ textAlign: "center", padding: "24px 0 8px", fontSize: 11, color: "rgba(255,255,255,0.2)" }}>
            {isDemo ? "Demo data Â· " : `${fileName} Â· `}ROAS = Reported conversion value Ã· spend Â· Better Herbs Dashboard v2
          </div>
        </>
      )}
    </div>
  );
}
