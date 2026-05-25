import { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTS ─────────────────────────────────────────────────────────────
const PRODUCTS = ["All Products", "Lactify", "Brainify Powder", "Brainify Drops", "Mamafy", "Flow Joy Drops"];

const PRODUCT_COLORS = {
  "All Products": "#c8a2f8",
  "Lactify": "#60d394",
  "Brainify Powder": "#5ca4f7",
  "Brainify Drops": "#f7a75c",
  "Mamafy": "#f76b8a",
  "Flow Joy Drops": "#5cf7e0",
  "Other": "#888",
};

// ─── DEMO DATA ──────────────────────────────────────────────────────────────
const DEMO_ADS = [
  { product: "Brainify Drops", name: "Brainify drops -Maninder Kaur (Blessings) Video", spend: 67969.55, sales: 221287.18, purchases: 282, roas: 3.26, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Divya Bajpai", spend: 64246.65, sales: 231873.0, purchases: 266, roas: 3.61, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr. Vinod Video", spend: 52376.8, sales: 147822.2, purchases: 200, roas: 2.82, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Static - 13 (Ingredients)", spend: 49863.56, sales: 137496.4, purchases: 164, roas: 2.76, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr Vinod 2", spend: 34658.46, sales: 115426.0, purchases: 154, roas: 3.33, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr.sajid", spend: 32878.27, sales: 79342.0, purchases: 108, roas: 2.41, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Ankit Jha  Video", spend: 32495.09, sales: 115801.0, purchases: 149, roas: 3.56, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify_Dr. Nayana", spend: 26811.34, sales: 99339.0, purchases: 113, roas: 3.71, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr. Sajid", spend: 24048.87, sales: 55767.0, purchases: 77, roas: 2.32, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy -  Static - 13", spend: 23724.17, sales: 55242.0, purchases: 69, roas: 2.33, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static 7", spend: 22743.13, sales: 68917.0, purchases: 87, roas: 3.03, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Doctor Compilation Video - 2  | 07/05/2026", spend: 21392.4, sales: 60229.0, purchases: 78, roas: 2.82, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy - Dr.Shaifali Ad code", spend: 20628.52, sales: 60443.0, purchases: 77, roas: 2.93, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy -  Static - 4 - Do not buy", spend: 18427.24, sales: 49252.0, purchases: 55, roas: 2.67, type: "Static", date: "" },
  { product: "Flow Joy Drops", name: "Flowjoy Drop - Dispatch video", spend: 18162.81, sales: 34447.0, purchases: 47, roas: 1.9, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify - Dr Vinod - Post ID", spend: 18037.31, sales: 42490.0, purchases: 55, roas: 2.36, type: "Doctor", date: "" },
  { product: "Flow Joy Drops", name: "Flowjoy Drops Static 2", spend: 17914.65, sales: 31689.4, purchases: 36, roas: 1.77, type: "Static", date: "" },
  { product: "Mamafy", name: "Mamafy - Dr. Garima", spend: 17630.1, sales: 51597.0, purchases: 64, roas: 2.93, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify - Customer Review - Lactation", spend: 16193.81, sales: 37134.0, purchases: 40, roas: 2.29, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Prachi Mahajan", spend: 15666.91, sales: 40520.0, purchases: 43, roas: 2.59, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Rajmani", spend: 15066.62, sales: 37219.0, purchases: 54, roas: 2.47, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Pushpendra 2nd video", spend: 14784.0, sales: 33402.0, purchases: 46, roas: 2.26, type: "Doctor", date: "" },
  { product: "Flow Joy Drops", name: "Flow Drop - Dr. Ankit", spend: 14451.06, sales: 18739.0, purchases: 24, roas: 1.3, type: "Doctor", date: "" },
  { product: "Mamafy", name: "mamafy - Dr. Samra - Edit 2 (Direct) | 04/05/2026", spend: 12788.34, sales: 28639.0, purchases: 34, roas: 2.24, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Brainify Powder Static - 12   | 07/05/2026", spend: 12309.36, sales: 45744.0, purchases: 56, roas: 3.72, type: "Static", date: "" },
  { product: "Flow Joy Drops", name: "flowjoy drops - Dr. Garima | 13/05/2026", spend: 12049.17, sales: 38292.0, purchases: 56, roas: 3.18, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify - 2nd Doctor Compilation", spend: 10843.5, sales: 14092.0, purchases: 17, roas: 1.3, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mothers Day - Drops | 1/04/20206", spend: 10533.15, sales: 40082.0, purchases: 43, roas: 3.81, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Rajmani Patel - Edited | 07/05/2026", spend: 10080.6, sales: 19795.0, purchases: 22, roas: 1.96, type: "UGC", date: "" },
  { product: "Mamafy", name: "Mothers Day -Powder | 1/04/20206", spend: 9705.77, sales: 28433.0, purchases: 31, roas: 2.93, type: "UGC", date: "" },
  { product: "Mamafy", name: "mamafy - Static - 14| 13/05/2026", spend: 9613.21, sales: 26801.0, purchases: 30, roas: 2.79, type: "Static", date: "" },
  { product: "Mamafy", name: "mamafy - Founder's  Dispatch Video May MOF | 13/05/2026", spend: 9552.32, sales: 27110.0, purchases: 35, roas: 2.84, type: "UGC", date: "" },
  { product: "Mamafy", name: "Mamafy - Doctor Compilation Video", spend: 9274.86, sales: 27753.0, purchases: 37, roas: 2.99, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy -  Dr. Smriti - Edited", spend: 9199.56, sales: 20831.0, purchases: 27, roas: 2.26, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Dr. Ankit Jha | 04/05/2026", spend: 8025.41, sales: 20111.0, purchases: 28, roas: 2.51, type: "Doctor", date: "" },
  { product: "Flow Joy Drops", name: "Flow drops - Dr. Sushma Mogri", spend: 7863.78, sales: 13991.0, purchases: 19, roas: 1.78, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Drops-Dr. Vinod Video", spend: 7805.03, sales: 13150.0, purchases: 17, roas: 1.68, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static - 3", spend: 7551.16, sales: 16959.0, purchases: 20, roas: 2.25, type: "Static", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Naaz | 13/05/2026", spend: 7492.65, sales: 17982.0, purchases: 23, roas: 2.4, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static - 8", spend: 7171.39, sales: 20869.0, purchases: 23, roas: 2.91, type: "Static", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Rohit Bharadwaj", spend: 6237.82, sales: 4989.0, purchases: 7, roas: 0.8, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static 11", spend: 5840.19, sales: 9986.0, purchases: 13, roas: 1.71, type: "Static", date: "" },
  { product: "Mamafy", name: "mamafy - Dr. Samra - Edit 4 (Nutrition) | 04/05/2026", spend: 5478.25, sales: 11714.0, purchases: 14, roas: 2.14, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Static 19  | 04/05/2026", spend: 5067.23, sales: 11378.0, purchases: 12, roas: 2.25, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Rajmani Patel  | 04/05/2026", spend: 5063.75, sales: 5712.0, purchases: 7, roas: 1.13, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Static 20  | 04/05/2026", spend: 5062.61, sales: 8797.0, purchases: 12, roas: 1.74, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Static 18  | 04/05/2026", spend: 5031.59, sales: 12055.0, purchases: 14, roas: 2.4, type: "Static", date: "" },
  { product: "Mamafy", name: "Mothers Day - Mamafy | 1/04/20206", spend: 4392.87, sales: 12466.0, purchases: 14, roas: 2.84, type: "UGC", date: "" },
  { product: "Mamafy", name: "mamafy - Dr. Priya Soni| 13/05/2026", spend: 4179.31, sales: 6061.0, purchases: 9, roas: 1.45, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify - Dr Vinod 2", spend: 4062.07, sales: 11545.0, purchases: 15, roas: 2.84, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mothers Day - Drops | Static 2", spend: 3714.32, sales: 9918.0, purchases: 12, roas: 2.67, type: "Static", date: "" },
  { product: "Mamafy", name: "Mothers Day -Powder | Static 2 |", spend: 3713.58, sales: 9991.0, purchases: 9, roas: 2.69, type: "Static", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Garima | 13/05/2026", spend: 3596.55, sales: 3305.0, purchases: 5, roas: 0.92, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify - Static - 27 | 13/05/2026", spend: 3594.98, sales: 6071.0, purchases: 8, roas: 1.69, type: "Static", date: "" },
  { product: "Mamafy", name: "Mothers Day - Mamafy - Retargeting | 1/04/20206", spend: 3587.21, sales: 19657.0, purchases: 20, roas: 5.48, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder Static - 22 | 19/05/2026", spend: 3458.88, sales: 10841.0, purchases: 12, roas: 3.13, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Maninder Kaur - March - Edited 19/05/2026", spend: 3456.78, sales: 6320.0, purchases: 9, roas: 1.83, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr. Rajeswari (Telugu) 19/05/2026", spend: 3450.49, sales: 10025.0, purchases: 14, roas: 2.91, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static - 17  19/05/2026", spend: 3436.48, sales: 9073.0, purchases: 11, roas: 2.64, type: "Static", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Priyanka (Ped.)", spend: 3402.97, sales: 8444.0, purchases: 12, roas: 2.48, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static 12", spend: 3273.36, sales: 12712.0, purchases: 16, roas: 3.88, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder Dr. Shiba (Odia)i) | 19/05/2026", spend: 3145.09, sales: 1907.0, purchases: 3, roas: 0.61, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Dr. Taran (Punjabi) | 19/05/2026", spend: 3137.31, sales: 1638.0, purchases: 2, roas: 0.52, type: "Doctor", date: "" },
  { product: "Lactify", name: "Mothers Day - Lactify | 1/04/20206", spend: 3135.28, sales: 7452.0, purchases: 8, roas: 2.38, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Dr. Mallika (Telugu)   | 07/05/2026", spend: 2975.07, sales: 5662.0, purchases: 8, roas: 1.9, type: "Doctor", date: "" },
  { product: "Mamafy", name: "mamafy - Founder's  Dispatch Video May MOF", spend: 2968.58, sales: 9470.0, purchases: 11, roas: 3.19, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Prachi Mahajan | 07/05/2026", spend: 2804.62, sales: 1997.0, purchases: 3, roas: 0.71, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr. Padma 07/05/2026", spend: 2645.07, sales: 1098.0, purchases: 2, roas: 0.42, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Smriti   | 07/05/2026", spend: 2614.28, sales: 2946.0, purchases: 3, roas: 1.13, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy - Static 2", spend: 2569.87, sales: 7566.0, purchases: 10, roas: 2.94, type: "Static", date: "" },
  { product: "Mamafy", name: "mamafy - Dr. Soniya Gupta 1st Edit | 04/05/2026", spend: 2510.47, sales: 649.0, purchases: 1, roas: 0.26, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops -Brainify Drops Static - 16 | 07/05/2026", spend: 2253.78, sales: 3156.0, purchases: 4, roas: 1.4, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Brainify Powder Static - 17   | 07/05/2026", spend: 2253.59, sales: 4204.0, purchases: 5, roas: 1.87, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Brainify Powder Static - 16   | 07/05/2026", spend: 2229.56, sales: 2396.0, purchases: 3, roas: 1.07, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Brainify Powder Static - 14   | 07/05/2026", spend: 2212.35, sales: 2586.0, purchases: 4, roas: 1.17, type: "Static", date: "" },
  { product: "Mamafy", name: "Mamafy - Mamafy Static - 6 - USP | 23/03/2026", spend: 2193.49, sales: 5004.0, purchases: 6, roas: 2.28, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Divyani Bhagat   | 07/05/2026", spend: 2161.87, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder -Dr. Sonal | 04/05/2026", spend: 2111.65, sales: 4543.0, purchases: 4, roas: 2.15, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Pushpendra  | 07/05/2026", spend: 2101.8, sales: 1489.0, purchases: 1, roas: 0.71, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Sunil (Kannad)  | 21/05/2026", spend: 2050.93, sales: 1450.0, purchases: 2, roas: 0.71, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder April Dispatch Video  | 21/05/2026", spend: 2047.16, sales: 4343.0, purchases: 7, roas: 2.12, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr. Padma Tamil 21/05/2026", spend: 2036.85, sales: 2376.0, purchases: 4, roas: 1.17, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Manisha  | 21/05/2026", spend: 2012.61, sales: 5298.0, purchases: 5, roas: 2.63, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify - USP Static", spend: 2001.53, sales: 7490.0, purchases: 8, roas: 3.74, type: "Static", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Gunjan - 2nd Video | 21/05/2026", spend: 1999.73, sales: 3464.0, purchases: 4, roas: 1.73, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Priyanka Deswal", spend: 1921.48, sales: 1958.0, purchases: 2, roas: 1.02, type: "Doctor", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Brainify Powder Static - 15   | 07/05/2026", spend: 1835.38, sales: 649.0, purchases: 1, roas: 0.35, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify - Google Static", spend: 1782.66, sales: 5163.0, purchases: 7, roas: 2.9, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Ankit", spend: 1759.25, sales: 3355.0, purchases: 5, roas: 1.91, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr. Neha", spend: 1599.95, sales: 3096.0, purchases: 4, roas: 1.94, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Rina Arthaba", spend: 1597.39, sales: 3855.0, purchases: 5, roas: 2.41, type: "UGC", date: "" },
  { product: "Mamafy", name: "Mothers Day -Powder - RETARGETING | 1/04/20206", spend: 1504.58, sales: 1998.0, purchases: 2, roas: 1.33, type: "UGC", date: "" },
  { product: "Mamafy", name: "Mothers Day - Mamafy | Static 2", spend: 1393.15, sales: 5115.0, purchases: 5, roas: 3.67, type: "Static", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Riya", spend: 1340.95, sales: 4495.0, purchases: 5, roas: 3.35, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy_Dr. Tanya Video", spend: 1307.86, sales: 3802.0, purchases: 4, roas: 2.91, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy -  Static - 16", spend: 1302.76, sales: 709.0, purchases: 1, roas: 0.54, type: "Static", date: "" },
  { product: "Mamafy", name: "Mamafy -  Dispatch video - April", spend: 1283.93, sales: 2876.0, purchases: 4, roas: 2.24, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify - Divya Bajpai", spend: 1255.71, sales: 2098.0, purchases: 2, roas: 1.67, type: "UGC", date: "" },
  { product: "Mamafy", name: "Mothers Day - Drops -  Retargeting | 1/04/20206", spend: 1156.79, sales: 949.0, purchases: 1, roas: 0.82, type: "UGC", date: "" },
  { product: "Brainify Powder", name: "Brainify Powder - Dr. Suryakamal", spend: 1063.4, sales: 549.0, purchases: 1, roas: 0.52, type: "Doctor", date: "" },
  { product: "Lactify", name: "Lactify - Dr. Srimukhi (Telugu)", spend: 1047.0, sales: 649.0, purchases: 1, roas: 0.62, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Dr. Pillai (Tamil) (Partnership ad) - Tamil", spend: 1030.61, sales: 599.0, purchases: 1, roas: 0.58, type: "Doctor", date: "" },
  { product: "Lactify", name: "Mothers Day - Lactify | Static 2", spend: 946.52, sales: 2597.0, purchases: 3, roas: 2.74, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static - 7", spend: 885.39, sales: 1309.0, purchases: 1, roas: 1.48, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify - Static - 13 (Ingredients)", spend: 844.66, sales: 619.0, purchases: 1, roas: 0.73, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static - 5", spend: 834.67, sales: 599.0, purchases: 1, roas: 0.72, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify Drops -Maninder Kaur (Blessings) Video", spend: 762.2, sales: 2787.0, purchases: 3, roas: 3.66, type: "UGC", date: "" },
  { product: "Other", name: "Dr. Tanya - Ad code | 23/05/2026", spend: 681.79, sales: 1835.0, purchases: 3, roas: 2.69, type: "Doctor", date: "" },
  { product: "Mamafy", name: "mamafy - Mamafy Static - 17 | 23/05/2026", spend: 665.84, sales: 1400.0, purchases: 2, roas: 2.1, type: "Static", date: "" },
  { product: "Mamafy", name: "mamafy - Doctor Compilation video 1 - Mamafy- April | 23/05/2026", spend: 664.76, sales: 1450.0, purchases: 2, roas: 2.18, type: "Doctor", date: "" },
  { product: "Mamafy", name: "mamafy - Pooja Shah | 23/05/2026", spend: 657.23, sales: 1450.0, purchases: 2, roas: 2.21, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify - Dr. Ankit Jha  Video", spend: 607.92, sales: 2158.0, purchases: 2, roas: 3.55, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify - Maninder Kaur", spend: 553.92, sales: 619.0, purchases: 1, roas: 1.12, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Drops- Static 3", spend: 541.16, sales: 549.0, purchases: 1, roas: 1.01, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Drops - Dr. Rajeswari (Telugu) 19/05/2026", spend: 522.43, sales: 609.0, purchases: 1, roas: 1.17, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy - Dr. Soniya Gupta", spend: 515.64, sales: 2033.0, purchases: 3, roas: 3.94, type: "Doctor", date: "" },
  { product: "Mamafy", name: "Mamafy - static 2", spend: 502.05, sales: 649.0, purchases: 1, roas: 1.29, type: "Static", date: "" },
  { product: "Mamafy", name: "Mamafy - Doctor Compilation Video - March - 1", spend: 458.42, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Lactify", name: "Mothers Day - Lactify - Retargeting | 1/04/20206", spend: 425.43, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" },
  { product: "Mamafy", name: "Mamafy - Static - 6 - USP", spend: 402.78, sales: 1259.0, purchases: 1, roas: 3.13, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Drops-Maninder Kaur (Blessings) Video", spend: 346.46, sales: 619.0, purchases: 1, roas: 1.79, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Drops - Dr.sajid", spend: 327.75, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Other", name: "Aanchal Naherwa - Ad code | 23/05/2026", spend: 322.11, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify - Blessing Video", spend: 246.81, sales: 2896.0, purchases: 4, roas: 11.73, type: "UGC", date: "" },
  { product: "Mamafy", name: "Mamafy - Drishti 3rd Video", spend: 183.04, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify - INA_Hindi - Post ID", spend: 148.93, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Drops - Static 12", spend: 114.0, sales: 0.0, purchases: 0, roas: 0, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Drops- Static 7", spend: 92.32, sales: 0.0, purchases: 0, roas: 0, type: "Static", date: "" },
  { product: "Mamafy", name: "mamafy - Dr. Samra - Edit 2 (Direct)", spend: 89.07, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Drops-Dr Maninder Kaur (Copycat Brand)", spend: 70.28, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Drops - Doctor compilation video 2 | 19/04/2026", spend: 66.49, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Drops - Static 8", spend: 65.44, sales: 0.0, purchases: 0, roas: 0, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Drops-Static 11 Us Vs Them", spend: 65.21, sales: 0.0, purchases: 0, roas: 0, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Brainify drops - Static 16", spend: 60.39, sales: 0.0, purchases: 0, roas: 0, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Drops - Static 7", spend: 57.17, sales: 0.0, purchases: 0, roas: 0, type: "Static", date: "" },
  { product: "Brainify Drops", name: "Drops - Rajmani", spend: 43.07, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Drops - Dr. Padmavathi - kannada - 23/03/2026", spend: 37.7, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Brainify Drops", name: "Brainify - Prachi Mahajan", spend: 9.63, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Brainify drops -Neha", spend: 9.37, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" },
  { product: "Brainify Drops", name: "Drops-Dr Sonal", spend: 6.16, sales: 0.0, purchases: 0, roas: 0, type: "Doctor", date: "" },
  { product: "Mamafy", name: "mamafy - Founders dispatch", spend: 3.92, sales: 0.0, purchases: 0, roas: 0, type: "UGC", date: "" }
];

const DEMO_GOOGLE = null;

const DEMO_DAILY_BLENDED = [
  { date: "2026-05-23", spend: 43830.66, sales: 94526.0, purchases: 123, roas: 2.16, cpa: 356.35 },
  { date: "2026-05-22", spend: 39249.93, sales: 103475.0, purchases: 125, roas: 2.64, cpa: 314.0 },
  { date: "2026-05-21", spend: 38065.49, sales: 102297.0, purchases: 127, roas: 2.69, cpa: 299.73 },
  { date: "2026-05-20", spend: 36471.78, sales: 106799.0, purchases: 126, roas: 2.93, cpa: 289.46 },
  { date: "2026-05-19", spend: 38548.98, sales: 99435.0, purchases: 124, roas: 2.58, cpa: 310.88 },
  { date: "2026-05-18", spend: 35721.35, sales: 91701.0, purchases: 118, roas: 2.57, cpa: 302.72 },
  { date: "2026-05-17", spend: 36874.72, sales: 84342.4, purchases: 106, roas: 2.29, cpa: 347.87 },
  { date: "2026-05-16", spend: 33252.27, sales: 98164.0, purchases: 123, roas: 2.95, cpa: 270.34 },
  { date: "2026-05-15", spend: 42432.98, sales: 103787.0, purchases: 128, roas: 2.45, cpa: 331.51 },
  { date: "2026-05-14", spend: 44390.93, sales: 128652.0, purchases: 173, roas: 2.9, cpa: 256.59 },
  { date: "2026-05-13", spend: 47867.89, sales: 117977.0, purchases: 154, roas: 2.46, cpa: 310.83 },
  { date: "2026-05-12", spend: 43030.63, sales: 121527.0, purchases: 158, roas: 2.82, cpa: 272.35 },
  { date: "2026-05-11", spend: 46935.94, sales: 120852.0, purchases: 149, roas: 2.57, cpa: 315.01 },
  { date: "2026-05-10", spend: 53418.04, sales: 105984.0, purchases: 129, roas: 1.98, cpa: 414.09 },
  { date: "2026-05-09", spend: 47565.09, sales: 128505.0, purchases: 163, roas: 2.7, cpa: 291.81 },
  { date: "2026-05-08", spend: 50798.55, sales: 134546.0, purchases: 163, roas: 2.65, cpa: 311.65 },
  { date: "2026-05-07", spend: 48453.26, sales: 149250.0, purchases: 184, roas: 3.08, cpa: 263.33 },
  { date: "2026-05-06", spend: 45393.77, sales: 126607.4, purchases: 160, roas: 2.79, cpa: 283.71 },
  { date: "2026-05-05", spend: 45702.66, sales: 139363.0, purchases: 172, roas: 3.05, cpa: 265.71 },
  { date: "2026-05-04", spend: 48429.97, sales: 101234.18, purchases: 122, roas: 2.09, cpa: 396.97 },
  { date: "2026-05-03", spend: 42434.81, sales: 125170.2, purchases: 147, roas: 2.95, cpa: 288.67 },
  { date: "2026-05-02", spend: 36756.61, sales: 103453.0, purchases: 134, roas: 2.81, cpa: 274.3 },
  { date: "2026-05-01", spend: 41699.08, sales: 115057.0, purchases: 135, roas: 2.76, cpa: 308.88 }
];

const DEMO_DAILY_BY_PRODUCT = [
  { date: "2026-05-23", product: "Brainify Drops", spend: 10124.28, sales: 24331.0, purchases: 32, roas: 2.4, cpa: 316.38 },
  { date: "2026-05-23", product: "Brainify Powder", spend: 17078.09, sales: 42383.0, purchases: 53, roas: 2.48, cpa: 322.23 },
  { date: "2026-05-23", product: "Flow Joy Drops", spend: 1930.66, sales: 4177.0, purchases: 7, roas: 2.16, cpa: 275.81 },
  { date: "2026-05-23", product: "Lactify", spend: 5314.95, sales: 5728.0, purchases: 8, roas: 1.08, cpa: 664.37 },
  { date: "2026-05-23", product: "Mamafy", spend: 8378.78, sales: 16072.0, purchases: 20, roas: 1.92, cpa: 418.94 },
  { date: "2026-05-23", product: "Other", spend: 1003.9, sales: 1835.0, purchases: 3, roas: 1.83, cpa: 334.63 },
  { date: "2026-05-22", product: "Brainify Drops", spend: 9383.37, sales: 19968.0, purchases: 27, roas: 2.13, cpa: 347.53 },
  { date: "2026-05-22", product: "Brainify Powder", spend: 16325.34, sales: 51916.0, purchases: 64, roas: 3.18, cpa: 255.08 },
  { date: "2026-05-22", product: "Flow Joy Drops", spend: 1961.44, sales: 1759.0, purchases: 3, roas: 0.9, cpa: 653.81 },
  { date: "2026-05-22", product: "Lactify", spend: 6021.9, sales: 14942.0, purchases: 15, roas: 2.48, cpa: 401.46 },
  { date: "2026-05-22", product: "Mamafy", spend: 5557.88, sales: 14890.0, purchases: 16, roas: 2.68, cpa: 347.37 },
  { date: "2026-05-21", product: "Brainify Drops", spend: 10402.19, sales: 26071.0, purchases: 35, roas: 2.51, cpa: 297.21 },
  { date: "2026-05-21", product: "Brainify Powder", spend: 15382.09, sales: 43217.0, purchases: 53, roas: 2.81, cpa: 290.23 },
  { date: "2026-05-21", product: "Flow Joy Drops", spend: 1875.24, sales: 5762.0, purchases: 9, roas: 3.07, cpa: 208.36 },
  { date: "2026-05-21", product: "Lactify", spend: 5881.37, sales: 13267.0, purchases: 14, roas: 2.26, cpa: 420.1 },
  { date: "2026-05-21", product: "Mamafy", spend: 4524.6, sales: 13980.0, purchases: 16, roas: 3.09, cpa: 282.79 },
  { date: "2026-05-20", product: "Brainify Drops", spend: 10876.67, sales: 35197.0, purchases: 43, roas: 3.24, cpa: 252.95 },
  { date: "2026-05-20", product: "Brainify Powder", spend: 15118.71, sales: 41206.0, purchases: 48, roas: 2.73, cpa: 314.97 },
  { date: "2026-05-20", product: "Flow Joy Drops", spend: 2124.61, sales: 4985.0, purchases: 5, roas: 2.35, cpa: 424.92 },
  { date: "2026-05-20", product: "Lactify", spend: 3437.06, sales: 9314.0, purchases: 11, roas: 2.71, cpa: 312.46 },
  { date: "2026-05-20", product: "Mamafy", spend: 4914.73, sales: 16097.0, purchases: 19, roas: 3.28, cpa: 258.67 },
  { date: "2026-05-19", product: "Brainify Drops", spend: 11474.37, sales: 25133.0, purchases: 32, roas: 2.19, cpa: 358.57 },
  { date: "2026-05-19", product: "Brainify Powder", spend: 15993.13, sales: 37146.0, purchases: 48, roas: 2.32, cpa: 333.19 },
  { date: "2026-05-19", product: "Flow Joy Drops", spend: 2165.93, sales: 5750.0, purchases: 7, roas: 2.65, cpa: 309.42 },
  { date: "2026-05-19", product: "Lactify", spend: 3506.62, sales: 13921.0, purchases: 17, roas: 3.97, cpa: 206.27 },
  { date: "2026-05-19", product: "Mamafy", spend: 5408.93, sales: 17485.0, purchases: 20, roas: 3.23, cpa: 270.45 },
  { date: "2026-05-18", product: "Brainify Drops", spend: 11417.5, sales: 32823.0, purchases: 43, roas: 2.87, cpa: 265.52 },
  { date: "2026-05-18", product: "Brainify Powder", spend: 10620.06, sales: 33529.0, purchases: 41, roas: 3.16, cpa: 259.03 },
  { date: "2026-05-18", product: "Flow Joy Drops", spend: 3652.84, sales: 8031.0, purchases: 12, roas: 2.2, cpa: 304.4 },
  { date: "2026-05-18", product: "Lactify", spend: 4025.49, sales: 5703.0, purchases: 7, roas: 1.42, cpa: 575.07 },
  { date: "2026-05-18", product: "Mamafy", spend: 6005.46, sales: 11615.0, purchases: 15, roas: 1.93, cpa: 400.36 },
  { date: "2026-05-17", product: "Brainify Drops", spend: 11344.38, sales: 25076.0, purchases: 28, roas: 2.21, cpa: 405.16 },
  { date: "2026-05-17", product: "Brainify Powder", spend: 11229.49, sales: 23687.0, purchases: 33, roas: 2.11, cpa: 340.29 },
  { date: "2026-05-17", product: "Flow Joy Drops", spend: 3450.72, sales: 9234.4, purchases: 12, roas: 2.68, cpa: 287.56 },
  { date: "2026-05-17", product: "Lactify", spend: 3974.52, sales: 13837.0, purchases: 16, roas: 3.48, cpa: 248.41 },
  { date: "2026-05-17", product: "Mamafy", spend: 6875.61, sales: 12508.0, purchases: 17, roas: 1.82, cpa: 404.45 },
  { date: "2026-05-16", product: "Brainify Drops", spend: 11193.23, sales: 25553.0, purchases: 34, roas: 2.28, cpa: 329.21 },
  { date: "2026-05-16", product: "Brainify Powder", spend: 9551.08, sales: 31910.0, purchases: 40, roas: 3.34, cpa: 238.78 },
  { date: "2026-05-16", product: "Flow Joy Drops", spend: 2195.72, sales: 5836.0, purchases: 6, roas: 2.66, cpa: 365.95 },
  { date: "2026-05-16", product: "Lactify", spend: 3649.15, sales: 13344.0, purchases: 16, roas: 3.66, cpa: 228.07 },
  { date: "2026-05-16", product: "Mamafy", spend: 6663.09, sales: 21521.0, purchases: 27, roas: 3.23, cpa: 246.78 },
  { date: "2026-05-15", product: "Brainify Drops", spend: 14933.93, sales: 25908.0, purchases: 32, roas: 1.73, cpa: 466.69 },
  { date: "2026-05-15", product: "Brainify Powder", spend: 11134.2, sales: 40870.0, purchases: 52, roas: 3.67, cpa: 214.12 },
  { date: "2026-05-15", product: "Flow Joy Drops", spend: 3893.22, sales: 6028.0, purchases: 6, roas: 1.55, cpa: 648.87 },
  { date: "2026-05-15", product: "Lactify", spend: 4368.42, sales: 11545.0, purchases: 14, roas: 2.64, cpa: 312.03 },
  { date: "2026-05-15", product: "Mamafy", spend: 8103.21, sales: 19436.0, purchases: 24, roas: 2.4, cpa: 337.63 },
  { date: "2026-05-14", product: "Brainify Drops", spend: 14259.64, sales: 36273.0, purchases: 51, roas: 2.54, cpa: 279.6 },
  { date: "2026-05-14", product: "Brainify Powder", spend: 13205.97, sales: 46453.0, purchases: 62, roas: 3.52, cpa: 213.0 },
  { date: "2026-05-14", product: "Flow Joy Drops", spend: 4147.24, sales: 8454.0, purchases: 12, roas: 2.04, cpa: 345.6 },
  { date: "2026-05-14", product: "Lactify", spend: 4632.45, sales: 5363.0, purchases: 7, roas: 1.16, cpa: 661.78 },
  { date: "2026-05-14", product: "Mamafy", spend: 8145.63, sales: 32109.0, purchases: 41, roas: 3.94, cpa: 198.67 },
  { date: "2026-05-13", product: "Brainify Drops", spend: 15190.83, sales: 40352.0, purchases: 52, roas: 2.66, cpa: 292.13 },
  { date: "2026-05-13", product: "Brainify Powder", spend: 12783.27, sales: 38431.0, purchases: 49, roas: 3.01, cpa: 260.88 },
  { date: "2026-05-13", product: "Flow Joy Drops", spend: 3972.43, sales: 12291.0, purchases: 19, roas: 3.09, cpa: 209.08 },
  { date: "2026-05-13", product: "Lactify", spend: 5327.93, sales: 6651.0, purchases: 9, roas: 1.25, cpa: 591.99 },
  { date: "2026-05-13", product: "Mamafy", spend: 10593.43, sales: 20252.0, purchases: 25, roas: 1.91, cpa: 423.74 },
  { date: "2026-05-12", product: "Brainify Drops", spend: 15548.41, sales: 40390.0, purchases: 56, roas: 2.6, cpa: 277.65 },
  { date: "2026-05-12", product: "Brainify Powder", spend: 10987.5, sales: 32400.0, purchases: 44, roas: 2.95, cpa: 249.72 },
  { date: "2026-05-12", product: "Flow Joy Drops", spend: 2469.0, sales: 3331.0, purchases: 5, roas: 1.35, cpa: 493.8 },
  { date: "2026-05-12", product: "Lactify", spend: 3217.75, sales: 10307.0, purchases: 12, roas: 3.2, cpa: 268.15 },
  { date: "2026-05-12", product: "Mamafy", spend: 10807.97, sales: 35099.0, purchases: 41, roas: 3.25, cpa: 263.61 },
  { date: "2026-05-11", product: "Brainify Drops", spend: 15832.76, sales: 38875.0, purchases: 52, roas: 2.46, cpa: 304.48 },
  { date: "2026-05-11", product: "Brainify Powder", spend: 13528.02, sales: 42641.0, purchases: 49, roas: 3.15, cpa: 276.08 },
  { date: "2026-05-11", product: "Flow Joy Drops", spend: 3613.56, sales: 4063.0, purchases: 7, roas: 1.12, cpa: 516.22 },
  { date: "2026-05-11", product: "Lactify", spend: 3120.88, sales: 3854.0, purchases: 5, roas: 1.23, cpa: 624.18 },
  { date: "2026-05-11", product: "Mamafy", spend: 10840.72, sales: 31419.0, purchases: 36, roas: 2.9, cpa: 301.13 },
  { date: "2026-05-10", product: "Brainify Drops", spend: 16656.82, sales: 34582.0, purchases: 47, roas: 2.08, cpa: 354.4 },
  { date: "2026-05-10", product: "Brainify Powder", spend: 19206.94, sales: 39616.0, purchases: 47, roas: 2.06, cpa: 408.66 },
  { date: "2026-05-10", product: "Flow Joy Drops", spend: 3542.58, sales: 4546.0, purchases: 6, roas: 1.28, cpa: 590.43 },
  { date: "2026-05-10", product: "Lactify", spend: 3404.12, sales: 7898.0, purchases: 8, roas: 2.32, cpa: 425.51 },
  { date: "2026-05-10", product: "Mamafy", spend: 10607.58, sales: 19342.0, purchases: 21, roas: 1.82, cpa: 505.12 },
  { date: "2026-05-09", product: "Brainify Drops", spend: 14100.16, sales: 34277.0, purchases: 47, roas: 2.43, cpa: 300.0 },
  { date: "2026-05-09", product: "Brainify Powder", spend: 18963.9, sales: 44540.0, purchases: 55, roas: 2.35, cpa: 344.8 },
  { date: "2026-05-09", product: "Flow Joy Drops", spend: 2957.8, sales: 10768.0, purchases: 14, roas: 3.64, cpa: 211.27 },
  { date: "2026-05-09", product: "Lactify", spend: 2795.92, sales: 7461.0, purchases: 8, roas: 2.67, cpa: 349.49 },
  { date: "2026-05-09", product: "Mamafy", spend: 8747.31, sales: 31459.0, purchases: 39, roas: 3.6, cpa: 224.29 },
  { date: "2026-05-08", product: "Brainify Drops", spend: 15281.2, sales: 50870.0, purchases: 63, roas: 3.33, cpa: 242.56 },
  { date: "2026-05-08", product: "Brainify Powder", spend: 18863.02, sales: 39326.0, purchases: 48, roas: 2.08, cpa: 392.98 },
  { date: "2026-05-08", product: "Flow Joy Drops", spend: 3247.53, sales: 4630.0, purchases: 5, roas: 1.43, cpa: 649.51 },
  { date: "2026-05-08", product: "Lactify", spend: 3253.03, sales: 11087.0, purchases: 12, roas: 3.41, cpa: 271.09 },
  { date: "2026-05-08", product: "Mamafy", spend: 10153.77, sales: 28633.0, purchases: 35, roas: 2.82, cpa: 290.11 },
  { date: "2026-05-07", product: "Brainify Drops", spend: 16060.33, sales: 53222.0, purchases: 71, roas: 3.31, cpa: 226.2 },
  { date: "2026-05-07", product: "Brainify Powder", spend: 14456.23, sales: 35713.0, purchases: 44, roas: 2.47, cpa: 328.55 },
  { date: "2026-05-07", product: "Flow Joy Drops", spend: 3460.07, sales: 7879.0, purchases: 9, roas: 2.28, cpa: 384.45 },
  { date: "2026-05-07", product: "Lactify", spend: 3199.01, sales: 10080.0, purchases: 10, roas: 3.15, cpa: 319.9 },
  { date: "2026-05-07", product: "Mamafy", spend: 11277.62, sales: 42356.0, purchases: 50, roas: 3.76, cpa: 225.55 },
  { date: "2026-05-06", product: "Brainify Drops", spend: 14785.16, sales: 43767.0, purchases: 61, roas: 2.96, cpa: 242.38 },
  { date: "2026-05-06", product: "Brainify Powder", spend: 13758.07, sales: 43699.4, purchases: 52, roas: 3.18, cpa: 264.58 },
  { date: "2026-05-06", product: "Flow Joy Drops", spend: 3301.72, sales: 3541.0, purchases: 5, roas: 1.07, cpa: 660.34 },
  { date: "2026-05-06", product: "Lactify", spend: 3106.65, sales: 10793.0, purchases: 13, roas: 3.47, cpa: 238.97 },
  { date: "2026-05-06", product: "Mamafy", spend: 10442.17, sales: 24807.0, purchases: 29, roas: 2.38, cpa: 360.07 },
  { date: "2026-05-05", product: "Brainify Drops", spend: 14674.04, sales: 43178.0, purchases: 57, roas: 2.94, cpa: 257.44 },
  { date: "2026-05-05", product: "Brainify Powder", spend: 13959.86, sales: 50180.0, purchases: 55, roas: 3.59, cpa: 253.82 },
  { date: "2026-05-05", product: "Flow Joy Drops", spend: 3248.78, sales: 5030.0, purchases: 6, roas: 1.55, cpa: 541.46 },
  { date: "2026-05-05", product: "Lactify", spend: 3320.09, sales: 6142.0, purchases: 8, roas: 1.85, cpa: 415.01 },
  { date: "2026-05-05", product: "Mamafy", spend: 10499.89, sales: 34833.0, purchases: 46, roas: 3.32, cpa: 228.26 },
  { date: "2026-05-04", product: "Brainify Drops", spend: 14866.2, sales: 31614.18, purchases: 41, roas: 2.13, cpa: 362.59 },
  { date: "2026-05-04", product: "Brainify Powder", spend: 13877.6, sales: 33057.0, purchases: 40, roas: 2.38, cpa: 346.94 },
  { date: "2026-05-04", product: "Flow Joy Drops", spend: 3362.18, sales: 625.0, purchases: 1, roas: 0.19, cpa: 3362.18 },
  { date: "2026-05-04", product: "Lactify", spend: 3744.26, sales: 12643.0, purchases: 16, roas: 3.38, cpa: 234.02 },
  { date: "2026-05-04", product: "Mamafy", spend: 12579.73, sales: 23295.0, purchases: 24, roas: 1.85, cpa: 524.16 },
  { date: "2026-05-03", product: "Brainify Drops", spend: 14750.04, sales: 48950.2, purchases: 60, roas: 3.32, cpa: 245.83 },
  { date: "2026-05-03", product: "Brainify Powder", spend: 10327.43, sales: 35178.0, purchases: 40, roas: 3.41, cpa: 258.19 },
  { date: "2026-05-03", product: "Flow Joy Drops", spend: 3379.42, sales: 5353.0, purchases: 7, roas: 1.58, cpa: 482.77 },
  { date: "2026-05-03", product: "Lactify", spend: 4287.98, sales: 9248.0, purchases: 11, roas: 2.16, cpa: 389.82 },
  { date: "2026-05-03", product: "Mamafy", spend: 9689.94, sales: 26441.0, purchases: 29, roas: 2.73, cpa: 334.14 },
  { date: "2026-05-02", product: "Brainify Drops", spend: 13499.33, sales: 41622.0, purchases: 55, roas: 3.08, cpa: 245.44 },
  { date: "2026-05-02", product: "Brainify Powder", spend: 7858.41, sales: 23058.0, purchases: 31, roas: 2.93, cpa: 253.5 },
  { date: "2026-05-02", product: "Flow Joy Drops", spend: 2966.5, sales: 8183.0, purchases: 11, roas: 2.76, cpa: 269.68 },
  { date: "2026-05-02", product: "Lactify", spend: 3671.87, sales: 5852.0, purchases: 7, roas: 1.59, cpa: 524.55 },
  { date: "2026-05-02", product: "Mamafy", spend: 8760.5, sales: 24738.0, purchases: 30, roas: 2.82, cpa: 292.02 },
  { date: "2026-05-01", product: "Brainify Drops", spend: 16116.34, sales: 46941.0, purchases: 57, roas: 2.91, cpa: 282.74 },
  { date: "2026-05-01", product: "Brainify Powder", spend: 8710.82, sales: 41988.0, purchases: 48, roas: 4.82, cpa: 181.48 },
  { date: "2026-05-01", product: "Flow Joy Drops", spend: 3522.28, sales: 6902.0, purchases: 8, roas: 1.96, cpa: 440.29 },
  { date: "2026-05-01", product: "Lactify", spend: 4451.18, sales: 5244.0, purchases: 6, roas: 1.18, cpa: 741.86 },
  { date: "2026-05-01", product: "Mamafy", spend: 8898.46, sales: 13982.0, purchases: 16, roas: 1.57, cpa: 556.15 }
];

// ─── PARSING HELPERS ────────────────────────────────────────────────────────
function extractProduct(adName, campaignName) {
  for (const src of [(adName || "").toLowerCase(), (campaignName || "").toLowerCase()]) {
    if (src.includes("lactify")) return "Lactify";
    if (src.includes("brainify powder") || (src.includes("brainify") && src.includes("powder"))) return "Brainify Powder";
    if (src.includes("brainify drops") || (src.includes("brainify") && src.includes("drop"))) return "Brainify Drops";
    if (src.includes("mamafy") || src.includes("mothers day") || src.includes("mother's day")) return "Mamafy";
    if (src.includes("flow joy") || src.includes("flowjoy") || src.includes("flow drop") || src.includes("flow drops")) return "Flow Joy Drops";
    if ((src.includes("drops") || src.includes("drop")) && !src.includes("mamafy")) return "Brainify Drops";
    if (src.includes("brainify")) return "Brainify Drops";
  }
  return "Other";
}

function extractAdType(adName) {
  const n = (adName || "").toLowerCase();
  if (n.includes("dr.") || n.includes("dr ") || n.includes("doctor") || n.includes("compilation")) return "Doctor";
  if (n.includes("static") || n.includes("usp") || n.includes("ingredient")) return "Static";
  return "UGC";
}


function extractCreator(adName) {
  const n = (adName || '').trim();
  if (/static/i.test(n) && !/dr\.?\s/i.test(n)) return 'Static Creative';
  if (/dispatch/i.test(n) && !/dr\.?\s/i.test(n)) return 'Dispatch Video';
  if (/compilation/i.test(n) && !/dr\.?\s/i.test(n)) return 'Compilation';
  if (/founder/i.test(n)) return "Founder's Video";
  const drMatch = n.match(/Dr\.?\s+([A-Za-z]+(?:\s+[A-Za-z]+)?)/i);
  if (drMatch) return drMatch[0].replace(/\s+/g, ' ').trim();
  const parts = n.split(/\s*[-–]\s*/);
  for (let i = 1; i < parts.length; i++) {
    const p = parts[i].replace(/\s*\|.*$/, '').replace(/\s*\(.*$/, '').replace(/\d.*$/, '').trim();
    if (p && p.length > 2 && p.length < 35 && !/^(static|video|drops|powder|ad|code|usp|google|ingredient)/i.test(p)) return p;
  }
  return 'Other';
}

function computeCreatorTotals(ads) {
  const totals = {};
  for (const ad of ads) {
    const creator = extractCreator(ad.name);
    if (!totals[creator]) totals[creator] = { creator, spend: 0, sales: 0, purchases: 0, count: 0 };
    totals[creator].spend += ad.spend;
    totals[creator].sales += ad.sales;
    totals[creator].purchases += ad.purchases;
    totals[creator].count += 1;
  }
  return Object.values(totals)
    .map(d => ({ ...d, roas: d.spend > 0 && d.sales > 0 ? d.sales / d.spend : 0, cpa: d.purchases > 0 ? d.spend / d.purchases : 0 }))
    .sort((a, b) => b.spend - a.spend);
}

function exportCSV(data, filename) {
  const headers = ['Ad Name','Product','Creator','Type','Date','Spend (INR)','Revenue (INR)','Purchases','ROAS'];
  const rows = data.map(ad => [ad.name, ad.product, extractCreator(ad.name), ad.type, ad.date || '', ad.spend.toFixed(2), ad.sales.toFixed(2), ad.purchases, ad.roas.toFixed(2)]);
  const csv = [headers, ...rows].map(r => r.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',')).join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = filename || 'ads-export.csv'; a.click();
  URL.revokeObjectURL(url);
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
  const s = String(val).replace(/[₹,\s]/g, "").replace(/[^0-9.-]/g, "");
  return parseFloat(s) || 0;
}

function parseMetaRows(headers, rows) {
  const nameIdx = findColIdx(headers, ["ad name", "name"]);
  const spendIdx = findColIdx(headers, ["amount spent", "spend", "cost"]);
  const salesIdx = findColIdx(headers, ["purchase conversion value", "conversion value", "revenue", "value"]);
  const purchasesIdx = findColIdx(headers, ["website purchases", "purchases", "conversions"]);
  const campaignIdx = findColIdx(headers, ["campaign name", "campaign"]);
  const dateIdx = findColIdx(headers, ["reporting starts", "day", "date", "report date", "reporting"]);

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
    const date = dateIdx >= 0 ? String(row[dateIdx] || "").trim() : "";
    ads.push({ name, spend, sales, purchases, roas, product, type, date });
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

// ─── AGGREGATION ────────────────────────────────────────────────────────────
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

// ─── DAILY AGGREGATION (LOOKER STUDIO) ──────────────────────────────────────
function computeDailyBlended(ads) {
  const byDate = {};
  for (const ad of ads) {
    const d = ad.date || "—";
    if (!byDate[d]) byDate[d] = { date: d, spend: 0, sales: 0, purchases: 0 };
    byDate[d].spend += ad.spend;
    byDate[d].sales += ad.sales;
    byDate[d].purchases += ad.purchases;
  }
  return Object.values(byDate)
    .map((d) => ({
      ...d,
      roas: d.spend > 0 && d.sales > 0 ? d.sales / d.spend : 0,
      cpa: d.purchases > 0 ? d.spend / d.purchases : 0,
    }))
    .sort((a, b) => new Date(b.date) - new Date(a.date));
}

function computeDailyByProduct(ads) {
  const map = {};
  for (const ad of ads) {
    const d = ad.date || "—";
    const key = `${d}||${ad.product}`;
    if (!map[key]) map[key] = { date: d, product: ad.product, spend: 0, sales: 0, purchases: 0 };
    map[key].spend += ad.spend;
    map[key].sales += ad.sales;
    map[key].purchases += ad.purchases;
  }
  return Object.values(map)
    .map((d) => ({
      ...d,
      roas: d.spend > 0 && d.sales > 0 ? d.sales / d.spend : 0,
      cpa: d.purchases > 0 ? d.spend / d.purchases : 0,
    }))
    .sort((a, b) => new Date(b.date) - new Date(a.date) || a.product.localeCompare(b.product));
}

// ─── FORMATTERS ─────────────────────────────────────────────────────────────
const fmt = (n) => {
  if (n >= 100000) return `₹${(n / 100000).toFixed(2)}L`;
  if (n >= 1000) return `₹${(n / 1000).toFixed(1)}K`;
  return `₹${n.toFixed(0)}`;
};
const fmtFull = (n) => `₹${n.toLocaleString("en-IN", { maximumFractionDigits: 0 })}`;

// ─── UI: ROAS BADGE ─────────────────────────────────────────────────────────
function RoasBadge({ value }) {
  const color = value >= 3 ? "#34d399" : value >= 2 ? "#fbbf24" : "#f87171";
  const bg = value >= 3 ? "rgba(52,211,153,0.12)" : value >= 2 ? "rgba(251,191,36,0.12)" : "rgba(248,113,113,0.12)";
  return (
    <span style={{ background: bg, color, padding: "3px 10px", borderRadius: 6, fontWeight: 700, fontSize: 13, fontFamily: "'JetBrains Mono', monospace" }}>
      {value.toFixed(2)}x
    </span>
  );
}

// ─── UI: MINI BAR ───────────────────────────────────────────────────────────
function MiniBar({ value, max }) {
  const pct = max > 0 ? Math.min((value / max) * 100, 100) : 0;
  return (
    <div style={{ width: "100%", height: 6, background: "rgba(255,255,255,0.06)", borderRadius: 3 }}>
      <div style={{ width: `${pct}%`, height: "100%", background: "linear-gradient(90deg, #5ca4f7, #c8a2f8)", borderRadius: 3, transition: "width 0.5s ease" }} />
    </div>
  );
}

// ─── UI: UPLOAD ZONE ────────────────────────────────────────────────────────
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
          <div style={{ color: "#c8a2f8", fontSize: 14, fontWeight: 600 }}>⏳ Reading file…</div>
        ) : (
          <>
            <div style={{ fontSize: 36, marginBottom: 12 }}>📊</div>
            <div style={{ fontSize: 15, fontWeight: 700, color: "#fff", marginBottom: 8 }}>
              Drop your Excel file here, or click to browse
            </div>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,0.35)", lineHeight: 1.8 }}>
              Accepts <strong style={{ color: "rgba(255,255,255,0.6)" }}>.xlsx</strong> files
              <br />
              <strong style={{ color: "rgba(255,255,255,0.5)" }}>Sheet 1</strong> → Meta Ads data &nbsp;·&nbsp;
              <strong style={{ color: "rgba(255,255,255,0.5)" }}>Sheet 2</strong> → Google Ads data (optional)
              <br />
              Required columns: <em>Ad Name · Amount Spent · Purchase Conversion Value · Website Purchases</em>
            </div>
          </>
        )}
      </div>

      {error && (
        <div style={{ marginTop: 10, padding: "10px 14px", borderRadius: 8, background: "rgba(248,113,113,0.1)", border: "1px solid rgba(248,113,113,0.25)", color: "#f87171", fontSize: 13 }}>
          ⚠️ {error}
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
          Load Demo Data (May 1–23, 2026)
        </button>
      </div>
    </div>
  );
}


// ─── UI: LOOKER STUDIO VIEW ──────────────────────────────────────────────────
function LookerStudioView({ metaAds, preBlended, preByProduct }) {
  const [productFilter, setProductFilter] = useState("All Products");
  const availableProducts = useMemo(() => ["All Products", ...Array.from(new Set((preByProduct || []).map(r => r.product))).filter(Boolean).sort()], [preByProduct]);
  const blended = useMemo(() => preBlended || computeDailyBlended(metaAds), [metaAds, preBlended]);
  const byProduct = useMemo(() => preByProduct || computeDailyByProduct(metaAds), [metaAds, preByProduct]);
  const hasDate = blended.length > 0 && blended[0].date !== "Unknown";

  const thStyle = {
    padding: "10px 12px", fontSize: 11, fontWeight: 600, textTransform: "uppercase",
    letterSpacing: "0.5px", color: "rgba(255,255,255,0.35)", textAlign: "right",
    borderBottom: "1px solid rgba(255,255,255,0.06)", whiteSpace: "nowrap",
  };
  const thLeft = { ...thStyle, textAlign: "left" };
  const td = (align = "right") => ({
    padding: "10px 12px", fontSize: 13,
    fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)",
    textAlign: align,
  });


  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 28 }}>

      {/* ── Blended Day-wise ── */}
      <div style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 14, padding: 20 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: "#fff", marginBottom: 16 }}>
          Blended Day-wise Performance
          <span style={{ fontSize: 12, fontWeight: 400, color: "rgba(255,255,255,0.35)", marginLeft: 8 }}>All products combined</span>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 2px" }}>
            <thead>
              <tr>
                <th style={thLeft}>Date</th>
                <th style={thStyle}>Spend</th>
                <th style={thStyle}>Revenue</th>
                <th style={thStyle}>ROAS</th>
                <th style={thStyle}>Orders</th>
                <th style={thStyle}>CPA</th>
              </tr>
            </thead>
            <tbody>
              {blended.map((row, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.015)" : "transparent" }}>
                  <td style={{ ...td("left"), color: "#fff", fontWeight: 600 }}>{row.date}</td>
                  <td style={td()}>{fmtFull(row.spend)}</td>
                  <td style={td()}>{fmtFull(row.sales)}</td>
                  <td style={{ ...td(), textAlign: "right" }}><RoasBadge value={row.roas} /></td>
                  <td style={{ ...td(), color: "#fff", fontWeight: 600 }}>{row.purchases.toLocaleString("en-IN")}</td>
                  <td style={td()}>{row.cpa > 0 ? fmtFull(row.cpa) : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* ── Bifurcated Day-wise ── */}
      <div style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 14, padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: "#fff" }}>Day-wise by Product</div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {availableProducts.map(p => (
              <button key={p} onClick={() => setProductFilter(p)} style={{ padding: "4px 12px", borderRadius: 6, fontSize: 11, fontWeight: 600, cursor: "pointer", border: `1px solid ${productFilter === p ? (PRODUCT_COLORS[p] || "#c8a2f8") : "rgba(255,255,255,0.08)"}`, background: productFilter === p ? `${PRODUCT_COLORS[p] || "#c8a2f8"}18` : "transparent", color: productFilter === p ? (PRODUCT_COLORS[p] || "#c8a2f8") : "rgba(255,255,255,0.4)", fontFamily: "inherit" }}>{p}</button>
            ))}
          </div>
          <button onClick={() => exportCSV((preByProduct || []).filter(r => productFilter === "All Products" || r.product === productFilter).map(r => ({...r, name: r.product + " (" + r.date + ")", type: "", roas: r.roas, sales: r.sales, spend: r.spend, purchases: r.purchases})), "daily-product-export.csv")} style={{ marginLeft: "auto", padding: "5px 12px", borderRadius: 6, fontSize: 11, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(92,164,247,0.3)", background: "rgba(92,164,247,0.08)", color: "#5ca4f7", fontFamily: "inherit" }}>↓ Export</button>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 2px" }}>
            <thead>
              <tr>
                <th style={thLeft}>Date</th>
                <th style={thLeft}>Product</th>
                <th style={thStyle}>Spend</th>
                <th style={thStyle}>Revenue</th>
                <th style={thStyle}>ROAS</th>
                <th style={thStyle}>Orders</th>
                <th style={thStyle}>CPA</th>
              </tr>
            </thead>
            <tbody>
              {(byProduct.filter(r => productFilter === "All Products" || r.product === productFilter)).map((row, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.015)" : "transparent" }}>
                  <td style={{ ...td("left"), color: "rgba(255,255,255,0.6)" }}>{row.date}</td>
                  <td style={{ padding: "10px 12px" }}>
                    <span style={{ fontSize: 11, fontWeight: 600, color: PRODUCT_COLORS[row.product] || "#888", background: `${PRODUCT_COLORS[row.product] || "#888"}15`, padding: "3px 8px", borderRadius: 4 }}>{row.product}</span>
                  </td>
                  <td style={td()}>{fmtFull(row.spend)}</td>
                  <td style={td()}>{fmtFull(row.sales)}</td>
                  <td style={{ ...td(), textAlign: "right" }}><RoasBadge value={row.roas} /></td>
                  <td style={{ ...td(), color: "#fff", fontWeight: 600 }}>{row.purchases.toLocaleString("en-IN")}</td>
                  <td style={td()}>{row.cpa > 0 ? fmtFull(row.cpa) : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

    </div>
  );
}

// ─── MAIN DASHBOARD ─────────────────────────────────────────────────────────
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
  const [activeView, setActiveView] = useState("dashboard");
  const [demoBlended, setDemoBlended] = useState(null);
  const [demoByProduct, setDemoByProduct] = useState(null);
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [creatorFilter, setCreatorFilter] = useState("All");
  const [creatorProductFilter, setCreatorProductFilter] = useState("All Products");

  const hasData = metaAds !== null;

  const loadDemo = useCallback(() => {
    setMetaAds(DEMO_ADS);
    setGoogleData(DEMO_GOOGLE);
    setIsDemo(true);
    setFileName("");
    setSelectedProduct("All Products");
    setSearchQuery("");
    setAdTypeFilter("All");
    setDemoBlended(DEMO_DAILY_BLENDED);
    setDemoByProduct(DEMO_DAILY_BY_PRODUCT);
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
    setDemoBlended(null);
    setDemoByProduct(null);
  }, []);

  const productTotals = useMemo(() => metaAds ? computeProductTotals(metaAds) : {}, [metaAds]);

  const presentProducts = useMemo(() => PRODUCTS.filter((p) => p === "All Products" || productTotals[p]), [productTotals]);

  const productData = productTotals[selectedProduct] || { spend: 0, sales: 0, purchases: 0, roas: 0 };

  const filteredAds = useMemo(() => {
    if (!metaAds) return [];
    let ads = selectedProduct !== "All Products" ? metaAds.filter((a) => a.product === selectedProduct) : metaAds;
    if (searchQuery) { const q = searchQuery.toLowerCase(); ads = ads.filter((a) => a.name.toLowerCase().includes(q) || a.product.toLowerCase().includes(q)); }
    if (adTypeFilter !== "All") ads = ads.filter((a) => a.type === adTypeFilter);
    if (creatorFilter !== "All") ads = ads.filter((a) => extractCreator(a.name) === creatorFilter);
    if (dateFrom) ads = ads.filter((a) => a.date && a.date >= dateFrom);
    if (dateTo) ads = ads.filter((a) => a.date && a.date <= dateTo);
    return [...ads].sort((a, b) => sortDir === "desc" ? b[sortField] - a[sortField] : a[sortField] - b[sortField]);
  }, [metaAds, selectedProduct, searchQuery, sortField, sortDir, adTypeFilter, creatorFilter, dateFrom, dateTo]);

  const maxSpend = useMemo(() => filteredAds.reduce((m, a) => Math.max(m, a.spend), 0), [filteredAds]);
  const uniqueCreators = useMemo(() => { if (!metaAds) return []; const s = new Set(metaAds.map(a => extractCreator(a.name))); return ["All", ...Array.from(s).sort()]; }, [metaAds]);
  const creatorTotals = useMemo(() => metaAds ? computeCreatorTotals(creatorProductFilter !== "All Products" ? metaAds.filter(a => a.product === creatorProductFilter) : (dateFrom || dateTo ? filteredAds : metaAds)) : [], [metaAds, creatorProductFilter, filteredAds, dateFrom, dateTo]);
  const activeFiltersCount = [adTypeFilter !== "All", creatorFilter !== "All", dateFrom !== "", dateTo !== "", searchQuery !== ""].filter(Boolean).length;
  const clearAllFilters = () => { setAdTypeFilter("All"); setCreatorFilter("All"); setDateFrom(""); setDateTo(""); setSearchQuery(""); };

  const handleSort = (field) => {
    if (sortField === field) setSortDir(sortDir === "desc" ? "asc" : "desc");
    else { setSortField(field); setSortDir("desc"); }
  };

  const SortIcon = ({ field }) => {
    if (sortField !== field) return <span style={{ opacity: 0.3, fontSize: 10 }}>⇅</span>;
    return <span style={{ fontSize: 10 }}>{sortDir === "desc" ? "↓" : "↑"}</span>;
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

      {/* ── HEADER ── */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 8, height: 8, borderRadius: "50%", background: hasData ? "#34d399" : "#444", boxShadow: hasData ? "0 0 10px #34d399" : "none", transition: "all 0.3s" }} />
            <h1 style={{ fontSize: 22, fontWeight: 700, letterSpacing: "-0.5px", color: "#fff" }}>
              Better Herbs — Performance Dashboard
            </h1>
          </div>
          <p style={{ fontSize: 12, color: "rgba(255,255,255,0.35)", marginTop: 4, marginLeft: 18 }}>
            {isDemo
              ? "Demo data · May 1–23, 2026 · Meta + Google Ads"
              : hasData
              ? `Live data · ${fileName} · Meta${googleData ? " + Google" : " only"}`
              : "Upload your Excel file below to populate the dashboard"}
          </p>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {hasData && (
            <button onClick={clearData} style={{ padding: "6px 14px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(255,255,255,0.1)", background: "rgba(255,255,255,0.04)", color: "rgba(255,255,255,0.45)" }}>
              ↑ Change File
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
        {hasData && (
          <div style={{ display: "flex", gap: 4, background: "rgba(255,255,255,0.03)", borderRadius: 10, padding: 4, border: "1px solid rgba(255,255,255,0.06)" }}>
            {[{ id: "dashboard", label: "📊 Dashboard" }, { id: "creators", label: "👤 Creators" }, { id: "looker", label: "📅 Looker Studio" }].map(({ id, label }) => (
              <button
                key={id}
                onClick={() => setActiveView(id)}
                style={{
                  padding: "6px 16px", borderRadius: 7, fontSize: 12, fontWeight: 600,
                  cursor: "pointer", border: "none", fontFamily: "inherit",
                  background: activeView === id ? "rgba(200,162,248,0.15)" : "transparent",
                  color: activeView === id ? "#c8a2f8" : "rgba(255,255,255,0.4)",
                  transition: "all 0.15s",
                }}
              >{label}</button>
            ))}
          </div>
        )}
        {hasData && (
          <div style={{ display: "flex", gap: 10, padding: "10px 0 4px", alignItems: "center", flexWrap: "wrap" }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
              <label style={{ fontSize: 10, color: "rgba(255,255,255,0.35)", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px" }}>From</label>
              <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} style={{ padding: "5px 10px", borderRadius: 7, border: "1px solid rgba(255,255,255,0.1)", background: "rgba(255,255,255,0.04)", color: "#e8e6f0", fontSize: 12, outline: "none", cursor: "pointer" }} />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
              <label style={{ fontSize: 10, color: "rgba(255,255,255,0.35)", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px" }}>To</label>
              <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} style={{ padding: "5px 10px", borderRadius: 7, border: "1px solid rgba(255,255,255,0.1)", background: "rgba(255,255,255,0.04)", color: "#e8e6f0", fontSize: 12, outline: "none", cursor: "pointer" }} />
            </div>
            {(dateFrom || dateTo) && (
              <button onClick={() => { setDateFrom(""); setDateTo(""); }} style={{ alignSelf: "flex-end", padding: "6px 12px", borderRadius: 7, fontSize: 11, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(248,113,113,0.3)", background: "rgba(248,113,113,0.08)", color: "#f87171", fontFamily: "inherit" }}>✕ Clear dates</button>
            )}
            {(dateFrom || dateTo) && (
              <span style={{ alignSelf: "flex-end", fontSize: 11, color: "rgba(255,255,255,0.35)", paddingBottom: 6 }}>
                {dateFrom && dateTo ? dateFrom + " → " + dateTo : dateFrom ? "From " + dateFrom : "Until " + dateTo}
              </span>
            )}
          </div>
        )}
      </div>

      {/* ── UPLOAD ZONE (no data state) ── */}
      {!hasData && <UploadZone onData={handleFileData} onDemo={loadDemo} />}

      {/* ── LOOKER STUDIO VIEW ── */}
      {hasData && activeView === "creators" && (
          <div>
            <div style={{ marginBottom: 20 }}>
              <div style={{ display: "flex", gap: 12, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: "#fff" }}>Creator Performance</div>
                <span style={{ fontSize: 12, color: "rgba(255,255,255,0.35)" }}>{creatorTotals.length} creators</span>
                <button onClick={() => exportCSV(creatorTotals.flatMap(c => (creatorProductFilter !== "All Products" ? (metaAds || []).filter(a => a.product === creatorProductFilter) : (metaAds || [])).filter(a => extractCreator(a.name) === c.creator)), "creator-export.csv")} style={{ marginLeft: "auto", padding: "6px 14px", borderRadius: 7, fontSize: 11, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(92,164,247,0.3)", background: "rgba(92,164,247,0.08)", color: "#5ca4f7", fontFamily: "inherit" }}>↓ Export CSV</button>
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {["All Products", ...Array.from(new Set((metaAds || []).map(a => a.product))).filter(Boolean).sort()].map(p => (
                  <button key={p} onClick={() => { setCreatorProductFilter(p); setCreatorFilter("All"); }} style={{ padding: "4px 12px", borderRadius: 6, fontSize: 11, fontWeight: 600, cursor: "pointer", border: `1px solid ${creatorProductFilter === p ? (PRODUCT_COLORS[p] || "#c8a2f8") : "rgba(255,255,255,0.08)"}`, background: creatorProductFilter === p ? `${PRODUCT_COLORS[p] || "#c8a2f8"}18` : "transparent", color: creatorProductFilter === p ? (PRODUCT_COLORS[p] || "#c8a2f8") : "rgba(255,255,255,0.4)", fontFamily: "inherit" }}>{p}</button>
                ))}
              </div>
            </div>
            <div style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 14, padding: 20 }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 3px" }}>
                  <thead>
                    <tr>
                      {[["Creator","left"],["Ads","center"],["Spend","right"],["Revenue","right"],["ROAS","right"],["Orders","right"],["CPA","right"]].map(([h,a]) => (
                        <th key={h} style={{ padding: "10px 12px", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px", color: "rgba(255,255,255,0.35)", textAlign: a, borderBottom: "1px solid rgba(255,255,255,0.06)", whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {creatorTotals.map((c, i) => (
                      <tr key={c.creator} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.015)" : "transparent", cursor: "pointer" }} onClick={() => setCreatorFilter(c.creator === creatorFilter ? "All" : c.creator)}>
                        <td style={{ padding: "11px 12px", fontSize: 13, fontWeight: 600, color: creatorFilter === c.creator ? "#c8a2f8" : "#fff" }}>
                          {creatorFilter === c.creator && <span style={{ marginRight: 6 }}>▶</span>}{c.creator}
                        </td>
                        <td style={{ padding: "11px 12px", textAlign: "center", fontSize: 13, color: "rgba(255,255,255,0.5)" }}>{c.count}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmtFull(c.spend)}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmtFull(c.sales)}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right" }}><RoasBadge value={c.roas} /></td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontWeight: 600, color: "#fff" }}>{c.purchases}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{c.cpa > 0 ? fmtFull(c.cpa) : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            {creatorFilter !== "All" && (
              <div style={{ marginTop: 24 }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: "#fff", marginBottom: 12 }}>
                  Ads by <span style={{ color: "#c8a2f8" }}>{creatorFilter}</span>
                  <button onClick={() => setCreatorFilter("All")} style={{ marginLeft: 12, fontSize: 11, padding: "3px 10px", borderRadius: 5, border: "1px solid rgba(255,255,255,0.1)", background: "transparent", color: "rgba(255,255,255,0.4)", cursor: "pointer", fontFamily: "inherit" }}>✕ Clear</button>
                </div>
                <div style={{ background: "rgba(255,255,255,0.025)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 14, padding: 20 }}>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 2px" }}>
                      <thead><tr>
                        {[["Ad Name","left"],["Product","left"],["Spend","right"],["Revenue","right"],["ROAS","right"],["Orders","right"]].map(([h,a]) => (
                          <th key={h} style={{ padding: "10px 12px", fontSize: 11, fontWeight: 600, textTransform: "uppercase", color: "rgba(255,255,255,0.35)", textAlign: a, borderBottom: "1px solid rgba(255,255,255,0.06)" }}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {(creatorProductFilter !== "All Products" ? (metaAds || []).filter(a => a.product === creatorProductFilter) : filteredAds).filter(a => extractCreator(a.name) === creatorFilter).sort((a,b) => b.spend - a.spend).map((ad, i) => (
                          <tr key={i} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.015)" : "transparent" }}>
                            <td style={{ padding: "10px 12px", fontSize: 12, color: "#fff", maxWidth: 300, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{ad.name}</td>
                            <td style={{ padding: "10px 12px" }}><span style={{ fontSize: 11, fontWeight: 600, color: PRODUCT_COLORS[ad.product] || "#888", background: (PRODUCT_COLORS[ad.product] || "#888") + "15", padding: "2px 7px", borderRadius: 4 }}>{ad.product}</span></td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmt(ad.spend)}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmt(ad.sales)}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right" }}><RoasBadge value={ad.roas} /></td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontWeight: 600, color: "#fff" }}>{ad.purchases}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
        {hasData && activeView === "looker" && (
        <LookerStudioView metaAds={metaAds} preBlended={isDemo ? demoBlended : null} preByProduct={isDemo ? demoByProduct : null} />
      )}

      {/* ── DASHBOARD ── */}
      {hasData && activeView === "dashboard" && (
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
              { label: "Meta ROAS", value: `${productData.roas.toFixed(2)}x`, sub: productData.roas >= 2.5 ? "✓ Healthy" : "⚠ Needs Work", highlight: true },
              { label: "Purchases", value: productData.purchases.toLocaleString("en-IN"), sub: cpa > 0 ? `CPA: ${fmtFull(cpa)}` : "—" },
              { label: "Google Ads", value: gd.spend > 0 ? `${gd.roas.toFixed(2)}x` : "—", sub: gd.spend > 0 ? `${fmt(gd.spend)} spend` : "No Google data" },
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
            <div style={{ display: "flex", gap: 12, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 10, color: "rgba(255,255,255,0.35)", fontWeight: 600, textTransform: "uppercase" }}>Creator</label>
                <select value={creatorFilter} onChange={e => setCreatorFilter(e.target.value)} style={{ padding: "6px 10px", borderRadius: 7, border: "1px solid rgba(255,255,255,0.1)", background: "#1a1a2e", color: "#e8e6f0", fontSize: 12, outline: "none", cursor: "pointer", maxWidth: 180 }}>
                  {uniqueCreators.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              {activeFiltersCount > 0 && (
                <button onClick={clearAllFilters} style={{ alignSelf: "flex-end", padding: "7px 14px", borderRadius: 7, fontSize: 11, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(248,113,113,0.3)", background: "rgba(248,113,113,0.08)", color: "#f87171", fontFamily: "inherit" }}>
                  ✕ Clear {activeFiltersCount} filter{activeFiltersCount > 1 ? "s" : ""}
                </button>
              )}
              <button onClick={() => exportCSV(filteredAds, (isDemo ? "demo" : "live") + "-ads-export.csv")} style={{ alignSelf: "flex-end", padding: "7px 14px", borderRadius: 7, fontSize: 11, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(92,164,247,0.3)", background: "rgba(92,164,247,0.08)", color: "#5ca4f7", fontFamily: "inherit" }}>
                ↓ Export CSV ({filteredAds.length})
              </button>
            </div>
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
              <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", color: "rgba(255,255,255,0.3)", fontSize: 14 }}>🔍</span>
              <input
                type="text" value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search by creator, ad name, or product…"
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
                      { key: "purchases", label: "Orders", aligf: "right", sortable: true },
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
            {isDemo ? "Demo data · " : `${fileName} · `}ROAS = Reported conversion value ÷ spend · Better Herbs Dashboard v2
          </div>
        </>
      )}
    </div>
  );
}
