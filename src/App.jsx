import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTS ─────────────────────────────────────────────────────────────
const PRODUCTS = ["All Products", "Lactify", "Brainify Powder", "Brainify Drops", "Mamafy", "Flow Joy Drops","Flowjoy"];

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
const DEMO_ADS = [{"product":"Brainify Drops","name":"Brainify drops -Maninder Kaur (Blessings) Video","spend":3204.26,"sales":10932.0,"purchases":17,"roas":3.41,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":22407,"outbound_clicks":403,"add_to_cart":56,"initiate_checkout":49},{"product":"Brainify Powder","name":"Brainify Powder - Divya Bajpai","spend":3157.5,"sales":14319.0,"purchases":14,"roas":4.53,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":23389,"outbound_clicks":585,"add_to_cart":56,"initiate_checkout":39},{"product":"Brainify Powder","name":"Brainify Powder - Static - 13 (Ingredients)","spend":1935.03,"sales":5409.0,"purchases":5,"roas":2.8,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":18255,"outbound_clicks":292,"add_to_cart":16,"initiate_checkout":14},{"product":"Mamafy","name":"mamafy - Founder's  Dispatch Video May MOF | 13/05/2026","spend":1730.55,"sales":4159.0,"purchases":5,"roas":2.4,"type":"Video","date":"2026-06-01","start_date":"2026-05-13","impressions":11385,"outbound_clicks":307,"add_to_cart":21,"initiate_checkout":9},{"product":"Brainify Powder","name":"Brainify Powder - Dr Vinod 2","spend":1665.76,"sales":4050.0,"purchases":5,"roas":2.43,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":11898,"outbound_clicks":357,"add_to_cart":22,"initiate_checkout":15},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Ankit Jha  Video","spend":1652.42,"sales":4433.0,"purchases":6,"roas":2.68,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":9231,"outbound_clicks":314,"add_to_cart":29,"initiate_checkout":15},{"product":"Lactify","name":"Lactify_Dr. Nayana","spend":1636.81,"sales":6156.0,"purchases":8,"roas":3.76,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":12125,"outbound_clicks":303,"add_to_cart":32,"initiate_checkout":22},{"product":"Brainify Powder","name":"Brainify Powder - Doctor Compilation Video - 2  | 07/05/2026","spend":1437.4,"sales":2855.0,"purchases":4,"roas":1.99,"type":"Video","date":"2026-06-01","start_date":"2026-05-07","impressions":9712,"outbound_clicks":223,"add_to_cart":15,"initiate_checkout":10},{"product":"Brainify Powder","name":"Brainify Powder - Summer Season","spend":1219.73,"sales":1857.0,"purchases":3,"roas":1.52,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":9837,"outbound_clicks":334,"add_to_cart":15,"initiate_checkout":8},{"product":"Brainify Drops","name":"Brainify Drops - Summer Season","spend":1187.35,"sales":1099.0,"purchases":1,"roas":0.93,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":10794,"outbound_clicks":216,"add_to_cart":4,"initiate_checkout":3},{"product":"Flowjoy","name":"flowjoy drops - Dr. Garima | 13/05/2026","spend":1086.77,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-05-13","impressions":8490,"outbound_clicks":153,"add_to_cart":25,"initiate_checkout":17},{"product":"Brainify Drops","name":"Brainify drops - Static 7","spend":1031.89,"sales":1967.0,"purchases":3,"roas":1.91,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":8973,"outbound_clicks":224,"add_to_cart":12,"initiate_checkout":7},{"product":"Lactify","name":"Lactify - Customer Review - Lactation","spend":974.49,"sales":669.0,"purchases":1,"roas":0.69,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":7554,"outbound_clicks":144,"add_to_cart":3,"initiate_checkout":2},{"product":"Brainify Drops","name":"Drops-Dr Maninder Kaur (Copycat Brand)","spend":963.84,"sales":1314.0,"purchases":2,"roas":1.36,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":6740,"outbound_clicks":222,"add_to_cart":10,"initiate_checkout":6},{"product":"Brainify Powder","name":"Brainify Powder - Brainify Powder Static - 12   | 07/05/2026","spend":953.44,"sales":4779.0,"purchases":6,"roas":5.01,"type":"Static","date":"2026-06-01","start_date":"2026-05-07","impressions":7063,"outbound_clicks":177,"add_to_cart":24,"initiate_checkout":14},{"product":"Mamafy","name":"Mamafy - Dr.Shaifali Ad code","spend":868.99,"sales":3203.0,"purchases":3,"roas":3.69,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":7690,"outbound_clicks":177,"add_to_cart":11,"initiate_checkout":7},{"product":"Mamafy","name":"mamafy - B - Roll Video| 26/05/2026","spend":850.47,"sales":4985.0,"purchases":6,"roas":5.86,"type":"Video","date":"2026-06-01","start_date":"2026-05-26","impressions":5865,"outbound_clicks":117,"add_to_cart":21,"initiate_checkout":14},{"product":"Lactify","name":"Lactify - Dr. Naaz | 13/05/2026","spend":837.77,"sales":755.0,"purchases":1,"roas":0.9,"type":"Static","date":"2026-06-01","start_date":"2026-05-13","impressions":7285,"outbound_clicks":182,"add_to_cart":4,"initiate_checkout":2},{"product":"Flowjoy","name":"Flowjoy Drop - Dispatch video","spend":822.95,"sales":0.0,"purchases":0,"roas":0.0,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":4899,"outbound_clicks":113,"add_to_cart":15,"initiate_checkout":10},{"product":"Lactify","name":"Lactify - Summer Season","spend":818.45,"sales":695.0,"purchases":1,"roas":0.85,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":7308,"outbound_clicks":161,"add_to_cart":4,"initiate_checkout":2},{"product":"Flowjoy","name":"Flowjoy - Summer Season","spend":803.38,"sales":569.0,"purchases":1,"roas":0.71,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":5951,"outbound_clicks":149,"add_to_cart":4,"initiate_checkout":3},{"product":"Brainify Powder","name":"Brainify Powder Divya Bajpai 2nd video  | 28/05/2026","spend":802.38,"sales":2588.0,"purchases":3,"roas":3.23,"type":"Video","date":"2026-06-01","start_date":"2026-05-28","impressions":5244,"outbound_clicks":147,"add_to_cart":13,"initiate_checkout":8},{"product":"Mamafy","name":"mamafy - Dr. Garima 2nd video 28/05/2026","spend":800.09,"sales":3565.0,"purchases":5,"roas":4.46,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":4546,"outbound_clicks":141,"add_to_cart":23,"initiate_checkout":9},{"product":"Brainify Drops","name":"Brainify drops - Maninder Kaur - April - 1 30/05/2026","spend":774.46,"sales":2507.0,"purchases":3,"roas":3.24,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":6004,"outbound_clicks":114,"add_to_cart":10,"initiate_checkout":7},{"product":"Brainify Powder","name":"Brainify Powder Dr. Vinod (Brainify Powder)  | 28/05/2026","spend":773.51,"sales":1837.0,"purchases":3,"roas":2.37,"type":"Static","date":"2026-06-01","start_date":"2026-05-28","impressions":6906,"outbound_clicks":152,"add_to_cart":11,"initiate_checkout":8},{"product":"Brainify Powder","name":"Brainify Powder -Dr. Ankit Jha | 04/05/2026","spend":764.38,"sales":1288.0,"purchases":2,"roas":1.69,"type":"Static","date":"2026-06-01","start_date":"2026-05-04","impressions":5421,"outbound_clicks":168,"add_to_cart":9,"initiate_checkout":4},{"product":"Mamafy","name":"mamafy - Doctor Compilation video 1 - Mamafy- April | 23/05/2026","spend":760.87,"sales":1510.0,"purchases":2,"roas":1.98,"type":"Video","date":"2026-06-01","start_date":"2026-05-23","impressions":4726,"outbound_clicks":76,"add_to_cart":6,"initiate_checkout":5},{"product":"Lactify","name":"Lactify - Dr. Sunil (Kannad)  | 21/05/2026","spend":759.04,"sales":2265.0,"purchases":3,"roas":2.98,"type":"Static","date":"2026-06-01","start_date":"2026-05-21","impressions":5977,"outbound_clicks":102,"add_to_cart":10,"initiate_checkout":6},{"product":"Lactify","name":"Lactify - Dr. Pillai (Tamil) | 28/05/2026","spend":755.88,"sales":1510.0,"purchases":2,"roas":2.0,"type":"Static","date":"2026-06-01","start_date":"2026-05-28","impressions":6689,"outbound_clicks":154,"add_to_cart":8,"initiate_checkout":4},{"product":"Lactify","name":"Lactify - Dr. Manisha  | 21/05/2026","spend":745.82,"sales":1474.0,"purchases":1,"roas":1.98,"type":"Static","date":"2026-06-01","start_date":"2026-05-21","impressions":5737,"outbound_clicks":115,"add_to_cart":4,"initiate_checkout":2},{"product":"Mamafy","name":"Mamafy -  Static - 4 - Do not buy","spend":745.32,"sales":3917.0,"purchases":3,"roas":5.26,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":6715,"outbound_clicks":141,"add_to_cart":11,"initiate_checkout":7},{"product":"Brainify Drops","name":"Brainify drops - Ingredients Video 28/05/2026","spend":744.41,"sales":3675.0,"purchases":5,"roas":4.94,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":4353,"outbound_clicks":113,"add_to_cart":20,"initiate_checkout":12},{"product":"Lactify","name":"Lactify - Lactify Static - 29 28/05/2026","spend":733.72,"sales":1954.0,"purchases":2,"roas":2.66,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":5279,"outbound_clicks":153,"add_to_cart":9,"initiate_checkout":6},{"product":"Brainify Powder","name":"Brainify Powder Static - 22 | 19/05/2026","spend":733.1,"sales":2532.0,"purchases":4,"roas":3.45,"type":"Static","date":"2026-06-01","start_date":"2026-05-19","impressions":6665,"outbound_clicks":133,"add_to_cart":14,"initiate_checkout":8},{"product":"Lactify","name":"Lactify - Founder's Lactify - Dispatch Video May MOF 28/05/2026","spend":733.0,"sales":600.0,"purchases":1,"roas":0.82,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":4287,"outbound_clicks":111,"add_to_cart":4,"initiate_checkout":2},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Pillai (Tamil) 2nd Video | 26/05/2026","spend":726.6,"sales":669.0,"purchases":1,"roas":0.92,"type":"Video","date":"2026-06-01","start_date":"2026-05-26","impressions":4430,"outbound_clicks":84,"add_to_cart":3,"initiate_checkout":2},{"product":"Brainify Drops","name":"Brainify drops - Static - 18  28/05/2026","spend":726.15,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":5379,"outbound_clicks":134,"add_to_cart":16,"initiate_checkout":11},{"product":"Lactify","name":"Lactify - Dr. Gunjan - 2nd Video | 21/05/2026","spend":716.1,"sales":1199.0,"purchases":1,"roas":1.67,"type":"Video","date":"2026-06-01","start_date":"2026-05-21","impressions":4504,"outbound_clicks":153,"add_to_cart":5,"initiate_checkout":2},{"product":"Mamafy","name":"Mamafy - Summer Season","spend":708.74,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":4956,"outbound_clicks":164,"add_to_cart":15,"initiate_checkout":10},{"product":"Mamafy","name":"mamafy -  2nd founder video - Mamafy Powder 28/05/2026","spend":708.34,"sales":755.0,"purchases":1,"roas":1.07,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":4754,"outbound_clicks":114,"add_to_cart":4,"initiate_checkout":3},{"product":"Brainify Powder","name":"Brainify Powder Doctor Compilation Video - 1 (Brainify Powder)  | 28/05/2026","spend":707.07,"sales":669.0,"purchases":1,"roas":0.95,"type":"Video","date":"2026-06-01","start_date":"2026-05-28","impressions":4714,"outbound_clicks":118,"add_to_cart":4,"initiate_checkout":2},{"product":"Flowjoy","name":"flowjoy  - Henna Jain  - Ad code | 23/05/2026","spend":686.26,"sales":595.0,"purchases":1,"roas":0.87,"type":"Static","date":"2026-06-01","start_date":"2026-05-23","impressions":5320,"outbound_clicks":101,"add_to_cart":3,"initiate_checkout":2},{"product":"Brainify Powder","name":"Brainify Powder -  Powder Static - 21 | 26/05/2026","spend":681.02,"sales":1528.0,"purchases":2,"roas":2.24,"type":"Static","date":"2026-06-01","start_date":"2026-05-26","impressions":5771,"outbound_clicks":162,"add_to_cart":9,"initiate_checkout":5},{"product":"Brainify Powder","name":"Brainify Powder -Dr. Mallika (Telugu)   | 07/05/2026","spend":671.75,"sales":2007.0,"purchases":3,"roas":2.99,"type":"Static","date":"2026-06-01","start_date":"2026-05-07","impressions":5945,"outbound_clicks":137,"add_to_cart":11,"initiate_checkout":8},{"product":"Brainify Powder","name":"Brainify Powder April Dispatch Video  | 21/05/2026","spend":605.79,"sales":1638.0,"purchases":2,"roas":2.7,"type":"Video","date":"2026-06-01","start_date":"2026-05-21","impressions":3717,"outbound_clicks":67,"add_to_cart":7,"initiate_checkout":5},{"product":"Brainify Powder","name":"Brainify Powder Static - 23  | 28/05/2026","spend":588.86,"sales":1198.0,"purchases":1,"roas":2.03,"type":"Static","date":"2026-06-01","start_date":"2026-05-28","impressions":5353,"outbound_clicks":107,"add_to_cart":4,"initiate_checkout":2},{"product":"Brainify Drops","name":"Brainify - Dr Vinod 2","spend":553.68,"sales":2007.0,"purchases":3,"roas":3.62,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":4163,"outbound_clicks":96,"add_to_cart":11,"initiate_checkout":5},{"product":"Brainify Drops","name":"Brainify drops - Static 11","spend":537.8,"sales":1878.0,"purchases":2,"roas":3.49,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":4759,"outbound_clicks":109,"add_to_cart":8,"initiate_checkout":4},{"product":"Brainify Drops","name":"Brainify drops - Dr. Vinod Video","spend":483.77,"sales":669.0,"purchases":1,"roas":1.38,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":3141,"outbound_clicks":91,"add_to_cart":4,"initiate_checkout":3},{"product":"Flowjoy","name":"Flowjoy - Dr. Tanya - Ad code | 23/05/2026","spend":468.41,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-05-23","impressions":3470,"outbound_clicks":87,"add_to_cart":10,"initiate_checkout":7},{"product":"Mamafy","name":"Mamafy - Dr. Soniya Gupta","spend":460.06,"sales":3344.0,"purchases":4,"roas":7.27,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":3932,"outbound_clicks":106,"add_to_cart":17,"initiate_checkout":9},{"product":"Flowjoy","name":"Flowjoy - Aanchal Naherwa - Ad code | 23/05/2026","spend":398.41,"sales":1259.0,"purchases":1,"roas":3.16,"type":"Static","date":"2026-06-01","start_date":"2026-05-23","impressions":3239,"outbound_clicks":107,"add_to_cart":5,"initiate_checkout":3},{"product":"Brainify Drops","name":"Brainify drops - Doctor Compilation Video - March  18  28/05/2026","spend":187.88,"sales":1029.0,"purchases":1,"roas":5.48,"type":"Video","date":"2026-06-01","start_date":"2026-06-01","impressions":1139,"outbound_clicks":23,"add_to_cart":4,"initiate_checkout":2},{"product":"Brainify Drops","name":"Brainify drops - Static - 17  19/05/2026","spend":148.89,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":1095,"outbound_clicks":28,"add_to_cart":3,"initiate_checkout":2},{"product":"Brainify Drops","name":"Drops- Static 3","spend":143.92,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":1358,"outbound_clicks":22,"add_to_cart":4,"initiate_checkout":3},{"product":"Brainify Drops","name":"Brainify - USP Static","spend":129.72,"sales":1908.0,"purchases":2,"roas":14.71,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":920,"outbound_clicks":29,"add_to_cart":9,"initiate_checkout":4},{"product":"Brainify Drops","name":"Drops - Static 7","spend":117.66,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":823,"outbound_clicks":27,"add_to_cart":2,"initiate_checkout":2},{"product":"Brainify Drops","name":"Drops-Dr. Sandesh","spend":104.0,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-06-01","start_date":"2026-06-01","impressions":812,"outbound_clicks":15,"add_to_cart":2,"initiate_checkout":2}];
const DEMO_GOOGLE = null;

const DEMO_DAILY_BLENDED = [{"date":"2026-06-01","spend":52110.83,"sales":131600.0,"purchases":161,"roas":2.53,"cpa":323.67}];
const DEMO_DAILY_BY_PRODUCT = [{"date":"2026-06-01","product":"Brainify Drops","spend":12146.3,"sales":32806.0,"purchases":43,"roas":2.7,"cpa":282.47},{"date":"2026-06-01","product":"Brainify Powder","spend":19099.86,"sales":53656.0,"purchases":65,"roas":2.81,"cpa":293.84},{"date":"2026-06-01","product":"Flowjoy","spend":4266.18,"sales":2423.0,"purchases":3,"roas":0.57,"cpa":1422.06},{"date":"2026-06-01","product":"Lactify","spend":8711.08,"sales":17277.0,"purchases":21,"roas":1.98,"cpa":414.81},{"date":"2026-06-01","product":"Mamafy","spend":7887.41,"sales":25438.0,"purchases":29,"roas":3.23,"cpa":271.98}];
// Ad-level daily data for accurate per-date-range ROAS computation
// Format: [adNameIndex, date, spend, sales, purchases]
const DEMO_AD_NAMES = ["Brainify drops - Static 7","Lactify_Dr. Nayana","Brainify drops -Maninder Kaur (Blessings) Video","Brainify drops - Dr. Vinod Video","Brainify Powder - Dr. Ankit Jha  Video","Brainify Powder - Static - 13 (Ingredients)","Brainify Powder - Divya Bajpai","Drops- Static 3","Brainify Powder - Dr Vinod 2","Brainify - Dr Vinod - Post ID","Brainify - Blessing Video","Brainify - Divya Bajpai","Flow Drop - Dr. Ankit","Mamafy - Mamafy Static - 6 - USP | 23/03/2026","Brainify - INA_Hindi - Post ID","Brainify - Maninder Kaur","Mamafy -  Static - 4 - Do not buy","Mamafy - Dr.Shaifali Ad code","Brainify drops - Dr.sajid","Brainify drops - Static 11","Brainify drops - Dr. Pillai (Tamil) (Partnership ad) - Tamil","Brainify drops - Static 12","Brainify drops - Dr. Neha","Flowjoy Drop - Dispatch video","Brainify drops - Dr. Sajid","Brainify Powder - Dr. Prachi Mahajan","Brainify Powder - Dr. Ankit","Brainify Powder - Dr. Riya","Brainify Powder - Dr. Suryakamal","Brainify Powder - Dr. Pushpendra 2nd video","Lactify - Dr. Priyanka Deswal","Lactify - Dr. Srimukhi (Telugu)","Lactify - 2nd Doctor Compilation","Mamafy -  Dr. Smriti - Edited","Mamafy -  Static - 13","Mamafy -  Static - 16","Mamafy -  Dispatch video - April","Brainify drops - Rajmani","Brainify drops - Rina Arthaba","Drops- Static 7","Brainify - USP Static","Brainify - Dr. Ankit Jha  Video","Brainify drops - Static - 7","Brainify drops - Static - 5","Brainify drops - Static - 3","Brainify drops - Static - 8","Flow drops - Dr. Sushma Mogri","Mothers Day - Mamafy | 1/04/20206","Mothers Day - Mamafy - Retargeting | 1/04/20206","Mothers Day - Lactify | 1/04/20206","Mothers Day - Lactify - Retargeting | 1/04/20206","Mothers Day - Drops | 1/04/20206","Mothers Day - Drops -  Retargeting | 1/04/20206","Mothers Day -Powder | 1/04/20206","Mothers Day -Powder - RETARGETING | 1/04/20206","Lactify - Dr. Rohit Bharadwaj","Lactify - Dr. Priyanka (Ped.)","Lactify - Customer Review - Lactation","Mamafy - Dr. Soniya Gupta","Drops-Dr Maninder Kaur (Copycat Brand)","Mamafy - Static 2","Flowjoy Drops Static 2","Brainify drops - Static 16","Brainify Powder Static - 11","Mamafy - Dr. Manisha Meena Gupta - Mamafy (Edited Version) | 22/03/2026","Mamafy - Doctor Compilation Video - March - 1 | 23/03/2026","Lactify -Dr. Mayur","Mamafy - Dispatch Video - March - 1","Flowjoy Drops Static 7","Lactify - Dispatch Video","Brainify - Maninder Kaur (Blessings)","Mamafy - Dr. Garima","Drops-Maninder Kaur (Blessings) Video","Brainify - Google Static","Brainify - Dr Vinod 2","Lactify - 2nd Dispatch Video","Mamafy - static 2","Mamafy - Static - 6 - USP","Mamafy - Doctor Compilation Video","Mamafy_Dr. Tanya Video","Drops-Static 11 Us Vs Them","Brainify - Prachi Mahajan","Brainify Powder - Dr. Padma","mamafy - Dr. Soniya Gupta 1st Edit | 04/05/2026","mamafy - Dr. Samra - Edit 4 (Nutrition) | 04/05/2026","mamafy - Dr. Samra - Edit 2 (Direct) | 04/05/2026","Brainify Powder -Dr. Sonal | 04/05/2026","Brainify Powder -Dr. Ankit Jha | 04/05/2026","Brainify Powder -Rajmani Patel  | 04/05/2026","Brainify Powder -Static 20  | 04/05/2026","Brainify Powder -Static 18  | 04/05/2026","Brainify Powder -Static 19  | 04/05/2026","Drops-Dr. Vinod Video","Brainify Powder - Dr. Prachi Mahajan | 07/05/2026","Brainify Powder - Dr. Pushpendra  | 07/05/2026","Brainify Powder - Dr. Divyani Bhagat   | 07/05/2026","Brainify Powder - Doctor Compilation Video - 2  | 07/05/2026","Brainify Powder - Dr. Smriti   | 07/05/2026","Brainify Powder - Brainify Powder Static - 12   | 07/05/2026","Brainify Powder - Brainify Powder Static - 14   | 07/05/2026","Brainify Powder - Brainify Powder Static - 15   | 07/05/2026","Brainify Powder - Brainify Powder Static - 16   | 07/05/2026","Brainify Powder - Brainify Powder Static - 17   | 07/05/2026","Brainify drops - Dr. Padma 07/05/2026","Brainify drops - Rajmani Patel - Edited | 07/05/2026","Brainify drops -Brainify Drops Static - 16 | 07/05/2026","Mothers Day -Powder | Static 2 |","Mothers Day - Drops | Static 2","Mothers Day - Lactify | Static 2","Mothers Day - Mamafy | Static 2","mamafy - Dr. Samra - Edit 2 (Direct)","mamafy - Founder's  Dispatch Video May MOF | 13/05/2026","mamafy - Dr. Priya Soni| 13/05/2026","mamafy - Static - 14| 13/05/2026","Brainify Drops -Maninder Kaur (Blessings) Video","flowjoy drops - Dr. Garima | 13/05/2026","Lactify - Dr. Garima | 13/05/2026","Lactify - Dr. Naaz | 13/05/2026","Lactify - Static - 27 | 13/05/2026","Mamafy - Drishti 3rd Video","Brainify - Static - 13 (Ingredients)","mamafy - Founder's  Dispatch Video May MOF","Mamafy - Doctor Compilation Video - March - 1","mamafy - Founders dispatch","Brainify Powder -Dr. Taran (Punjabi) | 19/05/2026","Brainify Powder Dr. Shiba (Odia)i) | 19/05/2026","Brainify Powder Static - 22 | 19/05/2026","Brainify Powder -Dr. Mallika (Telugu)   | 07/05/2026","Brainify drops - Maninder Kaur - March - Edited 19/05/2026","Brainify drops - Dr. Rajeswari (Telugu) 19/05/2026","Brainify drops - Static - 17  19/05/2026","Brainify drops - Dr. Padma Tamil 21/05/2026","Brainify Powder April Dispatch Video  | 21/05/2026","Lactify - Dr. Manisha  | 21/05/2026","Lactify - Dr. Sunil (Kannad)  | 21/05/2026","Lactify - Dr. Gunjan - 2nd Video | 21/05/2026","mamafy - Doctor Compilation video 1 - Mamafy- April | 23/05/2026","mamafy - Mamafy Static - 17 | 23/05/2026","mamafy - Pooja Shah | 23/05/2026","Flowjoy - Dr. Tanya - Ad code | 23/05/2026","Drops - Dr.sajid","Drops - Static 7","Drops - Rajmani","Drops - Static 8","Drops - Dr. Rajeswari (Telugu) 19/05/2026","Drops - Doctor compilation video 2 | 19/04/2026","Drops - Static 12","Drops - Dr. Padmavathi - kannada - 23/03/2026","Drops-Dr Sonal","Brainify drops -Neha","Flowjoy - Aanchal Naherwa - Ad code | 23/05/2026","Brainify drops - Maninder Kaur - April - 1  25/05/2026","mamafy - B - Roll Video| 26/05/2026","Brainify Powder Static 11","Brainify Powder -Static 18","Brainify Powder - Dr. Pillai (Tamil) 2nd Video | 26/05/2026","Brainify Powder -  Powder Static - 21 | 26/05/2026","flowjoy  - Henna Jain  - Ad code | 23/05/2026","Brainify Powder Static - 22","mamafy -  2nd founder video - Mamafy Powder 28/05/2026","mamafy - Dr. Garima 2nd video 28/05/2026","Brainify Powder Doctor Compilation Video - 1 (Brainify Powder)  | 28/05/2026","Brainify Powder Dr. Vinod (Brainify Powder)  | 28/05/2026","Brainify Powder Divya Bajpai 2nd video  | 28/05/2026","Brainify Powder Static - 23  | 28/05/2026","Brainify drops - Static - 18  28/05/2026","Brainify drops - Doctor Compilation Video - March  18  28/05/2026","Brainify drops - Ingredients Video 28/05/2026","Lactify - Dr. Pillai (Tamil) | 28/05/2026","Lactify - Lactify Static - 29 28/05/2026","Lactify - Founder's Lactify - Dispatch Video May MOF 28/05/2026","Drops-Dr Priyanka","Drops-Doctor Compilation","Drops-Maninder Kaur (Blessings)","Brainify drops - Maninder Kaur - April - 1 30/05/2026","Drops-Dr. Sandesh","Brainify Powder - Summer Season"];
const DEMO_ADS_DAILY = [[0,"2026-05-01",1057.2,5393.0,7],[1,"2026-05-01",493.56,3946.0,4],[2,"2026-05-01",2706.83,7919.0,9],[3,"2026-05-01",2181.84,6931.0,7],[4,"2026-05-01",1119.38,3705.0,4],[5,"2026-05-01",1154.17,7989.0,11],[6,"2026-05-01",1356.29,15024.0,14],[7,"2026-05-01",0.55,0.0,0],[8,"2026-05-01",1238.44,4134.0,6],[9,"2026-05-01",1064.57,3825.0,4],[10,"2026-05-01",13.49,0.0,0],[11,"2026-05-01",0.08,0.0,0],[12,"2026-05-01",834.62,849.0,1],[13,"2026-05-01",1001.65,649.0,1],[14,"2026-05-01",2.76,0.0,0],[15,"2026-05-01",0.12,0.0,0],[16,"2026-05-01",573.92,3306.0,4],[17,"2026-05-01",833.2,0.0,0],[18,"2026-05-01",2159.47,7948.0,11],[19,"2026-05-01",795.51,1748.0,2],[20,"2026-05-01",12.92,0.0,0],[21,"2026-05-01",2.92,0.0,0],[22,"2026-05-01",13.04,0.0,0],[23,"2026-05-01",795.84,0.0,0],[24,"2026-05-01",2290.4,4065.0,5],[25,"2026-05-01",976.5,5443.0,6],[26,"2026-05-01",905.14,1747.0,3],[27,"2026-05-01",720.43,2148.0,2],[28,"2026-05-01",658.78,549.0,1],[29,"2026-05-01",581.69,1249.0,1],[30,"2026-05-01",779.81,0.0,0],[31,"2026-05-01",683.4,649.0,1],[32,"2026-05-01",688.66,0.0,0],[33,"2026-05-01",1137.55,4404.0,4],[34,"2026-05-01",817.8,649.0,1],[35,"2026-05-01",620.75,0.0,0],[36,"2026-05-01",569.22,1408.0,2],[37,"2026-05-01",732.05,5218.0,6],[38,"2026-05-01",807.78,599.0,1],[42,"2026-05-01",556.0,0.0,0],[43,"2026-05-01",543.95,0.0,0],[44,"2026-05-01",515.59,1498.0,2],[45,"2026-05-01",659.27,1797.0,3],[46,"2026-05-01",798.31,2127.0,3],[47,"2026-05-01",394.28,1259.0,1],[48,"2026-05-01",220.21,0.0,0],[49,"2026-05-01",441.48,0.0,0],[50,"2026-05-01",282.65,0.0,0],[51,"2026-05-01",428.33,0.0,0],[52,"2026-05-01",364.99,949.0,1],[53,"2026-05-01",371.05,0.0,0],[54,"2026-05-01",408.11,0.0,0],[55,"2026-05-01",368.2,649.0,1],[56,"2026-05-01",336.81,0.0,0],[57,"2026-05-01",376.61,0.0,0],[60,"2026-05-01",1157.4,1358.0,2],[61,"2026-05-01",1093.51,3926.0,4],[0,"2026-05-02",816.02,4094.0,6],[1,"2026-05-02",420.33,4554.0,5],[2,"2026-05-02",2698.4,3924.0,5],[3,"2026-05-02",2030.03,8457.0,13],[4,"2026-05-02",941.66,1638.0,2],[5,"2026-05-02",1200.4,3884.0,6],[6,"2026-05-02",1362.1,6560.0,9],[7,"2026-05-02",25.12,0.0,0],[8,"2026-05-02",1191.81,3265.0,5],[9,"2026-05-02",495.87,1358.0,2],[11,"2026-05-02",0.91,0.0,0],[12,"2026-05-02",696.09,1824.0,2],[13,"2026-05-02",1162.16,3706.0,4],[16,"2026-05-02",652.84,649.0,1],[17,"2026-05-02",807.92,3535.0,5],[18,"2026-05-02",1987.25,3795.0,5],[19,"2026-05-02",861.74,1898.0,2],[20,"2026-05-02",200.79,599.0,1],[21,"2026-05-02",348.67,0.0,0],[22,"2026-05-02",1.59,0.0,0],[23,"2026-05-02",706.69,1174.0,2],[24,"2026-05-02",937.29,3744.0,5],[25,"2026-05-02",823.18,2808.0,2],[26,"2026-05-02",807.23,1608.0,2],[27,"2026-05-02",587.25,1338.0,2],[28,"2026-05-02",404.62,0.0,0],[29,"2026-05-02",540.16,1957.0,3],[30,"2026-05-02",452.94,0.0,0],[31,"2026-05-02",363.6,0.0,0],[32,"2026-05-02",552.39,0.0,0],[33,"2026-05-02",565.21,1897.0,3],[34,"2026-05-02",835.73,2067.0,3],[35,"2026-05-02",662.91,709.0,1],[36,"2026-05-02",680.84,1468.0,2],[37,"2026-05-02",512.13,2147.0,3],[38,"2026-05-02",740.17,2707.0,3],[72,"2026-05-02",28.76,0.0,0],[39,"2026-05-02",3.25,0.0,0],[40,"2026-05-02",188.18,1848.0,2],[41,"2026-05-02",11.85,1559.0,1],[73,"2026-05-02",107.54,1298.0,2],[74,"2026-05-02",16.27,649.0,1],[42,"2026-05-02",329.39,1309.0,1],[43,"2026-05-02",290.72,599.0,1],[44,"2026-05-02",497.48,539.0,1],[45,"2026-05-02",369.91,1098.0,1],[46,"2026-05-02",765.59,1774.0,2],[47,"2026-05-02",492.88,0.0,0],[48,"2026-05-02",216.33,0.0,0],[49,"2026-05-02",396.04,0.0,0],[50,"2026-05-02",142.78,0.0,0],[51,"2026-05-02",608.68,5503.0,5],[52,"2026-05-02",192.3,0.0,0],[53,"2026-05-02",604.79,1798.0,2],[54,"2026-05-02",149.18,0.0,0],[55,"2026-05-02",434.85,1298.0,2],[56,"2026-05-02",482.87,0.0,0],[57,"2026-05-02",426.07,0.0,0],[60,"2026-05-02",1128.73,3406.0,4],[61,"2026-05-02",798.13,3411.0,5],[76,"2026-05-03",296.61,0.0,0],[77,"2026-05-03",80.21,1259.0,1],[34,"2026-05-03",1676.35,2068.0,2],[16,"2026-05-03",1438.59,2067.0,3],[78,"2026-05-03",141.0,0.0,0],[79,"2026-05-03",128.0,0.0,0],[71,"2026-05-03",218.86,649.0,1],[0,"2026-05-03",948.9,4405.0,5],[1,"2026-05-03",1043.47,6581.0,8],[2,"2026-05-03",2875.44,10797.0,13],[3,"2026-05-03",2251.88,12318.2,15],[4,"2026-05-03",1109.82,4838.0,5],[5,"2026-05-03",2037.55,1747.0,2],[6,"2026-05-03",3312.65,14273.0,14],[8,"2026-05-03",1467.92,8968.0,12],[9,"2026-05-03",1537.77,1298.0,2],[10,"2026-05-03",19.01,1099.0,1],[12,"2026-05-03",814.59,1308.0,2],[13,"2026-05-03",29.68,649.0,1],[14,"2026-05-03",3.72,0.0,0],[15,"2026-05-03",7.24,0.0,0],[17,"2026-05-03",987.23,2627.0,3],[18,"2026-05-03",2272.86,3445.0,5],[19,"2026-05-03",514.12,549.0,1],[21,"2026-05-03",52.31,649.0,1],[22,"2026-05-03",22.68,0.0,0],[23,"2026-05-03",868.49,0.0,0],[24,"2026-05-03",2138.71,6640.0,8],[25,"2026-05-03",1566.02,2496.0,4],[26,"2026-05-03",46.88,0.0,0],[27,"2026-05-03",33.27,1009.0,1],[29,"2026-05-03",753.32,1847.0,2],[30,"2026-05-03",498.03,1958.0,2],[32,"2026-05-03",628.97,709.0,1],[33,"2026-05-03",833.16,709.0,1],[35,"2026-05-03",19.1,0.0,0],[36,"2026-05-03",33.87,0.0,0],[37,"2026-05-03",1165.2,3595.0,5],[38,"2026-05-03",49.44,549.0,1],[80,"2026-05-03",7.16,0.0,0],[40,"2026-05-03",102.3,0.0,0],[73,"2026-05-03",91.15,649.0,1],[74,"2026-05-03",35.38,0.0,0],[44,"2026-05-03",583.76,2957.0,2],[45,"2026-05-03",69.62,0.0,0],[46,"2026-05-03",725.86,1198.0,2],[47,"2026-05-03",273.4,2997.0,3],[48,"2026-05-03",254.8,2977.0,3],[49,"2026-05-03",316.09,0.0,0],[51,"2026-05-03",1277.36,3706.0,4],[52,"2026-05-03",599.5,0.0,0],[53,"2026-05-03",792.39,4086.0,4],[54,"2026-05-03",570.98,1998.0,2],[55,"2026-05-03",640.17,0.0,0],[56,"2026-05-03",519.62,0.0,0],[57,"2026-05-03",641.63,0.0,0],[81,"2026-05-03",1.39,0.0,0],[60,"2026-05-03",38.85,649.0,1],[61,"2026-05-03",970.48,2847.0,3],[76,"2026-05-04",117.22,0.0,0],[77,"2026-05-04",192.35,0.0,0],[34,"2026-05-04",1224.83,649.0,1],[16,"2026-05-04",909.24,2567.0,3],[78,"2026-05-04",1376.11,3167.0,3],[79,"2026-05-04",474.71,1249.0,1],[71,"2026-05-04",13.95,0.0,0],[83,"2026-05-04",1026.26,0.0,0],[84,"2026-05-04",970.3,709.0,1],[85,"2026-05-04",924.29,0.0,0],[0,"2026-05-04",985.96,5579.0,6],[1,"2026-05-04",1044.15,1198.0,1],[2,"2026-05-04",2989.29,7222.18,9],[3,"2026-05-04",2398.63,5443.0,7],[4,"2026-05-04",1143.58,6960.0,9],[5,"2026-05-04",1911.37,5893.0,7],[6,"2026-05-04",2917.89,4904.0,6],[7,"2026-05-04",2.32,0.0,0],[8,"2026-05-04",1439.07,1847.0,3],[9,"2026-05-04",1680.2,1957.0,3],[10,"2026-05-04",18.4,0.0,0],[12,"2026-05-04",852.51,625.0,1],[14,"2026-05-04",3.05,0.0,0],[17,"2026-05-04",930.41,5814.0,6],[18,"2026-05-04",2360.69,5352.0,7],[19,"2026-05-04",500.79,599.0,1],[20,"2026-05-04",242.47,0.0,0],[21,"2026-05-04",64.66,0.0,0],[22,"2026-05-04",12.66,0.0,0],[23,"2026-05-04",754.03,0.0,0],[24,"2026-05-04",1826.05,2257.0,3],[25,"2026-05-04",1459.87,2458.0,2],[29,"2026-05-04",694.15,1248.0,2],[30,"2026-05-04",190.7,0.0,0],[32,"2026-05-04",751.1,0.0,0],[33,"2026-05-04",838.12,0.0,0],[37,"2026-05-04",1093.74,1198.0,2],[72,"2026-05-04",0.46,0.0,0],[39,"2026-05-04",12.05,0.0,0],[40,"2026-05-04",32.6,709.0,1],[41,"2026-05-04",4.84,0.0,0],[73,"2026-05-04",26.84,0.0,0],[44,"2026-05-04",605.98,1298.0,2],[46,"2026-05-04",713.7,0.0,0],[47,"2026-05-04",468.63,1848.0,2],[48,"2026-05-04",338.32,1248.0,1],[49,"2026-05-04",361.46,2458.0,2],[51,"2026-05-04",1170.8,1548.0,2],[53,"2026-05-04",1227.88,4496.0,4],[54,"2026-05-04",376.31,0.0,0],[55,"2026-05-04",575.02,0.0,0],[56,"2026-05-04",220.98,7689.0,11],[57,"2026-05-04",600.85,1298.0,2],[81,"2026-05-04",4.52,0.0,0],[86,"2026-05-04",673.9,0.0,0],[87,"2026-05-04",709.78,2796.0,3],[88,"2026-05-04",758.89,1298.0,2],[89,"2026-05-04",718.44,1807.0,2],[90,"2026-05-04",728.42,2158.0,2],[91,"2026-05-04",722.24,1688.0,2],[61,"2026-05-04",1041.94,0.0,0],[76,"2026-05-05",10.17,0.0,0],[77,"2026-05-05",0.08,709.0,1],[34,"2026-05-05",1219.89,5323.0,7],[16,"2026-05-05",1095.76,2007.0,3],[78,"2026-05-05",900.07,2057.0,3],[79,"2026-05-05",171.27,709.0,1],[71,"2026-05-05",446.45,0.0,0],[83,"2026-05-05",505.73,649.0,1],[84,"2026-05-05",509.58,649.0,1],[85,"2026-05-05",506.95,709.0,1],[0,"2026-05-05",1024.36,3445.0,4],[1,"2026-05-05",1073.18,4844.0,6],[2,"2026-05-05",3206.34,11424.0,15],[3,"2026-05-05",2411.65,6451.0,9],[4,"2026-05-05",1148.19,3914.0,6],[5,"2026-05-05",1914.05,4336.0,4],[6,"2026-05-05",2871.51,11616.0,12],[7,"2026-05-05",1.45,0.0,0],[8,"2026-05-05",1466.03,7541.0,8],[9,"2026-05-05",1356.69,3125.0,4],[12,"2026-05-05",848.29,1009.0,1],[15,"2026-05-05",5.53,0.0,0],[17,"2026-05-05",948.85,4475.0,5],[18,"2026-05-05",2366.26,8738.0,10],[19,"2026-05-05",465.62,1199.0,1],[20,"2026-05-05",87.11,0.0,0],[21,"2026-05-05",103.18,0.0,0],[23,"2026-05-05",683.99,575.0,1],[24,"2026-05-05",1862.73,2895.0,5],[25,"2026-05-05",1377.57,10319.0,10],[29,"2026-05-05",699.9,1499.0,1],[32,"2026-05-05",709.54,0.0,0],[33,"2026-05-05",833.23,7180.0,10],[37,"2026-05-05",813.2,4204.0,6],[72,"2026-05-05",17.57,0.0,0],[92,"2026-05-05",16.66,0.0,0],[39,"2026-05-05",14.37,0.0,0],[80,"2026-05-05",0.72,0.0,0],[40,"2026-05-05",89.88,599.0,1],[41,"2026-05-05",117.33,0.0,0],[73,"2026-05-05",69.09,0.0,0],[59,"2026-05-05",6.6,0.0,0],[74,"2026-05-05",67.02,0.0,0],[44,"2026-05-05",570.68,1098.0,2],[46,"2026-05-05",660.66,2747.0,3],[47,"2026-05-05",521.0,1358.0,2],[48,"2026-05-05",344.56,1248.0,1],[49,"2026-05-05",322.95,649.0,1],[51,"2026-05-05",1296.49,5783.0,7],[53,"2026-05-05",1189.81,2686.0,4],[55,"2026-05-05",594.3,649.0,1],[57,"2026-05-05",620.12,0.0,0],[86,"2026-05-05",701.35,1098.0,1],[87,"2026-05-05",853.8,2616.0,4],[88,"2026-05-05",728.22,649.0,1],[89,"2026-05-05",727.43,1448.0,2],[90,"2026-05-05",731.5,1698.0,2],[91,"2026-05-05",740.31,3446.0,4],[61,"2026-05-05",1055.84,699.0,1],[76,"2026-05-06",16.45,0.0,0],[77,"2026-05-06",27.03,0.0,0],[34,"2026-05-06",1264.64,4014.0,5],[16,"2026-05-06",903.91,549.0,1],[78,"2026-05-06",1111.94,1408.0,2],[79,"2026-05-06",79.18,0.0,0],[83,"2026-05-06",652.76,0.0,0],[84,"2026-05-06",623.44,5093.0,5],[85,"2026-05-06",643.8,649.0,1],[0,"2026-05-06",1029.09,3944.0,6],[1,"2026-05-06",954.16,5145.0,5],[2,"2026-05-06",3110.41,13282.0,18],[3,"2026-05-06",2354.44,5282.0,8],[4,"2026-05-06",1144.7,5791.0,9],[5,"2026-05-06",1911.44,7102.4,7],[6,"2026-05-06",2903.12,4800.0,6],[7,"2026-05-06",2.55,0.0,0],[8,"2026-05-06",1378.86,11156.0,14],[9,"2026-05-06",1582.56,6771.0,8],[10,"2026-05-06",25.48,649.0,1],[12,"2026-05-06",782.33,649.0,1],[14,"2026-05-06",12.07,0.0,0],[15,"2026-05-06",6.51,0.0,0],[17,"2026-05-06",950.25,3994.0,6],[18,"2026-05-06",2417.67,5752.0,8],[19,"2026-05-06",476.02,1148.0,1],[20,"2026-05-06",200.22,0.0,0],[21,"2026-05-06",67.84,0.0,0],[22,"2026-05-06",88.24,549.0,1],[23,"2026-05-06",803.13,1664.0,2],[24,"2026-05-06",1880.95,5142.0,8],[25,"2026-05-06",1395.96,1148.0,1],[29,"2026-05-06",692.39,1248.0,2],[32,"2026-05-06",727.7,2656.0,4],[33,"2026-05-06",834.01,0.0,0],[37,"2026-05-06",928.83,1248.0,2],[40,"2026-05-06",20.97,0.0,0],[41,"2026-05-06",0.06,0.0,0],[73,"2026-05-06",11.36,0.0,0],[44,"2026-05-06",569.89,0.0,0],[46,"2026-05-06",666.53,0.0,0],[47,"2026-05-06",476.19,1699.0,1],[48,"2026-05-06",325.87,2557.0,3],[49,"2026-05-06",301.19,1848.0,2],[51,"2026-05-06",1264.6,2997.0,3],[53,"2026-05-06",1268.1,1847.0,2],[55,"2026-05-06",568.68,1144.0,2],[57,"2026-05-06",554.92,0.0,0],[86,"2026-05-06",736.4,3445.0,3],[87,"2026-05-06",761.45,949.0,1],[88,"2026-05-06",685.6,1259.0,1],[89,"2026-05-06",724.41,1797.0,3],[90,"2026-05-06",724.01,3905.0,4],[91,"2026-05-06",699.73,1099.0,1],[61,"2026-05-06",1049.73,1228.0,2],[34,"2026-05-07",1199.44,2716.0,3],[78,"2026-05-07",855.25,2876.0,4],[79,"2026-05-07",254.38,1149.0,1],[71,"2026-05-07",1255.17,6930.0,9],[83,"2026-05-07",325.72,0.0,0],[84,"2026-05-07",708.58,0.0,0],[85,"2026-05-07",722.04,2408.0,2],[0,"2026-05-07",992.7,2296.0,4],[1,"2026-05-07",941.84,4745.0,5],[2,"2026-05-07",3111.2,9764.0,14],[3,"2026-05-07",2388.37,13600.0,19],[4,"2026-05-07",1081.99,5153.0,7],[5,"2026-05-07",1945.09,2726.0,4],[6,"2026-05-07",2830.73,7459.0,8],[7,"2026-05-07",35.53,0.0,0],[8,"2026-05-07",1460.64,7761.0,9],[9,"2026-05-07",1518.02,7471.0,9],[12,"2026-05-07",771.01,1839.0,3],[16,"2026-05-07",844.33,2497.0,3],[17,"2026-05-07",954.44,3366.0,4],[18,"2026-05-07",2383.16,8037.0,10],[19,"2026-05-07",540.29,599.0,1],[20,"2026-05-07",136.17,0.0,0],[21,"2026-05-07",64.54,0.0,0],[22,"2026-05-07",0.6,0.0,0],[23,"2026-05-07",905.82,1797.0,2],[24,"2026-05-07",1870.12,3944.0,6],[25,"2026-05-07",1398.86,4435.0,5],[29,"2026-05-07",730.49,1897.0,3],[32,"2026-05-07",720.31,2548.0,2],[33,"2026-05-07",818.22,2057.0,3],[37,"2026-05-07",930.19,2296.0,3],[72,"2026-05-07",34.87,0.0,0],[39,"2026-05-07",16.61,0.0,0],[80,"2026-05-07",0.08,0.0,0],[40,"2026-05-07",14.73,0.0,0],[41,"2026-05-07",55.21,0.0,0],[73,"2026-05-07",13.28,0.0,0],[74,"2026-05-07",0.97,0.0,0],[44,"2026-05-07",627.74,5215.0,5],[46,"2026-05-07",757.17,599.0,1],[47,"2026-05-07",507.33,2007.0,3],[48,"2026-05-07",333.21,3007.0,3],[49,"2026-05-07",294.01,1848.0,2],[51,"2026-05-07",1213.46,8848.0,10],[53,"2026-05-07",1286.05,4495.0,5],[55,"2026-05-07",629.03,0.0,0],[57,"2026-05-07",613.82,939.0,1],[87,"2026-05-07",680.03,1358.0,2],[88,"2026-05-07",654.37,0.0,0],[89,"2026-05-07",694.29,709.0,1],[90,"2026-05-07",674.31,709.0,1],[91,"2026-05-07",663.17,2148.0,2],[93,"2026-05-07",185.37,0.0,0],[94,"2026-05-07",165.15,0.0,0],[95,"2026-05-07",165.12,0.0,0],[96,"2026-05-07",179.79,0.0,0],[97,"2026-05-07",159.51,0.0,0],[98,"2026-05-07",179.18,1358.0,2],[99,"2026-05-07",148.88,0.0,0],[100,"2026-05-07",153.51,0.0,0],[101,"2026-05-07",152.48,0.0,0],[102,"2026-05-07",153.27,0.0,0],[103,"2026-05-07",481.51,0.0,0],[104,"2026-05-07",439.92,0.0,0],[105,"2026-05-07",404.52,0.0,0],[61,"2026-05-07",1026.07,3644.0,3],[76,"2026-05-08",2.95,0.0,0],[77,"2026-05-08",23.64,0.0,0],[34,"2026-05-08",1196.29,4044.0,6],[16,"2026-05-08",836.44,2627.0,3],[78,"2026-05-08",649.26,1358.0,2],[79,"2026-05-08",106.72,0.0,0],[71,"2026-05-08",1453.42,4625.0,5],[84,"2026-05-08",704.95,1408.0,2],[85,"2026-05-08",291.73,3217.0,3],[0,"2026-05-08",1029.49,4903.0,6],[1,"2026-05-08",980.08,2497.0,3],[2,"2026-05-08",2953.08,12813.0,15],[3,"2026-05-08",2354.3,12242.0,15],[4,"2026-05-08",1053.96,8948.0,10],[5,"2026-05-08",1883.74,709.0,1],[6,"2026-05-08",2569.51,7550.0,8],[8,"2026-05-08",1387.78,1248.0,2],[9,"2026-05-08",1025.38,2667.0,3],[10,"2026-05-08",32.33,649.0,1],[12,"2026-05-08",765.93,0.0,0],[14,"2026-05-08",16.2,0.0,0],[15,"2026-05-08",10.6,0.0,0],[17,"2026-05-08",950.36,3425.0,5],[18,"2026-05-08",2278.83,4384.0,6],[19,"2026-05-08",519.86,549.0,1],[20,"2026-05-08",45.64,0.0,0],[21,"2026-05-08",190.09,1757.0,2],[22,"2026-05-08",23.76,1009.0,1],[23,"2026-05-08",794.85,1298.0,1],[24,"2026-05-08",1097.87,4634.0,6],[25,"2026-05-08",1321.64,2348.0,2],[29,"2026-05-08",722.81,4953.0,6],[32,"2026-05-08",712.22,4664.0,5],[33,"2026-05-08",829.19,1248.0,2],[37,"2026-05-08",941.51,1658.0,2],[39,"2026-05-08",1.27,0.0,0],[80,"2026-05-08",0.88,0.0,0],[40,"2026-05-08",131.69,0.0,0],[73,"2026-05-08",90.16,0.0,0],[59,"2026-05-08",3.13,0.0,0],[74,"2026-05-08",74.13,0.0,0],[44,"2026-05-08",635.71,1148.0,2],[46,"2026-05-08",714.5,625.0,1],[47,"2026-05-08",518.62,649.0,1],[48,"2026-05-08",289.93,649.0,1],[49,"2026-05-08",281.32,649.0,1],[51,"2026-05-08",1131.49,3596.0,3],[53,"2026-05-08",1168.78,1787.0,2],[55,"2026-05-08",665.96,0.0,0],[57,"2026-05-08",613.45,3277.0,3],[87,"2026-05-08",628.6,0.0,0],[88,"2026-05-08",671.48,649.0,1],[89,"2026-05-08",664.7,1897.0,3],[90,"2026-05-08",661.29,0.0,0],[91,"2026-05-08",690.1,1298.0,2],[93,"2026-05-08",700.85,0.0,0],[94,"2026-05-08",695.69,0.0,0],[95,"2026-05-08",729.26,0.0,0],[96,"2026-05-08",768.19,3506.0,4],[97,"2026-05-08",717.34,0.0,0],[98,"2026-05-08",592.26,2546.0,4],[99,"2026-05-08",628.39,1877.0,3],[100,"2026-05-08",585.69,0.0,0],[101,"2026-05-08",591.16,649.0,1],[102,"2026-05-08",598.58,1148.0,1],[103,"2026-05-08",621.07,549.0,1],[104,"2026-05-08",629.72,1309.0,1],[105,"2026-05-08",574.5,599.0,1],[61,"2026-05-08",972.25,2707.0,3],[76,"2026-05-09",20.88,649.0,1],[77,"2026-05-09",0.28,0.0,0],[34,"2026-05-09",1148.02,5973.0,7],[16,"2026-05-09",830.35,1248.0,1],[78,"2026-05-09",396.47,2856.0,4],[79,"2026-05-09",17.73,0.0,0],[71,"2026-05-09",1341.93,8149.0,11],[84,"2026-05-09",613.19,709.0,1],[0,"2026-05-09",971.87,2296.0,3],[1,"2026-05-09",935.44,2946.0,3],[2,"2026-05-09",2682.61,11932.0,16],[3,"2026-05-09",1914.65,1897.0,3],[4,"2026-05-09",1007.49,4095.0,5],[5,"2026-05-09",1671.89,6463.0,7],[6,"2026-05-09",2543.76,8288.0,12],[7,"2026-05-09",15.47,0.0,0],[8,"2026-05-09",1213.47,6161.0,7],[9,"2026-05-09",370.02,1208.0,1],[10,"2026-05-09",0.19,0.0,0],[12,"2026-05-09",747.91,1059.0,1],[15,"2026-05-09",9.96,0.0,0],[17,"2026-05-09",929.89,3226.0,4],[18,"2026-05-09",1994.51,1747.0,3],[19,"2026-05-09",455.87,1148.0,2],[20,"2026-05-09",20.93,0.0,0],[21,"2026-05-09",448.67,1998.0,2],[22,"2026-05-09",1.01,0.0,0],[23,"2026-05-09",756.21,7186.0,10],[24,"2026-05-09",1297.56,3854.0,5],[25,"2026-05-09",1239.62,1248.0,2],[29,"2026-05-09",587.61,1947.0,3],[32,"2026-05-09",617.85,649.0,1],[33,"2026-05-09",817.43,1338.0,2],[37,"2026-05-09",817.03,3344.0,6],[72,"2026-05-09",9.77,0.0,0],[39,"2026-05-09",9.93,0.0,0],[40,"2026-05-09",451.04,1258.0,1],[41,"2026-05-09",68.61,599.0,1],[73,"2026-05-09",325.33,0.0,0],[44,"2026-05-09",576.23,0.0,0],[46,"2026-05-09",628.2,2523.0,3],[47,"2026-05-09",386.49,649.0,1],[48,"2026-05-09",212.92,3116.0,3],[49,"2026-05-09",217.03,0.0,0],[51,"2026-05-09",1009.01,0.0,0],[53,"2026-05-09",1022.72,3546.0,4],[55,"2026-05-09",497.56,1249.0,1],[57,"2026-05-09",528.04,2617.0,3],[81,"2026-05-09",2.17,0.0,0],[87,"2026-05-09",521.72,1358.0,2],[88,"2026-05-09",644.17,1148.0,1],[89,"2026-05-09",611.46,1139.0,1],[90,"2026-05-09",616.39,2936.0,4],[91,"2026-05-09",616.95,1699.0,1],[93,"2026-05-09",838.6,0.0,0],[94,"2026-05-09",861.98,0.0,0],[95,"2026-05-09",836.57,0.0,0],[96,"2026-05-09",782.52,649.0,1],[97,"2026-05-09",845.09,2297.0,2],[98,"2026-05-09",690.83,2506.0,3],[99,"2026-05-09",684.98,0.0,0],[100,"2026-05-09",724.91,649.0,1],[101,"2026-05-09",715.09,0.0,0],[102,"2026-05-09",708.8,1957.0,3],[103,"2026-05-09",597.89,549.0,1],[104,"2026-05-09",601.74,899.0,1],[105,"2026-05-09",457.1,1548.0,2],[61,"2026-05-09",825.48,0.0,0],[77,"2026-05-10",14.64,0.0,0],[34,"2026-05-10",1242.95,4015.0,5],[78,"2026-05-10",264.66,759.0,1],[71,"2026-05-10",1942.28,1258.0,1],[84,"2026-05-10",719.49,649.0,1],[85,"2026-05-10",521.67,1358.0,1],[0,"2026-05-10",1078.1,2917.0,3],[1,"2026-05-10",1082.84,5402.0,5],[2,"2026-05-10",2972.98,8807.0,12],[3,"2026-05-10",2294.55,5940.0,10],[4,"2026-05-10",1269.65,3657.0,3],[5,"2026-05-10",2120.76,6162.0,8],[6,"2026-05-10",2906.2,10764.0,12],[7,"2026-05-10",0.18,0.0,0],[8,"2026-05-10",1637.37,3265.0,5],[9,"2026-05-10",228.64,649.0,1],[12,"2026-05-10",881.88,0.0,0],[15,"2026-05-10",251.64,0.0,0],[16,"2026-05-10",865.56,2716.0,3],[17,"2026-05-10",988.63,649.0,1],[18,"2026-05-10",2289.09,5402.0,8],[19,"2026-05-10",539.11,549.0,1],[21,"2026-05-10",183.28,0.0,0],[23,"2026-05-10",805.04,1150.0,2],[24,"2026-05-10",2089.71,2876.0,4],[25,"2026-05-10",1348.56,2287.0,3],[29,"2026-05-10",694.34,649.0,1],[32,"2026-05-10",735.77,0.0,0],[33,"2026-05-10",842.64,1359.0,1],[37,"2026-05-10",703.07,0.0,0],[72,"2026-05-10",6.86,0.0,0],[80,"2026-05-10",5.93,0.0,0],[40,"2026-05-10",208.0,1109.0,1],[41,"2026-05-10",3.02,0.0,0],[73,"2026-05-10",114.55,0.0,0],[59,"2026-05-10",0.94,0.0,0],[74,"2026-05-10",1077.64,4725.0,5],[44,"2026-05-10",585.57,0.0,0],[46,"2026-05-10",716.95,0.0,0],[47,"2026-05-10",270.71,0.0,0],[48,"2026-05-10",308.01,1858.0,2],[49,"2026-05-10",202.41,0.0,0],[51,"2026-05-10",584.61,2148.0,2],[53,"2026-05-10",654.53,1524.0,2],[55,"2026-05-10",643.72,0.0,0],[57,"2026-05-10",601.31,2496.0,3],[87,"2026-05-10",762.34,2716.0,4],[88,"2026-05-10",691.32,709.0,1],[89,"2026-05-10",676.25,0.0,0],[90,"2026-05-10",705.18,0.0,0],[91,"2026-05-10",689.81,0.0,0],[93,"2026-05-10",780.12,1358.0,2],[94,"2026-05-10",378.98,1489.0,1],[95,"2026-05-10",430.92,0.0,0],[96,"2026-05-10",762.24,5911.0,6],[97,"2026-05-10",669.09,0.0,0],[98,"2026-05-10",591.35,649.0,1],[99,"2026-05-10",558.82,0.0,0],[100,"2026-05-10",371.27,0.0,0],[101,"2026-05-10",563.59,0.0,0],[102,"2026-05-10",598.78,0.0,0],[103,"2026-05-10",724.55,0.0,0],[104,"2026-05-10",667.1,1608.0,2],[105,"2026-05-10",632.31,0.0,0],[106,"2026-05-10",608.25,1049.0,1],[107,"2026-05-10",579.26,0.0,0],[108,"2026-05-10",138.07,0.0,0],[109,"2026-05-10",199.69,0.0,0],[61,"2026-05-10",1138.71,3396.0,4],[76,"2026-05-11",32.76,0.0,0],[77,"2026-05-11",19.77,0.0,0],[34,"2026-05-11",1152.17,2517.0,3],[16,"2026-05-11",835.33,709.0,1],[78,"2026-05-11",754.44,709.0,1],[79,"2026-05-11",0.62,0.0,0],[71,"2026-05-11",1363.57,1398.0,2],[84,"2026-05-11",628.72,2497.0,3],[85,"2026-05-11",846.65,3216.0,4],[0,"2026-05-11",1047.37,0.0,0],[1,"2026-05-11",948.28,2456.0,3],[2,"2026-05-11",3007.7,7498.0,11],[3,"2026-05-11",2404.43,6501.0,8],[4,"2026-05-11",1295.74,10617.0,11],[5,"2026-05-11",2293.78,7140.0,9],[6,"2026-05-11",3129.41,4244.0,5],[7,"2026-05-11",0.77,0.0,0],[8,"2026-05-11",1642.09,4553.0,7],[9,"2026-05-11",540.63,599.0,1],[10,"2026-05-11",10.1,0.0,0],[12,"2026-05-11",879.81,639.0,1],[14,"2026-05-11",1.39,0.0,0],[15,"2026-05-11",130.28,0.0,0],[17,"2026-05-11",925.8,3211.0,4],[18,"2026-05-11",2401.62,7659.0,10],[19,"2026-05-11",171.26,0.0,0],[20,"2026-05-11",0.96,0.0,0],[21,"2026-05-11",210.07,1558.0,2],[22,"2026-05-11",995.47,1538.0,2],[23,"2026-05-11",850.63,1649.0,3],[24,"2026-05-11",1239.26,2446.0,4],[25,"2026-05-11",404.21,3573.0,3],[29,"2026-05-11",724.19,1598.0,2],[32,"2026-05-11",677.6,0.0,0],[33,"2026-05-11",808.44,639.0,1],[37,"2026-05-11",850.47,2147.0,3],[40,"2026-05-11",91.29,709.0,1],[73,"2026-05-11",170.46,0.0,0],[74,"2026-05-11",818.09,2397.0,3],[44,"2026-05-11",569.56,2257.0,3],[46,"2026-05-11",686.69,1200.0,2],[47,"2026-05-11",63.53,0.0,0],[48,"2026-05-11",323.69,649.0,1],[49,"2026-05-11",1.3,0.0,0],[51,"2026-05-11",171.11,2617.0,3],[53,"2026-05-11",37.86,1009.0,1],[55,"2026-05-11",594.06,0.0,0],[57,"2026-05-11",591.79,709.0,1],[87,"2026-05-11",682.94,2517.0,3],[88,"2026-05-11",229.7,0.0,0],[89,"2026-05-11",245.63,0.0,0],[90,"2026-05-11",190.49,649.0,1],[91,"2026-05-11",244.92,0.0,0],[93,"2026-05-11",299.68,639.0,1],[96,"2026-05-11",730.96,2907.0,2],[97,"2026-05-11",223.25,649.0,1],[98,"2026-05-11",598.35,0.0,0],[99,"2026-05-11",191.28,709.0,1],[101,"2026-05-11",207.24,1747.0,2],[102,"2026-05-11",194.16,1099.0,1],[103,"2026-05-11",220.05,0.0,0],[104,"2026-05-11",766.18,2557.0,3],[105,"2026-05-11",185.35,1009.0,1],[106,"2026-05-11",1332.03,6095.0,5],[107,"2026-05-11",1116.6,4305.0,5],[108,"2026-05-11",307.85,689.0,1],[109,"2026-05-11",427.63,1848.0,2],[61,"2026-05-11",1196.43,575.0,1],[78,"2026-05-12",559.52,2777.0,3],[79,"2026-05-12",4.3,0.0,0],[71,"2026-05-12",1548.04,4184.0,6],[85,"2026-05-12",866.46,709.0,1],[110,"2026-05-12",22.91,0.0,0],[0,"2026-05-12",911.28,1138.0,2],[1,"2026-05-12",1361.88,6272.0,7],[2,"2026-05-12",3230.72,13042.0,17],[3,"2026-05-12",2423.23,8366.0,11],[4,"2026-05-12",1460.29,4943.0,7],[5,"2026-05-12",1708.27,5552.0,7],[6,"2026-05-12",2855.3,7897.0,9],[7,"2026-05-12",12.33,0.0,0],[8,"2026-05-12",1468.1,3205.0,5],[9,"2026-05-12",917.43,709.0,1],[10,"2026-05-12",13.76,0.0,0],[12,"2026-05-12",676.37,959.0,1],[15,"2026-05-12",91.22,0.0,0],[16,"2026-05-12",842.57,2667.0,3],[17,"2026-05-12",966.62,4734.0,6],[18,"2026-05-12",2429.1,6540.0,10],[20,"2026-05-12",79.64,0.0,0],[21,"2026-05-12",242.62,0.0,0],[22,"2026-05-12",60.93,0.0,0],[23,"2026-05-12",926.71,599.0,1],[24,"2026-05-12",1746.71,2246.0,4],[29,"2026-05-12",1092.81,3255.0,5],[32,"2026-05-12",712.2,709.0,1],[34,"2026-05-12",2107.8,4214.0,6],[33,"2026-05-12",42.36,0.0,0],[37,"2026-05-12",822.62,2296.0,4],[72,"2026-05-12",33.3,0.0,0],[39,"2026-05-12",4.67,0.0,0],[40,"2026-05-12",68.05,0.0,0],[41,"2026-05-12",53.41,0.0,0],[73,"2026-05-12",77.78,649.0,1],[74,"2026-05-12",236.55,549.0,1],[44,"2026-05-12",651.3,0.0,0],[46,"2026-05-12",29.62,1198.0,2],[47,"2026-05-12",19.81,0.0,0],[48,"2026-05-12",329.52,2348.0,2],[51,"2026-05-12",230.59,2787.0,3],[53,"2026-05-12",81.81,1159.0,1],[55,"2026-05-12",26.27,0.0,0],[57,"2026-05-12",792.47,1418.0,2],[81,"2026-05-12",0.79,0.0,0],[87,"2026-05-12",714.65,1887.0,3],[96,"2026-05-12",1033.79,4413.0,6],[98,"2026-05-12",654.29,1248.0,2],[104,"2026-05-12",792.41,2757.0,3],[106,"2026-05-12",1251.95,2847.0,3],[107,"2026-05-12",1414.61,3406.0,4],[108,"2026-05-12",324.93,1908.0,2],[109,"2026-05-12",519.1,3267.0,3],[61,"2026-05-12",836.3,575.0,1],[11,"2026-05-12",648.56,2098.0,2],[76,"2026-05-13",0.67,0.0,0],[77,"2026-05-13",8.19,0.0,0],[34,"2026-05-13",1931.71,4624.0,4],[78,"2026-05-13",625.78,3425.0,5],[71,"2026-05-13",1385.23,4724.0,6],[85,"2026-05-13",660.25,0.0,0],[110,"2026-05-13",42.09,0.0,0],[111,"2026-05-13",884.01,2007.0,2],[112,"2026-05-13",851.83,0.0,0],[113,"2026-05-13",811.73,649.0,1],[0,"2026-05-13",1108.89,2907.0,3],[1,"2026-05-13",1350.27,1148.0,2],[2,"2026-05-13",3086.58,9498.0,11],[3,"2026-05-13",2214.49,2845.0,5],[4,"2026-05-13",1655.87,2257.0,3],[5,"2026-05-13",2539.35,6591.0,8],[6,"2026-05-13",3173.75,11947.0,15],[8,"2026-05-13",1621.71,3755.0,5],[9,"2026-05-13",1512.45,4303.0,6],[10,"2026-05-13",42.49,499.0,1],[12,"2026-05-13",883.67,1224.0,2],[14,"2026-05-13",2.81,0.0,0],[15,"2026-05-13",20.36,0.0,0],[16,"2026-05-13",826.53,0.0,0],[17,"2026-05-13",957.02,2067.0,3],[18,"2026-05-13",2279.56,5491.0,8],[21,"2026-05-13",160.8,0.0,0],[22,"2026-05-13",180.22,0.0,0],[23,"2026-05-13",873.13,1725.0,3],[24,"2026-05-13",1060.12,4344.0,5],[29,"2026-05-13",1075.05,599.0,1],[32,"2026-05-13",758.87,0.0,0],[37,"2026-05-13",602.76,2746.0,4],[80,"2026-05-13",13.75,0.0,0],[40,"2026-05-13",43.73,1258.0,1],[73,"2026-05-13",67.12,709.0,1],[59,"2026-05-13",14.64,0.0,0],[74,"2026-05-13",50.39,599.0,1],[44,"2026-05-13",561.67,949.0,1],[48,"2026-05-13",89.84,0.0,0],[51,"2026-05-13",146.62,549.0,1],[57,"2026-05-13",983.84,1968.0,2],[87,"2026-05-13",526.18,1238.0,2],[96,"2026-05-13",1477.61,6640.0,9],[98,"2026-05-13",713.75,5404.0,6],[104,"2026-05-13",799.37,0.0,0],[106,"2026-05-13",521.35,0.0,0],[107,"2026-05-13",603.85,2207.0,3],[108,"2026-05-13",175.67,0.0,0],[109,"2026-05-13",246.73,0.0,0],[114,"2026-05-13",28.39,1249.0,1],[92,"2026-05-13",747.7,2955.0,4],[115,"2026-05-13",1194.5,8767.0,13],[116,"2026-05-13",611.68,0.0,0],[117,"2026-05-13",799.04,2177.0,3],[118,"2026-05-13",648.56,1358.0,2],[61,"2026-05-13",1021.13,575.0,1],[11,"2026-05-13",592.54,0.0,0],[16,"2026-05-14",926.02,2997.0,3],[78,"2026-05-14",706.79,2127.0,3],[79,"2026-05-14",30.44,0.0,0],[71,"2026-05-14",1174.19,4025.0,5],[85,"2026-05-14",737.05,2716.0,4],[110,"2026-05-14",20.76,0.0,0],[111,"2026-05-14",730.76,5883.0,7],[112,"2026-05-14",660.42,649.0,1],[113,"2026-05-14",736.64,4624.0,6],[0,"2026-05-14",1044.43,1747.0,3],[1,"2026-05-14",1359.61,3256.0,4],[2,"2026-05-14",3146.5,7738.0,11],[3,"2026-05-14",2392.61,3344.0,6],[4,"2026-05-14",1790.18,5711.0,9],[5,"2026-05-14",1963.44,7200.0,10],[6,"2026-05-14",3001.26,9796.0,13],[7,"2026-05-14",83.74,549.0,1],[8,"2026-05-14",1820.01,9078.0,12],[9,"2026-05-14",1464.55,1897.0,3],[11,"2026-05-14",2.93,0.0,0],[12,"2026-05-14",818.09,1224.0,2],[14,"2026-05-14",58.35,0.0,0],[17,"2026-05-14",971.12,7081.0,9],[18,"2026-05-14",2412.23,4503.0,6],[20,"2026-05-14",3.76,0.0,0],[21,"2026-05-14",637.55,5082.0,7],[22,"2026-05-14",3.93,0.0,0],[23,"2026-05-14",805.99,2849.0,3],[24,"2026-05-14",1905.02,6061.0,8],[29,"2026-05-14",1042.41,3145.0,5],[32,"2026-05-14",688.91,759.0,1],[34,"2026-05-14",1451.44,2007.0,3],[37,"2026-05-14",96.72,599.0,1],[72,"2026-05-14",2.77,0.0,0],[39,"2026-05-14",7.69,0.0,0],[40,"2026-05-14",38.63,0.0,0],[41,"2026-05-14",1.5,0.0,0],[73,"2026-05-14",91.75,599.0,1],[74,"2026-05-14",14.42,0.0,0],[57,"2026-05-14",838.55,0.0,0],[96,"2026-05-14",2998.06,6211.0,8],[98,"2026-05-14",590.61,5312.0,5],[104,"2026-05-14",659.53,2946.0,3],[114,"2026-05-14",21.92,0.0,0],[92,"2026-05-14",169.11,1208.0,1],[115,"2026-05-14",1500.76,3622.0,6],[116,"2026-05-14",572.54,0.0,0],[117,"2026-05-14",557.92,0.0,0],[118,"2026-05-14",614.92,1348.0,2],[61,"2026-05-14",1022.4,759.0,1],[76,"2026-05-15",0.74,0.0,0],[77,"2026-05-15",0.2,0.0,0],[34,"2026-05-15",1383.37,2736.0,4],[78,"2026-05-15",734.86,2796.0,4],[71,"2026-05-15",1143.3,1398.0,2],[85,"2026-05-15",666.74,0.0,0],[111,"2026-05-15",625.23,0.0,0],[112,"2026-05-15",643.54,1238.0,2],[113,"2026-05-15",657.3,3681.0,4],[0,"2026-05-15",966.37,4674.0,5],[1,"2026-05-15",1295.06,7051.0,9],[2,"2026-05-15",2865.84,6311.0,8],[3,"2026-05-15",2330.62,7544.0,8],[4,"2026-05-15",1852.18,5732.0,8],[5,"2026-05-15",1875.64,7939.0,9],[6,"2026-05-15",2728.26,12745.0,15],[7,"2026-05-15",361.15,0.0,0],[8,"2026-05-15",1624.91,2496.0,4],[9,"2026-05-15",1461.63,2077.0,3],[12,"2026-05-15",810.49,0.0,0],[15,"2026-05-15",3.09,0.0,0],[16,"2026-05-15",728.8,4841.0,4],[17,"2026-05-15",921.2,1398.0,2],[18,"2026-05-15",845.97,549.0,1],[21,"2026-05-15",496.16,1668.0,2],[22,"2026-05-15",195.82,0.0,0],[23,"2026-05-15",689.29,0.0,0],[24,"2026-05-15",806.37,619.0,1],[29,"2026-05-15",1004.79,1907.0,3],[32,"2026-05-15",690.64,669.0,1],[37,"2026-05-15",1478.18,1847.0,3],[39,"2026-05-15",22.48,0.0,0],[80,"2026-05-15",36.69,0.0,0],[40,"2026-05-15",11.34,0.0,0],[73,"2026-05-15",25.61,0.0,0],[59,"2026-05-15",34.82,0.0,0],[74,"2026-05-15",31.48,0.0,0],[45,"2026-05-15",900.49,619.0,1],[57,"2026-05-15",752.16,1858.0,2],[96,"2026-05-15",1485.21,4783.0,7],[98,"2026-05-15",563.21,5268.0,6],[119,"2026-05-15",25.55,0.0,0],[58,"2026-05-15",93.76,0.0,0],[104,"2026-05-15",654.39,0.0,0],[114,"2026-05-15",2.11,0.0,0],[92,"2026-05-15",1402.63,0.0,0],[115,"2026-05-15",1404.06,2346.0,3],[116,"2026-05-15",588.71,0.0,0],[117,"2026-05-15",448.0,0.0,0],[118,"2026-05-15",593.85,1967.0,2],[120,"2026-05-15",0.69,0.0,0],[121,"2026-05-15",478.62,1348.0,2],[61,"2026-05-15",989.38,3682.0,3],[77,"2026-05-16",31.39,0.0,0],[34,"2026-05-16",974.85,3266.0,4],[16,"2026-05-16",749.48,709.0,1],[78,"2026-05-16",89.24,1438.0,2],[79,"2026-05-16",0.41,0.0,0],[71,"2026-05-16",1372.64,5199.0,6],[85,"2026-05-16",558.72,2736.0,4],[110,"2026-05-16",1.68,0.0,0],[111,"2026-05-16",490.88,4234.0,6],[112,"2026-05-16",579.37,0.0,0],[113,"2026-05-16",526.77,1872.0,1],[0,"2026-05-16",813.19,3026.0,4],[1,"2026-05-16",1025.18,5304.0,6],[2,"2026-05-16",2580.27,7110.0,9],[3,"2026-05-16",2010.93,5432.0,7],[4,"2026-05-16",1395.28,4484.0,5],[5,"2026-05-16",2459.95,5862.0,6],[6,"2026-05-16",2332.95,9842.0,12],[8,"2026-05-16",1401.34,4783.0,7],[9,"2026-05-16",275.89,1907.0,3],[10,"2026-05-16",57.85,0.0,0],[11,"2026-05-16",6.58,0.0,0],[12,"2026-05-16",613.63,3239.0,3],[14,"2026-05-16",48.58,0.0,0],[15,"2026-05-16",17.37,619.0,1],[17,"2026-05-16",669.13,669.0,1],[23,"2026-05-16",618.76,969.0,1],[29,"2026-05-16",780.52,619.0,1],[32,"2026-05-16",470.77,729.0,1],[37,"2026-05-16",1177.34,1538.0,2],[40,"2026-05-16",340.9,0.0,0],[73,"2026-05-16",321.92,0.0,0],[74,"2026-05-16",441.12,0.0,0],[45,"2026-05-16",1021.36,1807.0,3],[57,"2026-05-16",582.32,3337.0,3],[81,"2026-05-16",0.21,0.0,0],[96,"2026-05-16",721.17,5042.0,7],[98,"2026-05-16",459.87,1278.0,2],[122,"2026-05-16",27.98,0.0,0],[104,"2026-05-16",524.77,2446.0,3],[114,"2026-05-16",103.64,0.0,0],[92,"2026-05-16",1069.87,1049.0,1],[116,"2026-05-16",567.89,2636.0,4],[117,"2026-05-16",520.33,669.0,1],[118,"2026-05-16",482.66,669.0,1],[115,"2026-05-16",215.64,619.0,1],[72,"2026-05-16",132.39,619.0,1],[120,"2026-05-16",132.75,0.0,0],[41,"2026-05-16",116.3,0.0,0],[121,"2026-05-16",530.47,0.0,0],[61,"2026-05-16",747.69,1009.0,1],[60,"2026-05-16",60.08,1398.0,2],[16,"2026-05-17",645.96,1997.0,3],[78,"2026-05-17",7.92,0.0,0],[79,"2026-05-17",0.59,0.0,0],[71,"2026-05-17",666.12,1458.0,2],[85,"2026-05-17",705.25,0.0,0],[111,"2026-05-17",677.98,669.0,1],[112,"2026-05-17",699.25,669.0,1],[113,"2026-05-17",704.22,1398.0,2],[0,"2026-05-17",989.25,3116.0,4],[1,"2026-05-17",1327.93,6181.0,7],[2,"2026-05-17",2962.07,6542.0,7],[3,"2026-05-17",2337.2,6596.0,9],[4,"2026-05-17",1672.18,3195.0,5],[5,"2026-05-17",2699.01,5202.0,8],[6,"2026-05-17",2835.49,6842.0,8],[8,"2026-05-17",1579.03,3964.0,6],[9,"2026-05-17",1005.0,669.0,1],[10,"2026-05-17",0.77,0.0,0],[12,"2026-05-17",860.66,0.0,0],[17,"2026-05-17",945.08,729.0,1],[23,"2026-05-17",798.08,3716.0,5],[29,"2026-05-17",1001.17,1878.0,2],[34,"2026-05-17",1377.74,3631.0,4],[37,"2026-05-17",689.6,569.0,1],[40,"2026-05-17",153.92,0.0,0],[41,"2026-05-17",175.79,0.0,0],[73,"2026-05-17",167.34,1259.0,1],[74,"2026-05-17",79.01,0.0,0],[45,"2026-05-17",576.27,3589.0,2],[57,"2026-05-17",794.77,3126.0,4],[81,"2026-05-17",0.55,0.0,0],[96,"2026-05-17",844.2,1268.0,2],[98,"2026-05-17",598.41,1338.0,2],[119,"2026-05-17",53.66,0.0,0],[58,"2026-05-17",33.1,0.0,0],[104,"2026-05-17",666.33,1148.0,1],[114,"2026-05-17",2.93,0.0,0],[92,"2026-05-17",1511.76,1588.0,2],[115,"2026-05-17",781.48,3481.0,5],[116,"2026-05-17",637.4,669.0,1],[117,"2026-05-17",601.28,3132.0,3],[118,"2026-05-17",613.14,729.0,1],[72,"2026-05-17",4.7,0.0,0],[120,"2026-05-17",21.89,0.0,0],[121,"2026-05-17",358.15,1957.0,3],[61,"2026-05-17",1010.5,2037.4,2],[60,"2026-05-17",0.59,0.0,0],[77,"2026-05-18",3.05,0.0,0],[34,"2026-05-18",1032.82,729.0,1],[78,"2026-05-18",26.0,0.0,0],[71,"2026-05-18",564.7,1259.0,1],[85,"2026-05-18",712.1,2657.0,3],[110,"2026-05-18",0.25,0.0,0],[111,"2026-05-18",694.84,3465.0,5],[112,"2026-05-18",263.47,1378.0,2],[113,"2026-05-18",677.13,729.0,1],[0,"2026-05-18",1052.59,4259.0,5],[1,"2026-05-18",1321.12,1878.0,2],[2,"2026-05-18",2956.91,10333.0,11],[3,"2026-05-18",2317.53,7188.0,11],[4,"2026-05-18",1567.74,6293.0,7],[5,"2026-05-18",2083.01,6588.0,7],[6,"2026-05-18",2812.86,11052.0,13],[8,"2026-05-18",1637.01,4583.0,7],[9,"2026-05-18",0.01,0.0,0],[10,"2026-05-18",12.94,0.0,0],[12,"2026-05-18",877.05,2292.0,3],[16,"2026-05-18",728.94,729.0,1],[17,"2026-05-18",905.77,0.0,0],[23,"2026-05-18",895.11,0.0,0],[29,"2026-05-18",1002.31,1907.0,3],[37,"2026-05-18",696.65,569.0,1],[40,"2026-05-18",14.28,0.0,0],[73,"2026-05-18",11.38,0.0,0],[45,"2026-05-18",583.79,1238.0,2],[57,"2026-05-18",851.96,1638.0,2],[96,"2026-05-18",910.63,2007.0,3],[98,"2026-05-18",606.5,1099.0,1],[122,"2026-05-18",86.4,0.0,0],[58,"2026-05-18",53.52,0.0,0],[104,"2026-05-18",733.66,1598.0,2],[114,"2026-05-18",20.05,0.0,0],[92,"2026-05-18",2482.15,6350.0,9],[74,"2026-05-18",67.36,669.0,1],[116,"2026-05-18",608.6,0.0,0],[117,"2026-05-18",624.04,2187.0,3],[118,"2026-05-18",619.77,0.0,0],[115,"2026-05-18",794.22,5120.0,8],[72,"2026-05-18",75.01,0.0,0],[120,"2026-05-18",393.19,619.0,1],[121,"2026-05-18",240.62,669.0,1],[61,"2026-05-18",1086.46,619.0,1],[60,"2026-05-18",15.85,0.0,0],[77,"2026-05-19",0.4,0.0,0],[34,"2026-05-19",1.88,1768.0,2],[79,"2026-05-19",0.51,0.0,0],[71,"2026-05-19",59.2,0.0,0],[85,"2026-05-19",762.67,1458.0,2],[111,"2026-05-19",1056.66,2107.0,3],[112,"2026-05-19",481.43,2127.0,3],[113,"2026-05-19",1068.72,3326.0,4],[123,"2026-05-19",3.92,0.0,0],[0,"2026-05-19",961.13,2267.0,3],[1,"2026-05-19",1667.49,8519.0,10],[2,"2026-05-19",3074.8,11173.0,14],[3,"2026-05-19",2504.13,3095.0,5],[4,"2026-05-19",1774.26,4693.0,7],[5,"2026-05-19",3099.41,4294.0,6],[6,"2026-05-19",3304.88,16285.0,20],[8,"2026-05-19",1639.76,5803.0,7],[12,"2026-05-19",36.13,0.0,0],[16,"2026-05-19",694.35,2343.0,2],[17,"2026-05-19",893.74,3781.0,4],[23,"2026-05-19",882.39,595.0,1],[29,"2026-05-19",40.2,0.0,0],[37,"2026-05-19",15.33,0.0,0],[45,"2026-05-19",630.53,6132.0,6],[57,"2026-05-19",983.94,1817.0,2],[96,"2026-05-19",1501.3,1188.0,2],[98,"2026-05-19",1372.88,1338.0,2],[104,"2026-05-19",683.14,659.0,1],[92,"2026-05-19",394.3,0.0,0],[74,"2026-05-19",434.34,0.0,0],[116,"2026-05-19",9.73,0.0,0],[117,"2026-05-19",823.38,3585.0,5],[118,"2026-05-19",22.08,0.0,0],[115,"2026-05-19",1215.19,5155.0,6],[11,"2026-05-19",3.87,0.0,0],[120,"2026-05-19",256.14,0.0,0],[121,"2026-05-19",322.52,2343.0,2],[124,"2026-05-19",773.1,0.0,0],[125,"2026-05-19",860.17,0.0,0],[126,"2026-05-19",826.27,2257.0,2],[127,"2026-05-19",800.9,1288.0,2],[128,"2026-05-19",804.07,619.0,1],[129,"2026-05-19",867.13,569.0,1],[130,"2026-05-19",845.46,619.0,1],[61,"2026-05-19",32.22,0.0,0],[60,"2026-05-19",62.93,0.0,0],[77,"2026-05-20",0.11,0.0,0],[34,"2026-05-20",0.09,0.0,0],[71,"2026-05-20",350.37,2566.0,2],[85,"2026-05-20",690.33,1398.0,2],[111,"2026-05-20",932.51,2067.0,3],[113,"2026-05-20",979.25,3944.0,4],[0,"2026-05-20",1028.02,1029.0,1],[1,"2026-05-20",1637.71,4579.0,6],[2,"2026-05-20",3067.74,16408.0,20],[3,"2026-05-20",2390.36,5333.0,6],[4,"2026-05-20",1588.04,2356.0,4],[5,"2026-05-20",2895.85,10360.0,11],[6,"2026-05-20",3275.45,19702.0,21],[8,"2026-05-20",1722.97,1288.0,2],[16,"2026-05-20",764.7,2657.0,3],[17,"2026-05-20",817.4,2127.0,3],[23,"2026-05-20",845.02,2987.0,3],[29,"2026-05-20",94.69,0.0,0],[45,"2026-05-20",643.32,2497.0,3],[57,"2026-05-20",948.19,2098.0,2],[96,"2026-05-20",1620.5,1288.0,2],[98,"2026-05-20",1398.0,4425.0,5],[122,"2026-05-20",4.48,0.0,0],[58,"2026-05-20",151.66,1338.0,2],[104,"2026-05-20",704.75,0.0,0],[114,"2026-05-20",96.62,0.0,0],[92,"2026-05-20",10.85,0.0,0],[74,"2026-05-20",617.9,1957.0,3],[117,"2026-05-20",851.16,2637.0,3],[115,"2026-05-20",1279.59,1998.0,2],[11,"2026-05-20",0.24,0.0,0],[120,"2026-05-20",40.0,0.0,0],[121,"2026-05-20",223.83,0.0,0],[124,"2026-05-20",662.83,0.0,0],[125,"2026-05-20",680.75,569.0,1],[126,"2026-05-20",697.02,619.0,1],[127,"2026-05-20",482.61,599.0,1],[128,"2026-05-20",740.69,2376.0,4],[129,"2026-05-20",833.53,2496.0,3],[130,"2026-05-20",702.65,3101.0,3],[76,"2026-05-21",3.58,0.0,0],[77,"2026-05-21",0.97,0.0,0],[34,"2026-05-21",0.95,0.0,0],[16,"2026-05-21",582.1,2709.0,3],[79,"2026-05-21",2.03,0.0,0],[71,"2026-05-21",161.38,0.0,0],[85,"2026-05-21",691.14,3203.0,3],[110,"2026-05-21",0.63,0.0,0],[111,"2026-05-21",953.51,1450.0,2],[113,"2026-05-21",944.82,3838.0,4],[0,"2026-05-21",993.88,1767.0,2],[1,"2026-05-21",1565.55,2737.0,3],[2,"2026-05-21",2855.76,10883.0,16],[3,"2026-05-21",2119.27,5452.0,7],[4,"2026-05-21",2039.34,5242.0,8],[5,"2026-05-21",2548.62,6917.0,7],[6,"2026-05-21",2946.34,13038.0,15],[8,"2026-05-21",1595.26,5358.0,6],[17,"2026-05-21",833.41,1390.0,2],[23,"2026-05-21",721.5,2170.0,3],[45,"2026-05-21",600.72,0.0,0],[56,"2026-05-21",427.58,0.0,0],[57,"2026-05-21",867.0,3722.0,4],[96,"2026-05-21",2063.24,4394.0,6],[98,"2026-05-21",996.24,3904.0,6],[122,"2026-05-21",81.13,0.0,0],[119,"2026-05-21",2.49,0.0,0],[58,"2026-05-21",24.26,695.0,1],[104,"2026-05-21",749.31,1868.0,2],[114,"2026-05-21",379.2,1538.0,2],[117,"2026-05-21",776.46,2145.0,3],[115,"2026-05-21",1153.74,3592.0,6],[121,"2026-05-21",166.47,695.0,1],[124,"2026-05-21",597.5,0.0,0],[125,"2026-05-21",694.85,669.0,1],[126,"2026-05-21",668.72,3076.0,3],[127,"2026-05-21",538.37,0.0,0],[128,"2026-05-21",704.86,1727.0,2],[129,"2026-05-21",609.35,2267.0,3],[130,"2026-05-21",666.88,0.0,0],[131,"2026-05-21",722.96,569.0,1],[132,"2026-05-21",693.61,619.0,1],[133,"2026-05-21",742.3,3404.0,3],[134,"2026-05-21",810.73,0.0,0],[135,"2026-05-21",691.75,1259.0,1],[60,"2026-05-21",75.73,0.0,0],[16,"2026-05-22",588.86,3468.0,3],[78,"2026-05-22",3.94,0.0,0],[79,"2026-05-22",36.97,695.0,1],[71,"2026-05-22",333.13,0.0,0],[85,"2026-05-22",645.61,1510.0,2],[111,"2026-05-22",1275.24,4583.0,5],[113,"2026-05-22",1307.72,2045.0,2],[0,"2026-05-22",958.67,1188.0,2],[1,"2026-05-22",1536.71,5907.0,6],[2,"2026-05-22",3033.29,7290.0,9],[3,"2026-05-22",2269.3,5408.0,8],[4,"2026-05-22",1853.02,4453.0,6],[5,"2026-05-22",3645.07,9244.0,11],[6,"2026-05-22",3295.61,8466.0,10],[8,"2026-05-22",1580.43,6112.0,8],[17,"2026-05-22",750.05,1390.0,2],[23,"2026-05-22",738.06,545.0,1],[45,"2026-05-22",562.1,0.0,0],[56,"2026-05-22",1007.57,755.0,1],[57,"2026-05-22",825.17,4816.0,4],[96,"2026-05-22",1515.7,5338.0,7],[98,"2026-05-22",878.74,6733.0,7],[122,"2026-05-22",34.79,0.0,0],[119,"2026-05-22",101.34,0.0,0],[58,"2026-05-22",32.3,0.0,0],[104,"2026-05-22",8.28,0.0,0],[114,"2026-05-22",39.88,0.0,0],[29,"2026-05-22",228.96,0.0,0],[117,"2026-05-22",705.01,755.0,1],[115,"2026-05-22",1223.38,1214.0,2],[121,"2026-05-22",447.93,1199.0,1],[124,"2026-05-22",753.16,1638.0,2],[125,"2026-05-22",626.84,669.0,1],[126,"2026-05-22",658.07,4270.0,5],[127,"2026-05-22",593.89,2437.0,3],[128,"2026-05-22",627.32,0.0,0],[129,"2026-05-22",558.35,2836.0,4],[130,"2026-05-22",632.47,2627.0,3],[131,"2026-05-22",693.71,619.0,1],[132,"2026-05-22",695.89,2556.0,4],[133,"2026-05-22",602.07,1199.0,1],[134,"2026-05-22",666.45,0.0,0],[135,"2026-05-22",678.95,1510.0,2],[76,"2026-05-23",0.02,0.0,0],[77,"2026-05-23",0.47,0.0,0],[34,"2026-05-23",483.41,0.0,0],[78,"2026-05-23",71.61,0.0,0],[71,"2026-05-23",836.37,3775.0,5],[85,"2026-05-23",634.9,695.0,1],[110,"2026-05-23",0.75,0.0,0],[111,"2026-05-23",1231.48,645.0,1],[113,"2026-05-23",1198.95,695.0,1],[136,"2026-05-23",664.76,1450.0,2],[137,"2026-05-23",665.84,1400.0,2],[138,"2026-05-23",657.23,1450.0,2],[0,"2026-05-23",935.09,2527.0,3],[1,"2026-05-23",1445.56,2193.0,3],[2,"2026-05-23",2795.22,9577.0,12],[3,"2026-05-23",2082.42,2157.0,3],[4,"2026-05-23",1533.35,7126.0,9],[5,"2026-05-23",2303.07,7596.0,8],[6,"2026-05-23",2981.42,8779.0,9],[8,"2026-05-23",1444.77,5102.0,7],[16,"2026-05-23",562.66,3193.0,3],[17,"2026-05-23",791.03,755.0,1],[23,"2026-05-23",644.05,1799.0,3],[25,"2026-05-23",2355.5,1957.0,3],[45,"2026-05-23",558.67,2092.0,2],[56,"2026-05-23",407.57,0.0,0],[57,"2026-05-23",804.85,0.0,0],[87,"2026-05-23",1185.49,2676.0,4],[96,"2026-05-23",1997.95,4684.0,6],[98,"2026-05-23",824.94,1338.0,2],[122,"2026-05-23",223.64,0.0,0],[58,"2026-05-23",127.04,0.0,0],[114,"2026-05-23",67.46,0.0,0],[29,"2026-05-23",0.04,0.0,0],[117,"2026-05-23",786.21,695.0,1],[115,"2026-05-23",1286.78,2378.0,4],[121,"2026-05-23",199.97,1259.0,1],[124,"2026-05-23",350.76,0.0,0],[125,"2026-05-23",282.48,0.0,0],[126,"2026-05-23",608.8,619.0,1],[127,"2026-05-23",559.3,1338.0,2],[128,"2026-05-23",579.84,1598.0,2],[129,"2026-05-23",582.38,1857.0,3],[130,"2026-05-23",589.07,2726.0,4],[131,"2026-05-23",620.18,1188.0,2],[132,"2026-05-23",657.66,1168.0,2],[133,"2026-05-23",668.79,695.0,1],[134,"2026-05-23",573.75,1450.0,2],[135,"2026-05-23",629.03,695.0,1],[139,"2026-05-23",681.79,1835.0,3],[140,"2026-05-23",327.8,0.0,0],[141,"2026-05-23",57.17,0.0,0],[142,"2026-05-23",43.07,0.0,0],[59,"2026-05-23",10.15,0.0,0],[143,"2026-05-23",65.44,0.0,0],[144,"2026-05-23",522.43,609.0,1],[145,"2026-05-23",66.49,0.0,0],[146,"2026-05-23",114.01,0.0,0],[62,"2026-05-23",60.39,0.0,0],[147,"2026-05-23",37.77,0.0,0],[148,"2026-05-23",6.16,0.0,0],[149,"2026-05-23",9.37,0.0,0],[150,"2026-05-23",322.11,0.0,0],[60,"2026-05-23",29.71,755.0,1],[76,"2026-05-24",0.89,0.0,0],[16,"2026-05-24",573.39,0.0,0],[79,"2026-05-24",47.68,0.0,0],[71,"2026-05-24",603.7,0.0,0],[85,"2026-05-24",734.69,3304.0,3],[111,"2026-05-24",1296.35,6164.0,6],[113,"2026-05-24",1346.84,0.0,0],[136,"2026-05-24",672.38,755.0,1],[137,"2026-05-24",725.25,695.0,1],[138,"2026-05-24",790.59,755.0,1],[0,"2026-05-24",1003.61,1637.0,3],[1,"2026-05-24",1587.84,3760.0,5],[2,"2026-05-24",2964.3,6940.0,10],[3,"2026-05-24",2224.27,5850.0,9],[4,"2026-05-24",1561.52,4014.0,6],[5,"2026-05-24",2019.47,6308.0,7],[6,"2026-05-24",3398.14,10636.0,11],[8,"2026-05-24",1609.7,7206.0,9],[17,"2026-05-24",898.72,3213.0,3],[23,"2026-05-24",846.27,2219.0,3],[25,"2026-05-24",1530.79,1238.0,2],[34,"2026-05-24",1028.16,755.0,1],[45,"2026-05-24",591.42,619.0,1],[57,"2026-05-24",904.28,3064.0,3],[87,"2026-05-24",754.88,3345.0,5],[96,"2026-05-24",1425.72,1857.0,3],[98,"2026-05-24",980.16,1578.0,2],[122,"2026-05-24",0.92,0.0,0],[58,"2026-05-24",2.5,0.0,0],[74,"2026-05-24",26.79,0.0,0],[29,"2026-05-24",1.11,0.0,0],[117,"2026-05-24",788.33,1510.0,2],[115,"2026-05-24",1496.94,1833.0,3],[121,"2026-05-24",260.84,1510.0,2],[126,"2026-05-24",680.06,1983.0,3],[127,"2026-05-24",615.34,1338.0,2],[128,"2026-05-24",687.38,1138.0,2],[129,"2026-05-24",628.64,1138.0,2],[130,"2026-05-24",683.68,1029.0,1],[131,"2026-05-24",621.07,1648.0,2],[132,"2026-05-24",684.28,2035.0,2],[133,"2026-05-24",699.05,1894.0,2],[134,"2026-05-24",725.45,1460.0,2],[135,"2026-05-24",742.19,2014.0,2],[139,"2026-05-24",709.84,595.0,1],[140,"2026-05-24",314.41,0.0,0],[141,"2026-05-24",91.78,0.0,0],[142,"2026-05-24",34.54,0.0,0],[59,"2026-05-24",2.96,0.0,0],[143,"2026-05-24",107.81,0.0,0],[144,"2026-05-24",100.09,619.0,1],[145,"2026-05-24",84.87,0.0,0],[146,"2026-05-24",76.49,1148.0,1],[62,"2026-05-24",181.89,549.0,1],[147,"2026-05-24",227.54,0.0,0],[148,"2026-05-24",8.49,0.0,0],[149,"2026-05-24",20.36,0.0,0],[150,"2026-05-24",630.37,4546.0,4],[120,"2026-05-24",145.18,0.0,0],[60,"2026-05-24",67.27,0.0,0],[77,"2026-05-25",0.09,0.0,0],[78,"2026-05-25",96.51,755.0,1],[71,"2026-05-25",800.26,755.0,1],[85,"2026-05-25",676.81,1390.0,2],[111,"2026-05-25",1225.15,4885.0,5],[113,"2026-05-25",339.3,685.0,1],[123,"2026-05-25",0.01,0.0,0],[136,"2026-05-25",749.92,1510.0,2],[137,"2026-05-25",734.62,0.0,0],[138,"2026-05-25",693.64,0.0,0],[0,"2026-05-25",860.78,3517.0,3],[1,"2026-05-25",1531.46,12383.0,12],[2,"2026-05-25",2856.59,13190.0,17],[3,"2026-05-25",2284.92,5542.0,8],[4,"2026-05-25",1557.07,3784.0,5],[5,"2026-05-25",1909.07,2343.0,3],[6,"2026-05-25",3073.9,12039.0,14],[8,"2026-05-25",1636.99,5592.0,8],[16,"2026-05-25",572.22,0.0,0],[17,"2026-05-25",842.07,2730.0,3],[18,"2026-05-25",2166.25,4513.0,7],[19,"2026-05-25",448.22,2152.0,2],[23,"2026-05-25",737.07,0.0,0],[25,"2026-05-25",1328.1,1328.0,2],[34,"2026-05-25",1085.66,2444.0,3],[45,"2026-05-25",531.52,569.0,1],[57,"2026-05-25",859.49,609.0,1],[87,"2026-05-25",678.7,3964.0,5],[96,"2026-05-25",1429.13,2307.0,3],[98,"2026-05-25",835.57,669.0,1],[122,"2026-05-25",47.5,0.0,0],[119,"2026-05-25",0.04,0.0,0],[58,"2026-05-25",554.42,1954.0,2],[114,"2026-05-25",32.8,569.0,1],[117,"2026-05-25",748.17,0.0,0],[115,"2026-05-25",1468.28,4806.0,8],[121,"2026-05-25",9.98,0.0,0],[126,"2026-05-25",664.28,1328.0,2],[127,"2026-05-25",649.85,669.0,1],[7,"2026-05-25",780.83,569.0,1],[59,"2026-05-25",758.12,1817.0,2],[21,"2026-05-25",603.67,0.0,0],[128,"2026-05-25",642.2,0.0,0],[129,"2026-05-25",689.43,2746.0,3],[130,"2026-05-25",628.26,0.0,0],[131,"2026-05-25",597.94,0.0,0],[132,"2026-05-25",647.62,1148.0,1],[133,"2026-05-25",682.0,1783.0,2],[134,"2026-05-25",694.49,695.0,1],[135,"2026-05-25",667.66,1510.0,2],[139,"2026-05-25",630.38,0.0,0],[140,"2026-05-25",883.91,549.0,1],[141,"2026-05-25",519.11,0.0,0],[142,"2026-05-25",121.72,0.0,0],[143,"2026-05-25",472.52,619.0,1],[144,"2026-05-25",148.0,0.0,0],[145,"2026-05-25",137.14,0.0,0],[146,"2026-05-25",577.73,569.0,1],[62,"2026-05-25",933.27,2157.0,3],[147,"2026-05-25",129.41,569.0,1],[148,"2026-05-25",0.15,0.0,0],[149,"2026-05-25",9.34,0.0,0],[150,"2026-05-25",580.28,0.0,0],[41,"2026-05-25",0.01,0.0,0],[74,"2026-05-25",56.71,0.0,0],[120,"2026-05-25",80.96,0.0,0],[72,"2026-05-25",102.24,0.0,0],[92,"2026-05-25",173.4,0.0,0],[11,"2026-05-25",7.44,0.0,0],[151,"2026-05-25",458.09,619.0,1],[76,"2026-05-26",1.34,0.0,0],[77,"2026-05-26",0.3,0.0,0],[16,"2026-05-26",604.02,5088.0,5],[78,"2026-05-26",80.42,755.0,1],[71,"2026-05-26",534.77,2769.0,3],[85,"2026-05-26",757.05,0.0,0],[110,"2026-05-26",4.21,0.0,0],[111,"2026-05-26",1508.63,2649.0,3],[123,"2026-05-26",0.5,0.0,0],[136,"2026-05-26",803.45,2179.0,3],[137,"2026-05-26",200.0,0.0,0],[138,"2026-05-26",666.7,0.0,0],[152,"2026-05-26",504.05,755.0,1],[0,"2026-05-26",1095.58,1748.0,2],[1,"2026-05-26",1652.82,7347.0,7],[2,"2026-05-26",3142.62,6419.0,9],[3,"2026-05-26",2387.08,5920.0,8],[4,"2026-05-26",1740.31,3775.0,5],[5,"2026-05-26",1961.43,3840.0,5],[6,"2026-05-26",3218.96,13416.0,15],[8,"2026-05-26",1674.75,3864.0,6],[17,"2026-05-26",992.64,4994.0,6],[18,"2026-05-26",3013.57,12029.0,14],[19,"2026-05-26",660.83,1288.0,2],[23,"2026-05-26",855.53,545.0,1],[25,"2026-05-26",988.41,0.0,0],[34,"2026-05-26",1217.12,2739.0,3],[45,"2026-05-26",662.24,1818.0,2],[57,"2026-05-26",1002.03,3213.0,3],[87,"2026-05-26",863.02,2576.0,4],[96,"2026-05-26",1509.08,5062.0,7],[98,"2026-05-26",872.88,3725.0,5],[122,"2026-05-26",32.78,0.0,0],[119,"2026-05-26",25.33,0.0,0],[58,"2026-05-26",582.65,755.0,1],[114,"2026-05-26",11.12,0.0,0],[74,"2026-05-26",110.11,619.0,1],[117,"2026-05-26",885.92,2769.0,3],[115,"2026-05-26",1568.99,5743.0,8],[126,"2026-05-26",778.69,3466.0,4],[127,"2026-05-26",569.73,1768.0,2],[7,"2026-05-26",1069.7,1757.0,3],[59,"2026-05-26",1194.58,5434.0,6],[21,"2026-05-26",920.8,0.0,0],[128,"2026-05-26",771.11,2557.0,3],[129,"2026-05-26",722.92,0.0,0],[130,"2026-05-26",787.92,2847.0,3],[131,"2026-05-26",721.93,0.0,0],[132,"2026-05-26",776.63,545.0,1],[133,"2026-05-26",664.95,695.0,1],[134,"2026-05-26",681.71,1510.0,2],[135,"2026-05-26",728.64,2205.0,3],[139,"2026-05-26",794.14,1190.0,2],[140,"2026-05-26",1492.54,1857.0,3],[141,"2026-05-26",250.99,569.0,1],[143,"2026-05-26",271.32,969.0,1],[144,"2026-05-26",7.34,0.0,0],[145,"2026-05-26",109.61,0.0,0],[146,"2026-05-26",141.74,0.0,0],[62,"2026-05-26",489.34,2107.0,3],[147,"2026-05-26",122.84,0.0,0],[148,"2026-05-26",0.7,0.0,0],[149,"2026-05-26",3.85,0.0,0],[150,"2026-05-26",761.9,595.0,1],[41,"2026-05-26",49.59,0.0,0],[120,"2026-05-26",69.39,0.0,0],[72,"2026-05-26",0.08,0.0,0],[92,"2026-05-26",3.48,0.0,0],[151,"2026-05-26",736.41,0.0,0],[153,"2026-05-26",38.15,0.0,0],[154,"2026-05-26",8.79,0.0,0],[44,"2026-05-26",2.95,0.0,0],[155,"2026-05-26",326.46,619.0,1],[156,"2026-05-26",304.12,0.0,0],[60,"2026-05-26",20.98,0.0,0],[78,"2026-05-27",199.52,695.0,1],[79,"2026-05-27",187.63,0.0,0],[71,"2026-05-27",1395.8,2145.0,3],[85,"2026-05-27",693.8,0.0,0],[110,"2026-05-27",0.03,0.0,0],[111,"2026-05-27",1606.56,1954.0,2],[113,"2026-05-27",12.68,0.0,0],[136,"2026-05-27",696.67,2245.0,3],[138,"2026-05-27",170.46,2629.0,3],[152,"2026-05-27",762.81,1199.0,1],[0,"2026-05-27",1094.75,1068.0,2],[1,"2026-05-27",1697.08,2973.0,3],[2,"2026-05-27",2982.28,5867.0,8],[3,"2026-05-27",2353.79,5004.0,6],[4,"2026-05-27",1532.33,3195.0,5],[5,"2026-05-27",2049.74,4230.0,6],[6,"2026-05-27",3480.5,15534.0,17],[8,"2026-05-27",1654.96,8378.0,11],[9,"2026-05-27",619.19,1957.0,3],[10,"2026-05-27",7.45,0.0,0],[11,"2026-05-27",19.79,0.0,0],[14,"2026-05-27",12.04,0.0,0],[16,"2026-05-27",707.32,2145.0,3],[17,"2026-05-27",946.23,2709.0,3],[18,"2026-05-27",2777.11,5397.0,6],[19,"2026-05-27",647.6,619.0,1],[23,"2026-05-27",824.82,1249.0,1],[34,"2026-05-27",1278.27,3228.0,3],[40,"2026-05-27",406.95,1409.0,1],[41,"2026-05-27",424.53,669.0,1],[73,"2026-05-27",346.27,755.0,1],[45,"2026-05-27",704.32,1029.0,1],[57,"2026-05-27",939.51,2628.0,4],[81,"2026-05-27",0.64,0.0,0],[87,"2026-05-27",928.56,1338.0,2],[96,"2026-05-27",1335.42,3605.0,5],[98,"2026-05-27",977.34,2357.0,2],[119,"2026-05-27",36.42,0.0,0],[58,"2026-05-27",586.78,2014.0,2],[29,"2026-05-27",110.97,0.0,0],[117,"2026-05-27",851.82,695.0,1],[115,"2026-05-27",1509.74,1574.0,2],[121,"2026-05-27",6.2,0.0,0],[126,"2026-05-27",772.28,2257.0,3],[127,"2026-05-27",593.85,1338.0,2],[7,"2026-05-27",576.45,1868.0,2],[59,"2026-05-27",1156.16,3066.0,4],[21,"2026-05-27",405.58,1409.0,1],[128,"2026-05-27",267.68,0.0,0],[129,"2026-05-27",689.27,619.0,1],[130,"2026-05-27",769.16,0.0,0],[131,"2026-05-27",331.09,569.0,1],[132,"2026-05-27",776.16,2163.0,3],[133,"2026-05-27",759.62,3258.0,4],[134,"2026-05-27",747.28,1350.0,1],[135,"2026-05-27",766.17,1450.0,2],[139,"2026-05-27",1032.01,595.0,1],[140,"2026-05-27",798.4,0.0,0],[141,"2026-05-27",73.49,0.0,0],[142,"2026-05-27",121.5,0.0,0],[143,"2026-05-27",46.9,0.0,0],[144,"2026-05-27",30.64,0.0,0],[145,"2026-05-27",10.3,0.0,0],[146,"2026-05-27",15.37,0.0,0],[62,"2026-05-27",358.27,0.0,0],[147,"2026-05-27",49.35,0.0,0],[150,"2026-05-27",764.88,0.0,0],[120,"2026-05-27",42.94,0.0,0],[92,"2026-05-27",275.68,0.0,0],[151,"2026-05-27",692.33,1029.0,1],[153,"2026-05-27",87.0,0.0,0],[154,"2026-05-27",0.25,0.0,0],[44,"2026-05-27",8.39,0.0,0],[157,"2026-05-27",1196.3,1290.0,2],[158,"2026-05-27",4.12,0.0,0],[155,"2026-05-27",1006.19,669.0,1],[156,"2026-05-27",834.08,669.0,1],[114,"2026-05-27",108.89,0.0,0],[148,"2026-05-27",0.82,0.0,0],[149,"2026-05-27",0.06,0.0,0],[27,"2026-05-27",14.08,0.0,0],[60,"2026-05-27",0.54,0.0,0],[76,"2026-05-28",0.54,0.0,0],[77,"2026-05-28",4.88,0.0,0],[34,"2026-05-28",1175.62,2398.0,2],[78,"2026-05-28",406.2,755.0,1],[79,"2026-05-28",5.29,0.0,0],[71,"2026-05-28",1981.78,4574.0,5],[85,"2026-05-28",525.42,755.0,1],[110,"2026-05-28",0.01,0.0,0],[111,"2026-05-28",1562.67,695.0,1],[123,"2026-05-28",0.17,0.0,0],[136,"2026-05-28",700.08,0.0,0],[152,"2026-05-28",633.23,3404.0,4],[159,"2026-05-28",187.21,0.0,0],[160,"2026-05-28",159.23,755.0,1],[0,"2026-05-28",983.63,1687.0,3],[1,"2026-05-28",1539.38,5644.0,8],[2,"2026-05-28",2952.35,4274.0,5],[3,"2026-05-28",2331.62,5571.0,9],[4,"2026-05-28",1446.48,1857.0,3],[5,"2026-05-28",1835.59,3555.0,3],[6,"2026-05-28",3012.52,9554.0,9],[8,"2026-05-28",1582.33,5152.0,7],[10,"2026-05-28",151.22,1473.0,1],[14,"2026-05-28",64.46,0.0,0],[15,"2026-05-28",117.31,0.0,0],[16,"2026-05-28",562.53,695.0,1],[17,"2026-05-28",988.62,7177.0,5],[18,"2026-05-28",2687.01,5232.0,7],[19,"2026-05-28",577.5,2407.0,2],[23,"2026-05-28",816.29,595.0,1],[40,"2026-05-28",391.8,619.0,1],[41,"2026-05-28",60.6,0.0,0],[73,"2026-05-28",319.36,619.0,1],[74,"2026-05-28",296.22,1288.0,2],[45,"2026-05-28",607.73,569.0,1],[57,"2026-05-28",886.49,695.0,1],[87,"2026-05-28",1750.7,3245.0,5],[96,"2026-05-28",1384.7,2911.0,3],[98,"2026-05-28",907.79,3615.0,5],[122,"2026-05-28",16.67,0.0,0],[58,"2026-05-28",533.61,735.0,1],[117,"2026-05-28",784.29,2920.0,3],[115,"2026-05-28",1463.67,1238.0,2],[126,"2026-05-28",653.29,2462.0,3],[127,"2026-05-28",640.52,1288.0,2],[7,"2026-05-28",226.38,0.0,0],[59,"2026-05-28",1036.94,1784.0,2],[129,"2026-05-28",712.82,0.0,0],[130,"2026-05-28",655.67,3506.0,4],[132,"2026-05-28",689.01,2278.0,2],[133,"2026-05-28",716.7,0.0,0],[134,"2026-05-28",698.05,695.0,1],[135,"2026-05-28",666.33,0.0,0],[139,"2026-05-28",581.93,0.0,0],[140,"2026-05-28",16.0,669.0,1],[143,"2026-05-28",178.06,549.0,1],[146,"2026-05-28",28.51,0.0,0],[62,"2026-05-28",145.8,1118.0,2],[150,"2026-05-28",741.76,0.0,0],[120,"2026-05-28",27.27,0.0,0],[92,"2026-05-28",265.73,1188.0,2],[151,"2026-05-28",621.24,0.0,0],[153,"2026-05-28",43.16,0.0,0],[44,"2026-05-28",8.19,0.0,0],[157,"2026-05-28",1181.0,1574.0,2],[155,"2026-05-28",608.17,1288.0,2],[156,"2026-05-28",771.47,669.0,1],[147,"2026-05-28",60.4,0.0,0],[145,"2026-05-28",31.45,0.0,0],[114,"2026-05-28",52.88,0.0,0],[141,"2026-05-28",245.25,0.0,0],[144,"2026-05-28",33.72,0.0,0],[142,"2026-05-28",65.48,0.0,0],[148,"2026-05-28",8.3,0.0,0],[149,"2026-05-28",2.69,0.0,0],[29,"2026-05-28",177.14,2373.0,2],[27,"2026-05-28",21.23,0.0,0],[161,"2026-05-28",257.88,0.0,0],[162,"2026-05-28",267.03,0.0,0],[163,"2026-05-28",270.56,0.0,0],[164,"2026-05-28",273.17,0.0,0],[165,"2026-05-28",249.5,0.0,0],[166,"2026-05-28",242.49,619.0,1],[167,"2026-05-28",304.01,569.0,1],[168,"2026-05-28",203.88,0.0,0],[169,"2026-05-28",248.29,0.0,0],[170,"2026-05-28",231.96,0.0,0],[60,"2026-05-28",7.4,0.0,0],[76,"2026-05-29",5.74,0.0,0],[16,"2026-05-29",620.51,1330.0,2],[78,"2026-05-29",0.08,0.0,0],[71,"2026-05-29",1084.52,4073.0,5],[110,"2026-05-29",2.72,0.0,0],[111,"2026-05-29",1423.07,3658.0,4],[113,"2026-05-29",1.04,0.0,0],[136,"2026-05-29",671.72,0.0,0],[152,"2026-05-29",749.86,1450.0,2],[159,"2026-05-29",752.06,1450.0,2],[160,"2026-05-29",747.59,2840.0,4],[0,"2026-05-29",964.44,2087.0,3],[1,"2026-05-29",1602.1,6025.0,6],[2,"2026-05-29",3016.05,14839.0,18],[3,"2026-05-29",2328.67,10467.0,13],[4,"2026-05-29",1435.93,7536.0,8],[5,"2026-05-29",1846.31,5125.0,5],[6,"2026-05-29",3012.58,12287.0,13],[8,"2026-05-29",1566.91,3914.0,6],[9,"2026-05-29",528.37,619.0,1],[10,"2026-05-29",174.09,0.0,0],[11,"2026-05-29",7.36,0.0,0],[14,"2026-05-29",20.33,0.0,0],[15,"2026-05-29",11.08,0.0,0],[17,"2026-05-29",975.34,2589.0,3],[18,"2026-05-29",2776.22,5389.0,7],[19,"2026-05-29",572.52,619.0,1],[23,"2026-05-29",758.14,545.0,1],[34,"2026-05-29",1046.69,735.0,1],[40,"2026-05-29",541.34,1338.0,2],[41,"2026-05-29",124.18,0.0,0],[73,"2026-05-29",222.46,0.0,0],[74,"2026-05-29",71.55,669.0,1],[45,"2026-05-29",218.9,619.0,1],[57,"2026-05-29",871.18,695.0,1],[81,"2026-05-29",1.84,0.0,0],[87,"2026-05-29",1442.59,4394.0,6],[96,"2026-05-29",1399.58,1947.0,3],[98,"2026-05-29",922.64,1818.0,2],[122,"2026-05-29",0.1,0.0,0],[58,"2026-05-29",462.6,1450.0,2],[117,"2026-05-29",753.99,3203.0,3],[115,"2026-05-29",1507.79,4977.0,7],[121,"2026-05-29",1.79,0.0,0],[126,"2026-05-29",661.77,1917.0,2],[127,"2026-05-29",581.74,619.0,1],[7,"2026-05-29",222.61,0.0,0],[59,"2026-05-29",1065.19,4281.0,3],[129,"2026-05-29",723.49,1857.0,3],[130,"2026-05-29",686.34,569.0,1],[132,"2026-05-29",652.03,669.0,1],[133,"2026-05-29",683.36,1933.0,3],[134,"2026-05-29",676.99,755.0,1],[135,"2026-05-29",662.6,1510.0,2],[139,"2026-05-29",556.87,1214.0,2],[150,"2026-05-29",695.51,545.0,1],[120,"2026-05-29",62.53,0.0,0],[72,"2026-05-29",1.94,0.0,0],[151,"2026-05-29",645.87,1029.0,1],[153,"2026-05-29",57.04,0.0,0],[154,"2026-05-29",0.38,0.0,0],[157,"2026-05-29",1185.18,2853.0,5],[155,"2026-05-29",577.61,2007.0,3],[156,"2026-05-29",599.75,0.0,0],[147,"2026-05-29",20.51,0.0,0],[92,"2026-05-29",128.93,1029.0,1],[143,"2026-05-29",223.5,569.0,1],[114,"2026-05-29",34.69,0.0,0],[141,"2026-05-29",306.4,969.0,1],[144,"2026-05-29",31.75,0.0,0],[140,"2026-05-29",49.1,0.0,0],[146,"2026-05-29",40.52,0.0,0],[62,"2026-05-29",235.64,1148.0,1],[142,"2026-05-29",28.6,619.0,1],[148,"2026-05-29",2.08,0.0,0],[149,"2026-05-29",2.53,0.0,0],[29,"2026-05-29",89.61,0.0,0],[27,"2026-05-29",5.58,0.0,0],[161,"2026-05-29",769.15,1188.0,2],[162,"2026-05-29",716.51,1168.0,2],[163,"2026-05-29",691.86,1338.0,2],[164,"2026-05-29",707.21,1424.0,1],[165,"2026-05-29",701.73,1637.0,3],[166,"2026-05-29",652.75,0.0,0],[167,"2026-05-29",754.86,1807.0,3],[168,"2026-05-29",842.14,0.0,0],[169,"2026-05-29",738.36,1894.0,2],[170,"2026-05-29",707.1,2205.0,2],[60,"2026-05-29",0.01,0.0,0],[77,"2026-05-30",0.41,0.0,0],[34,"2026-05-30",981.27,1350.0,1],[78,"2026-05-30",238.14,1259.0,1],[71,"2026-05-30",1662.36,4149.0,5],[111,"2026-05-30",1510.81,0.0,0],[123,"2026-05-30",0.08,0.0,0],[136,"2026-05-30",579.72,669.0,1],[152,"2026-05-30",708.26,4865.0,6],[159,"2026-05-30",571.18,695.0,1],[160,"2026-05-30",602.92,4069.0,5],[0,"2026-05-30",1007.65,5862.0,6],[1,"2026-05-30",1537.43,2825.0,3],[2,"2026-05-30",3042.15,9058.0,12],[3,"2026-05-30",2181.34,5344.0,6],[4,"2026-05-30",1378.08,3156.0,4],[5,"2026-05-30",1844.52,2506.0,4],[6,"2026-05-30",3104.57,11894.0,14],[8,"2026-05-30",1499.91,3201.0,5],[9,"2026-05-30",56.43,0.0,0],[10,"2026-05-30",2.88,0.0,0],[15,"2026-05-30",14.86,0.0,0],[16,"2026-05-30",787.71,695.0,1],[17,"2026-05-30",966.27,625.0,1],[18,"2026-05-30",2669.61,3514.0,6],[19,"2026-05-30",582.7,2336.0,3],[23,"2026-05-30",745.55,525.0,1],[40,"2026-05-30",737.19,619.0,1],[41,"2026-05-30",276.49,0.0,0],[73,"2026-05-30",255.59,0.0,0],[74,"2026-05-30",1057.46,2976.0,4],[45,"2026-05-30",0.33,0.0,0],[57,"2026-05-30",809.92,1304.0,2],[87,"2026-05-30",1106.95,1897.0,3],[96,"2026-05-30",1322.43,7572.0,8],[98,"2026-05-30",887.65,2826.0,4],[122,"2026-05-30",5.87,0.0,0],[58,"2026-05-30",504.72,695.0,1],[117,"2026-05-30",777.76,1510.0,2],[115,"2026-05-30",1479.87,3508.0,6],[121,"2026-05-30",9.42,0.0,0],[126,"2026-05-30",668.74,0.0,0],[127,"2026-05-30",540.36,619.0,1],[7,"2026-05-30",145.07,0.0,0],[59,"2026-05-30",1064.73,3312.0,3],[129,"2026-05-30",726.02,1238.0,2],[130,"2026-05-30",675.13,1249.0,1],[132,"2026-05-30",653.3,1768.0,2],[133,"2026-05-30",674.05,1390.0,2],[134,"2026-05-30",660.4,1510.0,2],[135,"2026-05-30",646.43,0.0,0],[139,"2026-05-30",580.27,2430.0,4],[62,"2026-05-30",562.1,2887.0,3],[150,"2026-05-30",703.4,0.0,0],[120,"2026-05-30",77.18,669.0,1],[72,"2026-05-30",15.1,0.0,0],[92,"2026-05-30",131.67,0.0,0],[151,"2026-05-30",289.86,0.0,0],[153,"2026-05-30",20.71,0.0,0],[157,"2026-05-30",1187.53,0.0,0],[155,"2026-05-30",489.12,0.0,0],[156,"2026-05-30",492.64,599.0,1],[143,"2026-05-30",168.47,0.0,0],[114,"2026-05-30",5.82,0.0,0],[141,"2026-05-30",307.38,549.0,1],[144,"2026-05-30",7.65,0.0,0],[146,"2026-05-30",44.11,1158.0,1],[142,"2026-05-30",13.17,0.0,0],[29,"2026-05-30",1.18,0.0,0],[27,"2026-05-30",0.2,0.0,0],[161,"2026-05-30",533.96,0.0,0],[162,"2026-05-30",581.19,3671.0,5],[163,"2026-05-30",600.52,1288.0,2],[164,"2026-05-30",582.36,0.0,0],[165,"2026-05-30",604.4,0.0,0],[166,"2026-05-30",657.11,1188.0,2],[167,"2026-05-30",491.85,1138.0,2],[168,"2026-05-30",504.29,0.0,0],[169,"2026-05-30",553.23,3264.0,4],[170,"2026-05-30",600.64,695.0,1],[171,"2026-05-30",13.87,0.0,0],[172,"2026-05-30",8.5,0.0,0],[173,"2026-05-30",5.79,0.0,0],[174,"2026-05-30",318.56,0.0,0],[175,"2026-05-30",307.33,1188.0,2],[76,"2026-05-31",0.72,0.0,0],[77,"2026-05-31",4.55,0.0,0],[34,"2026-05-31",442.44,755.0,1],[78,"2026-05-31",60.57,0.0,0],[79,"2026-05-31",3.14,0.0,0],[71,"2026-05-31",1140.96,2205.0,3],[110,"2026-05-31",43.39,0.0,0],[111,"2026-05-31",1697.19,3595.0,5],[123,"2026-05-31",0.12,0.0,0],[136,"2026-05-31",770.69,1374.0,2],[152,"2026-05-31",813.65,2075.0,3],[159,"2026-05-31",757.65,695.0,1],[160,"2026-05-31",761.14,695.0,1],[1,"2026-05-31",1707.07,3045.0,5],[4,"2026-05-31",1674.67,8504.0,9],[5,"2026-05-31",2035.59,2536.0,3],[6,"2026-05-31",3466.3,6659.0,9],[41,"2026-05-31",10.26,0.0,0],[9,"2026-05-31",608.66,0.0,0],[10,"2026-05-31",3.3,0.0,0],[11,"2026-05-31",5.37,0.0,0],[14,"2026-05-31",26.74,0.0,0],[15,"2026-05-31",84.27,0.0,0],[122,"2026-05-31",15.72,0.0,0],[18,"2026-05-31",1261.5,1717.0,2],[57,"2026-05-31",969.21,959.0,1],[81,"2026-05-31",0.08,0.0,0],[117,"2026-05-31",883.23,1460.0,2],[115,"2026-05-31",1467.0,1164.0,2],[129,"2026-05-31",420.43,569.0,1],[133,"2026-05-31",773.68,619.0,1],[134,"2026-05-31",765.04,1894.0,2],[135,"2026-05-31",769.55,0.0,0],[96,"2026-05-31",1557.73,7042.0,8],[87,"2026-05-31",875.08,2676.0,4],[154,"2026-05-31",0.36,0.0,0],[147,"2026-05-31",6.6,0.0,0],[92,"2026-05-31",48.34,0.0,0],[59,"2026-05-31",952.02,1278.0,2],[145,"2026-05-31",0.06,0.0,0],[114,"2026-05-31",5.39,0.0,0],[144,"2026-05-31",24.17,0.0,0],[140,"2026-05-31",34.3,0.0,0],[148,"2026-05-31",2.78,0.0,0],[149,"2026-05-31",1.13,0.0,0],[29,"2026-05-31",2.44,0.0,0],[168,"2026-05-31",788.63,2809.0,3],[169,"2026-05-31",740.57,695.0,1],[170,"2026-05-31",798.17,669.0,1],[0,"2026-05-31",1090.74,2956.0,4],[2,"2026-05-31",3161.45,9756.0,14],[3,"2026-05-31",2527.85,6569.0,10],[8,"2026-05-31",1781.93,4633.0,7],[40,"2026-05-31",473.93,0.0,0],[73,"2026-05-31",142.14,599.0,1],[74,"2026-05-31",228.76,0.0,0],[58,"2026-05-31",570.11,695.0,1],[16,"2026-05-31",777.38,4512.0,4],[17,"2026-05-31",930.08,5550.0,7],[19,"2026-05-31",546.26,1249.0,1],[23,"2026-05-31",926.1,525.0,1],[98,"2026-05-31",1017.32,1867.0,2],[126,"2026-05-31",770.24,1938.2,2],[127,"2026-05-31",687.51,2486.0,3],[130,"2026-05-31",779.46,569.0,1],[132,"2026-05-31",780.03,569.0,1],[139,"2026-05-31",850.51,595.0,1],[157,"2026-05-31",727.3,2243.0,3],[150,"2026-05-31",757.91,0.0,0],[120,"2026-05-31",105.6,0.0,0],[153,"2026-05-31",27.47,0.0,0],[7,"2026-05-31",221.67,0.0,0],[44,"2026-05-31",1.68,0.0,0],[155,"2026-05-31",826.8,1188.0,2],[156,"2026-05-31",787.17,619.0,1],[143,"2026-05-31",145.07,1239.0,1],[141,"2026-05-31",317.97,1249.0,1],[161,"2026-05-31",772.43,2356.0,2],[162,"2026-05-31",835.99,669.0,1],[163,"2026-05-31",778.13,1338.0,2],[164,"2026-05-31",730.08,0.0,0],[165,"2026-05-31",721.88,1518.0,2],[166,"2026-05-31",740.72,1288.0,2],[167,"2026-05-31",728.54,2207.0,3],[174,"2026-05-31",733.45,569.0,1],[60,"2026-05-31",51.28,0.0,0],[146,"2026-05-31",60.8,0.0,0],[62,"2026-05-31",435.7,2776.0,4],[171,"2026-05-31",10.25,0.0,0],[172,"2026-05-31",3.01,0.0,0],[173,"2026-05-31",5.84,0.0,0],[175,"2026-05-31",592.0,3126.0,4],[176,"2026-05-31",28.49,0.0,0]];
// ─── PARSING HELPERS ────────────────────────────────────────────────────────
function extractProduct(adName, campaignName) {
  for (const src of [(adName || "").toLowerCase(), (campaignName || "").toLowerCase()]) {
    if (src.includes("lactify")) return "Lactify";
    if (src.includes("brainify powder") || (src.includes("brainify") && src.includes("powder"))) return "Brainify Powder";
    if (src.includes("brainify drops") || (src.includes("brainify") && src.includes("drop"))) return "Brainify Drops";
    if (src.includes("mamafy")) return "Mamafy";
    // Mothers Day campaigns: identify product from name
    if (src.includes("mothers day") || src.includes("mother's day")) {
      if (src.includes("powder")) return "Brainify Powder";
      if (src.includes("drops") || src.includes("drop")) return "Brainify Drops";
      if (src.includes("lactify")) return "Lactify";
      return "Mamafy";
    }
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
  if (!adName) return 'Unknown';
  let name = adName.trim();

  // Handle underscore separator: "Lactify_Dr. Nayana" -> "Dr. Nayana"
  const underIdx = name.indexOf('_');
  if (underIdx > 0 && !name.slice(0, underIdx).includes(' ')) {
    name = name.slice(underIdx + 1).trim();
  }

  // Handle no-space dash prefix: "Drops-Dr. Vinod" -> "Dr. Vinod"
  if (/^[A-Za-z]+-[A-Z]/.test(name)) {
    name = name.slice(name.indexOf('-') + 1).trim();
  }

  // Strip product prefix using ' - ' or ' -'
  const dashIdx = name.indexOf(' - ');
  if (dashIdx > 0) {
    name = name.slice(dashIdx + 3).trim();
  } else {
    const dashIdx2 = name.indexOf(' -');
    if (dashIdx2 > 0) name = name.slice(dashIdx2 + 2).trim();
  }

  // Strip date suffix: "| 13/05/2026" or "| 1/04/20206"
  name = name.replace(/\s*\|\s*\d{1,2}\/\d{1,2}\/\d{4,6}.*$/, '').trim();

  // If another " - " remains (e.g. "Dr. Samra - Edit 2"), keep only first segment
  const nextDash = name.indexOf(' - ');
  if (nextDash > 0) name = name.slice(0, nextDash).trim();

  // Strip trailing noise tokens
  name = name
    .replace(/\s+Video$/i, '')
    .replace(/\s+\(Blessings\)/i, '')
    .replace(/\s+\(Copycat[^)]*\)/i, '')
    .replace(/\s+Ad\s*code\s*$/i, '')
    .replace(/\s+\(Direct\)\s*$/i, '')
    .replace(/\s+\(Nutrition\)\s*$/i, '')
    .trim();

  return name || adName.trim();
}

function computeCreatorTotals(ads) {
  const totals = {};
  for (const ad of ads) {
    const creator = extractCreator(ad.name);
    if (!totals[creator]) totals[creator] = { creator, spend: 0, sales: 0, purchases: 0, add_to_cart: 0, initiate_checkout: 0, impressions: 0, outbound_clicks: 0, count: 0, start_date: null };
    totals[creator].spend += ad.spend;
    totals[creator].sales += ad.sales;
    totals[creator].purchases += ad.purchases;
    totals[creator].add_to_cart += (ad.add_to_cart || 0);
    totals[creator].initiate_checkout += (ad.initiate_checkout || 0);
    totals[creator].impressions += (ad.impressions || 0);
    totals[creator].outbound_clicks += (ad.outbound_clicks || 0);
    totals[creator].count += 1;
    if (ad.start_date && (!totals[creator].start_date || ad.start_date < totals[creator].start_date)) totals[creator].start_date = ad.start_date;
  }
  return Object.values(totals)
    .map(d => ({
      ...d,
      roas: d.spend > 0 && d.sales > 0 ? d.sales / d.spend : 0,
      cpa: d.purchases > 0 ? d.spend / d.purchases : 0,
      cpm: d.impressions > 0 ? (d.spend / d.impressions) * 1000 : 0,
      outbound_ctr: d.impressions > 0 ? (d.outbound_clicks / d.impressions) * 100 : 0,
    }))
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
          Load Demo Data (Jun 1–1, 2026)
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
  const [showCalendar, setShowCalendar] = useState(false);
  const [calendarMonth, setCalendarMonth] = useState(() => { const d = new Date(); return { year: d.getFullYear(), month: d.getMonth() }; });
  const calendarRef = useRef(null);
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

  const dateFilteredAds = useMemo(() => {
    if (!metaAds) return null;
    let ads = metaAds;
    if (dateFrom) ads = ads.filter(a => !a.date || a.date >= dateFrom);
    if (dateTo) ads = ads.filter(a => !a.date || a.date <= dateTo);
    return ads;
  }, [metaAds, dateFrom, dateTo]);

  const productTotals = useMemo(() => {
    // In demo mode: derive KPIs from daily by-product data (respects date filter properly)
    if (isDemo && demoByProduct) {
      const rows = demoByProduct.filter(r =>
        (!dateFrom || r.date >= dateFrom) && (!dateTo || r.date <= dateTo)
      );
      const result = {};
      for (const row of rows) {
        if (!result[row.product]) result[row.product] = { spend: 0, sales: 0, purchases: 0 };
        result[row.product].spend += row.spend;
        result[row.product].sales += row.sales;
        result[row.product].purchases += row.purchases;
      }
      const allSpend = Object.values(result).reduce((s, d) => s + d.spend, 0);
      const allSales  = Object.values(result).reduce((s, d) => s + d.sales,  0);
      const allPurch  = Object.values(result).reduce((s, d) => s + d.purchases, 0);
      result["All Products"] = { spend: allSpend, sales: allSales, purchases: allPurch, roas: allSpend > 0 ? allSales / allSpend : 0 };
      for (const p in result) {
        if (p !== "All Products") result[p].roas = result[p].spend > 0 ? result[p].sales / result[p].spend : 0;
      }
      return result;
    }
    // For uploaded Excel data: use ad-level rows filtered by date
    const base = dateFilteredAds || metaAds;
    return base ? computeProductTotals(base) : {};
  }, [isDemo, demoByProduct, dateFrom, dateTo, dateFilteredAds, metaAds]);

  const presentProducts = useMemo(() => PRODUCTS.filter((p) => p === "All Products" || productTotals[p]), [productTotals]);

  const productData = productTotals[selectedProduct] || { spend: 0, sales: 0, purchases: 0, roas: 0 };

  // In demo mode, DEMO_ADS are period-totals (all dated 2026-05-27).
  // When a date filter is active, we scale each ad proportionally using demoByProduct daily data.
  const scaledDemoAds = useMemo(() => {
    if (!metaAds) return null;
    if (!isDemo || !demoByProduct) return dateFilteredAds || metaAds;
    if (!dateFrom && !dateTo) return metaAds;
    // Compute per-ad spend/sales/purchases from actual daily data for selected range
    // This gives each creator their true ROAS for the period (not scaled monthly ratio)
    const adSpend = {}, adSales = {}, adPurchases = {};
    for (const [idx, date, sp, sa, pu] of DEMO_ADS_DAILY) {
      if (dateFrom && date < dateFrom) continue;
      if (dateTo && date > dateTo) continue;
      const name = DEMO_AD_NAMES[idx];
      adSpend[name] = (adSpend[name] || 0) + sp;
      adSales[name] = (adSales[name] || 0) + sa;
      adPurchases[name] = (adPurchases[name] || 0) + pu;
    }
    // Also scale non-ROAS metrics (ATC, IC, impressions) proportionally using product daily data
    const productScale = {};
    const products = [...new Set(demoByProduct.map(r => r.product))];
    for (const product of products) {
      const all = demoByProduct.filter(r => r.product === product);
      const inRange = all.filter(r => (!dateFrom || r.date >= dateFrom) && (!dateTo || r.date <= dateTo));
      const totalSpend = all.reduce((s, r) => s + r.spend, 0);
      const rangeSpend = inRange.reduce((s, r) => s + r.spend, 0);
      productScale[product] = totalSpend > 0 ? rangeSpend / totalSpend : 0;
    }
    const allTotal = demoByProduct.reduce((s, r) => s + r.spend, 0);
    const allRange = demoByProduct.filter(r => (!dateFrom || r.date >= dateFrom) && (!dateTo || r.date <= dateTo)).reduce((s, r) => s + r.spend, 0);
    const overallScale = allTotal > 0 ? allRange / allTotal : 0;
    return metaAds.map(ad => {
      const sp = adSpend[ad.name] || 0;
      const sa = adSales[ad.name] || 0;
      const pu = adPurchases[ad.name] || 0;
      const scale = productScale[ad.product] ?? overallScale;
      return { ...ad, spend: sp, sales: sa, purchases: pu,
        roas: sp > 0 ? sa / sp : 0,
        add_to_cart: Math.round((ad.add_to_cart || 0) * scale),
        initiate_checkout: Math.round((ad.initiate_checkout || 0) * scale),
        impressions: Math.round((ad.impressions || 0) * scale),
        outbound_clicks: Math.round((ad.outbound_clicks || 0) * scale) };
    });
  }, [isDemo, metaAds, demoByProduct, dateFrom, dateTo, dateFilteredAds]);

  const filteredAds = useMemo(() => {
    const base = scaledDemoAds || metaAds;
    if (!base) return [];
    let ads = (selectedProduct !== "All Products" ? base.filter((a) => a.product === selectedProduct) : base);
    if (searchQuery) { const q = searchQuery.toLowerCase(); ads = ads.filter((a) => a.name.toLowerCase().includes(q) || a.product.toLowerCase().includes(q)); }
    if (adTypeFilter !== "All") ads = ads.filter((a) => a.type === adTypeFilter);
    if (creatorFilter !== "All") ads = ads.filter((a) => extractCreator(a.name) === creatorFilter);
    return [...ads].sort((a, b) => sortDir === "desc" ? b[sortField] - a[sortField] : a[sortField] - b[sortField]);
  }, [scaledDemoAds, metaAds, selectedProduct, searchQuery, sortField, sortDir, adTypeFilter, creatorFilter]);

  const maxSpend = useMemo(() => filteredAds.reduce((m, a) => Math.max(m, a.spend), 0), [filteredAds]);
  const uniqueCreators = useMemo(() => { const base = scaledDemoAds || metaAds; if (!base) return []; const s = new Set(base.map(a => extractCreator(a.name))); return ["All", ...Array.from(s).sort()]; }, [scaledDemoAds, metaAds]);
  const creatorTotals = useMemo(() => { const base = scaledDemoAds || metaAds; if (!base) return []; const pool = (creatorProductFilter !== "All Products" ? base.filter(a => a.product === creatorProductFilter) : base); return computeCreatorTotals(pool); }, [scaledDemoAds, metaAds, creatorProductFilter]);
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
              ? "Demo data · Jun 1–1, 2026 · Meta + Google Ads"
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
{hasData && (() => {
          // Close calendar on outside click
          const handleOutsideClick = (e) => {
            if (calendarRef.current && !calendarRef.current.contains(e.target)) setShowCalendar(false);
          };
          // Quick preset helper
          const applyPreset = (days) => {
            if (days === 0) { setDateFrom(""); setDateTo(""); setShowCalendar(false); return; }
            const to = new Date(); to.setHours(0,0,0,0);
            const from = new Date(to); from.setDate(from.getDate() - (days - 1));
            const fmt = d => d.toISOString().slice(0,10);
            setDateFrom(fmt(from)); setDateTo(fmt(to)); setShowCalendar(false);
          };
          const presets = [
            { label: "All", days: 0 }, { label: "Last 7D", days: 7 },
            { label: "Last 14D", days: 14 }, { label: "Last 30D", days: 30 }, { label: "This Month", days: -1 }
          ];
          // Calendar helpers
          const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
          const { year, month } = calendarMonth;
          const firstDay = new Date(year, month, 1).getDay();
          const daysInMonth = new Date(year, month + 1, 0).getDate();
          const calCells = [];
          for (let i = 0; i < firstDay; i++) calCells.push(null);
          for (let d = 1; d <= daysInMonth; d++) calCells.push(d);
          const fmt2 = (y,m,d) => `${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
          const applyThisMonth = () => {
            const now = new Date(); const y = now.getFullYear(); const m = now.getMonth();
            setDateFrom(fmt2(y,m,1)); setDateTo(fmt2(y,m,new Date(y,m+1,0).getDate())); setShowCalendar(false);
          };
          const handleDayClick = (d) => {
            const clicked = fmt2(year, month, d);
            if (!dateFrom || (dateFrom && dateTo)) { setDateFrom(clicked); setDateTo(""); }
            else if (clicked < dateFrom) { setDateTo(dateFrom); setDateFrom(clicked); }
            else { setDateTo(clicked); setShowCalendar(false); }
          };
          const isInRange = (d) => {
            const dt = fmt2(year, month, d);
            if (dateFrom && dateTo) return dt >= dateFrom && dt <= dateTo;
            if (dateFrom && !dateTo) return dt === dateFrom;
            return false;
          };
          const isStart = (d) => fmt2(year, month, d) === dateFrom;
          const isEnd = (d) => fmt2(year, month, d) === dateTo;
          const displayRange = dateFrom && dateTo ? `${dateFrom} → ${dateTo}` : dateFrom ? `From ${dateFrom}` : "All dates";
          return (
            <div ref={calendarRef} style={{ position: "relative", display: "inline-block", padding: "6px 0 4px" }}>
              {/* Trigger button */}
              <button
                onClick={() => { if (!showCalendar) { document.addEventListener('mousedown', handleOutsideClick, { once: true }); } setShowCalendar(v => !v); }}
                style={{ display:"flex", alignItems:"center", gap:8, padding:"7px 14px", borderRadius:9, border:"1px solid rgba(92,164,247,0.25)", background: (dateFrom||dateTo) ? "rgba(92,164,247,0.12)" : "rgba(255,255,255,0.04)", color:"#e8e6f0", fontSize:12, fontWeight:600, cursor:"pointer", fontFamily:"inherit", whiteSpace:"nowrap" }}
              >
                <span style={{ fontSize:14 }}>📅</span>
                <span style={{ color: (dateFrom||dateTo) ? "#5ca4f7" : "rgba(255,255,255,0.5)" }}>{displayRange}</span>
                <span style={{ fontSize:10, color:"rgba(255,255,255,0.3)", marginLeft:2 }}>{showCalendar ? "▲" : "▼"}</span>
              </button>
              {(dateFrom || dateTo) && (
                <button onClick={() => { setDateFrom(""); setDateTo(""); }} style={{ marginLeft:6, padding:"6px 10px", borderRadius:7, fontSize:11, fontWeight:600, cursor:"pointer", border:"1px solid rgba(248,113,113,0.3)", background:"rgba(248,113,113,0.08)", color:"#f87171", fontFamily:"inherit" }}>✕</button>
              )}
              {/* Dropdown panel */}
              {showCalendar && (
                <div style={{ position:"absolute", top:"calc(100% + 6px)", left:0, zIndex:999, background:"#1a1a2e", border:"1px solid rgba(255,255,255,0.12)", borderRadius:14, padding:16, boxShadow:"0 8px 32px rgba(0,0,0,0.5)", minWidth:300 }}>
                  {/* Preset buttons */}
                  <div style={{ display:"flex", gap:6, marginBottom:14, flexWrap:"wrap" }}>
                    {presets.map(p => (
                      <button key={p.label} onClick={() => p.days === -1 ? applyThisMonth() : applyPreset(p.days)}
                        style={{ padding:"4px 10px", borderRadius:6, fontSize:11, fontWeight:600, cursor:"pointer", fontFamily:"inherit",
                          border: "1px solid rgba(92,164,247,0.3)", background: "rgba(92,164,247,0.1)", color:"#5ca4f7" }}>
                        {p.label}
                      </button>
                    ))}
                  </div>
                  {/* Calendar header */}
                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:10 }}>
                    <button onClick={() => setCalendarMonth(m => { let mo = m.month-1, yr = m.year; if(mo<0){mo=11;yr--;} return {year:yr,month:mo}; })}
                      style={{ background:"none", border:"none", color:"#e8e6f0", cursor:"pointer", fontSize:16, padding:"0 6px" }}>‹</button>
                    <span style={{ fontWeight:700, fontSize:13, color:"#e8e6f0" }}>{monthNames[month]} {year}</span>
                    <button onClick={() => setCalendarMonth(m => { let mo = m.month+1, yr = m.year; if(mo>11){mo=0;yr++;} return {year:yr,month:mo}; })}
                      style={{ background:"none", border:"none", color:"#e8e6f0", cursor:"pointer", fontSize:16, padding:"0 6px" }}>›</button>
                  </div>
                  {/* Day headers */}
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)", gap:2, marginBottom:4 }}>
                    {["Su","Mo","Tu","We","Th","Fr","Sa"].map(d => (
                      <div key={d} style={{ textAlign:"center", fontSize:10, color:"rgba(255,255,255,0.3)", fontWeight:700, padding:"2px 0" }}>{d}</div>
                    ))}
                  </div>
                  {/* Day cells */}
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(7,1fr)", gap:2 }}>
                    {calCells.map((d, i) => d === null ? (
                      <div key={i} />
                    ) : (
                      <button key={i} onClick={() => handleDayClick(d)}
                        style={{ padding:"6px 0", textAlign:"center", fontSize:12, fontWeight: isStart(d)||isEnd(d) ? 700 : 400,
                          borderRadius: isStart(d)||isEnd(d) ? 6 : isInRange(d) ? 0 : 6,
                          background: isStart(d)||isEnd(d) ? "#5ca4f7" : isInRange(d) ? "rgba(92,164,247,0.18)" : "transparent",
                          color: isStart(d)||isEnd(d) ? "#fff" : isInRange(d) ? "#5ca4f7" : "#e8e6f0",
                          border: "none", cursor:"pointer", fontFamily:"inherit" }}>
                        {d}
                      </button>
                    ))}
                  </div>
                  {/* Hint text */}
                  <div style={{ marginTop:10, fontSize:10, color:"rgba(255,255,255,0.25)", textAlign:"center" }}>
                    {!dateFrom ? "Click to set start date" : !dateTo ? "Click to set end date" : `${dateFrom} → ${dateTo}`}
                  </div>
                </div>
              )}
            </div>
          );
        })()}
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
                {(dateFrom || dateTo) && (
                  <span style={{ padding: "3px 10px", borderRadius: 6, fontSize: 11, fontWeight: 600, background: "rgba(92,164,247,0.12)", color: "#5ca4f7", border: "1px solid rgba(92,164,247,0.25)" }}>
                    📅 {dateFrom && dateTo ? dateFrom + " → " + dateTo : dateFrom ? "From " + dateFrom : "Until " + dateTo}
                  </span>
                )}
                <button onClick={() => exportCSV(creatorTotals.flatMap(c => (creatorProductFilter !== "All Products" ? (scaledDemoAds || metaAds || []).filter(a => a.product === creatorProductFilter) : (scaledDemoAds || metaAds || [])).filter(a => extractCreator(a.name) === c.creator)), "creator-export.csv")} style={{ marginLeft: "auto", padding: "6px 14px", borderRadius: 7, fontSize: 11, fontWeight: 600, cursor: "pointer", border: "1px solid rgba(92,164,247,0.3)", background: "rgba(92,164,247,0.08)", color: "#5ca4f7", fontFamily: "inherit" }}>↓ Export CSV</button>
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
                      {[["Creator","left"],["Ads","center"],["Spend","right"],["Revenue","right"],["ROAS","right"],["Orders","right"],["CPA","right"],["Add to Cart","right"],["Init. Checkout","right"],["Impressions","right"],["CPM","right"],["Outbound CTR","right"],["Started","right"]].map(([h,a]) => (
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
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, color: "rgba(255,255,255,0.7)" }}>{c.add_to_cart > 0 ? c.add_to_cart.toLocaleString("en-IN") : "—"}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, color: "rgba(255,255,255,0.7)" }}>{c.initiate_checkout > 0 ? c.initiate_checkout.toLocaleString("en-IN") : "—"}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.5)" }}>{c.impressions > 0 ? c.impressions.toLocaleString("en-IN") : "—"}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{c.cpm > 0 ? "₹" + c.cpm.toFixed(0) : "—"}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: c.outbound_ctr >= 2 ? "#34d399" : c.outbound_ctr >= 1 ? "#fbbf24" : "rgba(255,255,255,0.7)" }}>{c.outbound_ctr > 0 ? c.outbound_ctr.toFixed(2) + "%" : "—"}</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.5)", whiteSpace: "nowrap" }}>{c.start_date ? new Date(c.start_date).toLocaleDateString("en-IN", {day:"2-digit",month:"short"}) : "—"}</td>
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
                        {[["Ad Name","left"],["Product","left"],["Spend","right"],["Revenue","right"],["ROAS","right"],["Orders","right"],["Add to Cart","right"],["Init. Checkout","right"],["Impressions","right"],["CPM","right"],["Out. CTR","right"],["Started","right"]].map(([h,a]) => (
                          <th key={h} style={{ padding: "10px 12px", fontSize: 11, fontWeight: 600, textTransform: "uppercase", color: "rgba(255,255,255,0.35)", textAlign: a, borderBottom: "1px solid rgba(255,255,255,0.06)", whiteSpace: "nowrap" }}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {(creatorProductFilter !== "All Products" ? (scaledDemoAds || metaAds || []).filter(a => a.product === creatorProductFilter) : (scaledDemoAds || metaAds || [])).filter(a => extractCreator(a.name) === creatorFilter).sort((a,b) => b.spend - a.spend).map((ad, i) => (
                          <tr key={i} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.015)" : "transparent" }}>
                            <td style={{ padding: "10px 12px", fontSize: 12, color: "#fff", maxWidth: 300, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{ad.name}</td>
                            <td style={{ padding: "10px 12px" }}><span style={{ fontSize: 11, fontWeight: 600, color: PRODUCT_COLORS[ad.product] || "#888", background: (PRODUCT_COLORS[ad.product] || "#888") + "15", padding: "2px 7px", borderRadius: 4 }}>{ad.product}</span></td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmt(ad.spend)}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{fmt(ad.sales)}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right" }}><RoasBadge value={ad.roas} /></td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontWeight: 600, color: "#fff" }}>{ad.purchases}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, color: "rgba(255,255,255,0.7)" }}>{ad.add_to_cart > 0 ? ad.add_to_cart : "—"}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, color: "rgba(255,255,255,0.7)" }}>{ad.initiate_checkout > 0 ? ad.initiate_checkout : "—"}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.5)" }}>{ad.impressions > 0 ? ad.impressions.toLocaleString("en-IN") : "—"}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{ad.impressions > 0 ? "₹" + ((ad.spend / ad.impressions) * 1000).toFixed(0) : "—"}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: "rgba(255,255,255,0.7)" }}>{ad.impressions > 0 ? ((ad.outbound_clicks / ad.impressions) * 100).toFixed(2) + "%" : "—"}</td>
                            <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12, color: "rgba(255,255,255,0.5)", whiteSpace: "nowrap" }}>{ad.start_date ? new Date(ad.start_date).toLocaleDateString("en-IN", {day:"2-digit",month:"short"}) : "—"}</td>
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
        <LookerStudioView
          metaAds={metaAds}
          preBlended={(() => { const b = isDemo ? demoBlended : null; if (!b) return null; if (dateFrom || dateTo) return b.filter(r => (!dateFrom || r.date >= dateFrom) && (!dateTo || r.date <= dateTo)); return b; })()}
          preByProduct={(() => { const bp = isDemo ? demoByProduct : null; if (!bp) return null; if (dateFrom || dateTo) return bp.filter(r => (!dateFrom || r.date >= dateFrom) && (!dateTo || r.date <= dateTo)); return bp; })()}
        />
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
