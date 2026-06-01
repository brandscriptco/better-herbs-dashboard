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
const DEMO_ADS = [{"product":"Brainify Drops","name":"Brainify drops -Maninder Kaur (Blessings) Video","spend":92087.77,"sales":291630.18,"purchases":375,"roas":3.17,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":808473,"outbound_clicks":14553,"add_to_cart":961,"initiate_checkout":724},{"product":"Brainify Powder","name":"Brainify Powder - Divya Bajpai","spend":90014.21,"sales":323892.0,"purchases":368,"roas":3.6,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":971737,"outbound_clicks":24293,"add_to_cart":951,"initiate_checkout":806},{"product":"Brainify Drops","name":"Brainify drops - Dr. Vinod Video","spend":70996.4,"sales":198089.2,"purchases":269,"roas":2.79,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":962781,"outbound_clicks":27921,"add_to_cart":745,"initiate_checkout":488},{"product":"Brainify Powder","name":"Brainify Powder - Static - 13 (Ingredients)","spend":65366.65,"sales":167939.4,"purchases":200,"roas":2.57,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":560944,"outbound_clicks":8975,"add_to_cart":577,"initiate_checkout":471},{"product":"Brainify Drops","name":"Brainify drops - Dr.sajid","spend":50229.54,"sales":117133.0,"purchases":157,"roas":2.33,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":457632,"outbound_clicks":11441,"add_to_cart":404,"initiate_checkout":280},{"product":"Brainify Powder","name":"Brainify Powder - Dr Vinod 2","spend":47666.26,"sales":157366.0,"purchases":213,"roas":3.3,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":692615,"outbound_clicks":20778,"add_to_cart":709,"initiate_checkout":538},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Ankit Jha  Video","spend":44824.28,"sales":151622.0,"purchases":194,"roas":3.38,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":407929,"outbound_clicks":13870,"add_to_cart":474,"initiate_checkout":362},{"product":"Lactify","name":"Lactify_Dr. Nayana","spend":39666.58,"sales":143341.0,"purchases":162,"roas":3.61,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":468609,"outbound_clicks":11715,"add_to_cart":475,"initiate_checkout":369},{"product":"Brainify Powder","name":"Brainify Powder - Doctor Compilation Video - 2  | 07/05/2026","spend":32756.85,"sales":92532.0,"purchases":118,"roas":2.82,"type":"Video","date":"2026-05-31","start_date":"2026-05-07","impressions":355634,"outbound_clicks":8180,"add_to_cart":270,"initiate_checkout":227},{"product":"Mamafy","name":"Mamafy -  Static - 13","spend":31979.4,"sales":71414.0,"purchases":86,"roas":2.23,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":406397,"outbound_clicks":8534,"add_to_cart":234,"initiate_checkout":199},{"product":"Brainify Drops","name":"Brainify drops - Static 7","spend":30845.03,"sales":89479.0,"purchases":113,"roas":2.9,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":272783,"outbound_clicks":6820,"add_to_cart":279,"initiate_checkout":206},{"product":"Mamafy","name":"Mamafy - Dr.Shaifali Ad code","spend":28168.52,"sales":90030.0,"purchases":108,"roas":3.2,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":570322,"outbound_clicks":13117,"add_to_cart":353,"initiate_checkout":297},{"product":"Mamafy","name":"Mamafy - Dr. Garima","spend":26834.45,"sales":72267.0,"purchases":89,"roas":2.69,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":800257,"outbound_clicks":12804,"add_to_cart":292,"initiate_checkout":258},{"product":"Flowjoy","name":"Flowjoy Drop - Dispatch video","spend":24672.58,"sales":40650.0,"purchases":56,"roas":1.65,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":186257,"outbound_clicks":4284,"add_to_cart":147,"initiate_checkout":95},{"product":"Brainify Drops","name":"Brainify drops - Dr. Sajid","spend":24048.87,"sales":55767.0,"purchases":77,"roas":2.32,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":218935,"outbound_clicks":6568,"add_to_cart":232,"initiate_checkout":155},{"product":"Flowjoy","name":"flowjoy drops - Dr. Garima | 13/05/2026","spend":24011.62,"sales":63135.0,"purchases":94,"roas":2.63,"type":"Static","date":"2026-05-31","start_date":"2026-05-13","impressions":414557,"outbound_clicks":7462,"add_to_cart":245,"initiate_checkout":178},{"product":"Mamafy","name":"Mamafy -  Static - 4 - Do not buy","spend":23632.32,"sales":63717.0,"purchases":71,"roas":2.7,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":280999,"outbound_clicks":5901,"add_to_cart":179,"initiate_checkout":162},{"product":"Lactify","name":"Lactify - Customer Review - Lactation","spend":23435.94,"sales":50301.0,"purchases":56,"roas":2.15,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":145581,"outbound_clicks":2766,"add_to_cart":136,"initiate_checkout":116},{"product":"Mamafy","name":"mamafy - Founder's  Dispatch Video May MOF | 13/05/2026","spend":21383.53,"sales":50710.0,"purchases":61,"roas":2.37,"type":"Video","date":"2026-05-31","start_date":"2026-05-13","impressions":241198,"outbound_clicks":6512,"add_to_cart":220,"initiate_checkout":179},{"product":"Brainify Powder","name":"Brainify Powder - Brainify Powder Static - 12   | 07/05/2026","spend":19710.76,"sales":64199.0,"purchases":79,"roas":3.26,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":188449,"outbound_clicks":4711,"add_to_cart":199,"initiate_checkout":163},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Prachi Mahajan","spend":19514.79,"sales":45828.0,"purchases":50,"roas":2.35,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":150587,"outbound_clicks":2259,"add_to_cart":108,"initiate_checkout":80},{"product":"Flowjoy","name":"Flowjoy Drops Static 2","spend":17914.65,"sales":31689.4,"purchases":36,"roas":1.77,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":127370,"outbound_clicks":4203,"add_to_cart":100,"initiate_checkout":77},{"product":"Brainify Powder","name":"Brainify Powder -Dr. Ankit Jha | 04/05/2026","spend":16427.46,"sales":43546.0,"purchases":62,"roas":2.65,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":206879,"outbound_clicks":6413,"add_to_cart":201,"initiate_checkout":154},{"product":"Mamafy","name":"mamafy - Dr. Samra - Edit 2 (Direct) | 04/05/2026","spend":16176.12,"sales":36022.0,"purchases":42,"roas":2.23,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":227931,"outbound_clicks":5242,"add_to_cart":130,"initiate_checkout":95},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Pushpendra 2nd video","spend":15166.45,"sales":35775.0,"purchases":48,"roas":2.36,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":151902,"outbound_clicks":2279,"add_to_cart":118,"initiate_checkout":99},{"product":"Brainify Drops","name":"Brainify drops - Rajmani","spend":15066.62,"sales":37219.0,"purchases":54,"roas":2.47,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":183142,"outbound_clicks":4579,"add_to_cart":141,"initiate_checkout":96},{"product":"Flowjoy","name":"Flow Drop - Dr. Ankit","spend":14451.06,"sales":18739.0,"purchases":24,"roas":1.3,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":102921,"outbound_clicks":2985,"add_to_cart":98,"initiate_checkout":75},{"product":"Lactify","name":"Lactify - Dr. Naaz | 13/05/2026","spend":13966.34,"sales":32049.0,"purchases":39,"roas":2.29,"type":"Static","date":"2026-05-31","start_date":"2026-05-13","impressions":238923,"outbound_clicks":5973,"add_to_cart":150,"initiate_checkout":114},{"product":"Mamafy","name":"mamafy - Static - 14| 13/05/2026","spend":11313.11,"sales":28241.0,"purchases":32,"roas":2.5,"type":"Static","date":"2026-05-31","start_date":"2026-05-13","impressions":145306,"outbound_clicks":4650,"add_to_cart":84,"initiate_checkout":78},{"product":"Lactify","name":"Lactify - 2nd Doctor Compilation","spend":10843.5,"sales":14092.0,"purchases":17,"roas":1.3,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":113986,"outbound_clicks":2166,"add_to_cart":86,"initiate_checkout":60},{"product":"Brainify Drops","name":"Brainify drops - Static - 8","spend":10492.51,"sales":26991.0,"purchases":31,"roas":2.57,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":96520,"outbound_clicks":1641,"add_to_cart":97,"initiate_checkout":79},{"product":"Mamafy","name":"Mamafy - Doctor Compilation Video","spend":10356.3,"sales":31972.0,"purchases":42,"roas":3.09,"type":"Video","date":"2026-05-31","start_date":"2026-05-03","impressions":409121,"outbound_clicks":6955,"add_to_cart":150,"initiate_checkout":121},{"product":"Brainify Drops","name":"Brainify drops - Rajmani Patel - Edited | 07/05/2026","spend":10080.6,"sales":19795.0,"purchases":22,"roas":1.96,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":65699,"outbound_clicks":1183,"add_to_cart":53,"initiate_checkout":37},{"product":"Brainify Drops","name":"Brainify drops - Static 11","spend":9875.82,"sales":20656.0,"purchases":25,"roas":2.09,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":85256,"outbound_clicks":1961,"add_to_cart":56,"initiate_checkout":45},{"product":"Mamafy","name":"Mamafy -  Dr. Smriti - Edited","spend":9199.56,"sales":20831.0,"purchases":27,"roas":2.26,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":136586,"outbound_clicks":4234,"add_to_cart":62,"initiate_checkout":50},{"product":"Brainify Powder","name":"Brainify Powder Static - 22 | 19/05/2026","spend":9108.23,"sales":26192.2,"purchases":31,"roas":2.88,"type":"Static","date":"2026-05-31","start_date":"2026-05-19","impressions":73915,"outbound_clicks":1478,"add_to_cart":70,"initiate_checkout":67},{"product":"Brainify Drops","name":"Brainify drops - Static - 17  19/05/2026","spend":9102.15,"sales":18842.0,"purchases":22,"roas":2.07,"type":"Static","date":"2026-05-31","start_date":"2026-05-19","impressions":60683,"outbound_clicks":1578,"add_to_cart":63,"initiate_checkout":49},{"product":"Brainify Drops","name":"Drops-Dr. Vinod Video","spend":8832.26,"sales":15367.0,"purchases":20,"roas":1.74,"type":"Video","date":"2026-05-31","start_date":"2026-05-05","impressions":91650,"outbound_clicks":2658,"add_to_cart":60,"initiate_checkout":38},{"product":"Brainify Drops","name":"Brainify drops - Dr. Rajeswari (Telugu) 19/05/2026","spend":8763.76,"sales":18192.0,"purchases":26,"roas":2.08,"type":"Static","date":"2026-05-31","start_date":"2026-05-19","impressions":91465,"outbound_clicks":2470,"add_to_cart":67,"initiate_checkout":51},{"product":"Flowjoy","name":"Flow drops - Dr. Sushma Mogri","spend":7863.78,"sales":14640.0,"purchases":20,"roas":1.86,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":95469,"outbound_clicks":1528,"add_to_cart":49,"initiate_checkout":39},{"product":"Brainify Powder","name":"Brainify Powder -Dr. Mallika (Telugu)   | 07/05/2026","spend":7853.97,"sales":15787.0,"purchases":22,"roas":2.01,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":76635,"outbound_clicks":1763,"add_to_cart":63,"initiate_checkout":41},{"product":"Brainify Powder","name":"Brainify Powder April Dispatch Video  | 21/05/2026","spend":7706.22,"sales":15518.0,"purchases":20,"roas":2.01,"type":"Video","date":"2026-05-31","start_date":"2026-05-21","impressions":45855,"outbound_clicks":825,"add_to_cart":44,"initiate_checkout":31},{"product":"Lactify","name":"Lactify - Dr. Sunil (Kannad)  | 21/05/2026","spend":7700.34,"sales":11319.0,"purchases":14,"roas":1.47,"type":"Static","date":"2026-05-31","start_date":"2026-05-21","impressions":108351,"outbound_clicks":1842,"add_to_cart":61,"initiate_checkout":49},{"product":"Lactify","name":"Lactify - Dr. Manisha  | 21/05/2026","spend":7666.57,"sales":16870.0,"purchases":20,"roas":2.2,"type":"Static","date":"2026-05-31","start_date":"2026-05-21","impressions":43711,"outbound_clicks":874,"add_to_cart":53,"initiate_checkout":39},{"product":"Lactify","name":"Lactify - Dr. Gunjan - 2nd Video | 21/05/2026","spend":7649.3,"sales":12153.0,"purchases":15,"roas":1.59,"type":"Video","date":"2026-05-31","start_date":"2026-05-21","impressions":84891,"outbound_clicks":2886,"add_to_cart":65,"initiate_checkout":58},{"product":"Brainify Drops","name":"Brainify drops - Static - 3","spend":7572.37,"sales":16959.0,"purchases":20,"roas":2.24,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":63070,"outbound_clicks":1829,"add_to_cart":43,"initiate_checkout":35},{"product":"Brainify Drops","name":"Drops-Dr Maninder Kaur (Copycat Brand)","spend":7300.98,"sales":20972.0,"purchases":22,"roas":2.87,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":57057,"outbound_clicks":1883,"add_to_cart":72,"initiate_checkout":49},{"product":"Flowjoy","name":"Flowjoy - Dr. Tanya - Ad code | 23/05/2026","spend":6417.74,"sales":8454.0,"purchases":14,"roas":1.32,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":84447,"outbound_clicks":2111,"add_to_cart":39,"initiate_checkout":33},{"product":"Mamafy","name":"mamafy - Doctor Compilation video 1 - Mamafy- April | 23/05/2026","spend":6309.39,"sales":10182.0,"purchases":14,"roas":1.61,"type":"Video","date":"2026-05-31","start_date":"2026-05-23","impressions":86576,"outbound_clicks":1385,"add_to_cart":43,"initiate_checkout":31},{"product":"Lactify","name":"Lactify - Dr. Rohit Bharadwaj","spend":6237.82,"sales":4989.0,"purchases":7,"roas":0.8,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":43763,"outbound_clicks":919,"add_to_cart":24,"initiate_checkout":17},{"product":"Flowjoy","name":"Flowjoy - Aanchal Naherwa - Ad code | 23/05/2026","spend":5958.12,"sales":5686.0,"purchases":6,"roas":0.95,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":78235,"outbound_clicks":2582,"add_to_cart":15,"initiate_checkout":13},{"product":"Brainify Drops","name":"Brainify - Dr Vinod 2","spend":5909.67,"sales":18406.0,"purchases":24,"roas":3.11,"type":"Static","date":"2026-05-31","start_date":"2026-05-02","impressions":69840,"outbound_clicks":1606,"add_to_cart":68,"initiate_checkout":56},{"product":"Brainify Drops","name":"Brainify drops - Maninder Kaur - March - Edited 19/05/2026","spend":5825.15,"sales":10015.0,"purchases":14,"roas":1.72,"type":"Static","date":"2026-05-31","start_date":"2026-05-19","impressions":46322,"outbound_clicks":695,"add_to_cart":61,"initiate_checkout":29},{"product":"Mamafy","name":"mamafy - Dr. Samra - Edit 4 (Nutrition) | 04/05/2026","spend":5478.25,"sales":12313.0,"purchases":15,"roas":2.25,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":75313,"outbound_clicks":2410,"add_to_cart":53,"initiate_checkout":38},{"product":"Flowjoy","name":"flowjoy  - Henna Jain  - Ad code | 23/05/2026","spend":5477.31,"sales":7960.0,"purchases":12,"roas":1.45,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":53411,"outbound_clicks":1015,"add_to_cart":30,"initiate_checkout":18},{"product":"Brainify Drops","name":"Brainify drops - Static 12","spend":5203.41,"sales":14121.0,"purchases":17,"roas":2.71,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":40392,"outbound_clicks":1050,"add_to_cart":34,"initiate_checkout":32},{"product":"Brainify Powder","name":"Brainify Powder -Static 19  | 04/05/2026","spend":5067.23,"sales":11378.0,"purchases":12,"roas":2.25,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":44076,"outbound_clicks":1322,"add_to_cart":27,"initiate_checkout":20},{"product":"Brainify Powder","name":"Brainify Powder -Rajmani Patel  | 04/05/2026","spend":5063.75,"sales":5712.0,"purchases":7,"roas":1.13,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":39725,"outbound_clicks":1351,"add_to_cart":28,"initiate_checkout":23},{"product":"Brainify Powder","name":"Brainify Powder -Static 20  | 04/05/2026","spend":5062.61,"sales":10744.0,"purchases":14,"roas":2.12,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":44241,"outbound_clicks":1504,"add_to_cart":32,"initiate_checkout":21},{"product":"Brainify Powder","name":"Brainify Powder -Static 18  | 04/05/2026","spend":5031.59,"sales":13004.0,"purchases":15,"roas":2.58,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":46422,"outbound_clicks":743,"add_to_cart":23,"initiate_checkout":27},{"product":"Brainify Drops","name":"Brainify - USP Static","spend":4552.74,"sales":13683.0,"purchases":15,"roas":3.01,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":36114,"outbound_clicks":1120,"add_to_cart":44,"initiate_checkout":35},{"product":"Mamafy","name":"Mamafy - Dr. Soniya Gupta","spend":4313.03,"sales":10331.0,"purchases":13,"roas":2.4,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":62873,"outbound_clicks":1698,"add_to_cart":31,"initiate_checkout":23},{"product":"Brainify Drops","name":"Brainify drops - Dr. Padma Tamil 21/05/2026","spend":4308.88,"sales":4593.0,"purchases":7,"roas":1.07,"type":"Static","date":"2026-05-31","start_date":"2026-05-21","impressions":29362,"outbound_clicks":734,"add_to_cart":20,"initiate_checkout":15},{"product":"Mamafy","name":"mamafy - Dr. Priya Soni| 13/05/2026","spend":4179.31,"sales":6756.0,"purchases":10,"roas":1.62,"type":"Static","date":"2026-05-31","start_date":"2026-05-13","impressions":58215,"outbound_clicks":1572,"add_to_cart":31,"initiate_checkout":31},{"product":"Mamafy","name":"mamafy - B - Roll Video| 26/05/2026","spend":4171.86,"sales":13748.0,"purchases":17,"roas":3.3,"type":"Video","date":"2026-05-31","start_date":"2026-05-26","impressions":57181,"outbound_clicks":1144,"add_to_cart":46,"initiate_checkout":45},{"product":"Brainify Drops","name":"Drops - Dr.sajid","spend":3916.46,"sales":3075.0,"purchases":5,"roas":0.79,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":32025,"outbound_clicks":833,"add_to_cart":19,"initiate_checkout":17},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Pillai (Tamil) 2nd Video | 26/05/2026","spend":3834.35,"sales":5771.0,"purchases":9,"roas":1.51,"type":"Video","date":"2026-05-31","start_date":"2026-05-26","impressions":29079,"outbound_clicks":553,"add_to_cart":36,"initiate_checkout":26},{"product":"Brainify Powder","name":"Brainify Powder -  Powder Static - 21 | 26/05/2026","spend":3789.23,"sales":2556.0,"purchases":4,"roas":0.67,"type":"Static","date":"2026-05-31","start_date":"2026-05-26","impressions":27392,"outbound_clicks":767,"add_to_cart":12,"initiate_checkout":14},{"product":"Brainify Drops","name":"Drops- Static 3","spend":3783.87,"sales":4743.0,"purchases":7,"roas":1.25,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":33925,"outbound_clicks":543,"add_to_cart":23,"initiate_checkout":17},{"product":"Lactify","name":"Lactify - Dr. Garima | 13/05/2026","spend":3596.55,"sales":3305.0,"purchases":5,"roas":0.92,"type":"Static","date":"2026-05-31","start_date":"2026-05-13","impressions":21059,"outbound_clicks":590,"add_to_cart":18,"initiate_checkout":10},{"product":"Lactify","name":"Lactify - Static - 27 | 13/05/2026","spend":3594.98,"sales":6071.0,"purchases":8,"roas":1.69,"type":"Static","date":"2026-05-31","start_date":"2026-05-13","impressions":22646,"outbound_clicks":498,"add_to_cart":28,"initiate_checkout":19},{"product":"Brainify Drops","name":"Brainify drops - Maninder Kaur - April - 1  25/05/2026","spend":3443.8,"sales":3706.0,"purchases":4,"roas":1.08,"type":"Static","date":"2026-05-31","start_date":"2026-05-25","impressions":28878,"outbound_clicks":520,"add_to_cart":25,"initiate_checkout":19},{"product":"Lactify","name":"Lactify - Dr. Priyanka (Ped.)","spend":3403.0,"sales":9934.0,"purchases":14,"roas":2.92,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":24214,"outbound_clicks":436,"add_to_cart":17,"initiate_checkout":12},{"product":"Brainify Drops","name":"Brainify drops - Static 16","spend":3402.4,"sales":13900.0,"purchases":18,"roas":4.09,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":26992,"outbound_clicks":729,"add_to_cart":30,"initiate_checkout":22},{"product":"Mamafy","name":"mamafy - Founder's  Dispatch Video May MOF","spend":3256.81,"sales":10980.0,"purchases":13,"roas":3.37,"type":"Video","date":"2026-05-31","start_date":"2026-05-15","impressions":28796,"outbound_clicks":720,"add_to_cart":48,"initiate_checkout":36},{"product":"Brainify Powder","name":"Brainify Powder Dr. Shiba (Odia)i) | 19/05/2026","spend":3145.09,"sales":1907.0,"purchases":3,"roas":0.61,"type":"Static","date":"2026-05-31","start_date":"2026-05-19","impressions":21769,"outbound_clicks":479,"add_to_cart":12,"initiate_checkout":15},{"product":"Brainify Powder","name":"Brainify Powder -Dr. Taran (Punjabi) | 19/05/2026","spend":3137.35,"sales":1638.0,"purchases":2,"roas":0.52,"type":"Static","date":"2026-05-31","start_date":"2026-05-19","impressions":21393,"outbound_clicks":620,"add_to_cart":5,"initiate_checkout":6},{"product":"Brainify Drops","name":"Brainify - Google Static","spend":3068.48,"sales":7136.0,"purchases":10,"roas":2.33,"type":"Static","date":"2026-05-31","start_date":"2026-05-02","impressions":24461,"outbound_clicks":660,"add_to_cart":21,"initiate_checkout":17},{"product":"Mamafy","name":"mamafy - Pooja Shah | 23/05/2026","spend":2978.62,"sales":4834.0,"purchases":6,"roas":1.62,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":28959,"outbound_clicks":463,"add_to_cart":18,"initiate_checkout":15},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Prachi Mahajan | 07/05/2026","spend":2804.62,"sales":1997.0,"purchases":3,"roas":0.71,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":22637,"outbound_clicks":475,"add_to_cart":18,"initiate_checkout":15},{"product":"Mamafy","name":"Mamafy - Static 2","spend":2717.35,"sales":10189.0,"purchases":13,"roas":3.75,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":32237,"outbound_clicks":613,"add_to_cart":37,"initiate_checkout":21},{"product":"Brainify Drops","name":"Brainify drops - Dr. Padma 07/05/2026","spend":2645.07,"sales":1098.0,"purchases":2,"roas":0.42,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":19414,"outbound_clicks":311,"add_to_cart":6,"initiate_checkout":4},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Smriti   | 07/05/2026","spend":2614.28,"sales":2946.0,"purchases":3,"roas":1.13,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":16466,"outbound_clicks":412,"add_to_cart":7,"initiate_checkout":5},{"product":"Mamafy","name":"mamafy - Dr. Soniya Gupta 1st Edit | 04/05/2026","spend":2510.47,"sales":649.0,"purchases":1,"roas":0.26,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":47516,"outbound_clicks":855,"add_to_cart":10,"initiate_checkout":7},{"product":"Brainify Powder","name":"Brainify Powder Dr. Vinod (Brainify Powder)  | 28/05/2026","spend":2400.72,"sales":5508.0,"purchases":8,"roas":2.29,"type":"Static","date":"2026-05-31","start_date":"2026-05-28","impressions":18352,"outbound_clicks":404,"add_to_cart":28,"initiate_checkout":20},{"product":"Brainify Powder","name":"Brainify Powder Divya Bajpai 2nd video  | 28/05/2026","spend":2341.07,"sales":3964.0,"purchases":6,"roas":1.69,"type":"Video","date":"2026-05-31","start_date":"2026-05-28","impressions":23225,"outbound_clicks":650,"add_to_cart":14,"initiate_checkout":15},{"product":"Lactify","name":"Lactify - Dr. Pillai (Tamil) | 28/05/2026","spend":2338.94,"sales":2809.0,"purchases":3,"roas":1.2,"type":"Static","date":"2026-05-31","start_date":"2026-05-28","impressions":17260,"outbound_clicks":397,"add_to_cart":11,"initiate_checkout":9},{"product":"Lactify","name":"Lactify - Founder's Lactify - Dispatch Video May MOF 28/05/2026","spend":2337.87,"sales":3569.0,"purchases":4,"roas":1.53,"type":"Video","date":"2026-05-31","start_date":"2026-05-28","impressions":9751,"outbound_clicks":254,"add_to_cart":15,"initiate_checkout":9},{"product":"Brainify Powder","name":"Brainify Powder Doctor Compilation Video - 1 (Brainify Powder)  | 28/05/2026","spend":2333.42,"sales":3544.0,"purchases":4,"roas":1.52,"type":"Video","date":"2026-05-31","start_date":"2026-05-28","impressions":17656,"outbound_clicks":441,"add_to_cart":20,"initiate_checkout":9},{"product":"Mamafy","name":"mamafy - Mamafy Static - 17 | 23/05/2026","spend":2325.71,"sales":2095.0,"purchases":3,"roas":0.9,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":20375,"outbound_clicks":387,"add_to_cart":16,"initiate_checkout":11},{"product":"Brainify Drops","name":"Brainify drops - Doctor Compilation Video - March  18  28/05/2026","spend":2293.07,"sales":3095.0,"purchases":5,"roas":1.35,"type":"Video","date":"2026-05-31","start_date":"2026-05-28","impressions":18070,"outbound_clicks":361,"add_to_cart":10,"initiate_checkout":10},{"product":"Brainify Powder","name":"Brainify Powder Static - 23  | 28/05/2026","spend":2292.82,"sales":1424.0,"purchases":1,"roas":0.62,"type":"Static","date":"2026-05-31","start_date":"2026-05-28","impressions":16154,"outbound_clicks":323,"add_to_cart":14,"initiate_checkout":9},{"product":"Lactify","name":"Lactify - Lactify Static - 29 28/05/2026","spend":2280.45,"sales":5853.0,"purchases":7,"roas":2.57,"type":"Static","date":"2026-05-31","start_date":"2026-05-28","impressions":17220,"outbound_clicks":499,"add_to_cart":14,"initiate_checkout":9},{"product":"Brainify Drops","name":"Brainify drops - Ingredients Video 28/05/2026","spend":2279.26,"sales":5721.0,"purchases":9,"roas":2.51,"type":"Video","date":"2026-05-31","start_date":"2026-05-28","impressions":18408,"outbound_clicks":479,"add_to_cart":21,"initiate_checkout":16},{"product":"Brainify Drops","name":"Brainify drops - Static - 18  28/05/2026","spend":2277.51,"sales":3155.0,"purchases":5,"roas":1.39,"type":"Static","date":"2026-05-31","start_date":"2026-05-28","impressions":15945,"outbound_clicks":399,"add_to_cart":16,"initiate_checkout":15},{"product":"Mamafy","name":"mamafy - Dr. Garima 2nd video 28/05/2026","spend":2270.88,"sales":8359.0,"purchases":11,"roas":3.68,"type":"Video","date":"2026-05-31","start_date":"2026-05-28","impressions":29782,"outbound_clicks":923,"add_to_cart":27,"initiate_checkout":22},{"product":"Mamafy","name":"mamafy -  2nd founder video - Mamafy Powder 28/05/2026","spend":2268.1,"sales":2840.0,"purchases":4,"roas":1.25,"type":"Video","date":"2026-05-31","start_date":"2026-05-28","impressions":16112,"outbound_clicks":387,"add_to_cart":10,"initiate_checkout":7},{"product":"Brainify Drops","name":"Brainify drops -Brainify Drops Static - 16 | 07/05/2026","spend":2253.78,"sales":3156.0,"purchases":4,"roas":1.4,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":14673,"outbound_clicks":396,"add_to_cart":12,"initiate_checkout":9},{"product":"Brainify Powder","name":"Brainify Powder - Brainify Powder Static - 17   | 07/05/2026","spend":2253.59,"sales":4853.0,"purchases":6,"roas":2.15,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":19928,"outbound_clicks":598,"add_to_cart":9,"initiate_checkout":10},{"product":"Brainify Powder","name":"Brainify Powder - Brainify Powder Static - 16   | 07/05/2026","spend":2229.56,"sales":3035.0,"purchases":4,"roas":1.36,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":16672,"outbound_clicks":250,"add_to_cart":13,"initiate_checkout":11},{"product":"Brainify Powder","name":"Brainify Powder - Brainify Powder Static - 14   | 07/05/2026","spend":2212.35,"sales":2586.0,"purchases":4,"roas":1.17,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":18116,"outbound_clicks":308,"add_to_cart":6,"initiate_checkout":8},{"product":"Mamafy","name":"Mamafy - Mamafy Static - 6 - USP | 23/03/2026","spend":2193.49,"sales":5603.0,"purchases":7,"roas":2.55,"type":"Static","date":"2026-05-31","start_date":"2026-03-23","impressions":26053,"outbound_clicks":860,"add_to_cart":25,"initiate_checkout":16},{"product":"Brainify Drops","name":"Drops - Static 7","spend":2169.54,"sales":3336.0,"purchases":4,"roas":1.54,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":23602,"outbound_clicks":779,"add_to_cart":20,"initiate_checkout":13},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Divyani Bhagat   | 07/05/2026","spend":2161.87,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":17736,"outbound_clicks":319,"add_to_cart":3,"initiate_checkout":2},{"product":"Brainify Powder","name":"Brainify Powder -Dr. Sonal | 04/05/2026","spend":2111.65,"sales":4543.0,"purchases":4,"roas":2.15,"type":"Static","date":"2026-05-31","start_date":"2026-05-04","impressions":16418,"outbound_clicks":296,"add_to_cart":7,"initiate_checkout":6},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Pushpendra  | 07/05/2026","spend":2101.8,"sales":1489.0,"purchases":1,"roas":0.71,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":12856,"outbound_clicks":347,"add_to_cart":8,"initiate_checkout":3},{"product":"Lactify","name":"Lactify - Dr. Priyanka Deswal","spend":1921.48,"sales":1958.0,"purchases":2,"roas":1.02,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":19732,"outbound_clicks":375,"add_to_cart":19,"initiate_checkout":10},{"product":"Brainify Powder","name":"Brainify Powder - Brainify Powder Static - 15   | 07/05/2026","spend":1835.38,"sales":649.0,"purchases":1,"roas":0.35,"type":"Static","date":"2026-05-31","start_date":"2026-05-07","impressions":15215,"outbound_clicks":426,"add_to_cart":4,"initiate_checkout":6},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Ankit","spend":1759.25,"sales":4454.0,"purchases":6,"roas":2.53,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":15723,"outbound_clicks":252,"add_to_cart":23,"initiate_checkout":17},{"product":"Brainify Drops","name":"Drops - Static 8","spend":1679.09,"sales":3945.0,"purchases":5,"roas":2.35,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":14842,"outbound_clicks":297,"add_to_cart":19,"initiate_checkout":9},{"product":"Brainify Drops","name":"Brainify drops - Dr. Neha","spend":1599.95,"sales":3096.0,"purchases":4,"roas":1.94,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":10078,"outbound_clicks":262,"add_to_cart":12,"initiate_checkout":5},{"product":"Brainify Drops","name":"Brainify drops - Rina Arthaba","spend":1597.39,"sales":3855.0,"purchases":5,"roas":2.41,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":17053,"outbound_clicks":426,"add_to_cart":21,"initiate_checkout":17},{"product":"Brainify Drops","name":"Brainify - Dr. Ankit Jha  Video","spend":1553.58,"sales":2827.0,"purchases":3,"roas":1.82,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":12622,"outbound_clicks":404,"add_to_cart":13,"initiate_checkout":12},{"product":"Mamafy","name":"Mamafy_Dr. Tanya Video","spend":1551.6,"sales":3802.0,"purchases":4,"roas":2.45,"type":"Video","date":"2026-05-31","start_date":"2026-05-03","impressions":28352,"outbound_clicks":567,"add_to_cart":14,"initiate_checkout":9},{"product":"Brainify Drops","name":"Brainify - Static - 13 (Ingredients)","spend":1455.71,"sales":1288.0,"purchases":2,"roas":0.88,"type":"Static","date":"2026-05-31","start_date":"2026-05-15","impressions":16194,"outbound_clicks":518,"add_to_cart":7,"initiate_checkout":7},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Riya","spend":1382.04,"sales":6144.0,"purchases":6,"roas":4.45,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":10883,"outbound_clicks":196,"add_to_cart":9,"initiate_checkout":9},{"product":"Mamafy","name":"Mamafy -  Static - 16","spend":1302.76,"sales":709.0,"purchases":1,"roas":0.54,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":19375,"outbound_clicks":446,"add_to_cart":11,"initiate_checkout":7},{"product":"Brainify Drops","name":"Brainify - Divya Bajpai","spend":1295.67,"sales":2098.0,"purchases":2,"roas":1.62,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":5357,"outbound_clicks":145,"add_to_cart":9,"initiate_checkout":14},{"product":"Mamafy","name":"Mamafy -  Dispatch video - April","spend":1283.93,"sales":2876.0,"purchases":4,"roas":2.24,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":18200,"outbound_clicks":601,"add_to_cart":9,"initiate_checkout":8},{"product":"Brainify Drops","name":"Drops - Static 12","spend":1099.28,"sales":2875.0,"purchases":3,"roas":2.62,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":10094,"outbound_clicks":222,"add_to_cart":9,"initiate_checkout":6},{"product":"Brainify Powder","name":"Brainify Powder - Dr. Suryakamal","spend":1063.4,"sales":1748.0,"purchases":2,"roas":1.64,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":8401,"outbound_clicks":168,"add_to_cart":15,"initiate_checkout":7},{"product":"Brainify Drops","name":"Brainify drops - Maninder Kaur - April - 1 30/05/2026","spend":1052.01,"sales":569.0,"purchases":1,"roas":0.54,"type":"Static","date":"2026-05-31","start_date":"2026-05-30","impressions":9295,"outbound_clicks":177,"add_to_cart":12,"initiate_checkout":12},{"product":"Lactify","name":"Lactify - Dr. Srimukhi (Telugu)","spend":1047.0,"sales":649.0,"purchases":1,"roas":0.62,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":8512,"outbound_clicks":179,"add_to_cart":6,"initiate_checkout":1},{"product":"Brainify Drops","name":"Brainify drops - Dr. Pillai (Tamil) (Partnership ad) - Tamil","spend":1030.61,"sales":599.0,"purchases":1,"roas":0.58,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":9475,"outbound_clicks":265,"add_to_cart":4,"initiate_checkout":5},{"product":"Brainify Drops","name":"Brainify Drops -Maninder Kaur (Blessings) Video","spend":1013.79,"sales":3356.0,"purchases":4,"roas":3.31,"type":"Video","date":"2026-05-31","start_date":"2026-05-13","impressions":7650,"outbound_clicks":207,"add_to_cart":6,"initiate_checkout":6},{"product":"Brainify Drops","name":"Drops - Dr. Rajeswari (Telugu) 19/05/2026","spend":905.79,"sales":1228.0,"purchases":2,"roas":1.36,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":8070,"outbound_clicks":234,"add_to_cart":7,"initiate_checkout":4},{"product":"Brainify Drops","name":"Drops-Dr. Sandesh","spend":899.33,"sales":4314.0,"purchases":6,"roas":4.8,"type":"Static","date":"2026-05-31","start_date":"2026-05-30","impressions":7685,"outbound_clicks":138,"add_to_cart":14,"initiate_checkout":11},{"product":"Brainify Drops","name":"Brainify drops - Static - 7","spend":885.39,"sales":1998.0,"purchases":2,"roas":2.26,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":8012,"outbound_clicks":256,"add_to_cart":3,"initiate_checkout":2},{"product":"Brainify Drops","name":"Brainify drops - Static - 5","spend":834.67,"sales":1148.0,"purchases":2,"roas":1.38,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":6872,"outbound_clicks":227,"add_to_cart":5,"initiate_checkout":2},{"product":"Brainify Drops","name":"Brainify - Maninder Kaur","spend":781.44,"sales":619.0,"purchases":1,"roas":0.79,"type":"Static","date":"2026-05-31","start_date":"2026-05-01","impressions":6772,"outbound_clicks":223,"add_to_cart":2,"initiate_checkout":2},{"product":"Brainify Drops","name":"Drops - Dr. Padmavathi - kannada - 23/03/2026","spend":654.42,"sales":569.0,"purchases":1,"roas":0.87,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":4437,"outbound_clicks":102,"add_to_cart":6,"initiate_checkout":5},{"product":"Brainify Drops","name":"Brainify - Blessing Video","spend":585.75,"sales":4369.0,"purchases":5,"roas":7.46,"type":"Video","date":"2026-05-31","start_date":"2026-05-01","impressions":3644,"outbound_clicks":55,"add_to_cart":9,"initiate_checkout":7},{"product":"Mamafy","name":"Mamafy - Doctor Compilation Video - March - 1","spend":577.98,"sales":0.0,"purchases":0,"roas":0.0,"type":"Video","date":"2026-05-31","start_date":"2026-05-16","impressions":8004,"outbound_clicks":272,"add_to_cart":2,"initiate_checkout":2},{"product":"Mamafy","name":"Mamafy - static 2","spend":511.28,"sales":649.0,"purchases":1,"roas":1.27,"type":"Static","date":"2026-05-31","start_date":"2026-05-03","impressions":5709,"outbound_clicks":103,"add_to_cart":2,"initiate_checkout":1},{"product":"Brainify Drops","name":"Drops-Maninder Kaur (Blessings) Video","spend":465.82,"sales":619.0,"purchases":1,"roas":1.33,"type":"Video","date":"2026-05-31","start_date":"2026-05-02","impressions":3020,"outbound_clicks":88,"add_to_cart":1,"initiate_checkout":1},{"product":"Brainify Drops","name":"Drops - Doctor compilation video 2 | 19/04/2026","spend":439.92,"sales":0.0,"purchases":0,"roas":0.0,"type":"Video","date":"2026-05-31","start_date":"2026-04-19","impressions":3493,"outbound_clicks":84,"add_to_cart":1,"initiate_checkout":0},{"product":"Brainify Drops","name":"Drops - Rajmani","spend":428.08,"sales":619.0,"purchases":1,"roas":1.45,"type":"Static","date":"2026-05-31","start_date":"2026-05-23","impressions":2945,"outbound_clicks":71,"add_to_cart":5,"initiate_checkout":5},{"product":"Mamafy","name":"Mamafy - Static - 6 - USP","spend":413.01,"sales":1968.0,"purchases":2,"roas":4.77,"type":"Static","date":"2026-05-31","start_date":"2026-05-03","impressions":4866,"outbound_clicks":127,"add_to_cart":5,"initiate_checkout":4},{"product":"Brainify Powder","name":"Brainify Powder Static 11","spend":273.53,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-05-31","start_date":"2026-05-26","impressions":4032,"outbound_clicks":129,"add_to_cart":1,"initiate_checkout":1},{"product":"Mamafy","name":"Mamafy - Drishti 3rd Video","spend":244.83,"sales":0.0,"purchases":0,"roas":0.0,"type":"Video","date":"2026-05-31","start_date":"2026-05-15","impressions":2053,"outbound_clicks":33,"add_to_cart":0,"initiate_checkout":0},{"product":"Mamafy","name":"mamafy - Dr. Samra - Edit 2 (Direct)","spend":139.43,"sales":0.0,"purchases":0,"roas":0.0,"type":"Static","date":"2026-05-31","start_date":"2026-05-12","impressions":1624,"outbound_clicks":31,"add_to_cart":4,"initiate_checkout":1}];
const DEMO_GOOGLE = null;

const DEMO_DAILY_BLENDED = [{"date":"2026-05-01","spend":41699.08,"sales":119721.0,"purchases":140,"roas":2.87,"cpa":297.85},{"date":"2026-05-02","spend":36756.61,"sales":104162.0,"purchases":135,"roas":2.83,"cpa":272.27},{"date":"2026-05-03","spend":42434.81,"sales":127607.2,"purchases":150,"roas":3.01,"cpa":282.9},{"date":"2026-05-04","spend":48429.97,"sales":101234.18,"purchases":122,"roas":2.09,"cpa":396.97},{"date":"2026-05-05","spend":45702.66,"sales":144878.0,"purchases":177,"roas":3.17,"cpa":258.21},{"date":"2026-05-06","spend":45393.77,"sales":127915.4,"purchases":162,"roas":2.82,"cpa":280.21},{"date":"2026-05-07","spend":48453.26,"sales":149250.0,"purchases":184,"roas":3.08,"cpa":263.33},{"date":"2026-05-08","spend":50798.55,"sales":134546.0,"purchases":163,"roas":2.65,"cpa":311.65},{"date":"2026-05-09","spend":47565.09,"sales":129744.0,"purchases":164,"roas":2.73,"cpa":290.03},{"date":"2026-05-10","spend":53418.04,"sales":105984.0,"purchases":129,"roas":1.98,"cpa":414.09},{"date":"2026-05-11","spend":46935.94,"sales":120852.0,"purchases":149,"roas":2.57,"cpa":315.01},{"date":"2026-05-12","spend":43030.63,"sales":126310.0,"purchases":164,"roas":2.94,"cpa":262.38},{"date":"2026-05-13","spend":47867.89,"sales":119225.0,"purchases":156,"roas":2.49,"cpa":306.85},{"date":"2026-05-14","spend":44390.93,"sales":132099.0,"purchases":176,"roas":2.98,"cpa":252.22},{"date":"2026-05-15","spend":42432.98,"sales":103787.0,"purchases":128,"roas":2.45,"cpa":331.51},{"date":"2026-05-16","spend":33252.27,"sales":98164.0,"purchases":123,"roas":2.95,"cpa":270.34},{"date":"2026-05-17","spend":36874.72,"sales":84342.4,"purchases":106,"roas":2.29,"cpa":347.87},{"date":"2026-05-18","spend":35721.35,"sales":93375.0,"purchases":119,"roas":2.61,"cpa":300.18},{"date":"2026-05-19","spend":38548.98,"sales":101203.0,"purchases":126,"roas":2.63,"cpa":305.94},{"date":"2026-05-20","spend":36471.78,"sales":106799.0,"purchases":126,"roas":2.93,"cpa":289.46},{"date":"2026-05-21","spend":38065.49,"sales":102297.0,"purchases":127,"roas":2.69,"cpa":299.73},{"date":"2026-05-22","spend":39250.0,"sales":103475.0,"purchases":125,"roas":2.64,"cpa":314.0},{"date":"2026-05-23","spend":43846.44,"sales":94526.0,"purchases":123,"roas":2.16,"cpa":356.48},{"date":"2026-05-24","spend":44269.47,"sales":104594.0,"purchases":131,"roas":2.36,"cpa":337.93},{"date":"2026-05-25","spend":50130.44,"sales":115821.0,"purchases":146,"roas":2.31,"cpa":343.36},{"date":"2026-05-26","spend":54789.26,"sales":139844.0,"purchases":174,"roas":2.55,"cpa":314.88},{"date":"2026-05-27","spend":55760.21,"sales":117401.0,"purchases":144,"roas":2.11,"cpa":387.22},{"date":"2026-05-28","spend":53563.51,"sales":109961.0,"purchases":134,"roas":2.05,"cpa":399.73},{"date":"2026-05-29","spend":56724.49,"sales":147440.0,"purchases":184,"roas":2.6,"cpa":308.29},{"date":"2026-05-30","spend":55423.85,"sales":126238.0,"purchases":162,"roas":2.28,"cpa":342.12},{"date":"2026-05-31","spend":58470.68,"sales":124171.2,"purchases":162,"roas":2.12,"cpa":360.93}];
const DEMO_DAILY_BY_PRODUCT = [{"date":"2026-05-01","product":"Brainify Drops","spend":16909.66,"sales":51256.0,"purchases":61,"roas":3.03,"cpa":277.21},{"date":"2026-05-01","product":"Brainify Powder","spend":9489.98,"sales":41988.0,"purchases":48,"roas":4.42,"cpa":197.71},{"date":"2026-05-01","product":"Flowjoy","spend":3522.28,"sales":7551.0,"purchases":9,"roas":2.14,"cpa":391.36},{"date":"2026-05-01","product":"Lactify","spend":4451.18,"sales":5244.0,"purchases":6,"roas":1.18,"cpa":741.86},{"date":"2026-05-01","product":"Mamafy","spend":7325.98,"sales":13682.0,"purchases":16,"roas":1.87,"cpa":457.87},{"date":"2026-05-02","product":"Brainify Drops","spend":14300.31,"sales":47125.0,"purchases":60,"roas":3.3,"cpa":238.34},{"date":"2026-05-02","product":"Brainify Powder","spend":8612.38,"sales":24856.0,"purchases":33,"roas":2.89,"cpa":260.98},{"date":"2026-05-02","product":"Flowjoy","spend":2966.5,"sales":8183.0,"purchases":11,"roas":2.76,"cpa":269.68},{"date":"2026-05-02","product":"Lactify","spend":3671.87,"sales":5852.0,"purchases":7,"roas":1.59,"cpa":524.55},{"date":"2026-05-02","product":"Mamafy","spend":7205.55,"sales":18146.0,"purchases":24,"roas":2.52,"cpa":300.23},{"date":"2026-05-03","product":"Brainify Drops","spend":16626.9,"sales":53894.2,"purchases":66,"roas":3.24,"cpa":251.92},{"date":"2026-05-03","product":"Brainify Powder","spend":11690.8,"sales":42461.0,"purchases":47,"roas":3.63,"cpa":248.74},{"date":"2026-05-03","product":"Flowjoy","spend":3379.42,"sales":5353.0,"purchases":7,"roas":1.58,"cpa":482.77},{"date":"2026-05-03","product":"Lactify","spend":4287.98,"sales":9248.0,"purchases":11,"roas":2.16,"cpa":389.82},{"date":"2026-05-03","product":"Mamafy","spend":6449.71,"sales":16651.0,"purchases":19,"roas":2.58,"cpa":339.46},{"date":"2026-05-04","product":"Brainify Drops","spend":16037.0,"sales":33162.18,"purchases":43,"roas":2.07,"cpa":372.95},{"date":"2026-05-04","product":"Brainify Powder","spend":15481.79,"sales":37553.0,"purchases":44,"roas":2.43,"cpa":351.86},{"date":"2026-05-04","product":"Flowjoy","spend":3362.18,"sales":625.0,"purchases":1,"roas":0.19,"cpa":3362.18},{"date":"2026-05-04","product":"Lactify","spend":3744.26,"sales":12643.0,"purchases":16,"roas":3.38,"cpa":234.02},{"date":"2026-05-04","product":"Mamafy","spend":9804.74,"sales":17251.0,"purchases":18,"roas":1.76,"cpa":544.71},{"date":"2026-05-05","product":"Brainify Drops","spend":15970.53,"sales":49860.0,"purchases":65,"roas":3.12,"cpa":245.7},{"date":"2026-05-05","product":"Brainify Powder","spend":15149.67,"sales":55614.0,"purchases":61,"roas":3.67,"cpa":248.36},{"date":"2026-05-05","product":"Flowjoy","spend":3248.78,"sales":5030.0,"purchases":6,"roas":1.55,"cpa":541.46},{"date":"2026-05-05","product":"Lactify","spend":3320.09,"sales":6142.0,"purchases":8,"roas":1.85,"cpa":415.01},{"date":"2026-05-05","product":"Mamafy","spend":8013.59,"sales":28232.0,"purchases":37,"roas":3.52,"cpa":216.58},{"date":"2026-05-06","product":"Brainify Drops","spend":16049.76,"sales":46764.0,"purchases":64,"roas":2.91,"cpa":250.78},{"date":"2026-05-06","product":"Brainify Powder","spend":15026.17,"sales":45546.4,"purchases":54,"roas":3.03,"cpa":278.26},{"date":"2026-05-06","product":"Flowjoy","spend":3301.72,"sales":3541.0,"purchases":5,"roas":1.07,"cpa":660.34},{"date":"2026-05-06","product":"Lactify","spend":3106.65,"sales":10793.0,"purchases":13,"roas":3.47,"cpa":238.97},{"date":"2026-05-06","product":"Mamafy","spend":7909.47,"sales":21271.0,"purchases":26,"roas":2.69,"cpa":304.21},{"date":"2026-05-07","product":"Brainify Drops","spend":17273.79,"sales":62070.0,"purchases":81,"roas":3.59,"cpa":213.26},{"date":"2026-05-07","product":"Brainify Powder","spend":15742.28,"sales":40208.0,"purchases":49,"roas":2.55,"cpa":321.27},{"date":"2026-05-07","product":"Flowjoy","spend":3460.07,"sales":7879.0,"purchases":9,"roas":2.28,"cpa":384.45},{"date":"2026-05-07","product":"Lactify","spend":3199.01,"sales":10080.0,"purchases":10,"roas":3.15,"cpa":319.9},{"date":"2026-05-07","product":"Mamafy","spend":8778.11,"sales":29013.0,"purchases":35,"roas":3.31,"cpa":250.8},{"date":"2026-05-08","product":"Brainify Drops","spend":16412.69,"sales":54466.0,"purchases":66,"roas":3.32,"cpa":248.68},{"date":"2026-05-08","product":"Brainify Powder","spend":20031.8,"sales":41113.0,"purchases":50,"roas":2.05,"cpa":400.64},{"date":"2026-05-08","product":"Flowjoy","spend":3247.53,"sales":4630.0,"purchases":5,"roas":1.43,"cpa":649.51},{"date":"2026-05-08","product":"Lactify","spend":3253.03,"sales":11087.0,"purchases":12,"roas":3.41,"cpa":271.09},{"date":"2026-05-08","product":"Mamafy","spend":7853.5,"sales":23250.0,"purchases":30,"roas":2.96,"cpa":261.78},{"date":"2026-05-09","product":"Brainify Drops","spend":15109.17,"sales":34277.0,"purchases":47,"roas":2.27,"cpa":321.47},{"date":"2026-05-09","product":"Brainify Powder","spend":19986.62,"sales":48086.0,"purchases":59,"roas":2.41,"cpa":338.76},{"date":"2026-05-09","product":"Flowjoy","spend":2957.8,"sales":10768.0,"purchases":14,"roas":3.64,"cpa":211.27},{"date":"2026-05-09","product":"Lactify","spend":2795.92,"sales":7461.0,"purchases":8,"roas":2.67,"cpa":349.49},{"date":"2026-05-09","product":"Mamafy","spend":6715.58,"sales":29152.0,"purchases":36,"roas":4.34,"cpa":186.54},{"date":"2026-05-10","product":"Brainify Drops","spend":17820.69,"sales":36730.0,"purchases":49,"roas":2.06,"cpa":363.69},{"date":"2026-05-10","product":"Brainify Powder","spend":20469.72,"sales":42189.0,"purchases":50,"roas":2.06,"cpa":409.39},{"date":"2026-05-10","product":"Flowjoy","spend":3542.58,"sales":4546.0,"purchases":6,"roas":1.28,"cpa":590.43},{"date":"2026-05-10","product":"Lactify","spend":3404.12,"sales":7898.0,"purchases":8,"roas":2.32,"cpa":425.51},{"date":"2026-05-10","product":"Mamafy","spend":8180.93,"sales":14621.0,"purchases":16,"roas":1.79,"cpa":511.31},{"date":"2026-05-11","product":"Brainify Drops","spend":17120.47,"sales":45797.0,"purchases":60,"roas":2.67,"cpa":285.34},{"date":"2026-05-11","product":"Brainify Powder","spend":14897.91,"sales":49745.0,"purchases":55,"roas":3.34,"cpa":270.87},{"date":"2026-05-11","product":"Flowjoy","spend":3613.56,"sales":4063.0,"purchases":7,"roas":1.12,"cpa":516.22},{"date":"2026-05-11","product":"Lactify","spend":3120.88,"sales":3854.0,"purchases":5,"roas":1.23,"cpa":624.18},{"date":"2026-05-11","product":"Mamafy","spend":8183.12,"sales":17393.0,"purchases":22,"roas":2.13,"cpa":371.96},{"date":"2026-05-12","product":"Brainify Drops","spend":17193.61,"sales":46583.0,"purchases":63,"roas":2.71,"cpa":272.91},{"date":"2026-05-12","product":"Brainify Powder","spend":12321.26,"sales":41189.0,"purchases":54,"roas":3.34,"cpa":228.17},{"date":"2026-05-12","product":"Flowjoy","spend":2469.0,"sales":3331.0,"purchases":5,"roas":1.35,"cpa":493.8},{"date":"2026-05-12","product":"Lactify","spend":3217.75,"sales":10307.0,"purchases":12,"roas":3.2,"cpa":268.15},{"date":"2026-05-12","product":"Mamafy","spend":7829.01,"sales":24900.0,"purchases":30,"roas":3.18,"cpa":260.97},{"date":"2026-05-13","product":"Brainify Drops","spend":15941.3,"sales":43108.0,"purchases":56,"roas":2.7,"cpa":284.67},{"date":"2026-05-13","product":"Brainify Powder","spend":13304.62,"sales":38431.0,"purchases":49,"roas":2.89,"cpa":271.52},{"date":"2026-05-13","product":"Flowjoy","spend":3972.43,"sales":12940.0,"purchases":20,"roas":3.26,"cpa":198.62},{"date":"2026-05-13","product":"Lactify","spend":5327.93,"sales":6651.0,"purchases":9,"roas":1.25,"cpa":591.99},{"date":"2026-05-13","product":"Mamafy","spend":9321.61,"sales":18095.0,"purchases":22,"roas":1.94,"cpa":423.71},{"date":"2026-05-14","product":"Brainify Drops","spend":14259.64,"sales":36273.0,"purchases":51,"roas":2.54,"cpa":279.6},{"date":"2026-05-14","product":"Brainify Powder","spend":13205.97,"sales":48201.0,"purchases":64,"roas":3.65,"cpa":206.34},{"date":"2026-05-14","product":"Flowjoy","spend":4147.24,"sales":8454.0,"purchases":12,"roas":2.04,"cpa":345.6},{"date":"2026-05-14","product":"Lactify","spend":4632.45,"sales":5363.0,"purchases":7,"roas":1.16,"cpa":661.78},{"date":"2026-05-14","product":"Mamafy","spend":8145.63,"sales":33808.0,"purchases":42,"roas":4.15,"cpa":193.94},{"date":"2026-05-15","product":"Brainify Drops","spend":14933.93,"sales":25908.0,"purchases":32,"roas":1.73,"cpa":466.69},{"date":"2026-05-15","product":"Brainify Powder","spend":11134.2,"sales":40870.0,"purchases":52,"roas":3.67,"cpa":214.12},{"date":"2026-05-15","product":"Flowjoy","spend":3893.22,"sales":6028.0,"purchases":6,"roas":1.55,"cpa":648.87},{"date":"2026-05-15","product":"Lactify","spend":4368.42,"sales":11545.0,"purchases":14,"roas":2.64,"cpa":312.03},{"date":"2026-05-15","product":"Mamafy","spend":8103.21,"sales":19436.0,"purchases":24,"roas":2.4,"cpa":337.63},{"date":"2026-05-16","product":"Brainify Drops","spend":11193.23,"sales":25553.0,"purchases":34,"roas":2.28,"cpa":329.21},{"date":"2026-05-16","product":"Brainify Powder","spend":9551.08,"sales":31910.0,"purchases":40,"roas":3.34,"cpa":238.78},{"date":"2026-05-16","product":"Flowjoy","spend":2195.72,"sales":5836.0,"purchases":6,"roas":2.66,"cpa":365.95},{"date":"2026-05-16","product":"Lactify","spend":3649.15,"sales":13344.0,"purchases":16,"roas":3.66,"cpa":228.07},{"date":"2026-05-16","product":"Mamafy","spend":6663.09,"sales":21521.0,"purchases":27,"roas":3.23,"cpa":246.78},{"date":"2026-05-17","product":"Brainify Drops","spend":11344.38,"sales":25076.0,"purchases":28,"roas":2.21,"cpa":405.16},{"date":"2026-05-17","product":"Brainify Powder","spend":11229.49,"sales":23687.0,"purchases":33,"roas":2.11,"cpa":340.29},{"date":"2026-05-17","product":"Flowjoy","spend":3450.72,"sales":9234.4,"purchases":12,"roas":2.68,"cpa":287.56},{"date":"2026-05-17","product":"Lactify","spend":3974.52,"sales":13837.0,"purchases":16,"roas":3.48,"cpa":248.41},{"date":"2026-05-17","product":"Mamafy","spend":6875.61,"sales":12508.0,"purchases":17,"roas":1.82,"cpa":404.45},{"date":"2026-05-18","product":"Brainify Drops","spend":11417.5,"sales":32823.0,"purchases":43,"roas":2.87,"cpa":265.52},{"date":"2026-05-18","product":"Brainify Powder","spend":10620.06,"sales":33529.0,"purchases":41,"roas":3.16,"cpa":259.03},{"date":"2026-05-18","product":"Flowjoy","spend":3652.84,"sales":8031.0,"purchases":12,"roas":2.2,"cpa":304.4},{"date":"2026-05-18","product":"Lactify","spend":4025.49,"sales":5703.0,"purchases":7,"roas":1.42,"cpa":575.07},{"date":"2026-05-18","product":"Mamafy","spend":6005.46,"sales":13289.0,"purchases":16,"roas":2.21,"cpa":375.34},{"date":"2026-05-19","product":"Brainify Drops","spend":11474.37,"sales":25133.0,"purchases":32,"roas":2.19,"cpa":358.57},{"date":"2026-05-19","product":"Brainify Powder","spend":15993.13,"sales":37146.0,"purchases":48,"roas":2.32,"cpa":333.19},{"date":"2026-05-19","product":"Flowjoy","spend":2165.93,"sales":5750.0,"purchases":7,"roas":2.65,"cpa":309.42},{"date":"2026-05-19","product":"Lactify","spend":3506.62,"sales":13921.0,"purchases":17,"roas":3.97,"cpa":206.27},{"date":"2026-05-19","product":"Mamafy","spend":5408.93,"sales":19253.0,"purchases":22,"roas":3.56,"cpa":245.86},{"date":"2026-05-20","product":"Brainify Drops","spend":10876.67,"sales":35197.0,"purchases":43,"roas":3.24,"cpa":252.95},{"date":"2026-05-20","product":"Brainify Powder","spend":15118.71,"sales":41206.0,"purchases":48,"roas":2.73,"cpa":314.97},{"date":"2026-05-20","product":"Flowjoy","spend":2124.61,"sales":4985.0,"purchases":5,"roas":2.35,"cpa":424.92},{"date":"2026-05-20","product":"Lactify","spend":3437.06,"sales":9314.0,"purchases":11,"roas":2.71,"cpa":312.46},{"date":"2026-05-20","product":"Mamafy","spend":4914.73,"sales":16097.0,"purchases":19,"roas":3.28,"cpa":258.67},{"date":"2026-05-21","product":"Brainify Drops","spend":10402.19,"sales":26071.0,"purchases":35,"roas":2.51,"cpa":297.21},{"date":"2026-05-21","product":"Brainify Powder","spend":15382.09,"sales":43217.0,"purchases":53,"roas":2.81,"cpa":290.23},{"date":"2026-05-21","product":"Flowjoy","spend":1875.24,"sales":5762.0,"purchases":9,"roas":3.07,"cpa":208.36},{"date":"2026-05-21","product":"Lactify","spend":5881.37,"sales":13267.0,"purchases":14,"roas":2.26,"cpa":420.1},{"date":"2026-05-21","product":"Mamafy","spend":4524.6,"sales":13980.0,"purchases":16,"roas":3.09,"cpa":282.79},{"date":"2026-05-22","product":"Brainify Drops","spend":9383.37,"sales":19968.0,"purchases":27,"roas":2.13,"cpa":347.53},{"date":"2026-05-22","product":"Brainify Powder","spend":16325.38,"sales":51916.0,"purchases":64,"roas":3.18,"cpa":255.08},{"date":"2026-05-22","product":"Flowjoy","spend":1961.44,"sales":1759.0,"purchases":3,"roas":0.9,"cpa":653.81},{"date":"2026-05-22","product":"Lactify","spend":6021.93,"sales":14942.0,"purchases":15,"roas":2.48,"cpa":401.46},{"date":"2026-05-22","product":"Mamafy","spend":5557.88,"sales":14890.0,"purchases":16,"roas":2.68,"cpa":347.37},{"date":"2026-05-23","product":"Brainify Drops","spend":10130.58,"sales":24331.0,"purchases":32,"roas":2.4,"cpa":316.58},{"date":"2026-05-23","product":"Brainify Powder","spend":17085.53,"sales":42383.0,"purchases":53,"roas":2.48,"cpa":322.37},{"date":"2026-05-23","product":"Flowjoy","spend":2934.73,"sales":6012.0,"purchases":10,"roas":2.05,"cpa":293.47},{"date":"2026-05-23","product":"Lactify","spend":5315.76,"sales":5728.0,"purchases":8,"roas":1.08,"cpa":664.47},{"date":"2026-05-23","product":"Mamafy","spend":8379.84,"sales":16072.0,"purchases":20,"roas":1.92,"cpa":418.99},{"date":"2026-05-24","product":"Brainify Drops","spend":10827.57,"sales":22315.0,"purchases":33,"roas":2.06,"cpa":328.11},{"date":"2026-05-24","product":"Brainify Powder","spend":15261.17,"sales":41538.0,"purchases":52,"roas":2.72,"cpa":293.48},{"date":"2026-05-24","product":"Flowjoy","spend":3683.42,"sales":9193.0,"purchases":11,"roas":2.5,"cpa":334.86},{"date":"2026-05-24","product":"Lactify","spend":5447.14,"sales":13702.0,"purchases":16,"roas":2.52,"cpa":340.45},{"date":"2026-05-24","product":"Mamafy","spend":9050.17,"sales":17846.0,"purchases":19,"roas":1.97,"cpa":476.32},{"date":"2026-05-25","product":"Brainify Drops","spend":18692.68,"sales":40266.0,"purchases":53,"roas":2.15,"cpa":352.69},{"date":"2026-05-25","product":"Brainify Powder","spend":14410.28,"sales":35171.0,"purchases":45,"roas":2.44,"cpa":320.23},{"date":"2026-05-25","product":"Flowjoy","spend":3416.01,"sales":4806.0,"purchases":8,"roas":1.41,"cpa":427.0},{"date":"2026-05-25","product":"Lactify","spend":5183.27,"sales":18470.0,"purchases":20,"roas":3.56,"cpa":259.16},{"date":"2026-05-25","product":"Mamafy","spend":8428.2,"sales":17108.0,"purchases":20,"roas":2.03,"cpa":421.41},{"date":"2026-05-26","product":"Brainify Drops","spend":21024.28,"sales":47938.0,"purchases":61,"roas":2.28,"cpa":344.66},{"date":"2026-05-26","product":"Brainify Powder","spend":15631.41,"sales":42656.0,"purchases":55,"roas":2.73,"cpa":284.21},{"date":"2026-05-26","product":"Flowjoy","spend":3980.56,"sales":8073.0,"purchases":12,"roas":2.03,"cpa":331.71},{"date":"2026-05-26","product":"Lactify","spend":5616.07,"sales":17739.0,"purchases":19,"roas":3.16,"cpa":295.58},{"date":"2026-05-26","product":"Mamafy","spend":8536.94,"sales":23438.0,"purchases":27,"roas":2.75,"cpa":316.18},{"date":"2026-05-27","product":"Brainify Drops","spend":19225.43,"sales":33643.0,"purchases":41,"roas":1.75,"cpa":468.91},{"date":"2026-05-27","product":"Brainify Powder","spend":16157.83,"sales":45733.0,"purchases":58,"roas":2.83,"cpa":278.58},{"date":"2026-05-27","product":"Flowjoy","spend":5327.75,"sales":4708.0,"purchases":6,"roas":0.88,"cpa":887.96},{"date":"2026-05-27","product":"Lactify","spend":5761.48,"sales":12354.0,"purchases":15,"roas":2.14,"cpa":384.1},{"date":"2026-05-27","product":"Mamafy","spend":9287.72,"sales":20963.0,"purchases":24,"roas":2.26,"cpa":386.99},{"date":"2026-05-28","product":"Brainify Drops","spend":16759.59,"sales":33741.0,"purchases":46,"roas":2.01,"cpa":364.34},{"date":"2026-05-28","product":"Brainify Powder","spend":16592.74,"sales":40916.0,"purchases":48,"roas":2.47,"cpa":345.68},{"date":"2026-05-28","product":"Flowjoy","spend":4784.65,"sales":3407.0,"purchases":5,"roas":0.71,"cpa":956.93},{"date":"2026-05-28","product":"Lactify","spend":5975.37,"sales":9954.0,"purchases":13,"roas":1.67,"cpa":459.64},{"date":"2026-05-28","product":"Mamafy","spend":9451.16,"sales":21943.0,"purchases":22,"roas":2.32,"cpa":429.6},{"date":"2026-05-29","product":"Brainify Drops","spend":18200.96,"sales":52160.0,"purchases":66,"roas":2.87,"cpa":275.77},{"date":"2026-05-29","product":"Brainify Powder","spend":17736.78,"sales":47351.0,"purchases":57,"roas":2.67,"cpa":311.17},{"date":"2026-05-29","product":"Flowjoy","spend":4703.49,"sales":10134.0,"purchases":16,"roas":2.15,"cpa":293.97},{"date":"2026-05-29","product":"Lactify","spend":7537.82,"sales":18220.0,"purchases":20,"roas":2.42,"cpa":376.89},{"date":"2026-05-29","product":"Mamafy","spend":8545.44,"sales":19575.0,"purchases":25,"roas":2.29,"cpa":341.82},{"date":"2026-05-30","product":"Brainify Drops","spend":18525.55,"sales":44285.0,"purchases":56,"roas":2.39,"cpa":330.81},{"date":"2026-05-30","product":"Brainify Powder","spend":16308.39,"sales":42471.0,"purchases":54,"roas":2.6,"cpa":302.01},{"date":"2026-05-30","product":"Flowjoy","spend":4696.62,"sales":6463.0,"purchases":11,"roas":1.38,"cpa":426.97},{"date":"2026-05-30","product":"Lactify","spend":6764.15,"sales":12498.0,"purchases":16,"roas":1.85,"cpa":422.76},{"date":"2026-05-30","product":"Mamafy","spend":9129.14,"sales":20521.0,"purchases":25,"roas":2.25,"cpa":365.17},{"date":"2026-05-31","product":"Brainify Drops","spend":17270.17,"sales":40263.0,"purchases":55,"roas":2.33,"cpa":314.0},{"date":"2026-05-31","product":"Brainify Powder","spend":19435.76,"sales":45080.2,"purchases":56,"roas":2.32,"cpa":347.07},{"date":"2026-05-31","product":"Flowjoy","spend":4728.82,"sales":4527.0,"purchases":7,"roas":0.96,"cpa":675.55},{"date":"2026-05-31","product":"Lactify","spend":8195.15,"sales":12150.0,"purchases":16,"roas":1.48,"cpa":512.2},{"date":"2026-05-31","product":"Mamafy","spend":8840.78,"sales":22151.0,"purchases":28,"roas":2.51,"cpa":315.74}];
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
          Load Demo Data (May 1–31, 2026)
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
    // Live mode: use dateFilteredAds directly
    if (!isDemo || !demoByProduct) return dateFilteredAds || metaAds;
    // Demo mode, no date filter: return as-is
    if (!dateFrom && !dateTo) return metaAds;
    // Demo mode + date filter: compute per-product scale factor
    const productScale = {};
    const products = [...new Set(demoByProduct.map(r => r.product))];
    for (const product of products) {
      const all = demoByProduct.filter(r => r.product === product);
      const inRange = all.filter(r => (!dateFrom || r.date >= dateFrom) && (!dateTo || r.date <= dateTo));
      const totalSpend = all.reduce((s, r) => s + r.spend, 0);
      const rangeSpend = inRange.reduce((s, r) => s + r.spend, 0);
      productScale[product] = totalSpend > 0 ? rangeSpend / totalSpend : 0;
    }
    // Fallback scale for any unmapped product
    const allTotal = demoByProduct.reduce((s, r) => s + r.spend, 0);
    const allRange = demoByProduct.filter(r => (!dateFrom || r.date >= dateFrom) && (!dateTo || r.date <= dateTo)).reduce((s, r) => s + r.spend, 0);
    const overallScale = allTotal > 0 ? allRange / allTotal : 0;
    return metaAds.map(ad => {
      const scale = productScale[ad.product] ?? overallScale;
      return { ...ad, spend: ad.spend * scale, sales: ad.sales * scale, purchases: Math.round(ad.purchases * scale), add_to_cart: Math.round((ad.add_to_cart || 0) * scale), initiate_checkout: Math.round((ad.initiate_checkout || 0) * scale), impressions: Math.round((ad.impressions || 0) * scale), outbound_clicks: Math.round((ad.outbound_clicks || 0) * scale), roas: ad.spend > 0 ? (ad.sales * scale) / (ad.spend * scale) : 0 };
    });
  }, [isDemo, metaAds, demoByProduct, dateFrom, dateTo, dateFilteredAds]);

  const filteredAds = useMemo(() => {
    const base = scaledDemoAds || metaAds;
    if (!base) return [];
    let ads = (selectedProduct !== "All Products" ? base.filter((a) => a.product === selectedProduct) : base).filter(a => !a.name.toLowerCase().includes('post id') && !a.name.toLowerCase().includes('mothers day'));
    if (searchQuery) { const q = searchQuery.toLowerCase(); ads = ads.filter((a) => a.name.toLowerCase().includes(q) || a.product.toLowerCase().includes(q)); }
    if (adTypeFilter !== "All") ads = ads.filter((a) => a.type === adTypeFilter);
    if (creatorFilter !== "All") ads = ads.filter((a) => extractCreator(a.name) === creatorFilter);
    return [...ads].sort((a, b) => sortDir === "desc" ? b[sortField] - a[sortField] : a[sortField] - b[sortField]);
  }, [scaledDemoAds, metaAds, selectedProduct, searchQuery, sortField, sortDir, adTypeFilter, creatorFilter]);

  const maxSpend = useMemo(() => filteredAds.reduce((m, a) => Math.max(m, a.spend), 0), [filteredAds]);
  const uniqueCreators = useMemo(() => { const base = scaledDemoAds || metaAds; if (!base) return []; const s = new Set(base.map(a => extractCreator(a.name))); return ["All", ...Array.from(s).sort()]; }, [scaledDemoAds, metaAds]);
  const creatorTotals = useMemo(() => { const base = scaledDemoAds || metaAds; if (!base) return []; const pool = (creatorProductFilter !== "All Products" ? base.filter(a => a.product === creatorProductFilter) : base).filter(a => !a.name.toLowerCase().includes('post id') && !a.name.toLowerCase().includes('mothers day')); return computeCreatorTotals(pool); }, [scaledDemoAds, metaAds, creatorProductFilter]);
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
              ? "Demo data · May 1–31, 2026 · Meta + Google Ads"
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
