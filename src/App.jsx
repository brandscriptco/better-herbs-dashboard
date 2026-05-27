import { useState, useMemo, useCallback, useRef } from "react";
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
const DEMO_ADS = [{"product":"Other","name":"ASC Drops and Powder Full Funnel","spend":36767.12,"sales":78211,"purchases":98,"roas":2.13,"type":"Campaign","date":"2026-05-01"},{"product":"Brainify Drops","name":"Brainify Drops_Sandbox_ABO","spend":288957.75,"sales":764354.38,"purchases":994,"roas":2.65,"type":"Campaign","date":"2026-05-01"},{"product":"Other","name":"Brainify drops and powder retargeting","spend":3508.57,"sales":1857,"purchases":3,"roas":0.53,"type":"Campaign","date":"2026-05-01"},{"product":"Brainify Powder","name":"Brainify_Powder_Sandbox_ABO","spend":348925.11,"sales":996658.4,"purchases":1230,"roas":2.86,"type":"Campaign","date":"2026-05-01"},{"product":"Brainify Drops","name":"Branify Drops CBO - 27.05.2026","spend":96.01,"sales":0,"purchases":0,"roas":0.0,"type":"Campaign","date":"2026-05-01"},{"product":"Flowjoy","name":"Flow Joy Drops_Sandbox_ABO","spend":85484.73,"sales":166476.4,"purchases":223,"roas":1.95,"type":"Campaign","date":"2026-05-01"},{"product":"Brainify Drops","name":"Full Funnel - Brainify Drops - CBO","spend":33550.37,"sales":79643,"purchases":109,"roas":2.37,"type":"Campaign","date":"2026-05-01"},{"product":"Brainify Powder","name":"Full Funnel - Brainify Powder - CBO","spend":25886.72,"sales":71943,"purchases":90,"roas":2.78,"type":"Campaign","date":"2026-05-01"},{"product":"Lactify","name":"Lactify_Sandbox_ABO","spend":106154.91,"sales":258818,"purchases":300,"roas":2.44,"type":"Campaign","date":"2026-05-01"},{"product":"Lactify","name":"Lactify_Sandbox_Creative_testing_CBO","spend":0,"sales":0,"purchases":0,"roas":0,"type":"Campaign","date":"2026-05-01"},{"product":"Mamafy","name":"Mamafy Costcap CBO","spend":33402.29,"sales":95555,"purchases":120,"roas":2.86,"type":"Campaign","date":"2026-05-01"},{"product":"Mamafy","name":"Mamafy Sandbox ABO","spend":62702.35,"sales":136513,"purchases":165,"roas":2.18,"type":"Campaign","date":"2026-05-01"},{"product":"Mamafy","name":"Mamafy_Sandbox_ABO","spend":85302.18,"sales":225227,"purchases":275,"roas":2.64,"type":"Campaign","date":"2026-05-01"},{"product":"Other","name":"Mother's Day Sale - Drops & Powder","spend":30328.19,"sales":93668,"purchases":101,"roas":3.09,"type":"Campaign","date":"2026-05-01"},{"product":"Other","name":"Mother's Day Sale - Lactify & Mamafy","spend":13880.46,"sales":50660,"purchases":52,"roas":3.65,"type":"Campaign","date":"2026-05-01"},{"product":"Other","name":"Retargeting all products | 12/05/2026","spend":2416.48,"sales":4874,"purchases":5,"roas":2.02,"type":"Campaign","date":"2026-05-01"},{"product":"Mamafy","name":"mamafy - Retargeting","spend":6814.54,"sales":19889,"purchases":24,"roas":2.92,"type":"Campaign","date":"2026-05-01"}];
const DEMO_GOOGLE = null;

const DEMO_DAILY_BLENDED = [{"date":"2026-05-01","spend":41699.08,"sales":119072,"purchases":139,"roas":2.86,"cpa":299.99},{"date":"2026-05-02","spend":36756.61,"sales":104162,"purchases":135,"roas":2.83,"cpa":272.27},{"date":"2026-05-03","spend":42434.81,"sales":125120.2,"purchases":148,"roas":2.95,"cpa":286.72},{"date":"2026-05-04","spend":48429.97,"sales":101234.18,"purchases":122,"roas":2.09,"cpa":396.97},{"date":"2026-05-05","spend":45702.66,"sales":145427,"purchases":178,"roas":3.18,"cpa":256.76},{"date":"2026-05-06","spend":45393.77,"sales":127915.4,"purchases":162,"roas":2.82,"cpa":280.21},{"date":"2026-05-07","spend":48453.26,"sales":149250,"purchases":184,"roas":3.08,"cpa":263.33},{"date":"2026-05-08","spend":50798.55,"sales":134546,"purchases":163,"roas":2.65,"cpa":311.65},{"date":"2026-05-09","spend":47565.09,"sales":129744,"purchases":164,"roas":2.73,"cpa":290.03},{"date":"2026-05-10","spend":53418.04,"sales":105984,"purchases":129,"roas":1.98,"cpa":414.09},{"date":"2026-05-11","spend":46935.94,"sales":120852,"purchases":149,"roas":2.57,"cpa":315.01},{"date":"2026-05-12","spend":43030.63,"sales":126310,"purchases":164,"roas":2.94,"cpa":262.38},{"date":"2026-05-13","spend":47867.89,"sales":119225,"purchases":156,"roas":2.49,"cpa":306.85},{"date":"2026-05-14","spend":44390.93,"sales":132099,"purchases":176,"roas":2.98,"cpa":252.22},{"date":"2026-05-15","spend":42432.98,"sales":103787,"purchases":128,"roas":2.45,"cpa":331.51},{"date":"2026-05-16","spend":33252.27,"sales":98164,"purchases":123,"roas":2.95,"cpa":270.34},{"date":"2026-05-17","spend":36874.72,"sales":84342.4,"purchases":106,"roas":2.29,"cpa":347.87},{"date":"2026-05-18","spend":35721.35,"sales":93375,"purchases":119,"roas":2.61,"cpa":300.18},{"date":"2026-05-19","spend":38548.98,"sales":101203,"purchases":126,"roas":2.63,"cpa":305.94},{"date":"2026-05-20","spend":36471.78,"sales":106799,"purchases":126,"roas":2.93,"cpa":289.46},{"date":"2026-05-21","spend":38065.49,"sales":102297,"purchases":127,"roas":2.69,"cpa":299.73},{"date":"2026-05-22","spend":39250.0,"sales":103475,"purchases":125,"roas":2.64,"cpa":314.0},{"date":"2026-05-23","spend":43846.44,"sales":94526,"purchases":123,"roas":2.16,"cpa":356.48},{"date":"2026-05-24","spend":44269.47,"sales":104594,"purchases":131,"roas":2.36,"cpa":337.93},{"date":"2026-05-25","spend":50125.8,"sales":115821,"purchases":146,"roas":2.31,"cpa":343.33},{"date":"2026-05-26","spend":54737.67,"sales":139844,"purchases":174,"roas":2.55,"cpa":314.58},{"date":"2026-05-27","spend":27703.6,"sales":55179,"purchases":66,"roas":1.99,"cpa":419.75}];
const DEMO_DAILY_BY_PRODUCT = [{"date":"2026-05-01","product":"Brainify Drops","spend":15035.32,"sales":44274,"purchases":54,"roas":2.94,"cpa":278.43},{"date":"2026-05-01","product":"Brainify Powder","spend":9791.84,"sales":47372,"purchases":53,"roas":4.84,"cpa":184.75},{"date":"2026-05-01","product":"Flowjoy","spend":3522.28,"sales":7551,"purchases":9,"roas":2.14,"cpa":391.36},{"date":"2026-05-01","product":"Lactify","spend":3727.05,"sales":5244,"purchases":6,"roas":1.41,"cpa":621.18},{"date":"2026-05-01","product":"Mamafy","spend":6711.49,"sales":12423,"purchases":15,"roas":1.85,"cpa":447.43},{"date":"2026-05-01","product":"Other","spend":2911.1,"sales":2208,"purchases":2,"roas":0.76,"cpa":1455.55},{"date":"2026-05-02","product":"Brainify Drops","spend":12678.71,"sales":34910,"purchases":47,"roas":2.75,"cpa":269.76},{"date":"2026-05-02","product":"Brainify Powder","spend":8679.03,"sales":29770,"purchases":39,"roas":3.43,"cpa":222.54},{"date":"2026-05-02","product":"Flowjoy","spend":2966.5,"sales":8183,"purchases":11,"roas":2.76,"cpa":269.68},{"date":"2026-05-02","product":"Lactify","spend":3133.05,"sales":5852,"purchases":7,"roas":1.87,"cpa":447.58},{"date":"2026-05-02","product":"Mamafy","spend":6496.34,"sales":18146,"purchases":24,"roas":2.79,"cpa":270.68},{"date":"2026-05-02","product":"Other","spend":2802.98,"sales":7301,"purchases":7,"roas":2.6,"cpa":400.43},{"date":"2026-05-03","product":"Brainify Drops","spend":12952.08,"sales":47142.2,"purchases":58,"roas":3.64,"cpa":223.31},{"date":"2026-05-03","product":"Brainify Powder","spend":12125.39,"sales":36936,"purchases":43,"roas":3.05,"cpa":281.99},{"date":"2026-05-03","product":"Flowjoy","spend":3379.42,"sales":5353,"purchases":7,"roas":1.58,"cpa":482.77},{"date":"2026-05-03","product":"Lactify","spend":3971.89,"sales":9248,"purchases":11,"roas":2.33,"cpa":361.08},{"date":"2026-05-03","product":"Mamafy","spend":5921.51,"sales":10677,"purchases":13,"roas":1.8,"cpa":455.5},{"date":"2026-05-03","product":"Other","spend":4084.52,"sales":15764,"purchases":16,"roas":3.86,"cpa":255.28},{"date":"2026-05-04","product":"Brainify Drops","spend":13095.75,"sales":28948.18,"purchases":37,"roas":2.21,"cpa":353.94},{"date":"2026-05-04","product":"Brainify Powder","spend":15648.05,"sales":35723,"purchases":44,"roas":2.28,"cpa":355.64},{"date":"2026-05-04","product":"Flowjoy","spend":3362.18,"sales":625,"purchases":1,"roas":0.19,"cpa":3362.18},{"date":"2026-05-04","product":"Lactify","spend":3382.8,"sales":10185,"purchases":14,"roas":3.01,"cpa":241.63},{"date":"2026-05-04","product":"Mamafy","spend":8997.79,"sales":14155,"purchases":15,"roas":1.57,"cpa":599.85},{"date":"2026-05-04","product":"Other","spend":3943.4,"sales":11598,"purchases":11,"roas":2.94,"cpa":358.49},{"date":"2026-05-05","product":"Brainify Drops","spend":12968.5,"sales":40353,"purchases":53,"roas":3.11,"cpa":244.69},{"date":"2026-05-05","product":"Brainify Powder","spend":15665.4,"sales":56652,"purchases":62,"roas":3.62,"cpa":252.67},{"date":"2026-05-05","product":"Flowjoy","spend":3248.78,"sales":5030,"purchases":6,"roas":1.55,"cpa":541.46},{"date":"2026-05-05","product":"Lactify","spend":2997.14,"sales":5493,"purchases":7,"roas":1.83,"cpa":428.16},{"date":"2026-05-05","product":"Mamafy","spend":7148.03,"sales":25626,"purchases":34,"roas":3.59,"cpa":210.24},{"date":"2026-05-05","product":"Other","spend":3674.81,"sales":12273,"purchases":16,"roas":3.34,"cpa":229.68},{"date":"2026-05-06","product":"Brainify Drops","spend":13126.15,"sales":36347,"purchases":52,"roas":2.77,"cpa":252.43},{"date":"2026-05-06","product":"Brainify Powder","spend":15417.08,"sales":51119.4,"purchases":61,"roas":3.32,"cpa":252.74},{"date":"2026-05-06","product":"Flowjoy","spend":3301.72,"sales":3541,"purchases":5,"roas":1.07,"cpa":660.34},{"date":"2026-05-06","product":"Lactify","spend":2805.46,"sales":8945,"purchases":11,"roas":3.19,"cpa":255.04},{"date":"2026-05-06","product":"Mamafy","spend":7107.41,"sales":17015,"purchases":22,"roas":2.39,"cpa":323.06},{"date":"2026-05-06","product":"Other","spend":3635.95,"sales":10948,"purchases":11,"roas":3.01,"cpa":330.54},{"date":"2026-05-07","product":"Brainify Drops","spend":14458.12,"sales":45751,"purchases":62,"roas":3.16,"cpa":233.2},{"date":"2026-05-07","product":"Brainify Powder","spend":16058.44,"sales":43184,"purchases":53,"roas":2.69,"cpa":302.99},{"date":"2026-05-07","product":"Flowjoy","spend":3460.07,"sales":7879,"purchases":9,"roas":2.28,"cpa":384.45},{"date":"2026-05-07","product":"Lactify","spend":2905,"sales":8232,"purchases":8,"roas":2.83,"cpa":363.12},{"date":"2026-05-07","product":"Mamafy","spend":7937.57,"sales":23999,"purchases":29,"roas":3.02,"cpa":273.71},{"date":"2026-05-07","product":"Other","spend":3634.06,"sales":20205,"purchases":23,"roas":5.56,"cpa":158.0},{"date":"2026-05-08","product":"Brainify Drops","spend":13900.71,"sales":47554,"purchases":59,"roas":3.42,"cpa":235.61},{"date":"2026-05-08","product":"Brainify Powder","spend":20243.51,"sales":42642,"purchases":52,"roas":2.11,"cpa":389.3},{"date":"2026-05-08","product":"Flowjoy","spend":3247.53,"sales":4630,"purchases":5,"roas":1.43,"cpa":649.51},{"date":"2026-05-08","product":"Lactify","spend":2971.71,"sales":10438,"purchases":11,"roas":3.51,"cpa":270.16},{"date":"2026-05-08","product":"Mamafy","spend":7044.95,"sales":21952,"purchases":28,"roas":3.12,"cpa":251.61},{"date":"2026-05-08","product":"Other","spend":3390.14,"sales":7330,"purchases":8,"roas":2.16,"cpa":423.77},{"date":"2026-05-09","product":"Brainify Drops","spend":12872.84,"sales":31212,"purchases":44,"roas":2.42,"cpa":292.56},{"date":"2026-05-09","product":"Brainify Powder","spend":20191.22,"sales":47605,"purchases":58,"roas":2.36,"cpa":348.12},{"date":"2026-05-09","product":"Flowjoy","spend":2957.8,"sales":10768,"purchases":14,"roas":3.64,"cpa":211.27},{"date":"2026-05-09","product":"Lactify","spend":2578.89,"sales":7461,"purchases":8,"roas":2.89,"cpa":322.36},{"date":"2026-05-09","product":"Mamafy","spend":6116.17,"sales":25387,"purchases":32,"roas":4.15,"cpa":191.13},{"date":"2026-05-09","product":"Other","spend":2848.17,"sales":7311,"purchases":8,"roas":2.57,"cpa":356.02},{"date":"2026-05-10","product":"Brainify Drops","spend":14773.33,"sales":28099,"purchases":40,"roas":1.9,"cpa":369.33},{"date":"2026-05-10","product":"Brainify Powder","spend":21090.43,"sales":46099,"purchases":54,"roas":2.19,"cpa":390.56},{"date":"2026-05-10","product":"Flowjoy","spend":3542.58,"sales":4546,"purchases":6,"roas":1.28,"cpa":590.43},{"date":"2026-05-10","product":"Lactify","spend":3063.64,"sales":7898,"purchases":8,"roas":2.58,"cpa":382.95},{"date":"2026-05-10","product":"Mamafy","spend":7402.52,"sales":12763,"purchases":14,"roas":1.72,"cpa":528.75},{"date":"2026-05-10","product":"Other","spend":3545.54,"sales":6579,"purchases":7,"roas":1.86,"cpa":506.51},{"date":"2026-05-11","product":"Brainify Drops","spend":14070.52,"sales":35170,"purchases":47,"roas":2.5,"cpa":299.37},{"date":"2026-05-11","product":"Brainify Powder","spend":15290.26,"sales":46346,"purchases":54,"roas":3.03,"cpa":283.15},{"date":"2026-05-11","product":"Flowjoy","spend":3613.56,"sales":4063,"purchases":7,"roas":1.12,"cpa":516.22},{"date":"2026-05-11","product":"Lactify","spend":2811.73,"sales":3165,"purchases":4,"roas":1.13,"cpa":702.93},{"date":"2026-05-11","product":"Mamafy","spend":7368.27,"sales":14896,"purchases":19,"roas":2.02,"cpa":387.8},{"date":"2026-05-11","product":"Other","spend":3781.6,"sales":17212,"purchases":18,"roas":4.55,"cpa":210.09},{"date":"2026-05-12","product":"Brainify Drops","spend":13440.86,"sales":36385,"purchases":51,"roas":2.71,"cpa":263.55},{"date":"2026-05-12","product":"Brainify Powder","spend":12446.14,"sales":39090,"purchases":53,"roas":3.14,"cpa":234.83},{"date":"2026-05-12","product":"Flowjoy","spend":2469,"sales":3331,"purchases":5,"roas":1.35,"cpa":493.8},{"date":"2026-05-12","product":"Lactify","spend":2892.82,"sales":8399,"purchases":10,"roas":2.9,"cpa":289.28},{"date":"2026-05-12","product":"Mamafy","spend":6273.38,"sales":18576,"purchases":24,"roas":2.96,"cpa":261.39},{"date":"2026-05-12","product":"Other","spend":5508.43,"sales":20529,"purchases":21,"roas":3.73,"cpa":262.31},{"date":"2026-05-13","product":"Brainify Drops","spend":12082.85,"sales":28780,"purchases":37,"roas":2.38,"cpa":326.56},{"date":"2026-05-13","product":"Brainify Powder","spend":13168.93,"sales":41535,"purchases":54,"roas":3.15,"cpa":243.87},{"date":"2026-05-13","product":"Flowjoy","spend":3972.43,"sales":12940,"purchases":20,"roas":3.26,"cpa":198.62},{"date":"2026-05-13","product":"Lactify","spend":5152.26,"sales":6651,"purchases":9,"roas":1.29,"cpa":572.47},{"date":"2026-05-13","product":"Mamafy","spend":8497.21,"sales":16028,"purchases":20,"roas":1.89,"cpa":424.86},{"date":"2026-05-13","product":"Other","spend":4994.21,"sales":13291,"purchases":16,"roas":2.66,"cpa":312.14},{"date":"2026-05-14","product":"Brainify Drops","spend":12396.48,"sales":32569,"purchases":46,"roas":2.63,"cpa":269.49},{"date":"2026-05-14","product":"Brainify Powder","spend":12513.26,"sales":44695,"purchases":61,"roas":3.57,"cpa":205.14},{"date":"2026-05-14","product":"Flowjoy","spend":4147.24,"sales":8454,"purchases":12,"roas":2.04,"cpa":345.6},{"date":"2026-05-14","product":"Lactify","spend":4632.45,"sales":5363,"purchases":7,"roas":1.16,"cpa":661.78},{"date":"2026-05-14","product":"Mamafy","spend":8145.63,"sales":32109,"purchases":41,"roas":3.94,"cpa":198.67},{"date":"2026-05-14","product":"Other","spend":2555.87,"sales":8909,"purchases":9,"roas":3.49,"cpa":283.99},{"date":"2026-05-15","product":"Brainify Drops","spend":11995.35,"sales":23831,"purchases":29,"roas":1.99,"cpa":413.63},{"date":"2026-05-15","product":"Brainify Powder","spend":11602.32,"sales":39083,"purchases":49,"roas":3.37,"cpa":236.78},{"date":"2026-05-15","product":"Flowjoy","spend":3893.22,"sales":6028,"purchases":6,"roas":1.55,"cpa":648.87},{"date":"2026-05-15","product":"Lactify","spend":4368.42,"sales":11545,"purchases":14,"roas":2.64,"cpa":312.03},{"date":"2026-05-15","product":"Mamafy","spend":8103.21,"sales":19436,"purchases":24,"roas":2.4,"cpa":337.63},{"date":"2026-05-15","product":"Other","spend":2470.46,"sales":3864,"purchases":6,"roas":1.56,"cpa":411.74},{"date":"2026-05-16","product":"Brainify Drops","spend":8127.86,"sales":21359,"purchases":28,"roas":2.63,"cpa":290.28},{"date":"2026-05-16","product":"Brainify Powder","spend":9535.98,"sales":29692,"purchases":40,"roas":3.11,"cpa":238.4},{"date":"2026-05-16","product":"Flowjoy","spend":2195.72,"sales":5836,"purchases":6,"roas":2.66,"cpa":365.95},{"date":"2026-05-16","product":"Lactify","spend":3649.15,"sales":13344,"purchases":16,"roas":3.66,"cpa":228.07},{"date":"2026-05-16","product":"Mamafy","spend":6663.09,"sales":21521,"purchases":27,"roas":3.23,"cpa":246.78},{"date":"2026-05-16","product":"Other","spend":3080.47,"sales":6412,"purchases":6,"roas":2.08,"cpa":513.41},{"date":"2026-05-17","product":"Brainify Drops","spend":8220.72,"sales":21560,"purchases":24,"roas":2.62,"cpa":342.53},{"date":"2026-05-17","product":"Brainify Powder","spend":11674.18,"sales":23089,"purchases":31,"roas":1.98,"cpa":376.59},{"date":"2026-05-17","product":"Flowjoy","spend":3450.72,"sales":9234.4,"purchases":12,"roas":2.68,"cpa":287.56},{"date":"2026-05-17","product":"Lactify","spend":3974.52,"sales":13837,"purchases":16,"roas":3.48,"cpa":248.41},{"date":"2026-05-17","product":"Mamafy","spend":6875.61,"sales":12508,"purchases":17,"roas":1.82,"cpa":404.45},{"date":"2026-05-17","product":"Other","spend":2678.97,"sales":4114,"purchases":6,"roas":1.54,"cpa":446.49},{"date":"2026-05-18","product":"Brainify Drops","spend":8341.13,"sales":25185,"purchases":32,"roas":3.02,"cpa":260.66},{"date":"2026-05-18","product":"Brainify Powder","spend":10555.48,"sales":32500,"purchases":40,"roas":3.08,"cpa":263.89},{"date":"2026-05-18","product":"Flowjoy","spend":3652.84,"sales":8031,"purchases":12,"roas":2.2,"cpa":304.4},{"date":"2026-05-18","product":"Lactify","spend":4025.49,"sales":5703,"purchases":7,"roas":1.42,"cpa":575.07},{"date":"2026-05-18","product":"Mamafy","spend":6005.46,"sales":11615,"purchases":15,"roas":1.93,"cpa":400.36},{"date":"2026-05-18","product":"Other","spend":3140.95,"sales":10341,"purchases":13,"roas":3.29,"cpa":241.61},{"date":"2026-05-19","product":"Brainify Drops","spend":10385.72,"sales":25133,"purchases":32,"roas":2.42,"cpa":324.55},{"date":"2026-05-19","product":"Brainify Powder","spend":13974.16,"sales":34759,"purchases":45,"roas":2.49,"cpa":310.54},{"date":"2026-05-19","product":"Flowjoy","spend":2165.93,"sales":5750,"purchases":7,"roas":2.65,"cpa":309.42},{"date":"2026-05-19","product":"Lactify","spend":3506.62,"sales":13921,"purchases":17,"roas":3.97,"cpa":206.27},{"date":"2026-05-19","product":"Mamafy","spend":5408.93,"sales":19253,"purchases":22,"roas":3.56,"cpa":245.86},{"date":"2026-05-19","product":"Other","spend":3107.62,"sales":2387,"purchases":3,"roas":0.77,"cpa":1035.87},{"date":"2026-05-20","product":"Brainify Drops","spend":10111.06,"sales":33240,"purchases":40,"roas":3.29,"cpa":252.78},{"date":"2026-05-20","product":"Brainify Powder","spend":13309.36,"sales":33841,"purchases":41,"roas":2.54,"cpa":324.62},{"date":"2026-05-20","product":"Flowjoy","spend":2124.61,"sales":4985,"purchases":5,"roas":2.35,"cpa":424.92},{"date":"2026-05-20","product":"Lactify","spend":3437.06,"sales":9314,"purchases":11,"roas":2.71,"cpa":312.46},{"date":"2026-05-20","product":"Mamafy","spend":4914.73,"sales":16097,"purchases":19,"roas":3.28,"cpa":258.67},{"date":"2026-05-20","product":"Other","spend":2574.96,"sales":9322,"purchases":10,"roas":3.62,"cpa":257.5},{"date":"2026-05-21","product":"Brainify Drops","spend":10022.99,"sales":24533,"purchases":33,"roas":2.45,"cpa":303.73},{"date":"2026-05-21","product":"Brainify Powder","spend":13300.31,"sales":39856,"purchases":48,"roas":3.0,"cpa":277.09},{"date":"2026-05-21","product":"Flowjoy","spend":1875.24,"sales":5762,"purchases":9,"roas":3.07,"cpa":208.36},{"date":"2026-05-21","product":"Lactify","spend":5881.37,"sales":13267,"purchases":14,"roas":2.26,"cpa":420.1},{"date":"2026-05-21","product":"Mamafy","spend":4524.6,"sales":13980,"purchases":16,"roas":3.09,"cpa":282.79},{"date":"2026-05-21","product":"Other","spend":2460.98,"sales":4899,"purchases":7,"roas":1.99,"cpa":351.57},{"date":"2026-05-22","product":"Brainify Drops","spend":9343.49,"sales":19968,"purchases":27,"roas":2.14,"cpa":346.06},{"date":"2026-05-22","product":"Brainify Powder","spend":13989.9,"sales":45319,"purchases":56,"roas":3.24,"cpa":249.82},{"date":"2026-05-22","product":"Flowjoy","spend":1961.44,"sales":1759,"purchases":3,"roas":0.9,"cpa":653.81},{"date":"2026-05-22","product":"Lactify","spend":6021.93,"sales":14942,"purchases":15,"roas":2.48,"cpa":401.46},{"date":"2026-05-22","product":"Mamafy","spend":5557.88,"sales":14890,"purchases":16,"roas":2.68,"cpa":347.37},{"date":"2026-05-22","product":"Other","spend":2375.36,"sales":6597,"purchases":8,"roas":2.78,"cpa":296.92},{"date":"2026-05-23","product":"Brainify Drops","spend":8742.87,"sales":23722,"purchases":31,"roas":2.71,"cpa":282.03},{"date":"2026-05-23","product":"Brainify Powder","spend":15663.88,"sales":38513,"purchases":49,"roas":2.46,"cpa":319.67},{"date":"2026-05-23","product":"Flowjoy","spend":2934.73,"sales":6012,"purchases":10,"roas":2.05,"cpa":293.47},{"date":"2026-05-23","product":"Lactify","spend":5315.76,"sales":5728,"purchases":8,"roas":1.08,"cpa":664.47},{"date":"2026-05-23","product":"Mamafy","spend":8379.84,"sales":16072,"purchases":20,"roas":1.92,"cpa":418.99},{"date":"2026-05-23","product":"Other","spend":2809.36,"sales":4479,"purchases":5,"roas":1.59,"cpa":561.87},{"date":"2026-05-24","product":"Brainify Drops","spend":9404.37,"sales":19999,"purchases":30,"roas":2.13,"cpa":313.48},{"date":"2026-05-24","product":"Brainify Powder","spend":15077.01,"sales":41538,"purchases":52,"roas":2.76,"cpa":289.94},{"date":"2026-05-24","product":"Flowjoy","spend":3683.42,"sales":9193,"purchases":11,"roas":2.5,"cpa":334.86},{"date":"2026-05-24","product":"Lactify","spend":5447.14,"sales":13702,"purchases":16,"roas":2.52,"cpa":340.45},{"date":"2026-05-24","product":"Mamafy","spend":9050.17,"sales":17846,"purchases":19,"roas":1.97,"cpa":476.32},{"date":"2026-05-24","product":"Other","spend":1607.36,"sales":2316,"purchases":3,"roas":1.44,"cpa":535.79},{"date":"2026-05-25","product":"Brainify Drops","spend":14303.48,"sales":35234,"purchases":45,"roas":2.46,"cpa":317.86},{"date":"2026-05-25","product":"Brainify Powder","spend":14208.04,"sales":34552,"purchases":44,"roas":2.43,"cpa":322.91},{"date":"2026-05-25","product":"Flowjoy","spend":3415.99,"sales":4806,"purchases":8,"roas":1.41,"cpa":427.0},{"date":"2026-05-25","product":"Lactify","spend":5183.2,"sales":18470,"purchases":20,"roas":3.56,"cpa":259.16},{"date":"2026-05-25","product":"Mamafy","spend":8428.02,"sales":17108,"purchases":20,"roas":2.03,"cpa":421.4},{"date":"2026-05-25","product":"Other","spend":4587.07,"sales":5651,"purchases":9,"roas":1.23,"cpa":509.67},{"date":"2026-05-26","product":"Brainify Drops","spend":17820.17,"sales":41817,"purchases":52,"roas":2.35,"cpa":342.7},{"date":"2026-05-26","product":"Brainify Powder","spend":15340.48,"sales":42656,"purchases":55,"roas":2.78,"cpa":278.92},{"date":"2026-05-26","product":"Flowjoy","spend":3977.25,"sales":8073,"purchases":12,"roas":2.03,"cpa":331.44},{"date":"2026-05-26","product":"Lactify","spend":5612,"sales":17739,"purchases":19,"roas":3.16,"cpa":295.37},{"date":"2026-05-26","product":"Mamafy","spend":8532.07,"sales":23438,"purchases":27,"roas":2.75,"cpa":316.0},{"date":"2026-05-26","product":"Other","spend":3455.7,"sales":6121,"purchases":9,"roas":1.77,"cpa":383.97},{"date":"2026-05-27","product":"Brainify Drops","spend":7932.7,"sales":10922,"purchases":13,"roas":1.38,"cpa":610.21},{"date":"2026-05-27","product":"Brainify Powder","spend":8251.75,"sales":24435,"purchases":29,"roas":2.96,"cpa":284.54},{"date":"2026-05-27","product":"Flowjoy","spend":2962.53,"sales":4113,"purchases":5,"roas":1.39,"cpa":592.51},{"date":"2026-05-27","product":"Lactify","spend":2706.36,"sales":4732,"purchases":6,"roas":1.75,"cpa":451.06},{"date":"2026-05-27","product":"Mamafy","spend":4609.48,"sales":9668,"purchases":12,"roas":2.1,"cpa":384.12},{"date":"2026-05-27","product":"Other","spend":1240.78,"sales":1309,"purchases":1,"roas":1.05,"cpa":1240.78}];
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
  return (adName || 'Unknown').trim();
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
          Load Demo Data (May 1–25, 2026)
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

  const filteredAds = useMemo(() => {
    const base = dateFilteredAds || metaAds;
    if (!base) return [];
    let ads = selectedProduct !== "All Products" ? base.filter((a) => a.product === selectedProduct) : base;
    if (searchQuery) { const q = searchQuery.toLowerCase(); ads = ads.filter((a) => a.name.toLowerCase().includes(q) || a.product.toLowerCase().includes(q)); }
    if (adTypeFilter !== "All") ads = ads.filter((a) => a.type === adTypeFilter);
    if (creatorFilter !== "All") ads = ads.filter((a) => extractCreator(a.name) === creatorFilter);
    return [...ads].sort((a, b) => sortDir === "desc" ? b[sortField] - a[sortField] : a[sortField] - b[sortField]);
  }, [dateFilteredAds, metaAds, selectedProduct, searchQuery, sortField, sortDir, adTypeFilter, creatorFilter]);

  const maxSpend = useMemo(() => filteredAds.reduce((m, a) => Math.max(m, a.spend), 0), [filteredAds]);
  const uniqueCreators = useMemo(() => { const base = dateFilteredAds || metaAds; if (!base) return []; const s = new Set(base.map(a => extractCreator(a.name))); return ["All", ...Array.from(s).sort()]; }, [dateFilteredAds, metaAds]);
  const creatorTotals = useMemo(() => { const base = dateFilteredAds || metaAds; if (!base) return []; const pool = creatorProductFilter !== "All Products" ? base.filter(a => a.product === creatorProductFilter) : base; return computeCreatorTotals(pool); }, [dateFilteredAds, metaAds, creatorProductFilter]);
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
              ? "Demo data · May 1–25, 2026 · Meta + Google Ads"
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
