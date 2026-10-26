import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const sourcePath = "E:/battery-report.html";
const outputDir = "C:/inetpub/wwwroot/O365/outputs/battery_report_20260806";
const outputPath = outputDir + "/Battery_Health_Report_TNLX0006_2026-08-06.xlsx";
const html = await fs.readFile(sourcePath, "utf8");

function clean(v) {
  return v.replace(/<[^>]+>/g, " ").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&")
    .replace(/&#39;/g, "'").replace(/&quot;/gi, '"').replace(/\s+/g, " ").trim();
}
function section(heading) {
  const safe = heading.replace(/[.*+?^$\{\}()|[\]\\]/g, "\\$&");
  const m = html.match(new RegExp("<h2[^>]*>\\s*" + safe + "\\s*</h2>([\\s\\S]*?)(?=<h2|</body>)", "i"));
  return m ? m[1] : "";
}
function rows(s) {
  return [...s.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)].map(r =>
    [...r[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map(c => clean(c[1]))
  ).filter(r => r.length);
}
function parseMwh(v) {
  const m = String(v || "").match(/[\d,]+/);
  return m ? Number(m[0].replace(/,/g, "")) : null;
}
const systemPart = html.split(/<h2[^>]*>\s*Installed batteries\s*<\/h2>/i)[0];
function systemValue(label) {
  const m = systemPart.match(new RegExp(label + "\\s*</td>\\s*<td[^>]*>([\\s\\S]*?)</td>", "i"));
  return m ? clean(m[1]) : "";
}

const batteryMap = new Map(rows(section("Installed batteries")).filter(r => r.length >= 2).map(r => [r[0].toUpperCase(), r[1]]));
const capacityRows = rows(section("Battery capacity history")).slice(1).filter(r => r.length >= 3).map(r => ({
  period: r[0], full: parseMwh(r[1]), design: parseMwh(r[2]),
})).filter(r => r.full != null && r.design != null);
const lifeRows = rows(section("Battery life estimates"));
const sinceOs = lifeRows.find(r => String(r[0]).toLowerCase().includes("since os install")) || [];

const computerName = systemValue("COMPUTER NAME");
const productName = systemValue("SYSTEM PRODUCT NAME");
const bios = systemValue("BIOS");
const osBuild = systemValue("OS BUILD");
const reportTime = systemValue("REPORT TIME");
const batteryName = batteryMap.get("NAME") || "";
const manufacturer = batteryMap.get("MANUFACTURER") || "";
const serial = batteryMap.get("SERIAL NUMBER") || "";
const chemistry = batteryMap.get("CHEMISTRY") || "";
const designCapacity = parseMwh(batteryMap.get("DESIGN CAPACITY"));
const fullChargeCapacity = parseMwh(batteryMap.get("FULL CHARGE CAPACITY"));
const cycleCount = batteryMap.get("CYCLE COUNT") || "-";
const latestHistory = capacityRows[capacityRows.length - 1];
const initialHistory = capacityRows[0];
if (!designCapacity || !fullChargeCapacity || !latestHistory) throw new Error("Unable to parse required capacity values");

const wb = Workbook.create();
const summary = wb.worksheets.add("Summary");
const history = wb.worksheets.add("Capacity History");
const sourceSheet = wb.worksheets.add("Source & Notes");

const navy="#17365D", blue="#1F4E78", paleBlue="#DDEBF7", green="#E2F0D9";
const amber="#FFF2CC", red="#FCE4D6", gray="#E7E6E6", white="#FFFFFF", text="#1F2937";

function titleBand(sheet, address, value) {
  sheet.getRange(address).merge();
  sheet.getRange(address).values=[[value]];
  sheet.getRange(address).format={fill:navy,font:{bold:true,color:white,size:18},verticalAlignment:"center"};
  sheet.getRange(address).format.rowHeight=32;
}
function card(sheet,labelRange,valueRange,label,formula,fill,numberFormat) {
  sheet.getRange(labelRange).merge();
  sheet.getRange(labelRange).values=[[label]];
  sheet.getRange(labelRange).format={fill,font:{bold:true,color:navy},horizontalAlignment:"center",verticalAlignment:"center"};
  sheet.getRange(valueRange).merge();
  sheet.getRange(valueRange).formulas=[[formula]];
  sheet.getRange(valueRange).format={fill,font:{bold:true,color:navy,size:22},horizontalAlignment:"center",verticalAlignment:"center",numberFormat};
}

summary.showGridLines=false;
titleBand(summary,"A1:H2","Battery Health Report - "+computerName);
summary.getRange("A3:H3").merge();
summary.getRange("A3:H3").values=[[productName+" | Report time: "+reportTime]];
summary.getRange("A3:H3").format={fill:paleBlue,font:{italic:true,color:navy},verticalAlignment:"center"};
card(summary,"A5:B5","A6:B7","Current battery health","='Source & Notes'!$B$17/'Source & Notes'!$B$16",green,"0.0%");
card(summary,"C5:D5","C6:D7","Latest history health","='Capacity History'!$D$"+(capacityRows.length+1),amber,"0.0%");
card(summary,"E5:F5","E6:F7","Current capacity loss","=1-A6",amber,"0.0%");
card(summary,"G5:H5","G6:H7","Replacement threshold","='Source & Notes'!$B$21",gray,"0.0%");

summary.getRange("A9:H9").merge();
summary.getRange("A9:H9").values=[["คำแนะนำ"]];
summary.getRange("A9:H9").format={fill:blue,font:{bold:true,color:white,size:12}};
summary.getRange("A10:H12").merge();
summary.getRange("A10:H12").formulas=[['=IF(AND(A6>=\'Source & Notes\'!$B$21,C6>=\'Source & Notes\'!$B$21),"ยังไม่จำเป็นต้องเปลี่ยน",IF(A6<70%,"ควรเปลี่ยนโดยเร็ว","ควรวางแผนเปลี่ยน และตรวจยืนยันอีกครั้ง"))']];
summary.getRange("A10:H12").format={fill:amber,font:{bold:true,color:"#9C6500",size:18},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"outside",style:"medium",color:"#D6B656"}};

summary.getRange("A14:H14").merge();
summary.getRange("A14:H14").values=[["สรุปการประเมิน"]];
summary.getRange("A14:H14").format={fill:blue,font:{bold:true,color:white}};
summary.getRange("A15:H20").merge();
summary.getRange("A15:H20").values=[[
  "• Full Charge Capacity ปัจจุบัน "+fullChargeCapacity.toLocaleString("en-US")+" mWh จาก Design Capacity "+designCapacity.toLocaleString("en-US")+" mWh\n"+
  "• สุขภาพจากค่าปัจจุบัน "+(fullChargeCapacity/designCapacity*100).toFixed(1)+"% และเสื่อม "+((1-fullChargeCapacity/designCapacity)*100).toFixed(1)+"%\n"+
  "• Capacity history ล่าสุด "+latestHistory.full.toLocaleString("en-US")+" mWh หรือ "+(latestHistory.full/latestHistory.design*100).toFixed(1)+"% ซึ่งต่ำกว่าเกณฑ์ 80% เล็กน้อย\n"+
  "• ค่ายังแกว่งระหว่างค่าปัจจุบันกับ history จึงยังไม่จำเป็นต้องเปลี่ยนแบบเร่งด่วน แต่ควรวางแผนเปลี่ยนและรัน powercfg /batteryreport ตรวจซ้ำหลังชาร์จเต็ม\n"+
  "• หากใช้งานได้สั้นผิดปกติ ดับเอง ชาร์จไม่เข้า หรือแบตบวม ควรหยุดใช้งานแบตและเปลี่ยนทันที"
]];
summary.getRange("A15:H20").format={fill:"#F4F7FA",font:{color:text,size:11},wrapText:true,verticalAlignment:"top",borders:{preset:"outside",style:"thin",color:"#CBD5E1"}};
summary.getRange("A15:H20").format.rowHeight=24;
summary.getRange("A1:H20").format.columnWidth=17;
summary.freezePanes.freezeRows(3);

history.showGridLines=false;
history.getRange("A1:E1").values=[["Period","Full Charge Capacity (mWh)","Design Capacity (mWh)","Battery Health","Capacity Loss"]];
history.getRange("A1:E1").format={fill:blue,font:{bold:true,color:white},wrapText:true,verticalAlignment:"center"};
history.getRange("A1:E1").format.rowHeight=32;
history.getRangeByIndexes(1,0,capacityRows.length,3).values=capacityRows.map(r=>[r.period,r.full,r.design]);
history.getRange("D2").formulas=[["=B2/C2"]];
history.getRange("D2:D"+(capacityRows.length+1)).fillDown();
history.getRange("E2").formulas=[["=1-D2"]];
history.getRange("E2:E"+(capacityRows.length+1)).fillDown();
history.getRange("B2:C"+(capacityRows.length+1)).format.numberFormat="#,##0";
history.getRange("D2:E"+(capacityRows.length+1)).format.numberFormat="0.0%";
history.getRange("A2:E"+(capacityRows.length+1)).format={font:{color:text,size:10},borders:{insideHorizontal:{style:"thin",color:"#E5E7EB"}}};
history.getRange("D2:D"+(capacityRows.length+1)).conditionalFormats.add("cellIs",{operator:"lessThan",formula:"='Source & Notes'!$B$21",format:{fill:red,font:{color:"#9C0006",bold:true}}});
const historyTable=history.tables.add("A1:E"+(capacityRows.length+1),true,"BatteryCapacityHistory");
historyTable.style="TableStyleMedium2";
historyTable.showFilterButton=true;
history.freezePanes.freezeRows(1);
[24,24,22,16,16].forEach((w,i)=>history.getRangeByIndexes(0,i,capacityRows.length+1,1).format.columnWidth=w);
const chart=history.charts.add("line",history.getRange("A1:C"+(capacityRows.length+1)));
chart.title="Battery capacity trend (mWh)";
chart.hasLegend=true;
chart.yAxis={numberFormatCode:"#,##0"};
chart.xAxis={axisType:"textAxis",textStyle:{fontSize:9}};
chart.setPosition("G2","O21");

sourceSheet.showGridLines=false;
titleBand(sourceSheet,"A1:D2","Source Data & Assumptions");
const sourceRows=[
 ["ข้อมูลเครื่อง",""],["Computer name",computerName],["System product name",productName],["BIOS",bios],["OS build",osBuild],["Report time",reportTime],["",""],
 ["ข้อมูลแบตเตอรี่",""],["Battery name",batteryName],["Manufacturer",manufacturer],["Serial number",serial],["Chemistry",chemistry],
 ["Design capacity (mWh)",designCapacity],["Full charge capacity (mWh)",fullChargeCapacity],["Cycle count",cycleCount],["",""],
 ["เกณฑ์วิเคราะห์",""],["Replacement threshold",0.8],["Source file",sourcePath],["Initial history capacity",initialHistory.full],
 ["Latest history capacity",latestHistory.full],["Since OS install active estimate",sinceOs[1]||"-"],["At design capacity estimate",sinceOs[4]||"-"]
];
sourceSheet.getRangeByIndexes(3,0,sourceRows.length,2).values=sourceRows;
for(const r of [4,11,20]) sourceSheet.getRange("A"+r+":D"+r).format={fill:blue,font:{bold:true,color:white}};
sourceSheet.getRange("A4:D26").format.borders={insideHorizontal:{style:"thin",color:"#E5E7EB"}};
sourceSheet.getRange("A4:A26").format.font={bold:true,color:navy};
for(const r of [4,11,20]) sourceSheet.getRange("A"+r+":D"+r).format={fill:blue,font:{bold:true,color:white}};
sourceSheet.getRange("B16:B17").format.numberFormat="#,##0";
sourceSheet.getRange("B21").format={fill:amber,font:{bold:true,color:"#9C6500"},numberFormat:"0.0%"};
sourceSheet.getRange("B23:B24").format.numberFormat="#,##0";
sourceSheet.getRange("A1:A26").format.columnWidth=30;
sourceSheet.getRange("B1:B26").format.columnWidth=48;
sourceSheet.getRange("C1:D26").format.columnWidth=14;
sourceSheet.getRange("B22:B26").format.wrapText=true;
sourceSheet.freezePanes.freezeRows(2);

await fs.mkdir(outputDir,{recursive:true});
const check=await wb.inspect({kind:"table",range:"Summary!A1:H20",include:"values,formulas",tableMaxRows:22,tableMaxCols:10,maxChars:6000});
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:200},summary:"formula error scan",maxChars:2500});
await fs.writeFile(outputDir+"/verification.txt",check.ndjson+"\n"+errors.ndjson);
for(const [sheetName,range,file] of [["Summary","A1:H20","summary.png"],["Capacity History","A1:O24","capacity_history.png"],["Source & Notes","A1:D26","source_notes.png"]]){
 const preview=await wb.render({sheetName,range,scale:1.25,format:"png"});
 await fs.writeFile(outputDir+"/"+file,new Uint8Array(await preview.arrayBuffer()));
}
const exported=await SpreadsheetFile.exportXlsx(wb);
await exported.save(outputPath);
const reopened=await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const reopenedCheck=await reopened.inspect({kind:"table",range:"Summary!A5:H12",include:"values,formulas",tableMaxRows:10,tableMaxCols:10,maxChars:3000});
const reopenedErrors=await reopened.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},summary:"reopened formula errors",maxChars:2000});
await fs.writeFile(outputDir+"/reopen_verification.txt",reopenedCheck.ndjson+"\n"+reopenedErrors.ndjson);
console.log(JSON.stringify({outputPath,computerName,productName,designCapacity,fullChargeCapacity,currentHealth:fullChargeCapacity/designCapacity,latestHistoryHealth:latestHistory.full/latestHistory.design,capacityHistoryRows:capacityRows.length,formulaErrors:reopenedErrors.ndjson}));
