import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workDir = "C:/inetpub/wwwroot/O365/.codex-work/endpoint_hdd_report";
const outputDir = "C:/inetpub/wwwroot/O365/outputs/endpoint_hdd_health_20260806";
const data = JSON.parse(await fs.readFile(`${workDir}/data.json`, "utf8"));
const exactClientData = JSON.parse(await fs.readFile(`${workDir}/exact_seqrite_client.json`, "utf8"));
const exactClientByHost = new Map(exactClientData.items.map(item => [String(item.hostname).toUpperCase(), item]));
const clientStatus = hostname => exactClientByHost.has(String(hostname).toUpperCase()) ? "Installed - Client 7.6" : "Not found - Client 7.6";
const clientVersion = hostname => exactClientByHost.get(String(hostname).toUpperCase())?.software_version || "";

const wb = Workbook.create();
const summary = wb.worksheets.add("Summary");
const devices = wb.worksheets.add("All Devices");
const disks = wb.worksheets.add("Disk Detail");
const missing = wb.worksheets.add("Missing Client 7.6");

const navy = "#16324F";
const blue = "#1F4E78";
const cyan = "#DDEBF7";
const pale = "#F4F7FA";
const green = "#E2F0D9";
const amber = "#FFF2CC";
const red = "#FCE4D6";
const gray = "#E7E6E6";
const white = "#FFFFFF";
const text = "#1F2937";

function titleBand(sheet, range, value) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[value]];
  sheet.getRange(range).format = {
    fill: navy,
    font: { bold: true, color: white, size: 18 },
    verticalAlignment: "center",
  };
  sheet.getRange(range).format.rowHeight = 32;
}

function styleHeader(range) {
  range.format = {
    fill: blue,
    font: { bold: true, color: white },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#B4C6E7" },
  };
  range.format.rowHeight = 30;
}

function styleTableBody(range) {
  range.format = {
    font: { color: text, size: 10 },
    verticalAlignment: "center",
    borders: { insideHorizontal: { style: "thin", color: "#E5E7EB" } },
  };
}

function addTable(sheet, range, name) {
  const t = sheet.tables.add(range, true, name);
  t.style = "TableStyleMedium2";
  t.showFilterButton = true;
  t.showBandedRows = true;
  return t;
}

// Summary
summary.showGridLines = false;
titleBand(summary, "A1:H2", "TNLX Endpoint - HDD Health & Hardware Inventory");
summary.getRange("A3:H3").merge();
summary.getRange("A3:H3").values = [[`Snapshot: ${data.generated_at.replace("T", " ")}  |  Source: ${data.source}`]];
summary.getRange("A3:H3").format = { fill: cyan, font: { color: navy, italic: true }, verticalAlignment: "center" };
summary.getRange("A3:H3").format.rowHeight = 23;

const cards = [
  ["A5:B5", "A6:B7", "Total endpoints", `=COUNTA('All Devices'!$A$2:$A$${data.total_devices + 1})`, cyan],
  ["C5:D5", "C6:D7", "Healthy", `=COUNTIF('All Devices'!$L$2:$L$${data.total_devices + 1},"Healthy")`, green],
  ["E5:F5", "E6:F7", "Warning", `=COUNTIF('All Devices'!$L$2:$L$${data.total_devices + 1},"Warning")`, amber],
  ["G5:H5", "G6:H7", "Critical", `=COUNTIF('All Devices'!$L$2:$L$${data.total_devices + 1},"Critical")`, red],
];
for (const [labelRange, valueRange, label, formula, fill] of cards) {
  summary.getRange(labelRange).merge();
  summary.getRange(labelRange).values = [[label]];
  summary.getRange(labelRange).format = { fill, font: { bold: true, color: navy }, horizontalAlignment: "center", verticalAlignment: "center" };
  summary.getRange(valueRange).merge();
  summary.getRange(valueRange).formulas = [[formula]];
  summary.getRange(valueRange).format = { fill, font: { bold: true, color: navy, size: 24 }, horizontalAlignment: "center", verticalAlignment: "center", numberFormat: "#,##0" };
}

summary.getRange("A9:D9").merge();
summary.getRange("A9:D9").values = [["Seqrite Endpoint Security Client 7.6"]];
summary.getRange("A9:D9").format = { fill: blue, font: { bold: true, color: white }, verticalAlignment: "center" };
summary.getRange("A10:B10").values = [["Installed Client 7.6", "Client 7.6 not found"]];
summary.getRange("A10:B10").format = { fill: cyan, font: { bold: true, color: navy }, horizontalAlignment: "center" };
summary.getRange("A11").formulas = [[`=COUNTIF('All Devices'!$M$2:$M$${data.total_devices + 1},"Installed - Client 7.6")`]];
summary.getRange("B11").formulas = [[`=COUNTIF('All Devices'!$M$2:$M$${data.total_devices + 1},"Not found - Client 7.6")`]];
summary.getRange("A11:B11").format = { font: { bold: true, size: 18, color: navy }, horizontalAlignment: "center", numberFormat: "#,##0", borders: { preset: "outside", style: "thin", color: "#B4C6E7" } };
summary.getRange("E9:F9").merge();
summary.getRange("E9:F9").values = [["Unknown HDD status"]];
summary.getRange("E9:F9").format = { fill: gray, font: { bold: true, color: navy }, horizontalAlignment: "center" };
summary.getRange("E10:F11").merge();
summary.getRange("E10:F11").formulas = [[`=COUNTIF('All Devices'!$L$2:$L$${data.total_devices + 1},"Unknown")`]];
summary.getRange("E10:F11").format = { fill: gray, font: { bold: true, size: 18, color: navy }, horizontalAlignment: "center", verticalAlignment: "center", numberFormat: "#,##0" };
summary.getRange("G9:H9").merge();
summary.getRange("G9:H9").values = [["Logical disks"]];
summary.getRange("G9:H9").format = { fill: cyan, font: { bold: true, color: navy }, horizontalAlignment: "center" };
summary.getRange("G10:H11").merge();
summary.getRange("G10:H11").formulas = [[`=COUNTA('Disk Detail'!$A$2:$A$${data.total_disks + 1})`]];
summary.getRange("G10:H11").format = { fill: cyan, font: { bold: true, size: 18, color: navy }, horizontalAlignment: "center", verticalAlignment: "center", numberFormat: "#,##0" };

summary.getRange("A13:H13").merge();
summary.getRange("A13:H13").values = [["Rules and important notes"]];
summary.getRange("A13:H13").format = { fill: blue, font: { bold: true, color: white } };
summary.getRange("A14:H18").merge();
summary.getRange("A14:H18").values = [[
  "- Per-device HDD Health uses the most severe result across all disks and considers both SMART and capacity status.\n" +
  "- Severity order: Critical/Danger > Warning/Caution > Healthy > Unknown.\n" +
  "- Client status is matched against the exact software name Seqrite Endpoint Security Client 7.6, not the shortened word Seqrite.\n" +
  "- Not found - Client 7.6 means the exact Client 7.6 product was absent from the latest installed-software inventory; it is not definitive proof of uninstall.\n" +
  "- Per-disk records are in Disk Detail; affected hostnames are in Missing Client 7.6.\n" +
  "- This report is a point-in-time inventory snapshot."
]];
summary.getRange("A14:H18").format = { fill: pale, font: { color: text, size: 11 }, wrapText: true, verticalAlignment: "top", borders: { preset: "outside", style: "thin", color: "#CBD5E1" } };
summary.getRange("A14:H18").format.rowHeight = 24;
summary.getRange("A1:H18").format.columnWidth = 16;
summary.getRange("A1:A18").format.columnWidth = 20;
summary.getRange("B1:B18").format.columnWidth = 20;
summary.freezePanes.freezeRows(3);

// All Devices
devices.showGridLines = false;
const deviceHeaders = ["Hostname","IP Address","Username","Employee","CPU","RAM (GB)","Windows","Disk Count","Disk Type","Disk Models","Max Used %","HDD Health","Client 7.6 Status","Client Version","Updated At"];
const deviceRows = data.devices.map(d => [
  d.hostname,d.ip_address,d.username,[d.employee_name,d.nickname ? `(${d.nickname})` : ""].filter(Boolean).join(" "),
  d.cpu_name,d.ram_total_gb,d.windows_version,d.disk_count,d.disk_types,d.disk_models,
  d.max_used_percent == null ? null : d.max_used_percent/100,d.disk_health_status,clientStatus(d.hostname),clientVersion(d.hostname),
  d.updated_at ? new Date(d.updated_at + "Z") : null
]);
devices.getRangeByIndexes(0,0,1,deviceHeaders.length).values = [deviceHeaders];
devices.getRangeByIndexes(1,0,deviceRows.length,deviceHeaders.length).values = deviceRows;
styleHeader(devices.getRange("A1:O1"));
styleTableBody(devices.getRange(`A2:O${deviceRows.length+1}`));
devices.getRange(`F2:F${deviceRows.length+1}`).format.numberFormat = "0.00";
devices.getRange(`K2:K${deviceRows.length+1}`).format.numberFormat = "0.0%";
devices.getRange(`O2:O${deviceRows.length+1}`).format.numberFormat = "yyyy-mm-dd hh:mm";
devices.getRange(`L2:L${deviceRows.length+1}`).conditionalFormats.add("containsText",{text:"Healthy",format:{fill:green,font:{color:"#375623",bold:true}}});
devices.getRange(`L2:L${deviceRows.length+1}`).conditionalFormats.add("containsText",{text:"Warning",format:{fill:amber,font:{color:"#9C6500",bold:true}}});
devices.getRange(`L2:L${deviceRows.length+1}`).conditionalFormats.add("containsText",{text:"Critical",format:{fill:red,font:{color:"#9C0006",bold:true}}});
devices.getRange(`L2:L${deviceRows.length+1}`).conditionalFormats.add("containsText",{text:"Unknown",format:{fill:gray,font:{color:"#595959",bold:true}}});
devices.getRange(`M2:M${deviceRows.length+1}`).conditionalFormats.add("containsText",{text:"Not found",format:{fill:red,font:{color:"#9C0006",bold:true}}});
devices.getRange(`M2:M${deviceRows.length+1}`).conditionalFormats.add("containsText",{text:"Installed",format:{fill:green,font:{color:"#375623",bold:true}}});
addTable(devices,`A1:O${deviceRows.length+1}`,"AllDevicesTable");
devices.freezePanes.freezeRows(1);
devices.freezePanes.freezeColumns(1);
const deviceWidths = [15,15,24,22,42,11,38,11,14,45,12,13,24,15,20];
deviceWidths.forEach((w,i)=>devices.getRangeByIndexes(0,i,deviceRows.length+1,1).format.columnWidth=w);
devices.getRange(`D2:E${deviceRows.length+1}`).format.wrapText = true;
devices.getRange(`G2:J${deviceRows.length+1}`).format.wrapText = true;

// Disk Detail
disks.showGridLines = false;
const diskHeaders = ["Hostname","IP Address","Drive","Disk Model","Disk Type","SMART Health","Total (GB)","Free (GB)","Used (GB)","Used %","Capacity Status","CPU","RAM (GB)","Client 7.6 Status","Collected At"];
const diskRows = data.disk_details.map(d => [
  d.hostname,d.ip_address,d.drive_letter,d.model,d.disk_type,d.smart_health,d.total_gb,d.free_gb,d.used_gb,
  d.used_percent == null ? null : d.used_percent/100,d.capacity_status,d.cpu_name,d.ram_total_gb,clientStatus(d.hostname),
  d.collected_at ? new Date(d.collected_at + "Z") : null
]);
disks.getRangeByIndexes(0,0,1,diskHeaders.length).values=[diskHeaders];
disks.getRangeByIndexes(1,0,diskRows.length,diskHeaders.length).values=diskRows;
styleHeader(disks.getRange("A1:O1"));
styleTableBody(disks.getRange(`A2:O${diskRows.length+1}`));
disks.getRange(`G2:I${diskRows.length+1}`).format.numberFormat = "0.00";
disks.getRange(`J2:J${diskRows.length+1}`).format.numberFormat = "0.0%";
disks.getRange(`M2:M${diskRows.length+1}`).format.numberFormat = "0.00";
disks.getRange(`O2:O${diskRows.length+1}`).format.numberFormat = "yyyy-mm-dd hh:mm";
for (const col of ["F","K"]) {
  disks.getRange(`${col}2:${col}${diskRows.length+1}`).conditionalFormats.add("containsText",{text:"Healthy",format:{fill:green,font:{color:"#375623"}}});
  disks.getRange(`${col}2:${col}${diskRows.length+1}`).conditionalFormats.add("containsText",{text:"Warning",format:{fill:amber,font:{color:"#9C6500",bold:true}}});
  disks.getRange(`${col}2:${col}${diskRows.length+1}`).conditionalFormats.add("containsText",{text:"Caution",format:{fill:amber,font:{color:"#9C6500",bold:true}}});
  disks.getRange(`${col}2:${col}${diskRows.length+1}`).conditionalFormats.add("containsText",{text:"Critical",format:{fill:red,font:{color:"#9C0006",bold:true}}});
  disks.getRange(`${col}2:${col}${diskRows.length+1}`).conditionalFormats.add("containsText",{text:"Danger",format:{fill:red,font:{color:"#9C0006",bold:true}}});
}
addTable(disks,`A1:O${diskRows.length+1}`,"DiskDetailTable");
disks.freezePanes.freezeRows(1);
disks.freezePanes.freezeColumns(1);
const diskWidths=[15,15,9,42,12,14,12,12,12,11,15,42,11,24,20];
diskWidths.forEach((w,i)=>disks.getRangeByIndexes(0,i,diskRows.length+1,1).format.columnWidth=w);
disks.getRange(`D2:D${diskRows.length+1}`).format.wrapText=true;
disks.getRange(`L2:L${diskRows.length+1}`).format.wrapText=true;

// Missing exact Seqrite Endpoint Security Client 7.6
missing.showGridLines = false;
const missingHeaders = ["Hostname","IP Address","Username","CPU","RAM (GB)","Disk Type","HDD Health","Max Used %","Windows","Updated At","Client 7.6 Status"];
const missingDevices = data.devices.filter(d=>!exactClientByHost.has(String(d.hostname).toUpperCase()));
const missingRows = missingDevices.map(d=>[
  d.hostname,d.ip_address,d.username,d.cpu_name,d.ram_total_gb,d.disk_types,d.disk_health_status,
  d.max_used_percent==null?null:d.max_used_percent/100,d.windows_version,d.updated_at?new Date(d.updated_at + "Z"):null,clientStatus(d.hostname)
]);
missing.getRangeByIndexes(0,0,1,missingHeaders.length).values=[missingHeaders];
missing.getRangeByIndexes(1,0,missingRows.length,missingHeaders.length).values=missingRows;
styleHeader(missing.getRange("A1:K1"));
styleTableBody(missing.getRange(`A2:K${missingRows.length+1}`));
missing.getRange(`E2:E${missingRows.length+1}`).format.numberFormat="0.00";
missing.getRange(`H2:H${missingRows.length+1}`).format.numberFormat="0.0%";
missing.getRange(`J2:J${missingRows.length+1}`).format.numberFormat="yyyy-mm-dd hh:mm";
missing.getRange(`G2:G${missingRows.length+1}`).conditionalFormats.add("containsText",{text:"Warning",format:{fill:amber,font:{color:"#9C6500",bold:true}}});
missing.getRange(`G2:G${missingRows.length+1}`).conditionalFormats.add("containsText",{text:"Critical",format:{fill:red,font:{color:"#9C0006",bold:true}}});
missing.getRange(`K2:K${missingRows.length+1}`).conditionalFormats.add("containsText",{text:"Not found",format:{fill:red,font:{color:"#9C0006",bold:true}}});
addTable(missing,`A1:K${missingRows.length+1}`,"MissingSeqriteTable");
missing.freezePanes.freezeRows(1);
missing.freezePanes.freezeColumns(1);
const missingWidths=[15,15,24,42,11,14,13,12,38,20,24];
missingWidths.forEach((w,i)=>missing.getRangeByIndexes(0,i,missingRows.length+1,1).format.columnWidth=w);
missing.getRange(`D2:D${missingRows.length+1}`).format.wrapText=true;
missing.getRange(`I2:I${missingRows.length+1}`).format.wrapText=true;

await fs.mkdir(outputDir,{recursive:true});

const summaryCheck = await wb.inspect({kind:"table",range:"Summary!A1:H18",include:"values,formulas",tableMaxRows:20,tableMaxCols:10,maxChars:5000});
const errorCheck = await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"formula errors",maxChars:3000});
await fs.writeFile(`${outputDir}/verification.txt`, summaryCheck.ndjson + "\n" + errorCheck.ndjson);

const previews = [
  ["Summary","A1:H18","summary.png"],
  ["All Devices","A1:O20","all_devices.png"],
  ["Disk Detail","A1:O20","disk_detail.png"],
  ["Missing Client 7.6","A1:K20","missing_seqrite.png"],
];
for (const [sheetName,range,file] of previews) {
  const blob = await wb.render({sheetName,range,scale:1.25,format:"png"});
  await fs.writeFile(`${outputDir}/${file}`,new Uint8Array(await blob.arrayBuffer()));
}

const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(`${outputDir}/TNLX_Endpoint_HDD_Health_Report_Seqrite_Client_7.6_2026-08-06.xlsx`);
const exportedPath = `${outputDir}/TNLX_Endpoint_HDD_Health_Report_Seqrite_Client_7.6_2026-08-06.xlsx`;
const reopened = await SpreadsheetFile.importXlsx(await FileBlob.load(exportedPath));
const reopenedSummary = await reopened.inspect({kind:"table",range:"Summary!A5:H11",include:"values,formulas",tableMaxRows:10,tableMaxCols:10,maxChars:3000});
const reopenedErrors = await reopened.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},summary:"reopened formula errors",maxChars:2000});
await fs.writeFile(`${outputDir}/reopen_verification.txt`,reopenedSummary.ndjson + "\n" + reopenedErrors.ndjson);
console.log(JSON.stringify({output:exportedPath,devices:deviceRows.length,disks:diskRows.length,missingSeqrite:missingRows.length,checks:errorCheck.ndjson,reopenedChecks:reopenedErrors.ndjson}));


