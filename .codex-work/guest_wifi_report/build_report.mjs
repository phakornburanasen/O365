import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const jobs = [
  ["2607300001","receiver@example.com",3,"สำเร็จ","รอดำเนินการ","2026-07-30T16:37:08"],
  ["2607300002","receiver@example.com",4,"สำเร็จ","รอดำเนินการ","2026-07-30T16:48:34"],
  ["2607300003","receiver@example.com",2,"สำเร็จ","รอดำเนินการ","2026-07-30T17:15:13"],
  ["2607310001","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-07-31T09:46:54"],
  ["2607310002","phakorn_bu@thanulux.com",5,"สำเร็จ","สำเร็จ","2026-07-31T09:50:47"],
  ["2607310003","phakorn_bu@thanulux.com",6,"สำเร็จ","สำเร็จ","2026-07-31T09:53:55"],
  ["2607310004","phakorn_bu@thanulux.com",6,"สำเร็จ","สำเร็จ","2026-07-31T09:56:31"],
  ["2607310005","phakorn_bu@thanulux.com",6,"สำเร็จ","สำเร็จ","2026-07-31T10:01:36"],
  ["2607310006","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-07-31T10:03:54"],
  ["2607310007","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-07-31T10:06:09"],
  ["2607310008","phakorn_bu@thanulux.com",4,"สำเร็จ","สำเร็จ","2026-07-31T10:08:24"],
  ["2607310009","phakorn_bu@thanulux.com",4,"ไม่สำเร็จ","รอดำเนินการ","2026-07-31T10:11:24"],
  ["2607310010","phakorn_bu@thanulux.com",4,"สำเร็จ","สำเร็จ","2026-07-31T10:14:20"],
  ["2607310011","phakorn_bu@thanulux.com",4,"สำเร็จ","สำเร็จ","2026-07-31T10:16:21"],
  ["2607310012","phakorn_bu@thanulux.com",5,"สำเร็จ","สำเร็จ","2026-07-31T10:54:52"],
  ["2607310013","phakorn_bu@thanulux.com",6,"สำเร็จ","สำเร็จ","2026-07-31T10:57:32"],
  ["2607310014","phakorn.7235@gmail.com",4,"สำเร็จ","สำเร็จ","2026-07-31T11:19:22"],
  ["2607310015","phakorn_bu@thanulux.com",2,"สำเร็จ","สำเร็จ","2026-07-31T11:31:21"],
  ["2607310016","phakorn_bu@thanulux.com",1,"สำเร็จ","สำเร็จ","2026-07-31T13:00:42"],
  ["2607310017","phakorn_bu@thanulux.com",2,"สำเร็จ","สำเร็จ","2026-07-31T13:19:38"],
  ["2607310018","phakorn_bu@thanulux.com",4,"สำเร็จ","สำเร็จ","2026-07-31T13:31:13"],
  ["2607310019","phakorn_bu@thanulux.com",10,"สำเร็จ","สำเร็จ","2026-07-31T13:42:09"],
  ["2607310020","phakorn_bu@thanulux.com",1,"สำเร็จ","สำเร็จ","2026-07-31T14:04:49"],
  ["2607310021","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-07-31T14:58:19"],
  ["2607310022","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-07-31T15:08:31"],
  ["2607310023","phakorn_bu@thanulux.com",4,"สำเร็จ","สำเร็จ","2026-07-31T15:43:41"],
  ["2607310024","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-07-31T16:41:11"],
  ["2608030001","phakorn_bu@thanulux.com",1,"สำเร็จ","สำเร็จ","2026-08-03T09:18:25"],
  ["2608030002","phakorn_bu@thanulux.com",2,"สำเร็จ","สำเร็จ","2026-08-03T09:42:51"],
  ["2608030003","phakorn_bu@thanulux.com",10,"สำเร็จ","สำเร็จ","2026-08-03T11:10:40"],
  ["2608030004","phakorn_bu@thanulux.com",1,"สำเร็จ","สำเร็จ","2026-08-03T12:56:29"],
  ["2608050001","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-08-05T11:43:49"],
  ["2608050002","phakorn_bu@thanulux.com",1,"สำเร็จ","สำเร็จ","2026-08-05T11:47:03"],
  ["2608050003","phakorn_bu@thanulux.com",1,"สำเร็จ","สำเร็จ","2026-08-05T14:43:13"],
  ["2608060001","phakorn_bu@thanulux.com",3,"สำเร็จ","สำเร็จ","2026-08-06T11:24:17"],
];

const outputDir = "C:/inetpub/wwwroot/O365/outputs/019fd54f-f1fd-7901-9c12-99b1821f224b";
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const summary = workbook.worksheets.add("สรุปผู้ขอ");
const detail = workbook.worksheets.add("รายละเอียดคำขอ");
summary.showGridLines = false;
detail.showGridLines = false;

const navy = "#16324F";
const teal = "#1F7A8C";
const lightTeal = "#DFF3F5";
const pale = "#F5F8FA";
const border = "#D5DEE5";
const red = "#B42318";
const amber = "#B54708";

summary.mergeCells("A1:H2");
summary.getRange("A1").values = [["รายงานผู้ขอใช้งาน Guest Wi‑Fi"]];
summary.getRange("A1:H2").format = {
  fill: navy,
  font: { bold: true, color: "#FFFFFF", size: 20 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
summary.mergeCells("A3:H3");
summary.getRange("A3").values = [["ข้อมูลคำขอทั้งหมดจากฐานข้อมูล ณ 6 สิงหาคม 2026 | ไม่รวม username และ password"]];
summary.getRange("A3:H3").format = { fill: "#E8EEF3", font: { color: "#355064", italic: true }, horizontalAlignment: "center" };

const cards = [["ผู้ขอทั้งหมด","จำนวนคำขอ","บัญชีที่ขอทั้งหมด","ส่งอีเมลสำเร็จ"]];
summary.getRange("A5:H5").values = [[cards[0][0],null,cards[0][1],null,cards[0][2],null,cards[0][3],null]];
for (const range of ["A5:B5","C5:D5","E5:F5","G5:H5","A6:B6","C6:D6","E6:F6","G6:H6"]) summary.getRange(range).merge();
summary.getRange("A6").formulas = [["=COUNTA(A9:A11)"]];
summary.getRange("C6").formulas = [["=COUNTA('รายละเอียดคำขอ'!$A$6:$A$40)"]];
summary.getRange("E6").formulas = [["=SUM('รายละเอียดคำขอ'!$C$6:$C$40)"]];
summary.getRange("G6").formulas = [["=COUNTIF('รายละเอียดคำขอ'!$E$6:$E$40,\"สำเร็จ\")"]];
summary.getRange("A5:H5").format = { fill: teal, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
summary.getRange("A6:H6").format = { fill: lightTeal, font: { bold: true, color: navy, size: 16 }, horizontalAlignment: "center", numberFormat: "#,##0", borders: { preset: "outside", style: "thin", color: border } };

summary.getRange("A8:H8").values = [["อีเมลผู้ขอ","จำนวนครั้งที่ขอ","จำนวนบัญชีที่ขอ","บัญชีที่สร้างสำเร็จ","คำขอไม่สำเร็จ","ส่งอีเมลสำเร็จ","ขอครั้งแรก","ขอล่าสุด"]];
summary.getRange("A9:A11").values = [["phakorn_bu@thanulux.com"],["receiver@example.com"],["phakorn.7235@gmail.com"]];
for (let row = 9; row <= 11; row++) {
  summary.getRange(`B${row}`).formulas = [[`=COUNTIF('รายละเอียดคำขอ'!$B$6:$B$40,A${row})`]];
  summary.getRange(`C${row}`).formulas = [[`=SUMIF('รายละเอียดคำขอ'!$B$6:$B$40,A${row},'รายละเอียดคำขอ'!$C$6:$C$40)`]];
  summary.getRange(`D${row}`).formulas = [[`=SUMIFS('รายละเอียดคำขอ'!$C$6:$C$40,'รายละเอียดคำขอ'!$B$6:$B$40,A${row},'รายละเอียดคำขอ'!$D$6:$D$40,\"สำเร็จ\")`]];
  summary.getRange(`E${row}`).formulas = [[`=COUNTIFS('รายละเอียดคำขอ'!$B$6:$B$40,A${row},'รายละเอียดคำขอ'!$D$6:$D$40,\"ไม่สำเร็จ\")`]];
  summary.getRange(`F${row}`).formulas = [[`=COUNTIFS('รายละเอียดคำขอ'!$B$6:$B$40,A${row},'รายละเอียดคำขอ'!$E$6:$E$40,\"สำเร็จ\")`]];
  summary.getRange(`G${row}`).formulas = [[`=MINIFS('รายละเอียดคำขอ'!$F$6:$F$40,'รายละเอียดคำขอ'!$B$6:$B$40,A${row})`]];
  summary.getRange(`H${row}`).formulas = [[`=MAXIFS('รายละเอียดคำขอ'!$F$6:$F$40,'รายละเอียดคำขอ'!$B$6:$B$40,A${row})`]];
}
summary.getRange("A8:H8").format = { fill: navy, font: { bold: true, color: "#FFFFFF" }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center" };
summary.getRange("A9:H11").format = { fill: "#FFFFFF", borders: { insideHorizontal: { style: "thin", color: border }, bottom: { style: "thin", color: border } } };
summary.getRange("B9:F11").format = { numberFormat: "#,##0", horizontalAlignment: "right" };
summary.getRange("G9:H11").format.numberFormat = "yyyy-mm-dd hh:mm";
summary.getRange("G9:H11").format.horizontalAlignment = "center";
summary.tables.add("A8:H11", true, "RequesterSummaryTable").style = "TableStyleMedium2";
summary.mergeCells("A14:H15");
summary.getRange("A14").values = [["หมายเหตุ: รายงานนี้รวมทุกระเบียนในตาราง job_ticket รวมถึง receiver@example.com ซึ่งมีลักษณะเป็นอีเมลตัวอย่าง และสรุปจำนวนตาม total_job ของแต่ละคำขอ"]];
summary.getRange("A14:H15").format = { fill: pale, font: { color: "#526777", italic: true }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: border } };
summary.freezePanes.freezeRows(3);

detail.mergeCells("A1:F2");
detail.getRange("A1").values = [["รายละเอียดคำขอ Guest Wi‑Fi ทั้งหมด"]];
detail.getRange("A1:F2").format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 18 }, horizontalAlignment: "center", verticalAlignment: "center" };
detail.mergeCells("A3:F3");
detail.getRange("A3").values = [["แหล่งข้อมูล: dbo.job_ticket | เวลาแสดงตามข้อมูลในระบบ | ไม่รวมข้อมูลรับรองการเข้าใช้งาน"]];
detail.getRange("A3:F3").format = { fill: "#E8EEF3", font: { color: "#355064", italic: true }, horizontalAlignment: "center" };
detail.getRange("A5:F5").values = [["Job ID","อีเมลผู้ขอ","จำนวนบัญชี","สถานะการสร้าง","สถานะส่งอีเมล","วันที่สร้าง"]];
detail.getRange("A6:F40").values = jobs.map(j => [j[0],j[1],j[2],j[3],j[4],new Date(`${j[5]}Z`)]);
detail.getRange("A5:F5").format = { fill: navy, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };
detail.getRange("A6:F40").format = { borders: { insideHorizontal: { style: "thin", color: border } } };
detail.getRange("A6:A40").format.numberFormat = "@";
detail.getRange("C6:C40").format = { numberFormat: "#,##0", horizontalAlignment: "right" };
detail.getRange("D6:E40").format.horizontalAlignment = "center";
detail.getRange("F6:F40").format = { numberFormat: "yyyy-mm-dd hh:mm", horizontalAlignment: "center" };
detail.getRange("D6:D40").conditionalFormats.add("containsText", { text: "ไม่สำเร็จ", format: { fill: "#FEE4E2", font: { color: red, bold: true } } });
detail.getRange("E6:E40").conditionalFormats.add("containsText", { text: "รอดำเนินการ", format: { fill: "#FEF0C7", font: { color: amber, bold: true } } });
detail.tables.add("A5:F40", true, "GuestWifiRequestsTable").style = "TableStyleMedium2";
detail.freezePanes.freezeRows(5);

summary.getRange("A1:H15").format.font.name = "Aptos";
detail.getRange("A1:F40").format.font.name = "Aptos";
summary.getRange("A1:H15").format.rowHeight = 22;
summary.getRange("A1:H2").format.rowHeight = 30;
summary.getRange("A8:H8").format.rowHeight = 34;
detail.getRange("A1:F40").format.rowHeight = 21;
detail.getRange("A1:F2").format.rowHeight = 30;
detail.getRange("A5:F5").format.rowHeight = 30;
summary.getRange("A:A").format.columnWidth = 31;
summary.getRange("B:F").format.columnWidth = 16;
summary.getRange("G:H").format.columnWidth = 21;
detail.getRange("A:A").format.columnWidth = 16;
detail.getRange("B:B").format.columnWidth = 34;
detail.getRange("C:C").format.columnWidth = 15;
detail.getRange("D:E").format.columnWidth = 19;
detail.getRange("F:F").format.columnWidth = 22;

const inspect = await workbook.inspect({ kind: "table", range: "สรุปผู้ขอ!A5:H11", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 8 });
console.log(inspect.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(errors.ndjson);

for (const [sheetName, fileName, range] of [["สรุปผู้ขอ","summary_preview.png","A1:H15"],["รายละเอียดคำขอ","detail_preview.png","A1:F40"]]) {
  const preview = await workbook.render({ sheetName, range, scale: 1.25, format: "png" });
  await fs.writeFile(`${outputDir}/${fileName}`, new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/guest_wifi_request_report_2026-08-06.xlsx`);
