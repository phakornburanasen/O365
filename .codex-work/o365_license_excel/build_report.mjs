import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const items = [
  {
    "emp_id": "55056",
    "name_email": "yothin_su@thanulux.com",
    "full_name": "โยธิน สุวรรณเกษร์",
    "disp_name": "yothin suwanket",
    "nick_name": "โจ้",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "arrow",
    "name_email": "arrow@thanulux.com",
    "full_name": "arrow",
    "disp_name": "arrow",
    "nick_name": "ส่วนกลาง",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "Lek",
    "name_email": "sompatsorn_taw@thanulux.com",
    "full_name": "sompatsorn_taw",
    "disp_name": "sompatsorn_taw",
    "nick_name": "เล็ก/ลาออก",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "Suchin",
    "name_email": "suchin_de@thanulux.com",
    "full_name": "suchin de",
    "disp_name": "suchin de",
    "nick_name": "สุชิน",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9030",
    "name_email": "pornanong_me@thanulux.com",
    "full_name": "พรอนงค์ มีภูมิรู้",
    "disp_name": "Pornanong Meephumroo",
    "nick_name": "กุ้ง",
    "depart_id": "2AM03",
    "depart_name": "ฝ่าย HR",
    "cost_center": "2AM03",
    "cost_name": "ฝ่าย HR",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9058",
    "name_email": "Phakorn_bu@thanulux.com",
    "full_name": "ภากร บูรณเสน",
    "disp_name": "Phakorn Buranasen",
    "nick_name": "โบ้",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9062",
    "name_email": "Witsanu_ke@thanulux.com",
    "full_name": "วิษณุ กีพิมาย",
    "disp_name": "Witsanu Keephimai",
    "nick_name": "นุ",
    "depart_id": "2AM05",
    "depart_name": "แผนกตรวจสอบภายใน",
    "cost_center": "2AM05",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9070",
    "name_email": "Saowanee_ti@thanulux.com",
    "full_name": "เสาวณี ติยะพานิชกุล",
    "disp_name": "Saowanee Tiyapanichkul",
    "nick_name": "พี่เจี๊ยบ",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9095",
    "name_email": "jeerachart_pc@thanulux.com",
    "full_name": "จีรฉัตร ปากชำนิ",
    "disp_name": "Jeerachart Pakchamni",
    "nick_name": "ก๊อฟ",
    "depart_id": "I8111",
    "depart_name": "INNER GUY LAROCHE",
    "cost_center": "I8111",
    "cost_name": "ฝ่าย Apparel 2",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9108",
    "name_email": "corp_chalida@thanulux.com",
    "full_name": "ชลิดา สุประหยัด",
    "disp_name": "Chalida Suprayad",
    "nick_name": "ยุ้ย",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9118",
    "name_email": "Waethanee@thanulux.com",
    "full_name": "เวทะนี ธนาศรัณยากร",
    "disp_name": "Waethanee Tanasarunyakorn",
    "nick_name": "เวย์",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9133",
    "name_email": "wassana_sa@thanulux.com",
    "full_name": "วาสนา สรรวัจนา",
    "disp_name": "Wassana Sanvatjana",
    "nick_name": "จิ๊บ",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9172",
    "name_email": "peerada_po@thanulux.com",
    "full_name": "พีรดา ศิริภักดีนันต์",
    "disp_name": "Peerada Siripakdeenan",
    "nick_name": "หลิน",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9173",
    "name_email": "Kannika_ka@thanulux.com",
    "full_name": "กรรณิกา แก้วตา",
    "disp_name": "Kannika Kaewta",
    "nick_name": "กรรณ",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9207",
    "name_email": "ben_pr@thanulux.com",
    "full_name": "กัญญาณัฏฐ์ สุภะ",
    "disp_name": "Kanyanatth Supa (Ben)",
    "nick_name": "เบน/ลาออก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9220",
    "name_email": "Sompob@thanulux.com",
    "full_name": "กิติภณ วิสุทธิ์เทพ",
    "disp_name": "Kitiphon wisutthep",
    "nick_name": "เค",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9232",
    "name_email": "Prapaporn_ka@thanulux.com",
    "full_name": "ประภาภรณ์ กำจรจรุงวิทย์",
    "disp_name": "Prapaporn Kumjornjarungwit",
    "nick_name": "ภรณ์",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9252",
    "name_email": "Sirapirom@thanulux.com",
    "full_name": "ศิราภิรมย์ ธีรประเสริฐ",
    "disp_name": "Sirapirom Teeraprasert",
    "nick_name": "พี่เตี๊ยบ",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9256",
    "name_email": "Tawatchai_pi@thanulux.com",
    "full_name": "ธวัชชัย ปิ่นแก้ว",
    "disp_name": "Tawatchai pinkaew",
    "nick_name": "วัช",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9265",
    "name_email": "Montree_le@thanulux.com",
    "full_name": "มนตรี เลี้ยงหล่ำ",
    "disp_name": "Montree Leanglam",
    "nick_name": "เอก",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9267",
    "name_email": "wannee_so@thanulux.com",
    "full_name": "วรรณี สมบูรณ์",
    "disp_name": "Wannee Somboon",
    "nick_name": "ขวัญ",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9280",
    "name_email": "kannika@thanulux.com",
    "full_name": "กรรณิกา บำรุงวงศ์ทอง",
    "disp_name": "Kannika Bumroongvongtong",
    "nick_name": "เจ",
    "depart_id": "2RE01",
    "depart_name": "ฝ่ายพัฒนาธุรกิจ",
    "cost_center": "1RE01",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9299",
    "name_email": "thanachai_ph@thanulux.com",
    "full_name": "ธนชัย พัวพันปัญโญภาส",
    "disp_name": "Thanachai Phuaphanpunyopas",
    "nick_name": "หนึ่ง",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9304",
    "name_email": "Avitra_nu@thanulux.com",
    "full_name": "อวิตรา นุชารัมย์",
    "disp_name": "Avitra Nucharam",
    "nick_name": "นก",
    "depart_id": "Y8111",
    "depart_name": "แผนก OLYMP",
    "cost_center": "Y8111",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9315",
    "name_email": "banjongruk_sa@thanulux.com",
    "full_name": "บรรจงรักษ์ ปุญญาคม",
    "disp_name": "Banjongruk Punyakom",
    "nick_name": "มด",
    "depart_id": "I8111",
    "depart_name": "INNER GUY LAROCHE",
    "cost_center": "I8111",
    "cost_name": "ฝ่าย Apparel 2",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9334",
    "name_email": "Naluedee_ph@thanulux.com",
    "full_name": "ณฤดี พีรบริรักษ์",
    "disp_name": "Naluedee Phiraborirak",
    "nick_name": "ฟรอยด์",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9344",
    "name_email": "Waleerat_sk@thanulux.com",
    "full_name": "วลีรัตน์ สกุลประเสริฐยิ่ง",
    "disp_name": "Waleerat Sakulprasertying",
    "nick_name": "น้อง",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9353",
    "name_email": "Purchase_patompon@thanulux.com",
    "full_name": "ปฐมพล ผลาพิบูลย์",
    "disp_name": "Patompon Bhlapibul",
    "nick_name": "ตู่",
    "depart_id": "2PC02",
    "depart_name": "จัดซื้อเครื่องหนัง",
    "cost_center": "2PC02",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9362",
    "name_email": "Waraluk_su@thanulux.com",
    "full_name": "วราลักษณ์ สุขพรรณเจริญ",
    "disp_name": "Waraluk Sukphancharoen",
    "nick_name": "ยุ้ย",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9370",
    "name_email": "ping@thanulux.com",
    "full_name": "กนกชัย มิ่งขวัญปิยะกุล",
    "disp_name": "Kanokchai Mingquanpiyakul",
    "nick_name": "พี่ปิง",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9384",
    "name_email": "Karnsiree@thanulux.com",
    "full_name": "กานต์สิรี เทศทองคำ",
    "disp_name": "Karnsiree Testhongkam",
    "nick_name": "วา",
    "depart_id": "C8111",
    "depart_name": "ABSORBA",
    "cost_center": "C8111",
    "cost_name": "ABSORBA",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9410",
    "name_email": "wanna_inner@thanulux.com",
    "full_name": "วรรณา ภู่รักแก้ว",
    "disp_name": "Wanna Purakkaew",
    "nick_name": "พี่ณา",
    "depart_id": "P8111",
    "depart_name": "แผนกธุรกิจ DAKS",
    "cost_center": "P8111",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9439",
    "name_email": "thanakrit_sa@thanulux.com",
    "full_name": "ธนกฤษ ไสยกิจ",
    "disp_name": "Thanakrit Saiyakit",
    "nick_name": "เอ",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9451",
    "name_email": "wiparat_ph@thanulux.com",
    "full_name": "วิภารัตน์ พันธ์ผาด",
    "disp_name": "Wiparat Phanphad",
    "nick_name": "แอม",
    "depart_id": "2AM03",
    "depart_name": "ฝ่าย HR",
    "cost_center": "2AM03",
    "cost_name": "ฝ่าย HR",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9463",
    "name_email": "Sunee@thanulux.com",
    "full_name": "สุนีย์ คงสงฆ์",
    "disp_name": "Sunee Kongsong",
    "nick_name": "อ้อย",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9490",
    "name_email": "janthana_su@thanulux.com",
    "full_name": "จันทนา สวนพฤกษา",
    "disp_name": "Janthana Suanprueksa",
    "nick_name": "หน่อย",
    "depart_id": "1RI01",
    "depart_name": "หน่วยวิจัยพัฒนา",
    "cost_center": "1RI01",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9495",
    "name_email": "Sakorn_pe@thanulux.com",
    "full_name": "สาคร เพิ่มสวัสดิ์",
    "disp_name": "Sakorn Permsawat",
    "nick_name": "โอ๋",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9515",
    "name_email": "Panmanat_ti@thanulux.com",
    "full_name": "ปานมนัส ทิตระกูล",
    "disp_name": "Panmanat Titragool",
    "nick_name": "เนม",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9521",
    "name_email": "Chanthita_si@thanulux.com",
    "full_name": "ฉันทิตา ศิริศักดิพร",
    "disp_name": "Chanthita Sirisakdiporn",
    "nick_name": "ตา",
    "depart_id": "2AM03",
    "depart_name": "ฝ่าย HR",
    "cost_center": "2AM03",
    "cost_name": "ฝ่าย HR",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9530",
    "name_email": "salila_kl@thanulux.com",
    "full_name": "สลิลา ไกลกล่อม",
    "disp_name": "Salila Klaiklom",
    "nick_name": "นุช",
    "depart_id": "H8111",
    "depart_name": "House Brand",
    "cost_center": "H8111",
    "cost_name": null,
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9571",
    "name_email": "pailin_lo@thanulux.com",
    "full_name": "ไพลิน โล่ห์บัณฑิตกุล",
    "disp_name": "Pailin Lohbanditkul",
    "nick_name": "ไพลิน",
    "depart_id": "P8111",
    "depart_name": "แผนกธุรกิจ DAKS",
    "cost_center": "P8111",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9603",
    "name_email": "panjaphol_ja@thanulux.com",
    "full_name": "ปัญจพล จันทร์ทรง",
    "disp_name": "Panjaphol Jansong",
    "nick_name": "หนึ่ง/ลาออก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9614",
    "name_email": "jueng@thanulux.com",
    "full_name": "สวรรณกมล กิจสุนทราภา",
    "disp_name": "Sawankamon Kijsoontrapa",
    "nick_name": "พี่จึง",
    "depart_id": "P8111",
    "depart_name": "แผนกธุรกิจ DAKS",
    "cost_center": "P8111",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9618",
    "name_email": "Laddawan_so@thanulux.com",
    "full_name": "ลัดดาวัลย์ โสทัศน์",
    "disp_name": "Laddawan Sothus",
    "nick_name": "เจี๊ยบ",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9622",
    "name_email": "atchara_guy@thanulux.com",
    "full_name": "อัจฉรา ธรรมนาวรรณ",
    "disp_name": "Atchara Thammanawan",
    "nick_name": "อุ๋ย",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9655",
    "name_email": "ekkaphan_ch@thanulux.com",
    "full_name": "เอกพรรณ ชุมนิกาย",
    "disp_name": "Ekkaphan Chumnikay",
    "nick_name": "ปอนด์",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9658",
    "name_email": "aoi_gl@thanulux.com",
    "full_name": "รัชนิต ลือพงศ์พาณิชย์",
    "disp_name": "Ratchanit Luepongpanich",
    "nick_name": "อ๋อย",
    "depart_id": "2AM01",
    "depart_name": "แผนกบัญชี",
    "cost_center": "2AM01",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9671",
    "name_email": "sirichai_th@thanulux.com",
    "full_name": "ศิริชัย ติระบาล",
    "disp_name": "Sirichai Thirabal",
    "nick_name": "ก๊อต/ลาออก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9672",
    "name_email": "apichart_ph@thanulux.com",
    "full_name": "อภิชาติ ผาลึกโรจน์",
    "disp_name": "Apichart Phaluekroj",
    "nick_name": "ชาติ",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9696",
    "name_email": "Siri@thanulux.com",
    "full_name": "ศิริ รัตนาไพบูลย์",
    "disp_name": "ศิริ รัตนาไพบูลย์",
    "nick_name": "หน่อง",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9710",
    "name_email": "Kesinee_su@thanulux.com",
    "full_name": "เกศินี สุภะชุณหะ",
    "disp_name": "Kesinee Supachunha",
    "nick_name": "จริง",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9711",
    "name_email": "chanida_la@thanulux.com",
    "full_name": "ชนิดา เหล่าประชาสกุล",
    "disp_name": "Chanida Laoprachasakul",
    "nick_name": "เก๋",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9716",
    "name_email": "Somyos_ch@thanulux.com",
    "full_name": "สมยศ ชูเผือก",
    "disp_name": "Somyos Chupuak",
    "nick_name": "เปี๊ยก",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9724",
    "name_email": "natjittakan_guy@thanulux.com",
    "full_name": "ณัฐจิตกานต์ รินทา",
    "disp_name": "Natjittakan Rinta",
    "nick_name": "นัท",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9767",
    "name_email": "prateep_so@thanulux.com",
    "full_name": "ประทีป สงสนั่น",
    "disp_name": "Prateep Songsanun",
    "nick_name": "ทีป",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9780",
    "name_email": "jakkaphan_se@thanulux.com",
    "full_name": "จักรพันธุ์ เสนาไชย",
    "disp_name": "Jakkaphan Senachai",
    "nick_name": "เจ",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9790",
    "name_email": "watcharaporn_guy@thanulux.com",
    "full_name": "วัชราพร จักร์ทิพย์",
    "disp_name": "Watcharaporn Jakthip",
    "nick_name": "เก๋",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9795",
    "name_email": "Nopporn_tr@thanulux.com",
    "full_name": "นพพร ตรัยตรึงศ์โกศล",
    "disp_name": "Nopporn Traitruengkosol",
    "nick_name": "บอย",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9800",
    "name_email": "Siriporn_t@thanulux.com",
    "full_name": "ศิริพร ธีรภาพวิวัฒน์",
    "disp_name": "Siriporn Thiraphapwiwat",
    "nick_name": "พี่พรใหญ่",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9803",
    "name_email": "Chanchira_da@thanulux.com",
    "full_name": "จันทร์จิรา บุญจามรสกุลรุ่ง",
    "disp_name": "Chanchira Danprasopsuk",
    "nick_name": "ฝัน",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9820",
    "name_email": "sirima_bu@thanulux.com",
    "full_name": "สิริมา บุญจึงมงคล",
    "disp_name": "Sirima Boomjungmongkon",
    "nick_name": "สิมา",
    "depart_id": "E8111",
    "depart_name": "แผนก Export Men's Wear",
    "cost_center": "E8111",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9851",
    "name_email": "Yupaporn@thanulux.com",
    "full_name": "ยุพาพร เจียรกุล",
    "disp_name": "Yupaporn Jearrakul",
    "nick_name": "พี่ตา",
    "depart_id": "2AM01",
    "depart_name": "แผนกบัญชี",
    "cost_center": "2AM01",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9855",
    "name_email": "saowaluk_po@thanulux.com",
    "full_name": "เสาวลักษณ์ ลาวัณย์รัตนากุล",
    "disp_name": "Saowaluk Lavanrattanakul",
    "nick_name": "ส้มโอ",
    "depart_id": "I8111",
    "depart_name": "INNER GUY LAROCHE",
    "cost_center": "I8111",
    "cost_name": "ฝ่าย Apparel 2",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9861",
    "name_email": "pharisa_pr@thanulux.com",
    "full_name": "พริศา ประชุมทอง",
    "disp_name": "Pharisa Prachumtong",
    "nick_name": "ยุ้ย",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9864",
    "name_email": "Anisa_tr@thanulux.com",
    "full_name": "อณิษา ตริยะพัฒนกำจร",
    "disp_name": "Anisa Triyapattanakamjorn",
    "nick_name": "ส้ม",
    "depart_id": "2AM01",
    "depart_name": "แผนกบัญชี",
    "cost_center": "2AM01",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9887",
    "name_email": "Somchai_we@thanulux.com",
    "full_name": "สมชาย วีรปฏิญญา",
    "disp_name": "Somchai Weerapatinya",
    "nick_name": "โจ๊ก",
    "depart_id": "2PC02",
    "depart_name": "จัดซื้อเครื่องหนัง",
    "cost_center": "2PC02",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9975",
    "name_email": "orawan_ud@thanulux.com",
    "full_name": "อรวรรณ อุดมประเสริฐ",
    "disp_name": "Orawan udomprasert",
    "nick_name": "อิ๋ว",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9989",
    "name_email": "kantima_pu@thanulux.com",
    "full_name": "กันทิมา ภู่เจริญ",
    "disp_name": "Kantima Pujaroen",
    "nick_name": "กันต์",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "T9996",
    "name_email": "Soontaree_ch@thanulux.com",
    "full_name": "สุนทรี จิตต์รุ่งเรืองชัย",
    "disp_name": "Soontaree Chitrongruangchai",
    "nick_name": "อ้อม",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "T9999",
    "name_email": "Marin@thanulux.com",
    "full_name": "มารินทร์ ลีลานุวัฒน์",
    "disp_name": "Marin Leelanuwatana",
    "nick_name": "คุณปุ่น",
    "depart_id": "1KA00",
    "depart_name": "ผู้บริหาร",
    "cost_center": "1KA00",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA115",
    "name_email": "jittraporn_pl@thanulux.com",
    "full_name": "จิตราภรณ์ เปลวเพลิง",
    "disp_name": "Jittraporn Plaewpleong",
    "nick_name": "แป้ง/ลาออก",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA193",
    "name_email": "nisakorn_ph@thanulux.com",
    "full_name": "นิศากร พลายงาม",
    "disp_name": "Nisakorn Phlayngam",
    "nick_name": "ก๊อฟ",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA233",
    "name_email": "thanchanok_si@thanulux.com",
    "full_name": "ธัญชนก ทองปอน",
    "disp_name": "Thanchanok Tongpon",
    "nick_name": "นุ่น",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA245",
    "name_email": "pattrawan_ke@thanulux.com",
    "full_name": "ภัทราวรรณ เกตุลักษณ์",
    "disp_name": "Pattrawan Kettuluk",
    "nick_name": "แพร",
    "depart_id": "2AM01",
    "depart_name": "แผนกบัญชี",
    "cost_center": "2AM01",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA250",
    "name_email": "thanat_sr@thanulux.com",
    "full_name": "ธนัช ศรีพรหม",
    "disp_name": "Thanat Sriprom",
    "nick_name": "เบียร์",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA287",
    "name_email": "Pattra_ti@thanulux.com",
    "full_name": "ภัทรา เธียรนภาพรโชค",
    "disp_name": "Pattra Tiennapapornchok",
    "nick_name": "ภัทร์",
    "depart_id": "2AM01",
    "depart_name": "แผนกบัญชี",
    "cost_center": "2AM01",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA361",
    "name_email": "sirikul_ng@thanulux.com",
    "full_name": "ศิริกุล งอกเงย",
    "disp_name": "Sirikul Ngockngoei",
    "nick_name": "เอ/ลาออก",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA366",
    "name_email": "morrakot_wi@thanulux.com",
    "full_name": "มรกต วิทยะประสาท",
    "disp_name": "Morrakot Wittayaprasart",
    "nick_name": "เหมียว",
    "depart_id": "Y8111",
    "depart_name": "แผนก OLYMP",
    "cost_center": "Y8111",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA400",
    "name_email": "kanyanut_kh@thanulux.com",
    "full_name": "กัญญาณัฐ ขุมทอง",
    "disp_name": "Kanyanut Khumthong",
    "nick_name": "เอิ้น",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA416",
    "name_email": "Nissa_sa@thanulux.com",
    "full_name": "นิสษา สอิ๊ดประเสริฐ",
    "disp_name": "Nissa Sareidprasert",
    "nick_name": "เอ",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA423",
    "name_email": "Chayaporn_pa@thanulux.com",
    "full_name": "ชยาภรณ์ พยุงศักดิ์สกุล",
    "disp_name": "Chayaporn Payungsaksakul",
    "nick_name": "ปิงปิง",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA445",
    "name_email": "Suthita_su@thanulux.com",
    "full_name": "สุธิตา สุทธิสอาด",
    "disp_name": "Suthita Sutthisaard",
    "nick_name": "มิกกี้",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA446",
    "name_email": "chutaporn_kh@thanulux.com",
    "full_name": "ชุตาภรณ์ เข็มราช",
    "disp_name": "chutaporn Khemrach",
    "nick_name": "ป่าน",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA450",
    "name_email": "Boonyathip_wu@thanulux.com",
    "full_name": "บุณยทิยพ์ หวู",
    "disp_name": "Boonyathip Wu",
    "nick_name": "ครีม/ลาออก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA451",
    "name_email": "natnicha_ha@thanulux.com",
    "full_name": "ณัฐณิชาช์ หาญยุทธ",
    "disp_name": "Natnicha Hanyut",
    "nick_name": "โพย/ลาออก",
    "depart_id": "H8111",
    "depart_name": "House Brand",
    "cost_center": "H8111",
    "cost_name": null,
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA453",
    "name_email": "chanita_pa@thanulux.com",
    "full_name": "ชนิตา พานิชเจริญ",
    "disp_name": "Chanita Panicharoen",
    "nick_name": "นุ่น",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA469",
    "name_email": "thanawin_th@thanulux.com",
    "full_name": "ธนะวินทร์ ธนนท์จิรานันท์",
    "disp_name": "Thanawin Thanontjiranan",
    "nick_name": "เอก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA472",
    "name_email": "buntharika_pur@thanulux.com",
    "full_name": "บุญฑริกา ปุริเส",
    "disp_name": "Buntharika",
    "nick_name": "เล็ก/ลาออก",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA473",
    "name_email": "watchanon_wat@thanulux.com",
    "full_name": "วัฒน์ชานนท์ วัฒนสิทธิ์กุล",
    "disp_name": "watchanon_wat",
    "nick_name": "เท็น",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA474",
    "name_email": "rapeephan_th@thanulux.com",
    "full_name": "ระพีพรรณ์ ธรรมมิยะ",
    "disp_name": "Rapeephan_th",
    "nick_name": "ป๋อม",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA475",
    "name_email": "jantima.kae@thanulux.com",
    "full_name": "จันทิมา แก้วพลอย",
    "disp_name": "jantima.kae",
    "nick_name": "ตุ๊กตา",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA476",
    "name_email": "daranee.isu@thanulux.com",
    "full_name": "ดารณี ไอศุริยวงศ์",
    "disp_name": "Daranee Isuriyawong",
    "nick_name": "หญิง",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA479",
    "name_email": "sonthaya_sr@thanulux.com",
    "full_name": "สนธยา ศรีไทยวัฒนา",
    "disp_name": "Sonthaya Srithaiwattana",
    "nick_name": "เจี๊ยบ",
    "depart_id": "L8111",
    "depart_name": "เครื่องหนัง GUY LAROCHE",
    "cost_center": "L8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนังในประเทศ",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA480",
    "name_email": "nutsuda_pe@thanulux.com",
    "full_name": "ณัฐสุดา เปี่ยมขำดี",
    "disp_name": "nutsuda_pe",
    "nick_name": "ตาล",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA483",
    "name_email": "katayut_bo@thanulux.com",
    "full_name": "คทายุทธ์ บุญสุข",
    "disp_name": "Katayut Bonsuk",
    "nick_name": "ยุทธ",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA486",
    "name_email": "taksaon.san@thanulux.com",
    "full_name": "ทักษอร สันทศนะสุวรรณ",
    "disp_name": "Taksaon Santasanasuwan",
    "nick_name": "โบว์",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA488",
    "name_email": "kittisak_ji@thanulux.com",
    "full_name": "กิตติศักดิ์ จินากุล",
    "disp_name": "Kittisak Jinnakul",
    "nick_name": "เบิร์ด",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA494",
    "name_email": "nantawan_vo@thanulux.com",
    "full_name": "นันทวรรณ วงศ์จิรภาส",
    "disp_name": "Nantawan Vongjirapas",
    "nick_name": "อร",
    "depart_id": "R8111",
    "depart_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "cost_center": "R8111",
    "cost_name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT",
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA498",
    "name_email": "thatthep_tr@thanulux.com",
    "full_name": "ทัตเทพ ตรีมงคล",
    "disp_name": "Thatthep Tremongkon",
    "nick_name": "เทอร์โบ",
    "depart_id": "2AM04",
    "depart_name": "ฝ่าย IT",
    "cost_center": "2AM04",
    "cost_name": "ฝ่าย IT",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA500",
    "name_email": "nudjanart_pro@thanulux.com",
    "full_name": "นุจนาฏ ปรองพิมาย",
    "disp_name": "Nudjanart Prongpimai",
    "nick_name": "อั้ม",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA501",
    "name_email": "sunisa_ka@thanulux.com",
    "full_name": "สุนิสา แก้วอุดม",
    "disp_name": "Sunisa Kaewudom",
    "nick_name": "ตุ๊กติ๊ก",
    "depart_id": "G8111",
    "depart_name": "ฝ่าย Guy Laroche",
    "cost_center": "G8111",
    "cost_name": "ฝ่าย Guy Laroche",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA502",
    "name_email": "pichai_ma@thanulux.com",
    "full_name": "พิชัย มาตุภูมานนท์",
    "disp_name": "Pichai Matupumanon",
    "nick_name": "พอร์ช/ลาออก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA505",
    "name_email": "nattakarn_ta@thanulux.com",
    "full_name": "ณัฐกาญจน์ ทวีรติธรรม",
    "disp_name": "Nattakarn Taweeratitham",
    "nick_name": "หมวย/ลาออก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA507",
    "name_email": "bordee_na@thanulux.com",
    "full_name": "บดี นาถนิติธาดา",
    "disp_name": "Bordee Natnititada",
    "nick_name": "ฟู/ลาออก",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA515",
    "name_email": "setthawat_ho@thanulux.com",
    "full_name": "เศรษฐวัจน์ หงษ์เทศ",
    "disp_name": "Setthawat Hongted",
    "nick_name": "แบงค์",
    "depart_id": "A8111",
    "depart_name": "ฝ่าย ARROW",
    "cost_center": "A8111",
    "cost_name": "ฝ่าย ARROW",
    "license_id": "OF002",
    "license_name": "Microsoft 365 Business Basic"
  },
  {
    "emp_id": "TA909",
    "name_email": "Weegul_ch@thanulux.com",
    "full_name": "วีกูล ชื่นใจดี",
    "disp_name": "Weegul Chuenjaidee",
    "nick_name": "แบงค์",
    "depart_id": "2AM05",
    "depart_name": "แผนกตรวจสอบภายใน",
    "cost_center": "2AM05",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  },
  {
    "emp_id": "TA966",
    "name_email": "supranee_ch@thanulux.com",
    "full_name": "สุปราณี จันทสูตร",
    "disp_name": "Supranee Chantasoot",
    "nick_name": "จอย",
    "depart_id": "3OS03",
    "depart_name": "ฝ่ายจัดซื้อและ Sourcing.",
    "cost_center": "3OS03",
    "cost_name": null,
    "license_id": "OF001",
    "license_name": "Microsoft 365 Business Standard"
  }
];
const departments = [
  {
    "id": "1KA00",
    "name": "ผู้บริหาร"
  },
  {
    "id": "1RI01",
    "name": "หน่วยวิจัยพัฒนา"
  },
  {
    "id": "2AM01",
    "name": "แผนกบัญชี"
  },
  {
    "id": "2AM03",
    "name": "ฝ่าย HR"
  },
  {
    "id": "2AM04",
    "name": "ฝ่าย IT"
  },
  {
    "id": "2AM05",
    "name": "แผนกตรวจสอบภายใน"
  },
  {
    "id": "2PC02",
    "name": "จัดซื้อเครื่องหนัง"
  },
  {
    "id": "2RE01",
    "name": "ฝ่ายพัฒนาธุรกิจ"
  },
  {
    "id": "3OS03",
    "name": "ฝ่ายจัดซื้อและ Sourcing."
  },
  {
    "id": "A8111",
    "name": "ฝ่าย ARROW"
  },
  {
    "id": "C8111",
    "name": "ABSORBA"
  },
  {
    "id": "E8111",
    "name": "แผนก Export Men's Wear"
  },
  {
    "id": "G8111",
    "name": "ฝ่าย Guy Laroche"
  },
  {
    "id": "H8111",
    "name": "House Brand"
  },
  {
    "id": "I8111",
    "name": "INNER GUY LAROCHE"
  },
  {
    "id": "L8111",
    "name": "เครื่องหนัง GUY LAROCHE"
  },
  {
    "id": "P8111",
    "name": "แผนกธุรกิจ DAKS"
  },
  {
    "id": "R8111",
    "name": "ฝ่ายธุรกิจเครื่องหนัง EXPORT"
  },
  {
    "id": "Y8111",
    "name": "แผนก OLYMP"
  }
];
const outputDir = "C:/inetpub/wwwroot/O365/outputs/019fd163-39d9-7303-8501-8b9ea8fadd38";
const outputPath = outputDir + "/o365_license_report_2026-08-05.xlsx";
const previewDir = outputDir + "/previews";

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const wb = Workbook.create();
const summary = wb.worksheets.add("สรุป");
const allSheet = wb.worksheets.add("รายชื่อทั้งหมด");
const basicSheet = wb.worksheets.add("Business Basic");
const standardSheet = wb.worksheets.add("Business Standard");

const colors = {
  navy: "#17365D",
  blue: "#2F75B5",
  lightBlue: "#D9EAF7",
  teal: "#0F6B78",
  green: "#548235",
  gold: "#BF9000",
  lightGray: "#F3F6F9",
  border: "#D6DEE8",
  white: "#FFFFFF",
  text: "#1F2937"
};

const headers = ["ลำดับ","รหัสพนักงาน","ชื่อ-นามสกุล","ชื่อภาษาอังกฤษ","ชื่อเล่น","อีเมล","รหัสแผนก","แผนก","Cost Center","ชื่อ Cost Center","รหัส License","ประเภท License"];

function makeRows(source) {
  return source.map((x, i) => [
    i + 1,
    x.emp_id ?? "",
    x.full_name ?? "",
    x.disp_name ?? "",
    x.nick_name ?? "",
    x.name_email ?? "",
    x.depart_id ?? "",
    x.depart_name ?? "",
    x.cost_center ?? "",
    x.cost_name ?? "",
    x.license_id ?? "",
    x.license_name ?? ""
  ]);
}

function styleListSheet(sheet, title, source, accent, tableName) {
  sheet.showGridLines = false;
  sheet.getRange("A1:L1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1:L1").format = {
    fill: accent,
    font: { bold: true, color: colors.white, size: 16 },
    verticalAlignment: "center"
  };
  sheet.getRange("A1:L1").format.rowHeight = 30;

  sheet.getRange("A2:L2").merge();
  sheet.getRange("A2").values = [["ข้อมูลจากระบบ O365 License (read-only) ณ วันที่ 5 สิงหาคม 2026 | จำนวน " + source.length + " รายการ"]];
  sheet.getRange("A2:L2").format = {
    fill: colors.lightGray,
    font: { color: "#475569", italic: true },
    verticalAlignment: "center"
  };
  sheet.getRange("A2:L2").format.rowHeight = 22;

  sheet.getRange("A4:L4").values = [headers];
  const rows = makeRows(source);
  if (rows.length) sheet.getRangeByIndexes(4, 0, rows.length, headers.length).values = rows;

  const lastRow = 4 + rows.length;
  const table = sheet.tables.add("A4:L" + lastRow, true, tableName);
  table.style = "TableStyleMedium2";
  table.showBandedRows = true;
  table.showFilterButton = true;

  sheet.getRange("A4:L4").format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true
  };
  sheet.getRange("A4:L4").format.rowHeight = 32;
  if (rows.length) {
    sheet.getRange("A5:A" + lastRow).format.horizontalAlignment = "center";
    sheet.getRange("B5:B" + lastRow).format.horizontalAlignment = "center";
    sheet.getRange("G5:G" + lastRow).format.horizontalAlignment = "center";
    sheet.getRange("I5:I" + lastRow).format.horizontalAlignment = "center";
    sheet.getRange("K5:K" + lastRow).format.horizontalAlignment = "center";
    sheet.getRange("A5:L" + lastRow).format.verticalAlignment = "center";
    sheet.getRange("A5:L" + lastRow).format.rowHeight = 21;
  }

  const widths = [8,14,24,24,13,30,13,28,14,27,14,34];
  widths.forEach((w, i) => sheet.getRangeByIndexes(0, i, lastRow, 1).format.columnWidth = w);
  sheet.freezePanes.freezeRows(4);
  sheet.freezePanes.freezeColumns(2);
}

const basic = items.filter(x => x.license_name === "Microsoft 365 Business Basic");
const standard = items.filter(x => x.license_name === "Microsoft 365 Business Standard");

styleListSheet(allSheet, "รายชื่อผู้ได้รับ Microsoft 365 License ทั้งหมด", items, colors.blue, "AllLicensesTable");
styleListSheet(basicSheet, "รายชื่อ Microsoft 365 Business Basic", basic, colors.teal, "BusinessBasicTable");
styleListSheet(standardSheet, "รายชื่อ Microsoft 365 Business Standard", standard, colors.green, "BusinessStandardTable");

summary.showGridLines = false;
summary.getRange("A1:E1").merge();
summary.getRange("A1").values = [["สรุป Microsoft 365 License แยกตามประเภทและแผนก"]];
summary.getRange("A1:E1").format = {
  fill: colors.navy,
  font: { bold: true, color: colors.white, size: 16 },
  verticalAlignment: "center"
};
summary.getRange("A1:E1").format.rowHeight = 32;
summary.getRange("A2:E2").merge();
summary.getRange("A2").values = [["ข้อมูล ณ วันที่ 5 สิงหาคม 2026 | แหล่งข้อมูล: O365 License (read-only)"]];
summary.getRange("A2:E2").format = { fill: colors.lightGray, font: { italic: true, color: "#475569" } };

summary.getRange("A4:B4").values = [["รายการ","จำนวน"]];
summary.getRange("A5:A7").values = [["License ทั้งหมด"],["Business Basic"],["Business Standard"]];
summary.getRange("B5").formulas = [["=COUNTA('รายชื่อทั้งหมด'!$B$5:$B$111)"]];
summary.getRange("B6").formulas = [["=COUNTIF('รายชื่อทั้งหมด'!$L$5:$L$111,\"Microsoft 365 Business Basic\")"]];
summary.getRange("B7").formulas = [["=COUNTIF('รายชื่อทั้งหมด'!$L$5:$L$111,\"Microsoft 365 Business Standard\")"]];
summary.getRange("A4:B4").format = { fill: colors.blue, font: { bold: true, color: colors.white }, horizontalAlignment: "center" };
summary.getRange("A5:B7").format.borders = { preset: "outside", style: "thin", color: colors.border };
summary.getRange("B5:B7").format = { numberFormat: "#,##0", font: { bold: true }, horizontalAlignment: "right" };

summary.getRange("A10:E10").values = [["รหัสแผนก","แผนก","Business Basic","Business Standard","รวม"]];
summary.getRange("A10:E10").format = {
  fill: colors.navy,
  font: { bold: true, color: colors.white },
  horizontalAlignment: "center",
  wrapText: true
};
const deptRows = departments.map(d => [d.id, d.name, null, null, null]);
summary.getRangeByIndexes(10, 0, deptRows.length, 5).values = deptRows;
const deptLast = 10 + deptRows.length;
for (let row = 11; row <= deptLast; row++) {
  summary.getRange("C" + row).formulas = [["=COUNTIFS('รายชื่อทั้งหมด'!$G$5:$G$111,A" + row + ",'รายชื่อทั้งหมด'!$L$5:$L$111,\"Microsoft 365 Business Basic\")"]];
  summary.getRange("D" + row).formulas = [["=COUNTIFS('รายชื่อทั้งหมด'!$G$5:$G$111,A" + row + ",'รายชื่อทั้งหมด'!$L$5:$L$111,\"Microsoft 365 Business Standard\")"]];
  summary.getRange("E" + row).formulas = [["=SUM(C" + row + ":D" + row + ")"]];
}
const totalRow = deptLast + 1;
summary.getRange("A" + totalRow + ":B" + totalRow).merge();
summary.getRange("A" + totalRow).values = [["รวมทั้งหมด"]];
summary.getRange("C" + totalRow).formulas = [["=SUM(C11:C" + deptLast + ")"]];
summary.getRange("D" + totalRow).formulas = [["=SUM(D11:D" + deptLast + ")"]];
summary.getRange("E" + totalRow).formulas = [["=SUM(E11:E" + deptLast + ")"]];
summary.getRange("A" + totalRow + ":E" + totalRow).format = {
  fill: colors.lightBlue,
  font: { bold: true, color: colors.navy },
  borders: { preset: "doubleBottom", style: "medium", color: colors.navy }
};
summary.getRange("C11:E" + totalRow).format.numberFormat = "#,##0";
summary.getRange("A10:E" + totalRow).format.borders = {
  insideHorizontal: { style: "thin", color: colors.border },
  outside: { style: "thin", color: colors.border }
};
summary.getRange("A11:A" + deptLast).format.horizontalAlignment = "center";
summary.getRange("C11:E" + totalRow).format.horizontalAlignment = "right";
summary.getRange("A1:A" + totalRow).format.columnWidth = 16;
summary.getRange("B1:B" + totalRow).format.columnWidth = 34;
summary.getRange("C1:D" + totalRow).format.columnWidth = 18;
summary.getRange("E1:E" + totalRow).format.columnWidth = 12;
summary.freezePanes.freezeRows(10);

const checks = [];
checks.push((await wb.inspect({kind:"table", range:"สรุป!A1:E31", include:"values,formulas", tableMaxRows:31, tableMaxCols:5, maxChars:8000})).ndjson);
checks.push((await wb.inspect({kind:"table", range:"รายชื่อทั้งหมด!A1:L10", include:"values,formulas", tableMaxRows:10, tableMaxCols:12, maxChars:5000})).ndjson);
checks.push((await wb.inspect({kind:"match", searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options:{useRegex:true,maxResults:300}, summary:"final formula error scan"})).ndjson);

for (const sheetName of ["สรุป","รายชื่อทั้งหมด","Business Basic","Business Standard"]) {
  const preview = await wb.render({sheetName, autoCrop:"all", scale:1, format:"png"});
  await fs.writeFile(previewDir + "/" + sheetName.replace(/[^a-zA-Z0-9ก-๙]+/g, "_") + ".png", new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(JSON.stringify({outputPath, total:items.length, basic:basic.length, standard:standard.length, departments:departments.length, checks}, null, 2));

