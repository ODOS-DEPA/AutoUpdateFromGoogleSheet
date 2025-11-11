import { google } from "googleapis";
import XLSX from "xlsx";
import axios from "axios";
import FormData from "form-data"; // ใช้ form-data สำหรับ Node.js
import cron from 'node-cron';

// Your Google Sheet ID
const sheetId = "1lzPLvx1PDQ9gG2xSPm_SeA7Js23qIv8ycuPnbIN5xhw";

// Authenticate using service account
const auth = new google.auth.GoogleAuth({
  keyFile: "./odos-470204-b0ff22e70862.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

// ฟังก์ชันสร้าง XLSX buffer จาก Google Sheet
async function exportSheetToXLSX() {
  try {
    const sheets = google.sheets({ version: "v4", auth });
    const range = "dump_db!A:X";

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range,
    });

    const rows = response.data.values;
    if (!rows || !rows.length) {
      console.log("No data found.");
      return null;
    }

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "dump_db");

    const xlsxBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "buffer",
    });

    console.log("XLSX buffer generated successfully!");
    return xlsxBuffer;

  } catch (error) {
    console.error("Error reading sheet:", error);
  }
}

// ฟังก์ชันส่งไฟล์และ confirm upload
async function data_transfering(fileBuffer) {
  const uploadEndpoints = [
    "/DataUpload",
    "/citizenID/upload",
    "/EnglishScore/upload",
    "/TechScore/upload"
  ];

  const dataConfirmUpload = { confirm: true };

  for (const endpoint of uploadEndpoints) {
    try {
      // สร้าง form-data
      const formData = new FormData();
      formData.append("file", fileBuffer, {
        filename: "data.xlsx",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      // ส่งไฟล์
      await axios.post(
        `https://odos.thaigov.go.th:8443${endpoint}`,
        formData,
        { headers: formData.getHeaders() } // สำคัญ ต้องใช้ getHeaders()
      );
      console.log(`✅ File uploaded successfully to ${endpoint}`);

      // ส่ง confirm upload เป็น JSON
      await axios.post(
        `https://odos.thaigov.go.th:8443${endpoint}/confirmUpload`,
        dataConfirmUpload,
        { headers: { "Content-Type": "application/json" } }
      );
      console.log(`✅ Confirm upload successful at ${endpoint}`);

    } catch (error) {
      console.error(`❌ Upload or confirm failed at ${endpoint}`, error.response?.data || error);
      throw error; // stop process หากเกิด error
    }
  }
}
// crontab : 0 0,12 * * * > update at midnight and noon
// MAIN execution
async function main() {
  cron.schedule("0 0,12 * * *", async () => { // <-- async ที่นี่
    let now = new Date();
    console.log("Current time : ", now.toLocaleTimeString());

    try {
      const xlsxBuffer = await exportSheetToXLSX();
      if (xlsxBuffer) {
        await data_transfering(xlsxBuffer);
      }
    } catch (error) {
      console.error("Error in cron process:", error);
    }
  });
}


// เรียก main
main();
