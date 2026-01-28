// const xlsx = require("xlsx");
// const multer = require("multer");
// const express = require("express");
// //const fs = require("fs");


// const app = express();
// const upload = multer({ storage: multer.memoryStorage() });

// app.post("/excel-to-json", upload.single("file"), (req, res) => {
//   try {
//     if (!req.file) {
//       return res.status(400).json({ message: "Excel file is required" });
//     }

  
//     //Read work book from buffer
//     // const workbook = xlsx.readFile("./analytics.xlsx");
//     const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
//     //const workbook = xlsx.readFile("D:/RBPractice/Task2/excelForConverter.xlsx");
//     const sheetName = workbook.SheetNames[0];
//     const sheet = workbook.Sheets[sheetName];

//     let jsonData = xlsx.utils.sheet_to_json(sheet);


//     // Manual date formatting - NO LIBRARIES!
//     const formatToIST = (jsDate) => {
//       const year = jsDate.getFullYear();
//       const month = String(jsDate.getMonth() + 1).padStart(2, "0");
//       const day = String(jsDate.getDate()).padStart(2, "0");
//       const hours = String(jsDate.getHours()).padStart(2, "0");
//       const minutes = String(jsDate.getMinutes()).padStart(2, "0");
//       const seconds = String(jsDate.getSeconds()).padStart(2, "0");
//       const milliseconds = String(jsDate.getMilliseconds()).padStart(3, "0");

//       return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}.${milliseconds}+05:30`;
//     };

//     const fromSerial = (serial) => {
//       // Excel serial date: days since 1899-12-30
//       const baseDate = new Date(1899, 11, 30); // Month is 0-indexed
//       const targetDate = new Date(
//         baseDate.getTime() + serial * 24 * 60 * 60 * 1000,
//       );
//       return formatToIST(targetDate);
//     };

//     // MANUAL DD/MM/YYYY parsing - NO LIBRARIES!
//     /*function fromDDMMYYYY(str) {
//       console.log("Input:", str);
//       const parts = str.split("/");
//       const day = parseInt(parts[0], 10);
//       const month = parseInt(parts[1], 10);
//       const year = parseInt(parts[2], 10);

//       // Create JavaScript Date: new Date(year, month-1, day)
//       const jsDate = new Date(year, month - 1, day);
//       const result = formatToIST(jsDate);
//       console.log("Output:", result);
//       return result;
//     }*/

//     const fromDDMMYYYY = (str) => {
//       const [day, month, year] = str.split("/").map(Number);
//       return formatToIST(new Date(year, month - 1, day));
//     };

//     // MANUAL DD/MM/YYYY HH:mm parsing
//     /*const fromDDMMYYYYHHmm = (str) => {
//       console.log("Input with time:", str);
//       const [datePart, timePart] = str.split(" ");
//       const [day, month, year] = datePart.split("/").map((x) => parseInt(x, 10));
//       const [hours, minutes] = timePart.split(":").map((x) => parseInt(x, 10));

//       const jsDate = new Date(year, month - 1, day, hours, minutes);
//       const result = formatToIST(jsDate);
//       console.log("Output:", result);
//       return result;
//     };*/

//     const fromDDMMYYYYHHmm = (str) => {
//       const [datePart, timePart] = str.split(" ");
//       const [day, month, year] = datePart.split("/").map(Number);
//       const [hours, minutes] = timePart.split(":").map(Number);
//         return formatToIST(new Date(year, month - 1, day, hours, minutes));
//     };

//     jsonData=jsonData.map((row)=>{
//       const newRow = { ...row  };

//       const handleDate = (value) => {
//       if (typeof value === "number") return fromSerial(value);
//       if (typeof value === "string") {
//         if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) return fromDDMMYYYY(value);
//         if (/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}$/.test(value))
//           return fromDDMMYYYYHHmm(value);
//       }
//         return value;
//       };
//     }

//     // Handle created_at
//     /*if (typeof newRow["created_at"] === "number") {
//       newRow["created_at"] = fromSerial(newRow["created_at"]);
//     } else if (typeof newRow["created_at"] === "string") {
//       if (/^\d{2}\/\d{2}\/\d{4}$/.test(newRow["created_at"])) {
//         newRow["created_at"] = fromDDMMYYYY(newRow["created_at"]);
//       } else if (
//         /^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}$/.test(newRow["created_at"])
//       ) {
//         newRow["created_at"] = fromDDMMYYYYHHmm(newRow["created_at"]);
//       }
//     }*/

//   // Handle other date columns
//   [
//     "created_date",
//     "due_date",
//     "start_date",
//     "expected_date_purchase",
//     "completed_at",
//   ].forEach((col) => {
//     if (typeof newRow[col] === "number") {
//       newRow[col] = fromSerial(newRow[col]);
//     } else if (typeof newRow[col] === "string") {
//       if (/^\d{2}\/\d{2}\/\d{4}$/.test(newRow[col])) {
//         newRow[col] = fromDDMMYYYY(newRow[col]);
//       } else if (/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}$/.test(newRow[col])) {
//         newRow[col] = fromDDMMYYYYHHmm(newRow[col]);
//       }
//     }
//   });

//   return newRow;
// });

// fs.writeFileSync("output.json", JSON.stringify(jsonData, null, 2));
// console.log("✅ Dates formatted manually - this WILL work!");
