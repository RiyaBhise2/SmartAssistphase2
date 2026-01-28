const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.post("/exceltojson", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: "Excel file is required" });
    }

    // Read workbook from buffer
    const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    let jsonData = xlsx.utils.sheet_to_json(sheet);

    const columnMap = {
      "Create Date": "created_at",
      "Retailer Name": "dealer_name",
      "Lead Owner": "lead_owner",
      "First Name": "fname",
      "Last Name": "lname",
      "Email": "email",
      "Mobile": "mobile",
      "Lead Source": "lead_source",
      "Brand": "brand",
      "Primary Model": "PMI",
      "Enquiry Type": "enquiry_type",
      "Lead Owne": "cxp_lead_code",
      "Lead Status": "status"
    };

    // ===== Manual date formatting - NO LIBRARIES!=====
    const formatToIST = (jsDate) => {
      const year = jsDate.getFullYear();
      const month = String(jsDate.getMonth() + 1).padStart(2, "0");
      const day = String(jsDate.getDate()).padStart(2, "0");
      const hours = String(jsDate.getHours()).padStart(2, "0");
      const minutes = String(jsDate.getMinutes()).padStart(2, "0");
      const seconds = String(jsDate.getSeconds()).padStart(2, "0");
      const milliseconds = String(jsDate.getMilliseconds()).padStart(3, "0");

      return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}.${milliseconds}+05:30`;
    };

    const fromSerial = (serial) => {
      const baseDate = new Date(1899, 11, 30);
      const targetDate = new Date(
        baseDate.getTime() + serial * 24 * 60 * 60 * 1000
      );
      return formatToIST(targetDate);
    };

    const fromDDMMYYYY = (str) => {
      const [day, month, year] = str.split("/").map(Number);
      return formatToIST(new Date(year, month - 1, day));
    };

    const fromDDMMYYYYHHmm = (str) => {
      const [datePart, timePart] = str.split(" ");
      const [day, month, year] = datePart.split("/").map(Number);
      const [hours, minutes] = timePart.split(":").map(Number);
      return formatToIST(new Date(year, month - 1, day, hours, minutes));
    };
    
    jsonData = jsonData.map((row) => {
      const mappedRow = {};

      const handleDate = (value) => {
        if (typeof value === "number") return fromSerial(value);
        if (typeof value === "string") {
          if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) return fromDDMMYYYY(value);
          if (/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}$/.test(value))
            return fromDDMMYYYYHHmm(value);
        }
        return value;
      };

      // Excel → JSON mapping
      Object.keys(columnMap).forEach((excelCol) => {
        const jsonKey = columnMap[excelCol];
        if (row[excelCol] !== undefined) {
          mappedRow[jsonKey] = row[excelCol];
        }
      });

      // fixed values
      mappedRow.purchase_type = "New Vehicle";
    

      // date conversion
      if (mappedRow.created_at) {
        mappedRow.created_at = handleDate(mappedRow.created_at);
      }

      return mappedRow;
    }); 

    //  JSON RESPONSE ONLY
    return res.status(200).json(jsonData);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ message: "Failed to process Excel file" });
  }
});

app.listen(3000, () => {
  console.log(" Server running on port 3000");
});

