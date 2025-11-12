const express = require("express");
const multer = require("multer");
const fs = require("fs");
const pdf = require("pdf-parse");
const ExcelJS = require("exceljs");
const path = require("path");
const database = require("better-sqlite3");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));
app.use(express.static("."));

function normalize(line) {
  return line.replace(/\s+/g, " ").trim();
}

function extractValue(lines, label) {
  const line = lines.find((l) => normalize(l).includes(label));
  if (!line) return "";
  const value = line.split(":")[1] || "";
  return value.trim();
}

function extractTwoValues(lines, label) {
  const line = lines.find((l) => normalize(l).includes(label));
  if (!line) return ["", ""];
  const parts = line.split(":")[1]?.split("/") || ["", ""];
  return [parts[0].trim(), parts[1]?.trim()];
}

function extractCustomerName(lines) {
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].trim() === "ITALY") {
      const block = lines.slice(i - 3, i + 1).map(normalize);
      return block.join(", ");
    }
  }
  return "";
}

function extractTermOfPayment(lines) {
  const line = lines.find((l) => normalize(l).startsWith("Term of payment:"));
  if (!line) return "";
  return line.split(":")[1]?.trim() || "";
}

function extractInvoiceValue(lines) {
  const line = lines.find((l) => normalize(l).includes("Sum of positions*"));
  if (!line) return "";
  const match = normalize(line).match(/Sum of positions\*\s*([\d.,]+)/);
  return match ? match[1] : "";
}

function convertToUSFormat(italianStr) {
  if (!italianStr) return "";
  // Rimuove il punto (migliaia), sostituisce la virgola (decimali)
  const number = parseFloat(italianStr.replace(/\./g, "").replace(",", "."));
  // Lo riconverte in stringa con migliaia separate da virgola e due decimali
  return number.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

app.post("/upload", upload.single("pdf"), async (req, res) => {
  const dataBuffer = fs.readFileSync(req.file.path);
  const data = await pdf(dataBuffer);
  const lines = data.text.split(/\r?\n/);

  try {
    const customerName = extractCustomerName(lines);
    const poNo = extractValue(lines, "PO / no");
    //const poDate = extractValue(lines, "PO / date");
    const poDate = extractValue(lines, "PO / date")?.replace(/\./g, "/") || "";

    const [orderNo, orderDateRaw] = extractTwoValues(lines, "Order no.");
    const orderDate = orderDateRaw ? orderDateRaw.replace(/\./g, "/") : "";

    const [deliveryNoteNo, deliveryDateRaw] = extractTwoValues(
      lines,
      "Deliv. note no."
    );
    const deliveryDate = deliveryDateRaw
      ? deliveryDateRaw.replace(/\./g, "/")
      : "";

    /* const [orderNo, orderDate] = extractTwoValues(lines, "Order no.");
    const [deliveryNoteNo, deliveryDate] = extractTwoValues(
      lines,
      "Deliv. note no."
    ); */
    const [invoiceNo, invoiceDateRaw] = extractTwoValues(lines, "Invoice no.");
    const invoiceDate = invoiceDateRaw
      ? invoiceDateRaw.replace(/\./g, "/")
      : "";

    const termOfPayment = extractTermOfPayment(lines);
    const invoiceValue = extractInvoiceValue(lines);
    const invoiceValueUS = convertToUSFormat(invoiceValue);

    const parsed = {
      "customer name": customerName,
      "PO no": poNo,
      "PO Date": poDate,
      "Order no": orderNo,
      "Order Date": orderDate,
      "Delivery note no": deliveryNoteNo,
      "Delivery Date": deliveryDate,
      "Invoice no": invoiceNo,
      "Invoice Date": invoiceDate,
      "Invoice Value": invoiceValueUS,
      "Term of payment": termOfPayment,
    };

    const obj = parsed;
    const values = Object.values(obj);
    console.log(values);

    const db = new database("invoices.db");

    db.exec(`
      CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY,
        Customer TEXT,
        Po_No TEXT,
        Po_Date TEXT,
        Order_No TEXT,
        Order_date TEXT,
        Delivery_Note_No TEXT,
        Delivery_Date TEXT,
        Invoice_No TEXT,
        Invoice_Date TEXT,
        Invoice_Value REAL,
        Term_of_Payment TEXT
      )        
    `);

    const insert = db.prepare(
      `INSERT INTO users (Customer,Po_No,Po_Date, Order_No, Order_date, Delivery_Note_No, Delivery_Date, Invoice_No, Invoice_Date, Invoice_Value,Term_of_Payment ) VALUES (?,?, ?, ?,?,?,?,?,?,?,?)`
    );

    insert.run(...values); // spread array elements into the placeholders

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data");

    // Replace with your actual table name
    const tableName = "users";
    const rows = db.prepare(`SELECT * FROM ${tableName}`).all();

    // Write header row
    if (rows.length > 0) {
      worksheet.columns = Object.keys(rows[0]).map((key) => ({
        header: key,
        key: key,
        width: 30,
      }));

      // Write data rows
      rows.forEach((row) => {
        worksheet.addRow(row);
      });
    }
    workbook.xlsx
      .writeFile("output.xlsx")
      .then(() => {
        console.log("Excel file created: output.xlsx");
      })
      .catch((err) => {
        console.error("Error writing Excel file:", err);
      });
    // Close the database connection
    db.close(); // ✅ Safe to close here
  } catch (err) {
    console.error("❌ Errore durante l'estrazione:", err);
    res.status(500).json({ error: "Errore durante l'estrazione dei dati." });
  }
});

app.listen(3000, () => {
  console.log("✅ Server attivo su http://localhost:3000");
});

app.get("/download", (req, res) => {
  const filePath = "./output.xlsx";
  res.download(filePath, "output.xlsx", (err) => {
    if (err) {
      console.error("Download error:", err);
    } else {
      // Delete both files after download
      fs.unlinkSync("output.xlsx");
      fs.unlinkSync("invoices.db");
      console.log("Files deleted after download");
    }
  });
});
