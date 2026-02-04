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

function parseItalianNumber(italianStr) {
  if (!italianStr) return null;
  // Rimuove il punto (migliaia), sostituisce la virgola (decimali)
  const number = parseFloat(italianStr.replace(/\./g, "").replace(",", "."));
  if (Number.isNaN(number)) return null;
  return Number(number.toFixed(2));
}

app.post("/upload", upload.array("pdfs"), async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res
      .status(400)
      .json({ error: "Nessun file caricato. Seleziona almeno un PDF." });
  }

  const db = new database("invoices.db");

  try {
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
      `INSERT INTO users (Customer,Po_No,Po_Date, Order_No, Order_date, Delivery_Note_No, Delivery_Date, Invoice_No, Invoice_Date, Invoice_Value,Term_of_Payment ) VALUES (?,?, ?, ?,?,?,?,?,?,?,?)`,
    );

    for (const file of req.files) {
      const dataBuffer = fs.readFileSync(file.path);
      const data = await pdf(dataBuffer);
      const lines = data.text.split(/\r?\n/);

      const customerName = extractCustomerName(lines);
      const poNo = extractValue(lines, "PO / no");
      const poDate =
        extractValue(lines, "PO / date")?.replace(/\./g, "/") || "";

      const [orderNo, orderDateRaw] = extractTwoValues(lines, "Order no.");
      const orderDate = orderDateRaw ? orderDateRaw.replace(/\./g, "/") : "";

      const [deliveryNoteNo, deliveryDateRaw] = extractTwoValues(
        lines,
        "Deliv. note no.",
      );
      const deliveryDate = deliveryDateRaw
        ? deliveryDateRaw.replace(/\./g, "/")
        : "";

      const [invoiceNo, invoiceDateRaw] = extractTwoValues(
        lines,
        "Invoice no.",
      );
      const invoiceDate = invoiceDateRaw
        ? invoiceDateRaw.replace(/\./g, "/")
        : "";

      const termOfPayment = extractTermOfPayment(lines);
      const invoiceValue = extractInvoiceValue(lines);
      const invoiceValueNumber = parseItalianNumber(invoiceValue);

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
        "Invoice Value": invoiceValueNumber,
        "Term of payment": termOfPayment,
      };

      const values = Object.values(parsed);
      console.log(values);

      insert.run(...values);

      fs.unlinkSync(file.path);
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data");

    const tableName = "users";
    const rows = db.prepare(`SELECT * FROM ${tableName}`).all();

    if (rows.length > 0) {
      worksheet.columns = Object.keys(rows[0]).map((key) => ({
        header: key,
        key: key,
        width: 30,
      }));

      rows.forEach((row) => {
        worksheet.addRow(row);
      });

      const invoiceValueColumn = worksheet.getColumn("Invoice_Value");
      if (invoiceValueColumn) {
        invoiceValueColumn.numFmt = '#.##0,00 "€"';
        invoiceValueColumn.eachCell({ includeEmpty: false }, (cell) => {
          if (typeof cell.value === "number") {
            cell.value = Number(cell.value.toFixed(2));
          } else if (cell.value) {
            const parsed = parseItalianNumber(String(cell.value));
            if (parsed !== null) {
              cell.value = Number(parsed.toFixed(2));
            }
          }
        });
      }
    }
    await workbook.xlsx.writeFile("output.xlsx");
    console.log("Excel file created: output.xlsx");

    res.redirect("/download");
  } catch (err) {
    console.error("❌ Errore durante l'estrazione:", err);
    res.status(500).json({ error: "Errore durante l'estrazione dei dati." });
  } finally {
    db.close();
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
