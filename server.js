const express = require("express");
const multer = require("multer");
const fs = require("fs");
const pdf = require("pdf-parse");
const ExcelJS = require("exceljs");
const path = require("path");
const database = require("better-sqlite3");
const open = require("open");

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

function parseInvoiceDate(dateStr) {
  if (!dateStr) return null;

  const parts = String(dateStr)
    .trim()
    .split(/[/.\-]/);
  if (parts.length !== 3) return null;

  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);
  let year = parseInt(parts[2], 10);

  if (!day || !month || !year) return null;
  if (year < 100) year += 2000;

  return new Date(year, month - 1, day, 12, 0, 0);
}

function formatDateForDb(date) {
  if (!date) return "";

  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();

  return `${dd}/${mm}/${yyyy}`;
}

function calculateTermOfPaymentDate(invoiceDateStr, termOfPayment) {
  const invoiceDate = parseInvoiceDate(invoiceDateStr);
  if (!invoiceDate || !termOfPayment) return "";

  const term = normalize(termOfPayment).toLowerCase();

  let daysToAdd = 0;
  if (/\b60\b/.test(term)) {
    daysToAdd = 60;
  } else if (/\b30\b/.test(term)) {
    daysToAdd = 30;
  } else {
    return "";
  }

  const calculatedDate = new Date(
    invoiceDate.getFullYear(),
    invoiceDate.getMonth(),
    invoiceDate.getDate(),
    12,
    0,
    0,
  );
  calculatedDate.setDate(calculatedDate.getDate() + daysToAdd);

  // Se c'è end oppure month, prevale sempre la regola fine mese, anche se c'è anche net.
  if (term.includes("end") || term.includes("month")) {
    const lastDayOfMonth = new Date(
      calculatedDate.getFullYear(),
      calculatedDate.getMonth() + 1,
      0,
      12,
      0,
      0,
    );
    return formatDateForDb(lastDayOfMonth);
  }

  // Se c'è net e non ci sono end/month, usa la data +30/+60 giorni.
  if (term.includes("net")) {
    return formatDateForDb(calculatedDate);
  }

  return "";
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
        Term_of_Payment TEXT,
        Term_of_Payment_Date TEXT
        )
    `);

    const columns = db
      .prepare(`PRAGMA table_info(users)`)
      .all()
      .map((col) => col.name);
    if (!columns.includes("Term_of_Payment_Date")) {
      db.exec(`ALTER TABLE users ADD COLUMN Term_of_Payment_Date TEXT`);
    }

    const insert = db.prepare(
      `INSERT INTO users (Customer,Po_No,Po_Date, Order_No, Order_date, Delivery_Note_No, Delivery_Date, Invoice_No, Invoice_Date, Invoice_Value, Term_of_Payment, Term_of_Payment_Date ) VALUES (?,?, ?, ?,?,?,?,?,?,?,?,?)`,
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
      const termOfPaymentDate = calculateTermOfPaymentDate(
        invoiceDate,
        termOfPayment,
      );
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
        "Term of payment Date": termOfPaymentDate,
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
        worksheet.addRow({
          ...row,
          Invoice_Date: parseInvoiceDate(row.Invoice_Date),
          Term_of_Payment_Date: parseInvoiceDate(row.Term_of_Payment_Date),
        });
      });

      const invoiceDateColumn = worksheet.getColumn("Invoice_Date");
      if (invoiceDateColumn) {
        invoiceDateColumn.numFmt = "dd/mm/yy";
      }

      const termOfPaymentDateColumn = worksheet.getColumn(
        "Term_of_Payment_Date",
      );
      if (termOfPaymentDateColumn) {
        termOfPaymentDateColumn.numFmt = "dd/mm/yy";
      }

      const invoiceValueColumn = worksheet.getColumn("Invoice_Value");
      if (invoiceValueColumn) {
        // Force two decimals in Excel while keeping Euro suffix.
        // Number format strings must use US separators internally.
        invoiceValueColumn.numFmt = '#,##0.00 "€"';
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

app.listen(3000, async () => {
  const url = "http://localhost:3000";

  console.log(`✅ Server attivo su ${url}`);

  try {
    await open.default(url, { app: { name: "chrome" } });
  } catch (err) {
    await open.default(url);
  }
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
