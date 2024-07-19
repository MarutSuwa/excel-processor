const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const fs = require("fs");
const csv = require("csv-parser");

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
    },
  });

  mainWindow.loadFile("index.html");
}

app.on("ready", createWindow);

ipcMain.handle("select-file", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [
      { name: "CSV Files", extensions: ["csv"] },
      { name: "Excel Files", extensions: ["xlsx"] },
    ],
  });
  return result.filePaths[0];
});

ipcMain.handle("save-file", async (event, defaultPath) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultPath,
    filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
  });
  return result.filePath;
});

ipcMain.handle(
  "process-file",
  async (event, filePath, savePath, openTimeColumn, closeTimeColumn) => {
    let data = [];
    let headers = [];

    if (filePath.endsWith(".csv")) {
      data = await new Promise((resolve, reject) => {
        const results = [];
        fs.createReadStream(filePath)
          .pipe(csv())
          .on("headers", (csvHeaders) => {
            headers = csvHeaders;
          })
          .on("data", (data) => results.push(Object.values(data)))
          .on("end", () => resolve(results))
          .on("error", (error) => reject(error));
      });
    } else if (filePath.endsWith(".xlsx")) {
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
      });
      headers = sheet[0];
      data = sheet.slice(1);
    } else {
      throw new Error("Unsupported file type.");
    }

    const openTimeIdx = XLSX.utils.decode_col(openTimeColumn.toUpperCase());
    const closeTimeIdx = XLSX.utils.decode_col(closeTimeColumn.toUpperCase());

    function isValidDatetime(dateStr) {
      const regex = /^\d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2}:\d{2}$/;
      return regex.test(dateStr);
    }

    function convertToDatetime(dateStr) {
      const [day, month, year, hour, minute, second] = dateStr.split(/[\/ :]/);
      return new Date(year, month - 1, day, hour, minute, second);
    }

    function formatDateTime(date) {
      const pad = (num) => num.toString().padStart(2, "0");
      const day = pad(date.getDate());
      const month = pad(date.getMonth() + 1);
      const year = date.getFullYear();
      const hours = pad(date.getHours());
      const minutes = pad(date.getMinutes());
      const seconds = pad(date.getSeconds());
      return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
    }

    // Convert data in F and G columns to datetime
    data.forEach((row) => {
      if (isValidDatetime(row[openTimeIdx])) {
        row[openTimeIdx] = convertToDatetime(row[openTimeIdx]);
      }
      if (isValidDatetime(row[closeTimeIdx])) {
        row[closeTimeIdx] = convertToDatetime(row[closeTimeIdx]);
      }
    });

    function calculateDuration(openTime, closeTime) {
      const delta = closeTime - openTime;
      const days = Math.floor(delta / (1000 * 60 * 60 * 24));
      const hours = Math.floor(
        (delta % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60)
      );
      const minutes = Math.floor((delta % (1000 * 60 * 60)) / (1000 * 60));
      const seconds = Math.floor((delta % (1000 * 60)) / 1000);
      const pad = (num) => num.toString().padStart(2, "0");
      if (days > 0) {
        return `${days} day ${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
      } else {
        return `${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
      }
    }

    function evaluateDuration(duration) {
      const parts = duration.split(" ");
      if (parts.length > 1 && parts[1] === "day") {
        return "≥ 5 นาที";
      } else {
        const [hours, minutes, seconds] = parts[0].split(":").map(Number);
        const totalMinutes = hours * 60 + minutes + seconds / 60;
        return totalMinutes >= 5 ? "≥ 5 นาที" : "< 5 นาที";
      }
    }

    // Add headers for Duration and Result
    headers.push("Duration", "Result");

    const processedData = [headers];
    const below5M = [headers];
    const above5M = [headers];

    data.forEach((row) => {
      if (
        row[openTimeIdx] instanceof Date &&
        row[closeTimeIdx] instanceof Date
      ) {
        const duration = calculateDuration(row[openTimeIdx], row[closeTimeIdx]);
        const result = evaluateDuration(duration);
        const newRow = [...row, duration, result];
        processedData.push(newRow);
        if (result === "< 5 นาที") {
          below5M.push(newRow);
        } else {
          above5M.push(newRow);
        }
      } else {
        const newRow = [...row, "", ""];
        processedData.push(newRow);
      }
    });

    // Convert columns L, M, and N from text to number
    processedData.forEach((row, rowIndex) => {
      if (rowIndex > 0) {
        // skip header
        row[11] = parseFloat(row[11]) || row[11]; // L column
        row[12] = parseFloat(row[12]) || row[12]; // M column
        row[13] = parseFloat(row[13]) || row[13]; // N column
      }
    });

    const newWorkbook = new ExcelJS.Workbook();
    const allDataSheet = newWorkbook.addWorksheet("All Data");
    const below5MSheet = newWorkbook.addWorksheet("Below 5M");
    const above5MSheet = newWorkbook.addWorksheet("5M Up");

    allDataSheet.addRows(
      processedData.map((row) =>
        row.map((cell) => (cell instanceof Date ? formatDateTime(cell) : cell))
      )
    );
    below5MSheet.addRows(
      below5M.map((row) =>
        row.map((cell) => (cell instanceof Date ? formatDateTime(cell) : cell))
      )
    );
    above5MSheet.addRows(
      above5M.map((row) =>
        row.map((cell) => (cell instanceof Date ? formatDateTime(cell) : cell))
      )
    );

    const greenFill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "90EE90" },
    };
    const redFill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFB6C1" },
    };

    allDataSheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const resultCell = row.getCell(row.cellCount);
        if (resultCell.value === "≥ 5 นาที") {
          resultCell.fill = greenFill;
        } else if (resultCell.value === "< 5 นาที") {
          resultCell.fill = redFill;
        }
      }
    });

    below5MSheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const resultCell = row.getCell(row.cellCount);
        if (resultCell.value === "≥ 5 นาที") {
          resultCell.fill = greenFill;
        } else if (resultCell.value === "< 5 นาที") {
          resultCell.fill = redFill;
        }
      }
    });

    above5MSheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const resultCell = row.getCell(row.cellCount);
        if (resultCell.value === "≥ 5 นาที") {
          resultCell.fill = greenFill;
        } else if (resultCell.value === "< 5 นาที") {
          resultCell.fill = redFill;
        }
      }
    });

    await newWorkbook.xlsx.writeFile(savePath);

    return savePath;
  }
);
