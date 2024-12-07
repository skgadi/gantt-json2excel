const fs = require("fs");
const ExcelJS = require("exceljs");

/**
 * Converts Gantt JSON data to an Excel file.
 * @param {Array} data - Array of Gantt JSON objects (e.g., [{ task, start, end, color }]).
 * @param {string} outputFile - Name of the output Excel file.
 */
function ganttToExcel(data, outputFile) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Gantt Chart");

  // Add headers
  sheet.columns = [
    { header: "Task", key: "task", width: 30 },
    { header: "Start Date", key: "start", width: 15 },
    { header: "End Date", key: "end", width: 15 },
    { header: "Progress (%)", key: "progress", width: 15 },
  ];

  // Add rows
  data.forEach((row) => {
    sheet.addRow({
      task: row.task,
      start: row.start,
      end: row.end,
      progress: row.progress,
    });
  });

  // Write to file
  workbook.xlsx.writeFile(outputFile)
    .then(() => {
      console.log(`Excel file created: ${outputFile}`);
    })
    .catch((err) => {
      console.error("Error creating Excel file:", err);
    });
}

module.exports = ganttToExcel;
