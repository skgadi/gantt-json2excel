import * as ExcelJS from "exceljs";
import moment from "moment";

// Define an interface for the Gantt task
interface GanttTask {
  task: string;
  start: string;
  end: string;
  color?: {
    argb: string;
  };
}

/**
 * Converts Gantt JSON data to an Excel file.
 * @param data - Array of Gantt task objects.
 * @param outputFile - The name of the output Excel file.
 */
export function ganttToExcel(
  data: GanttTask[],
  outputFile: string,
  title?: string,
  subTitle?: string,
  dateLang?: string
): void {
  if (!data || data.length === 0) {
    console.error("No data provided.");
    return;
  }

  // Get min and max dates
  const limits: { min: Date; max: Date } = data.reduce(
    (acc, row) => {
      const start = moment(row.start).toDate();
      const end = moment(row.end).toDate();
      return {
        min: start < acc.min ? start : acc.min,
        max: end > acc.max ? end : acc.max,
      };
    },
    { min: moment(data[0].start).toDate(), max: moment(data[0].end).toDate() }
  );

  const noOfDays = moment(limits.max).diff(moment(limits.min), "days") + 1;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Gantt Chart");

  // Define columns for Excel sheet
  sheet.columns = [
    { header: "Task", key: "task", width: 30 },
    { header: "Start", key: "start", width: 15 },
    { header: "End", key: "end", width: 15 },
    ...Array.from({ length: noOfDays }, (_, i) => {
      return {
        header: moment(limits.min).add(i, "days").format("D"),
        width: 3,
      };
    }),
  ];
  sheet.addRows(data);

  // Write to file
  workbook.xlsx
    .writeFile(outputFile)
    .then(() => {
      console.log(`Excel file created: ${outputFile}`);
    })
    .catch((err) => {
      console.error("Error creating Excel file:", err);
    });
}
