/**
 * Converts Gantt chart data to an Excel file.
 *
 * @param inSheets - An array of sheet details containing the Gantt chart data.
 * @param meta - Optional metadata for the Excel file.
 * @param options - Optional configuration options for the Gantt chart.
 * @returns A promise that resolves to an object containing the Excel buffer and a return code.
 *
 * @interface GSK_SHEET_DETAILS
 * @property sheetName - The name of the sheet.
 * @property title - Optional title for the sheet.
 * @property subTitle - Optional subtitle for the sheet.
 * @property data - An array of Gantt task data.
 *
 * @interface GSK_GANTT_TASK
 * @property task - The task name.
 * @property start - The start date of the task.
 * @property end - The end date of the task.
 * @property color - Optional color for the task in ARGB format.
 *
 * @interface GSK_GANTT_OPTIONS
 * @property leftPadding - Optional left padding in days.
 * @property rightPadding - Optional right padding in days.
 * @property minDaysForMonth - Optional minimum days to display for a month.
 * @property defaultColor - Optional default color for tasks.
 * @property language - Optional language for the Excel file.
 * @property borderStyle - Optional border style for the cells.
 *
 * @interface GSK_META
 * @property outputFileName - Optional output file name.
 * @property author - Optional author of the Excel file.
 * @property title - Optional title of the Excel file.
 * @property subTitle - Optional subtitle of the Excel file.
 *
 * @interface GSK_GANTT_OUTPUT
 * @property buffer - Optional buffer containing the Excel file data.
 * @property returnCode - Return code indicating the result of the operation (0 for success, 1 for failure due to no data, 2 for failure due to other reasons).
 * @property errorMessage - Optional error message if the operation failed.
 *
 * @throws Will throw an error if there is an issue during the conversion process.
 */

import * as ExcelJS from "exceljs";
import moment from "moment";

interface GSK_SHEET_DETAILS {
  sheetName: string;
  title?: string;
  subTitle?: string;
  data: GSK_GANTT_TASK[];
}

interface GSK_GANTT_TASK {
  task: string;
  start: string;
  end: string;
  color?: {
    argb: string;
  };
}

interface GSK_GANTT_OPTIONS {
  leftPadding?: number;
  rightPadding?: number;
  minDaysForMonth?: number;
  defaultColor?: string;
  language?: string;
  borderStyle?: "thick" | "thin" | "double";
}

interface GSK_META {
  outputFileName?: string;
  author?: string;
  title?: string;
  subTitle?: string;
}

interface GSK_GANTT_OUTPUT {
  buffer?: ExcelJS.Buffer;
  returnCode: number; // 0 for success, 1 for failure due to no data, 2 for failure due to other reasons
  errorMessage?: string;
}

/**
 * Converts Gantt chart data to an Excel file.
 *
 * @param inSheets - An array of sheet details containing the Gantt chart data.
 * @param meta - Optional metadata for the Excel file.
 * @param options - Optional configuration options for the Gantt chart.
 * @returns A promise that resolves to an object containing the Excel buffer and a return code.
 */
export async function ganttToExcel(
  inSheets: GSK_SHEET_DETAILS[],
  meta?: GSK_META,
  options?: GSK_GANTT_OPTIONS
): Promise<GSK_GANTT_OUTPUT> {
  try {
    if (!inSheets || inSheets.length === 0) {
      console.error("No data provided.");
      //Return empty buffer
      return {
        returnCode: 1,
        errorMessage: "No data provided.",
      };
    }

    // meta data
    const outputFile = meta?.outputFileName || "gantt.xlsx";
    const author = meta?.author || "GSK";
    const title = meta?.title || "";
    const subTitle = meta?.subTitle || "";

    const language = options?.language || "en";
    const borderStyle = options?.borderStyle || "thick";

    const workbook = new ExcelJS.Workbook();
    // set the meta data
    workbook.creator = author;
    workbook.lastModifiedBy = author;
    workbook.created = new Date();
    workbook.modified = new Date();
    // set the title and subtitle
    workbook.title = title;
    workbook.subject = subTitle;

    inSheets.forEach(async (inSheet) => {
      if (!inSheet.data || inSheet.data.length === 0) {
        console.error("No data provided.");
        //Return empty buffer
        return Buffer.from("");
      }

      const data = inSheet.data;
      // Get min and max dates
      const limits: { min: Date; max: Date } = data.reduce(
        (acc, row) => {
          let start = moment(row.start).toDate();
          let end = moment(row.end).toDate();
          if (start > end) {
            [start, end] = [end, start];
            row.start = start.toISOString();
            row.end = end.toISOString();
          }
          return {
            min: start < acc.min ? start : acc.min,
            max: end > acc.max ? end : acc.max,
          };
        },
        {
          min: moment(data[0].start).toDate(),
          max: moment(data[0].end).toDate(),
        }
      );
      //Adjust limits to start and end with few days into month
      const leftPadding = options?.leftPadding || 0;
      const rightPadding = options?.rightPadding || 0;
      const minDaysForMonth = options?.minDaysForMonth || 5;
      limits.min = moment(limits.min).subtract(leftPadding, "days").toDate();
      limits.max = moment(limits.max).add(rightPadding, "days").toDate();
      //add days to the start if the month as only fewer days than the minDaysForMonth to display
      const noOfDaysInLeftSide =
        moment(limits.min).daysInMonth() - moment(limits.min).date();
      if (noOfDaysInLeftSide < minDaysForMonth) {
        limits.min = moment(limits.min)
          .subtract(minDaysForMonth - noOfDaysInLeftSide, "days")
          .toDate();
      }
      //add days to the end if the month as only fewer days than the minDaysForMonth to display
      const noOfDaysInRightSide = moment(limits.max).date();
      if (noOfDaysInRightSide < minDaysForMonth) {
        limits.max = moment(limits.max)
          .add(minDaysForMonth - noOfDaysInRightSide, "days")
          .toDate();
      }

      const noOfDays = moment(limits.max).diff(moment(limits.min), "days") + 1;

      // sheet name should not exceed 31 characters
      // should contain only alphanumeric characters, spaces, and underscores
      // It should not previously exist in the workbook
      let sheetName = inSheet.sheetName || "Sheet";
      sheetName = sheetName.replace(/[^a-zA-Z0-9\s_]/g, "");
      sheetName = sheetName.substring(0, 25);
      // en sure the sheet name is unique
      let i = 1;
      while (workbook.getWorksheet(sheetName)) {
        sheetName = `${sheetName}_${i}`;
        i++;
      }
      const sheet = workbook.addWorksheet(sheetName);

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

      const dataToAdd = data.map((row) => {
        return {
          task: row.task,
          start: moment(row.start).format("YYYY-MM-DD"),
          end: moment(row.end).format("YYYY-MM-DD"),
        };
      });

      sheet.addRows(dataToAdd);
      sheet.getRow(1).alignment = { horizontal: "center" };

      // put  borders for all the cells
      for (let i = 0; i < data.length + 1; i++) {
        for (let j = 0; j < noOfDays + 3; j++) {
          sheet.getCell(i + 1, j + 1).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      }

      // add Month and year
      // Insert an empty row at the top
      sheet.insertRow(1, { task: "", start: "", end: "" });
      // set the first row style as center across selection
      sheet.getRow(1).alignment = { horizontal: "centerContinuous" };
      // change cell of the cell next to the last cell of the row to left align
      const totalColumns = noOfDays + 3 + 1;
      sheet.getCell(1, totalColumns).alignment = { horizontal: "left" };

      // add month and year to the first cell of everytime a new month starts
      let previousMonth = moment(new Date(0)).month();
      let previousYear = moment(new Date(0)).year();
      for (let i = 0; i < noOfDays; i++) {
        const currentDate = moment(limits.min).add(i, "days");
        if (
          currentDate.month() !== previousMonth ||
          currentDate.year() !== previousYear
        ) {
          sheet.getCell(1, i + 4).value = currentDate.format("MMM YYYY");
          previousMonth = currentDate.month();
          previousYear = currentDate.year();
        }
      }

      // put border around the newly added cells
      for (let i = 0; i < noOfDays; i++) {
        sheet.getCell(1, i + 4).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }

      // add the task color
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const start = moment(row.start).toDate();
        const end = moment(row.end).toDate();
        const startDiff = moment(start).diff(moment(limits.min), "days");
        const endDiff = moment(end).diff(moment(limits.min), "days");
        for (let j = startDiff; j <= endDiff; j++) {
          sheet.getCell(i + 3, j + 4).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: {
              argb: row.color?.argb || options?.defaultColor || "FF0000",
            },
          };
          // put thick border around all the cells
          sheet.getCell(i + 3, j + 4).border = {
            top: { style: borderStyle },
            bottom: { style: borderStyle },
          };
        }

        // put thick border around the first and last cell
        sheet.getCell(i + 3, startDiff + 4).border = {
          top: { style: borderStyle },
          left: { style: borderStyle },
          bottom: { style: borderStyle },
        };
        sheet.getCell(i + 3, endDiff + 4).border = {
          top: { style: borderStyle },
          bottom: { style: borderStyle },
          right: { style: borderStyle },
        };
      }

      // insert a row if there exists a title or subtitle
      if (inSheet.title || inSheet.subTitle) {
        sheet.insertRow(1, { task: "" });
      }

      // add the title and subtitle
      if (inSheet.title) {
        sheet.insertRow(1, { task: inSheet.title });
        sheet.getRow(1).alignment = { horizontal: "centerContinuous" };
        sheet.getCell(1, 4 + noOfDays).alignment = { horizontal: "left" };
        // Set bold and font size to the title
        sheet.getCell(1, 1).font = { bold: true, size: 16 };
      }
      if (inSheet.subTitle) {
        sheet.insertRow(2, { task: inSheet.subTitle });
        sheet.getRow(2).alignment = { horizontal: "centerContinuous" };
        sheet.getCell(2, 4 + noOfDays).alignment = { horizontal: "left" };
        // Set bold and font size to the subtitle
        sheet.getCell(2, 1).font = { bold: true, size: 14 };
      }
    });

    // Write to file
    //await workbook.xlsx.writeFile(outputFile);

    //prepare buffer
    const buffer = await workbook.xlsx.writeBuffer();
    return {
      buffer,
      returnCode: 0,
    };
  } catch (error: any) {
    console.error("Error while converting to Excel", error);
    return {
      returnCode: 2,
      errorMessage: error.message,
    };
  }
}
