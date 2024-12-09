import { ganttToExcel } from "../index";

const sheets = [
  {
    sheetName: "Sample Gantt",
    data: [
      {
        task: "Task 1",
        start: "2024-01-01",
        end: "2024-01-05",
        color: { argb: "FFFFF00" },
      },
      { task: "Task 2", start: "2024-01-06", end: "2024-01-10" },
      { task: "Task 3", start: "2024-01-11", end: "2024-01-15" },
    ],
    title: "Sample Gantt Chart",
    subTitle: "Q1 2024",
  },
  {
    sheetName: "Sample Gantt 2",
    data: [
      { task: "Task 1", start: "2024-01-01", end: "2024-02-07" },
      { task: "Task 2", start: "2024-03-06", end: "2024-02-12" },
      { task: "Task 3", start: "2024-01-11", end: null },
    ],
    title: "Sample Gantt Chart 2",
  },
];

// Convert to Excel
ganttToExcel(
  sheets,
  {
    outputFileName: "sample-gantt.xlsx",
    author: "John Doe",
    title: "Sample Gantt Chart",
    subTitle: "Q1 2024",
  },
  { leftPadding: 2, rightPadding: 0, minDaysForMonth: 0 }
);
