# gantt-json2excel

Convert Gantt JSON data to Excel format effortlessly with `gantt-json2excel`.

## Features

- Easily transform Gantt chart data into Excel files.
- Supports customizable formatting for Excel output.
- Simple and intuitive API.

## Installation

Install the package using npm:

```bash
npm install gantt-json2excel
```

## Usage

Here’s how you can use `gantt-json2excel`:

### Example

```javascript
import { ganttToExcel } from "gantt-json2excel";

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
      { task: "Task 3", start: "2024-01-11", end: "2024-03-15" },
    ],
    title: "Sample Gantt Chart 2",
  },
];

// Convert to Excel and obtain the ArrayBuffer
const outBuffer = ganttToExcel(
  sheets,
  {
    outputFileName: "sample-gantt.xlsx",
    author: "John Doe",
    title: "Sample Gantt Chart",
    subTitle: "Q1 2024",
  },
  { leftPadding: 2, rightPadding: 0, minDaysForMonth: 0 }
);
```

## API

### `ganttToExcel(data, outputFile)`

Converts Gantt JSON data into an Excel file.

- **Parameters**:

  - `data` (Array): Array of Gantt JSON objects with fields like `task`, `start`, `end`, and `progress`.
  - `outputFile` (String): Name of the output Excel file.

- **Returns**: None

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/my-feature`).
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin feature/my-feature`).
5. Open a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

Special thanks to the open-source community for their tools and inspiration.
