import { ganttToExcel } from '../index';

// Sample Gantt chart data
const sampleData = [
  { task: 'Task 1', start: '2024-01-01', end: '2024-01-05', color: {argb: 'FFFFF00'} },
  { task: 'Task 2', start: '2024-01-06', end: '2024-01-10' },
  { task: 'Task 3', start: '2024-01-11', end: '2024-01-15'},
];

// Convert to Excel
ganttToExcel(sampleData, 'output.xlsx');
