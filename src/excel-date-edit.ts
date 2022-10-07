import { Workbook, Column } from "exceljs";
import dayjs from "dayjs";
import path from "path";
import fs from "fs";

interface TurnAroundTime {
  hr: number;
  min: number;
  identifier: string;
  hr_to_sec: number;
  min_to_sec: number;
  total_sec: number;
}

const columns: Partial<Column>[] = [
  { header: "Identifier", key: "identifier", width: 30 },
  { header: "Hour", key: "hr", width: 30 },
  { header: "Minutes", key: "min", width: 30 },
  { header: "Hour to Sec", key: "hr_to_sec", width: 30 },
  { header: "Minutes to Sec", key: "min_to_sec", width: 30 },
  { header: "Total Sec", key: "total_sec", width: 30 },
];

const getExcelD = async (filename: string) => {
  let wb: Workbook = new Workbook();

  let datafile = path.join(__dirname, "../assets", filename);
  const turnAroundTime: TurnAroundTime[] = [];
  console.log("====================================");
  console.log(`Dir: ${datafile}`);
  console.log("====================================");
  await wb.csv.readFile(datafile).then(async (sh) => {
    for (let i = 2; i <= sh.actualRowCount; i++) {
      const row = sh.getRow(i);
      const createdAt = dayjs(row.getCell("C").value as string);
      const updatedAt = dayjs(row.getCell("D").value as string);
      const total_minutes = updatedAt.diff(createdAt, "minutes");
      const hr = Math.floor(total_minutes / 60);
      const min = total_minutes % 60;
      const identifier = `${hr.toString()} hours, ${min.toString()} minutes`;
      const hr_to_sec = hr * 60 * 60;
      const min_to_sec = min * 60;
      const total_sec = total_minutes * 60;
      turnAroundTime.push({
        hr,
        min,
        identifier,
        hr_to_sec,
        min_to_sec,
        total_sec,
      });
    }
  });
  return turnAroundTime;
};

const runner = async () => {
  console.log("Start...");
  const files = fs.readdirSync(path.join(__dirname, "../assets"));

  const promises = files.map(async (file) => {
    const turnAroundTimes = await getExcelD(file);
    return {
      file,
      turnAroundTimes,
    };
  });

  const res = await Promise.all(promises);
  let wb: Workbook = new Workbook();

  await Promise.all(
    res.map(async ({ file, turnAroundTimes }) => {
      const sh = wb.addWorksheet(file);
      sh.columns = columns;
      sh.addRows(turnAroundTimes);
    })
  );

  wb.xlsx.writeFile(path.join(__dirname, "../exports", `september.xlsx`));
  console.log("Saved!");
};

runner();
