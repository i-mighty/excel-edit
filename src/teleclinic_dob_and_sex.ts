import path from "path";
import { promises as fs } from "fs";
import { Workbook, Column } from "exceljs";

interface Customer {
  dob?: string;
  sex?: string;
}

const subStrBetween = (str: string, start: string, end: string): string => {
  return str.substring(str.indexOf(start) + 1, str.lastIndexOf(end));
};

const fetchCustomersFromFiles = async () => {
  const folder = path.join(__dirname, "../assets/teleclinic");
  var filePaths = await fs.readdir(folder);
  // now search through each of them and get the dob and sex

  const files = await Promise.all(
    filePaths.map(async (filePath) =>
      fs.readFile(path.join(folder, filePath), "utf-8")
    )
  );
  //ALl the files have been retrieved as strings into the array. Now we search.
  const sexAndDOB = files.reduce<Customer[]>((prevFiles, fileString) => {
    const dateRegExp = /Date of birth is:\s*\*(\d{4}-\d{2}-\d{2})\*/;
    const sexRegExp = /Your gender is:\s*\*(Male|Female)\*/;
    const dateMatch = fileString.match(dateRegExp);
    const sexMatch = fileString.match(sexRegExp);
    let customer: Customer = {};
    if (dateMatch) {
      customer.dob = subStrBetween(dateMatch[0], "*", "*");
    }
    if (sexMatch) {
      customer.sex = subStrBetween(sexMatch[0], "*", "*");
    }
    if (customer.sex || customer.dob) {
      prevFiles.push(customer);
    }
    return prevFiles;
  }, []);

  //Now sexAndDOB is the collection of all customers that properly provided their date of birth and their sex.

  return sexAndDOB;
};

const createExcelFile = (customers: Customer[]) => {
  const columns: Partial<Column>[] = [
    { header: "Date of Birth", key: "dob", width: 50 },
    { header: "Sex/Gender", key: "sex", width: 50 },
  ];
  let wb: Workbook = new Workbook();
  const sh = wb.addWorksheet("Teleclinic Users");
  sh.columns = columns;
  sh.addRows(customers);
  wb.xlsx.writeFile(
    path.join(__dirname, "../exports", `teleclinic_dob_and_sex.xlsx`)
  );
  console.log("Saved!");
};

const runner = async () => {
  const customers = await fetchCustomersFromFiles();

  createExcelFile(customers);
};

runner();
