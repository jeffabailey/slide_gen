import * as ExcelJS from "exceljs";
import * as path from "path";

async function readExcelFile(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.readFile(path.join(__dirname, "slides/Speaker Target System.xlsx"));
    return wb;
  } catch (error) {
    throw new Error("Error reading the Excel file: " + (error as Error).message);
  }
}

function extractSlidesFromWorksheet(workbook: ExcelJS.Workbook): (string | undefined)[] {
  const slides: (string | undefined)[] = [];
  const ws = workbook.getWorksheet("Inputs");
  const rows = ws?.getRows(2, 21);

  if (rows) {
    rows.forEach((row, index) => {
      slides.push(row.getCell(2).value?.toString());
    });
  }

  return slides;
}

async function writeTextToFile(text: (string | undefined), fileName: string): Promise<void> {
  const fs = require("fs");
  fs.writeFile(fileName, text, (err: Error) => {
    if (err) {
      throw new Error("Error writing to file: " + err.message);
    }
  });
}

async function main() {
  try {
    const workbook = await readExcelFile();
    const slides = extractSlidesFromWorksheet(workbook);
    let slidesOutput = '---\n';
    slides.forEach((slideText) => {

      if(typeof slideText === "string") {
        slidesOutput += `${slideText}\n---\n`;
      }
    });
    const fileName = `slides/slides.md`;    
    writeTextToFile(slidesOutput, fileName);
  } catch (error) {
    console.error((error as Error).message);
  }
}

main();
