const readline = require("readline");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

function prompt(question) {
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      resolve(answer);
    });
  });
}

async function main() {
  const directoryPath = await prompt(
    "Enter the directory location of the Excel files: "
  );

  if (!fs.existsSync(directoryPath)) {
    console.log("Directory does not exist!");
    rl.close();
    return;
  }

  const fileList = fs.readdirSync(directoryPath);
  const excelFiles = fileList.filter(
    (file) => file.endsWith(".xlsx") || file.endsWith(".xls")
  );

  console.log("Excel files found within the given directory:");
  excelFiles.forEach((file, index) => {
    console.log(`${index + 1}) ${path.parse(file).name}`);
  });

  const selectedFilesInput = await prompt(
    "Select the file(s) that you want to add to the new file (separated by commas): "
  );
  const selectedFiles = selectedFilesInput.split(",").map((index) => parseInt(index) - 1);

  const selectedFileNames = [];
  console.log("\nSelected files:");
  selectedFiles.forEach((fileIndex) => {
    if (fileIndex < 0 || fileIndex >= excelFiles.length) {
      console.log(`Invalid file number: ${fileIndex + 1}`);
      return;
    }
    const selectedFile = excelFiles[fileIndex];
    selectedFileNames.push(selectedFile);
    console.log(selectedFile);
  });

  const titles = [];

  console.log("\nViewing the titles within the selected files...");
  for (const selectedFile of selectedFileNames) {
    console.log(selectedFile);
    const wb = xlsx.readFile(path.join(directoryPath, selectedFile));
    const sheet = wb.Sheets[wb.SheetNames[0]];

    // Clear the titles array for each file
    titles.length = 0;

    const range = xlsx.utils.decode_range(sheet["!ref"]);
    for (let i = range.s.c; i <= range.e.c; i++) {
      const cellAddress = xlsx.utils.encode_cell({ r: range.s.r, c: i });
      titles.push(sheet[cellAddress].v);
      console.log(`${i + 1}) ${sheet[cellAddress].v}`);
    }

    const selectedTitlesInput = await prompt(
      "Select the titles of which you want to save the details (separated by commas): "
    );
    const selectedTitles = selectedTitlesInput.split(",").map((index) => parseInt(index) - 1);

    // Rest of the code...
  }

  const selectedTitlesDict = {};

  console.log("\nViewing the titles within the selected files...");
  for (const selectedFile of selectedFileNames) {
    console.log(selectedFile);
    const wb = xlsx.readFile(path.join(directoryPath, selectedFile));
    const sheet = wb.Sheets[wb.SheetNames[0]];

    // Clear the titles array for each file
    titles.length = 0;

    const range = xlsx.utils.decode_range(sheet["!ref"]);
    for (let i = range.s.c; i <= range.e.c; i++) {
      const cellAddress = xlsx.utils.encode_cell({ r: range.s.r, c: i });
      titles.push(sheet[cellAddress].v);
    }

    const selectedTitlesInput = await prompt(
      "Select the titles of which you want to save the details (separated by commas): "
    );
    const selectedTitles = selectedTitlesInput.split(",").map((index) => parseInt(index) - 1);

    selectedTitlesDict[selectedFile] = selectedTitles;
  }

  console.log("\nOverall, you have selected the following:");
  selectedFileNames.forEach((selectedFile) => {
    process.stdout.write(`\n${selectedFile} - `);
    const selectedTitles = selectedTitlesDict[selectedFile];
    selectedTitles.forEach((titleIndex, index) => {
      process.stdout.write(titles[titleIndex] + (index < selectedTitles.length - 1 ? ", " : ""));
    });
    console.log();
  });

  const mergeData = await prompt("\nDo you want to merge those data and write to a new Excel file? (y/n): ");
  if (mergeData.toLowerCase() !== "y") {
    console.log("Data merging cancelled.");
    rl.close();
    return;
  }

  const newFileName = await prompt("\nProvide the name for the new Excel file: ") + ".xlsx";
  const newFileLocation = await prompt("Where do you want to store the newly created file?: ");

  const headerRow = [];
  selectedFileNames.forEach((selectedFile) => {
    const selectedTitles = selectedTitlesDict[selectedFile];
    selectedTitles.forEach((titleIndex) => {
      headerRow.push(titles[titleIndex]);
    });
  });

  const dataDictRow = {};

  for (const selectedFile of selectedFileNames) {
    const wb = xlsx.readFile(path.join(directoryPath, selectedFile));
    const sheet = wb.Sheets[wb.SheetNames[0]];

    let rowNum = 2;
    while (sheet[xlsx.utils.encode_cell({ r: rowNum, c: 0 })]) {
      const dataRow = [];
      const selectedTitles = selectedTitlesDict[selectedFile];
      selectedTitles.forEach((titleIndex) => {
        const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: titleIndex });
        dataRow.push(sheet[cellAddress].v);
      });

      if (!dataDictRow[rowNum]) {
        dataDictRow[rowNum] = dataRow;
      } else {
        dataDictRow[rowNum].push(...dataRow);
      }

      rowNum++;
    }
  }

  const newWb = xlsx.utils.book_new();
  const newSheet = xlsx.utils.aoa_to_sheet([headerRow]);
  xlsx.utils.book_append_sheet(newWb, newSheet, "Merged Data");

  for (const [newSheetRow, dataRow] of Object.entries(dataDictRow)) {
    xlsx.utils.sheet_add_aoa(newSheet, [dataRow], {
      origin: -1,
      startRow: parseInt(newSheetRow) - 1,
    });
  }

  const newFilePath = path.join(newFileLocation, newFileName);
  xlsx.writeFile(newWb, newFilePath);

  console.log(`\nData written to "${newFileName}"`);
  rl.close();
}

main();
