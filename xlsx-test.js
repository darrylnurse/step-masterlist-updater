/*
* function to create objects
*
* const arrObj
*
* turn array to obj
* new object
* array => for(int i++, < datasize){
* object[keys[i]] = array[i]
*
* return object
* }
*
* mainObj[`${arrObj(data[i]).first-name}-${arrObj(data[i]).last-name`] = arrObj(data[i]);
*
* mainObj {
*   key (fname-lname) : {
*       fname: JOHN
*       LNAME: DOE
*       ...data
*       }
* }
*
*
* sorting names
* evaluate the first character
* turn to ascii
*
* sort in the excel sheet idiot
*
* build the export object in the form of the master list
*
* -birthdays are messed up, convert that field to a string //working using date
* -export is difficult
* -conjoin to larger list
* -match column headers of that list
*
* */

const {readFileSync} = require("fs");
const nodeXLSX = require("node-xlsx").default;
const xlsx = require("xlsx");

function newStudentSpreadsheet(newSheet = "", masterSheet = "", exportName = "") {
  function xlsxToObj(url = "") {

    const worksheet = nodeXLSX.parse(readFileSync(url));
    const data = worksheet[0].data;
    const keys = data[0];

    const excelSerialToDate = serial => {
      const days = Math.floor(serial - 25568);
      const value = days * 86400;
      const date =  new Date(value * 1000 + 1);
      return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`
    }

    const removeChar = (string = "", char = '-') => {
      if (!string.includes(char)) return string;
      return string.split(char).join('');
    }

    const formatProper = (string) => {
      const str = string.toLowerCase();
      return str[0].toUpperCase() +  str.slice(1, string.length);
    }

    let xlsxObject = {};
    const arrObj = array => {
      let dataObject = {};
      for (let i = 0; i < array.length; i++) {
        if (keys[i] === "Date of Birth") {
          dataObject[keys[i]] = excelSerialToDate(array[i]);
        } //todatestring optional depending on intended formatting
        else if (keys[i] === "date") {
          dataObject[keys[i]] = new Date(Date.parse(array[i])).toDateString();
        }
        else if (keys[i] === "Parent/Guardian's Phone Number") {
          dataObject[keys[i]] = typeof array[i] !== "number" ? removeChar(array[i], '-') : array[i];
        }
        else if (keys[i] === "New York State Resident" || /(psat | sat | act)/i.test(keys[i])) {
          dataObject[keys[i]] = /yes/i.test(array[i]) ? "TRUE" : "FALSE";
        }
        else if (keys[i] === "First Name" || keys[i] === "Last Name") {
          dataObject[keys[i]] = formatProper(array[i]).trim();
        } else {
          dataObject[keys[i]] = typeof array[i] === "string" ? array[i].trim() : array[i];
        }
      }
      return dataObject;
    }

    for (let i = 1; i < data.length; i++) {
      xlsxObject[`${arrObj(data[i])["First Name"]}-${arrObj(data[i])["Last Name"]}`] = arrObj(data[i]);
    }

    return xlsxObject;
  }

  const exportForm = xlsxToObj(newSheet);
  const masterList = xlsxToObj(masterSheet);

  // for (let key in exportForm) {
  //     let student = exportForm[key];
  //     console.log(key,":", student["Date of Birth"]);
  // }

  const arrayDifference = (toCheck, mainList) => {
    let difference = [];
    for (const element of toCheck) {
      if (!mainList.includes(element)) difference.push(element); //needs to account for lower and upper case
    }
    return difference;
  }

  //console.log(Object.keys(exportForm));

  let newStudents = arrayDifference(Object.keys(exportForm), Object.keys(masterList));

  let exportStudents = [];
  for (const student of newStudents) {
    exportStudents.push(exportForm[student]);
  }

  const exportXlsx = (object, name) => {
    let workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(object);
    xlsx.utils.book_append_sheet(workbook, worksheet, name);
    xlsx.writeFile(workbook, `${name}.xlsx`);
  }

  exportXlsx(exportStudents, exportName);
  console.log(exportName, "created successfully.")
}

newStudentSpreadsheet("export-form-aug.xlsx", "8.18.24 CORRECT Master List.xlsx", "new-august-export");
