'use strict';

function getKPI(tempExcel){
  //if(typeof require !== 'undefined') XLSX = require('xlsx');
  let XLSX;
  if(typeof require !== 'undefined') XLSX = require('xlsx');
  let workbook = XLSX.readFile(tempExcel);

  let safety;//work in progress
  //so i have some change for git
  let testmy;
  let first_sheet_name = workbook.SheetNames[0];
  let address_of_cell = 'A1';
  let worksheet = workbook.Sheets[first_sheet_name];
  let desired_cell = worksheet[address_of_cell];
  
  let workbookArray = new Array();
  let workbookJson = new Array();

  let sheet_name_list = workbook.SheetNames;

  sheet_name_list.forEach(function getSafetyKPI(tempSheetnames) { /* iterate through sheets */
    let worksheet = workbook.Sheets[tempSheetnames];
    let wantedPlant = "Warren";
    if(tempSheetnames===wantedPlant){
      for (let tempValueofCell in worksheet) {
      /* all keys that do not begin with "!" correspond to cell addresses */
        if(tempValueofCell[0] === '!') continue;
          workbookArray.push(worksheet[tempValueofCell].v);//must use v, can't change
      }
      workbookJson.push(workbookArray);
      //workbookArray = new Array();
    }
    //let myKPIJSON=JSON.stringify(workbookJson);
    return workbookJson;
    //return;
  });

  console.log(workbookJson);
  //XLSX.writeFile(workbooknew, 'out.xlsx');//write out object
}

let excelUsing = 'safetyTest.xlsx';
getKPI(excelUsing);
