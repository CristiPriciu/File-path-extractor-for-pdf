const fs = require('fs'); //pentru creare fisiere
const path = require('path'); // provides utilities for working with file and directory paths
const xlsx = require('xlsx') //pentru citire si scriere fisiere .xlsx

const myXLSX = xlsx.readFile('listapdf.xlsx'); //citeste fisierul .xlsx
const pdfSheet = myXLSX.Sheets[myXLSX.SheetNames[0]]; //selecteaza primul sheet
let currentCell = 1;

const directory = "./test";

fs.readdirSync(directory).forEach(file => {

  if (fs.lstatSync(path.resolve(directory, file)).isDirectory()) {

    fs.readdirSync(directory + "/" + file).forEach(file1 => {
      
      if(fs.lstatSync(directory + "/" + file + "/" + file1).isDirectory()){

        fs.readdirSync(directory + "/" + file + "/" + file1).forEach(file2 => {

          let last3Chars = file2.slice(file2.length - 3);

          if(last3Chars == "pdf"){
            pdfString = directory + "/" + file + "/" + file1 + "/" + file2;
            console.log(pdfString.slice(1));
            xlsx.utils.sheet_add_aoa(pdfSheet, [[pdfString.slice(1)]],{origin:'A' + currentCell});
            currentCell++;
          }
        });

      }else{

        let last3Chars = file1.slice(file1.length - 3);

        if(last3Chars == "pdf"){
          pdfString = directory + "/" + file + "/" + file1;
          console.log(pdfString.slice(1));
          xlsx.utils.sheet_add_aoa(pdfSheet, [[pdfString.slice(1)]],{origin:'A' + currentCell});
          currentCell++;
        }
      };
    });

  } else {

    let last3Chars = file.slice(file.length - 3);

    if(last3Chars == "pdf"){
      pdfString = directory + "/" + file;
      console.log(pdfString.slice(1));
      xlsx.utils.sheet_add_aoa(pdfSheet, [[pdfString.slice(1)]],{origin:'A' + currentCell});
      currentCell++;
    }
  }
});

xlsx.writeFile(myXLSX, 'listapdf.xlsx');