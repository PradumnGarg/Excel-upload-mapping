//Requiring modules
const xlsx = require('xlsx');
const fs = require('fs');
const { parse } = require('path');
const express=require('express');
const app=express();

//reading from excel file
const workbook = xlsx.readFile('./items1.xlsx'); //excel file name goes here
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// console.log(worksheet);

//To get the number of rows
var range = xlsx.utils.decode_range(worksheet['!ref']);
var num_rows = range.e.r - range.s.r ;

let posts = [];
let post = {};


//To ensure row count doesn't exceed 1000 and if does then send 404 error
if(num_rows>1000){
    app.use((req, res, next) => {
        res.status(404);
        if (req.accepts('json')) {
            res.send({ error: 'Not found' });
            return;
          }
      });
}
else{
//reading from json fike
var jsonObject=JSON.parse(fs.readFileSync('excel.json', 'utf8')); //json config file name goes here
// console.log(jsonObject);
// var keys = Object.keys( jsonObject );


//looping through each cell in the excel file
for (let cell in worksheet) {
 checkAt=1; //to ensure columns like AA are also considered

    // console.log(cell);
    const cellAsString = cell.toString();

if(cellAsString.length>2){
    checkAt=2;
}


    if (cellAsString[1] !== 'r' && cellAsString[1] !== 'm' && cellAsString[checkAt] > 1) {
              
         
            post[jsonObject[cellAsString.substring(0,checkAt)]] = worksheet[cell].v;

        //checks for last column so as to push into the array
        if (cellAsString.substring(0,checkAt) === Object.keys(jsonObject).reverse()[0]) {
            posts.push(post);
            post = {};
        }
    }
}

console.log(posts);
}

