//Requiring modules
const xlsx = require('xlsx');
const fs = require('fs');
const { parse } = require('path');
const express=require('express');
const app=express();

//reading from excel file
const workbook = xlsx.readFile('./test.xlsx'); //excel file name goes here
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

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
var jsonObject=JSON.parse(fs.readFileSync('excel1.json', 'utf8')); //json config file name goes here
// var keys = Object.keys( jsonObject );


//looping through each cell in the excel file
for (let cell in worksheet) {

    const cellAsString = cell.toString();
// console.log(jsonObject[cellAsString[0]]);

    if (cellAsString[1] !== 'r' && cellAsString[1] !== 'm' && cellAsString[1] > 1) {

            post[jsonObject[cellAsString[0]]] = worksheet[cell].v;

        
        //checks for last column so as to push into the array
        if (cellAsString[0] === Object.keys(jsonObject).sort().reverse()[0]) {
            posts.push(post);
            post = {};
        }
    }
}

console.log(posts);
}

