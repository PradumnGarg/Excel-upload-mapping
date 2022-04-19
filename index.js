//Requiring modules
const xlsx = require('xlsx');
const fs = require('fs');
const { parse } = require('path');

//reading from excel file
const workbook = xlsx.readFile('./test.xlsx'); //excel file name goes here
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

//reading from json fike
var jsonObject=JSON.parse(fs.readFileSync('excel.json', 'utf8')); //json config file name goes here
// var keys = Object.keys( jsonObject );


let posts = [];
let post = {};


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