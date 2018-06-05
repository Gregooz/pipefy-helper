'use strict';

const PIPEFY_TOKEN =
    'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJ1c2VyIjp7ImlkIjoxNDI3MzQsImVtYWlsIjoicGllZHBpcGVyQHNzY3RhMS5jb20iLCJhcHBsaWNhdGlvbiI6NDc3N319.0cEQkFkUhO1bxHBcqLjauqL31d8bJYBHOmhgMGmOM9T09yN7OiFcyMtJ8UcWuLE6rQ1WBzxESIstrwC9PKrusw';

const pipefy = require('../index')({
    accessToken: PIPEFY_TOKEN, logLevel: 'debug'
});

const XLSX = require('xlsx'); // Needed to read xlsx file
const fs = require('fs'); // Needed to write in a file
const workbook = XLSX.readFile("../extensions.xlsx"); // Location of our file

var db_json = []; // Empty JSONArray to fit Database records

pipefy.getTableRecords("8zgwlq_2").then(res => { // We  retrieve data from RFQ table
    res.forEach(element => { // For each record we get
        var jsonObj = {}; // We create a new empty JSON Object
        element.node.record_fields.forEach(e => { // We go through the list of fields of each record
            jsonObj[e.name] = e.value; // We fill the new attribute using the wanted field
        });
        db_json.push(jsonObj); // We add the newly created object in the array
    });
    // Writing the data in a file
    fs.writeFile("./db.json", JSON.stringify(db_json), err => {
        if (err) {
            return console.error(err);
        }
    });
});

var xlsx_json = XLSX.utils.sheet_to_json(workbook.Sheets.Data); // We take the Data sheet and transform it into JSON

// Writing the data in a file
fs.writeFile("./xlsx.json", JSON.stringify(xlsx_json), function (err) {
    if (err) {
        return console.error(err);
    }
});

// TODO use fusejs to search if each RFQ in xlsx.json is in db.json