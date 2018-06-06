'use strict';

const PIPEFY_TOKEN =
    'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJ1c2VyIjp7ImlkIjoxNDI3MzQsImVtYWlsIjoicGllZHBpcGVyQHNzY3RhMS5jb20iLCJhcHBsaWNhdGlvbiI6NDc3N319.0cEQkFkUhO1bxHBcqLjauqL31d8bJYBHOmhgMGmOM9T09yN7OiFcyMtJ8UcWuLE6rQ1WBzxESIstrwC9PKrusw';

const pipefy = require('../index')({
    accessToken: PIPEFY_TOKEN, logLevel: 'debug'
});

const XLSX = require('xlsx'); // Needed to read xlsx file
const fs = require('fs'); // Needed to write in a file
const Fuse = require('fuse-js-latest'); // Needed to search in the json file

const workbook = XLSX.readFile("../extensions.xlsx"); // Location of our file

function loadDbData() {
    return new Promise(resolve => { // Returning a promise to synchronize
        pipefy.getTableRecords("8zgwlq_2").then(res => { // We  retrieve data from RFQ table
            let db_array = [];
            res.forEach(element => { // For each record we get
                element.node.record_fields.forEach(e => { // We go through the list of fields of each record
                    if (e.name === "RFQ #") {
                        db_array.push({"RFQ": e.value}); // We push the RFQ number in our array
                    }
                });
            });
            // Writing the data in a file
            /*fs.writeFile("./db.json", JSON.stringify(db_json), err => {
                if (err) {
                    return console.error(err);
                }
            });*/
            resolve(db_array);
        });
    });
}

function loadXlsx() {
    let xlsx_json = XLSX.utils.sheet_to_json(workbook.Sheets.Data); // We take the Data sheet and transform it into JSON
    // Writing the data in a file
    /*fs.writeFile("./xlsx.json", JSON.stringify(xlsx_json), function (err) {
        if (err) {
            return console.error(err);
        }
    });*/
    return xlsx_json;
}

async function searchRFQs() {
    let options = {
        shouldSort: true,
        findAllMatches: true,
        threshold: 0,
        location: 0,
        distance: 100,
        maxPatternLength: 32,
        minMatchCharLength: 1,
        keys: ["RFQ"]
    };
    let result = [];
    let xlsx = loadXlsx();

    let fuse = new Fuse(await loadDbData(), options); // We have to wait for loadDb to give us the data from the database

    xlsx.forEach(e => { // We go through the xlsx JSON Array
        let search = fuse.search(e["RFQ / Project number"]); // For each element, we look if the RFQ # is in the data given by loadDb
        if (search.length !== 0) { // If we found something
            if (result.indexOf(search[0]) === -1) { // That has not already been found
                result.push(search[0].RFQ); //We add it to our list of existing RFQs
            }
        }
    });
    return result;
}

searchRFQs().then(existing => { // Once we have our list of existing
    let xlsx = loadXlsx(); // We retrieve our spreadsheet data
    let index_to_remove = [];
    for (let i = 0; i < xlsx.length - 1; i++) { // We go through it
        if (existing.indexOf(xlsx[i]["RFQ / Project number"]) !== -1) // If the RFQ number exists in our list of existing
            index_to_remove.push(i); // We keep its index for deletion
    }
    
    index_to_remove.reverse().forEach(index => { // We reverse it to handle index issues
        xlsx.splice(index, 1); // We remove the data possessing an RFQ already existing in our database
    });
});