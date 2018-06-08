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
const tableIDs = require('./tableIDs');

const tableToUpdate = "ilr6Z0J6";

async function loadDbData() {
    let db_array = [];
    await pipefy.getTableRecords(tableToUpdate).then(res => { // We  retrieve data from RFQ table
        res.forEach(element => { // For each record we get
            element.node.record_fields.forEach(e => { // We go through the list of fields of each record
                if (e.name === "RFQ #") {
                    db_array.push({"RFQ": e.value}); // We push the RFQ number in our array
                }
            });
        });
        // Writing the data in a file
        fs.writeFile("./db.json", JSON.stringify(db_array), err => {
            if (err) {
                return console.error(err);
            }
        });
    });
    return (db_array);
}

function loadXlsx() {
    return XLSX.utils.sheet_to_json(workbook.Sheets.Data);
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
    console.log("Spreadsheet: " + xlsx.length + " records");
    console.log("Loading database...");
    let db = await loadDbData();
    console.log("Database: " + db.length + " records");

    let fuse = new Fuse(db, options); // We have to wait for loadDb to give us the data from the database

    xlsx.forEach(async e => { // We go through the xlsx JSON Array
        let search = fuse.search(e["RFQ / Project number"]); // For each element, we look if the RFQ # is in the data given by loadDb
        if (search.length !== 0) { // If we found something
            if (result.indexOf(search[0]) === -1) { // That has not already been found
                result.push(search[0].RFQ); //We add it to our list of existing RFQs
            }
        }
    });
    console.log("Found " + result.length + " existing records in the database");
    return result;
}

searchRFQs().then(existing => { // Once we have our list of existing
    let xlsx = loadXlsx(); // We retrieve our spreadsheet data
    let index_to_remove = [];
    for (let i = 0; i < xlsx.length; i++) { // We go through it
        if (existing.indexOf(xlsx[i]["RFQ / Project number"]) !== -1) // If the RFQ number exists in our list of existing
            index_to_remove.push(i); // We keep its index for deletion
    }

    console.log("Deleting existing records from the list to add");
    index_to_remove.reverse().forEach(index => { // We reverse it to handle index issues
        xlsx.splice(index, 1); // We remove the data possessing an RFQ already existing in our database
        /*fs.writeFile("./xlsx.json", JSON.stringify(xlsx), function (err) {
        if (err) {
            return console.error(err);
        }
    });*/
    });
    console.log("Inserting " + xlsx.length + " new records...");
    xlsx.forEach(async e => {
        try {
            let params = {
                "table_id": tableToUpdate,
                "title": e["RFQ / Project number"],
                "fields_attributes": [
                    {"field_id": "rfq", "field_value": e["RFQ / Project number"]},
                    {"field_id": "date_received", "field_value": e["Date of last update to this row"]},
                    {"field_id": "estimated_start", "field_value": e["Start Date"]},
                    {"field_id": "estimated_end_date", "field_value": e["TO end date"]},
                    {"field_id": "estimated_duration_hours", "field_value": e["Budget  Hours"]},
                    {"field_id": "lcat", "field_value": e["Cat."]},
                    {"field_id": "level", "field_value": e["Level"]},
                    {"field_id": "sell_rate", "field_value": e["Sell rate 2018"]},
                    {"field_id": "rfq_status", "field_value": e["Status"]}
                ]
            };
            params.fields_attributes.push({
                "field_id": "acq_requester",
                "field_value": await getRecordIDByTitle(tableIDs["ACQ Contacts"], e["ACQ"])
            });
            params.fields_attributes.push(
                {
                    "field_id": "project_manager",
                    "field_value": await getRecordIDByTitle(tableIDs["Project Managers"], e["PM"])
                }
            );

            let wl = await getRecordIDByTitle(tableIDs["Work Locations"], e["Location"]);
            params.fields_attributes.push({
                "field_id": "work_location",
                "field_value": wl
            });

            await pipefy.createTableRecord(params);

        } catch (err) {
            console.error("Could not execute query: " + err);
        }
    });
});

async function getRecordIDByTitle(table_id, title) {
    if (title !== undefined) {
        title = title.trim().split(" ");
        let ok = false;

        await pipefy.getTableRecords(table_id).then(res => { // We  retrieve data from RFQ table
            let id = undefined;
            for (let e of res) {
                title.forEach(t => {
                    ok = e.node.title.includes(t);
                });
                if (ok) {
                    id = e.node.id;
                    break;
                }
            }
            return (id);
        });
    }
    return (undefined);
}