/**
 * @date June 2018
 * @author Grégoire DECAMP
 */
'use strict';

// Pipefy token
const PIPEFY_TOKEN =
    'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJ1c2VyIjp7ImlkIjoxNDI3MzQsImVtYWlsIjoicGllZHBpcGVyQHNzY3RhMS5jb20iLCJhcHBsaWNhdGlvbiI6NDc3N319.0cEQkFkUhO1bxHBcqLjauqL31d8bJYBHOmhgMGmOM9T09yN7OiFcyMtJ8UcWuLE6rQ1WBzxESIstrwC9PKrusw';

// Pipefy utility functions
const pipefy = require('../index')({
    accessToken: PIPEFY_TOKEN, logLevel: 'debug'
});

const XLSX = require('xlsx'); // Needed to read xlsx file
const fs = require('fs'); // Needed to write in a file
const Fuse = require('fuse-js-latest'); // Needed to search in the json file

const workbook = XLSX.readFile("../extensions.xlsx"); // Location of our file
const tableIDs = require('./tableIDs'); // JSON object to get the table IDs by their name

const tableToUpdate = "8zgwlq_2"; //  ID of the table where you want to add records

/**
 * Loads the existing data from the database in an array
 * @returns {Promise<Array>} An array with the existing data from the database
 */
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

/**
 * Loads the data from the spreadsheet in an array
 * @returns {any[]} An array with the data from the spreadsheet
 */
function loadXlsx() {
    return XLSX.utils.sheet_to_json(workbook.Sheets.Data);
}

/**
 * Main function, look for existing RFQs of the sreadsheet data in the database data
 * @returns {Promise<Array>} An array filled with existing RFQs in the spreadsheet AND the database
 */
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
    });
    console.log("Inserting " + xlsx.length + " new records...");
    xlsx.forEach(async e => {
        try {
            // We put the non-connected fields in our params Object
            let params = {
                "table_id": tableToUpdate,
                "title": e["RFQ / Project number"],
                "fields_attributes": [
                    {"field_id": "rfq", "field_value": e["RFQ / Project number"]},
                    {"field_id": "date_received", "field_value": e["Date of last update to this row"]},
                    {"field_id": "estimated_start", "field_value": e["Start Date"]},
                    {"field_id": "estimated_end_date", "field_value": e["TO end date"]},
                    {"field_id": "estimated_duration_hours", "field_value": e["Budget  Hours"]},
                    {"field_id": "sell_rate", "field_value": e["Sell rate 2018"]},
                    {"field_id": "rfq_status", "field_value": e["Status"]}
                ]
            };
            // For the connected fields, we use the getRecordIDByTitle method to find the id of the desired record
            params.fields_attributes.push({
                "field_id": "lcat",
                "field_value": await getRecordIDByTitle(tableIDs["LCAT"], e["Cat."])
            });
            params.fields_attributes.push({
                "field_id": "level_1",
                "field_value": await getRecordIDByTitle(tableIDs["Levels"], e["Level"])
            });
            params.fields_attributes.push({
                "field_id": "acq_requester",
                "field_value": await getRecordIDByTitle(tableIDs["ACQ Contacts"], e["ACQ"])
            });
            params.fields_attributes.push({
                "field_id": "project_manager",
                "field_value": await getRecordIDByTitle(tableIDs["Project Managers"], e["PM"])
            });
            params.fields_attributes.push({
                "field_id": "location_2",
                "field_value": await getRecordIDByTitle(tableIDs["Work Locations"], e["Location"])
            });

            if(e["RFQ / Project number"] === "783-1")
                console.log(params);


            // This allows us to have information about how the function went
            let ret = await pipefy.createTableRecord(params);

            // Gives us informations about the error if there is one (missing field, wrong id, etc...)
            if (ret.errors !== undefined) {
                ret.errors.forEach(el => {
                    console.error(el.message);
                    console.error(e["RFQ / Project number"]);
                });
            }
            else
                console.log(ret);

        } catch (err) {
            console.error("Could not execute query: " + err);
        }
    });
});

/**
 * Rertrieves the ID of a record given its title
 * @param table_id ID of the table to look for the ID
 * @param title The value that will allow to retrieve the ID
 * @returns {Promise<T>} The ID of the given record
 */
async function getRecordIDByTitle(table_id, title) {
    if (title !== undefined) { // We have to make sure that our title isn't undefined
        title = title.trim().split(" "); // We remove the white spaces ate the beginning and at the end and we split
        let ok = false; // Boolean value to make sure the record we found is the good one

        return await pipefy.getTableRecords(table_id).then(res => { // We  retrieve data from given table
            let id = undefined;
            for (let e of res) { // We loop through the records of the table
                title.forEach(t => { // We loop through our title, which has been split into an array to make sure evry part of it corresponds to a record
                    ok = e.node.title.includes(t); // Our boolean becomes true if the value of our title has been found in a record
                }); // We continue to loop through our title array; if not ALL the values of the title match the record, the boolean value will switch to false
                if (ok) { // If we have found a matching value
                    id = e.node.id; // We take the ID of the given record
                    break; // We stop on the first record we found
                }
            }
            return (id);
        });
    }
}