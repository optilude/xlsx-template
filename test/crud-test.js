/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, describe, before, after, it */
"use strict";

var buster       = require('buster'),
    XlsxTemplate = require('../lib'),
    fs           = require('fs'),
    path         = require('path'),
    zip          = require('node-zip'),
    etree        = require('elementtree');

buster.spec.expose();
buster.testRunner.timeout = 500;

describe("CRUD operations", function() {
    
    before(function(done) {
        done();
    });

    describe('XlsxTemplate', function() {

        it("can load data", function(done) {
            
            fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), function(err, data) {
                buster.expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                buster.expect(t.sharedStrings).toEqual([
                    "Name", "Role", "Plan table", "${table:planData.name}",
                    "${table:planData.role}", "${table:planData.days}",
                    "${dates}", "${revision}",
                    "Extracted on ${extractDate}"
                ]);

                done();
            });

        });

        it("can write changed shared strings", function(done) {
            
            fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), function(err, data) {
                buster.expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                
                t.replaceString("Plan table", "The plan");

                t.writeSharedStrings();
                
                var text = t.archive.file("xl/sharedStrings.xml").asText();
                buster.expect(text).not.toMatch("<si><t>Plan table</t></si>");
                buster.expect(text).toMatch("<si><t>The plan</t></si>");

                done();
            });

        });

        it("can substitute values and generate a file", function(done) {
            
            fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), function(err, data) {
                buster.expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                
                t.substitute(1, {
                    extractDate: new Date("2013-01-02"),
                    revision: 10,
                    dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
                    planData: [
                        {
                            name: "John Smith",
                            role: "Developer",
                            days: [8, 8, 4]
                        }, {
                            name: "James Smith",
                            role: "Analyst",
                            days: [4, 4, 4]
                        }, {
                            name: "Jim Smith",
                            role: "Manager",
                            days: [4, 4, 4]
                        }
                    ]
                });

                var newData = t.generate(),
                    archive = new zip(newData, {base64: false, checkCRC32: true});

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // extract date placeholder - interpolated into string referenced at B4
                buster.expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
                    ].find("t").text
                ).toEqual("Extracted on 41276");

                // revision placeholder - cell C4 changed from string to number
                buster.expect(sheet1.find("./sheetData/row/c[@r='C4']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

                // dates placeholder - added cells
                buster.expect(sheet1.find("./sheetData/row/c[@r='D6']").attrib.t).toEqual("d");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E6']").attrib.t).toEqual("d");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F6']").attrib.t).toEqual("d");

                buster.expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual("41275");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual("41276");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual("41277");

                // planData placeholder - added rows and cells
                buster.expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John Smith");
                buster.expect(sheet1.find("./sheetData/row/c[@r='B8']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text, 10)
                    ].find("t").text
                ).toEqual("James Smith");
                buster.expect(sheet1.find("./sheetData/row/c[@r='B9']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B9']/v").text, 10)
                    ].find("t").text
                ).toEqual("Jim Smith");

                buster.expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
                    ].find("t").text
                ).toEqual("Developer");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C8']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C8']/v").text, 10)
                    ].find("t").text
                ).toEqual("Analyst");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C9']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C9']/v").text, 10)
                    ].find("t").text
                ).toEqual("Manager");

                buster.expect(sheet1.find("./sheetData/row/c[@r='D7']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D8']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text).toEqual("4");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D9']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text).toEqual("4");

                buster.expect(sheet1.find("./sheetData/row/c[@r='E7']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E8']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text).toEqual("4");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E9']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text).toEqual("4");

                buster.expect(sheet1.find("./sheetData/row/c[@r='F7']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F8']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("4");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F9']").attrib.t).toEqual("n");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("4");

                // XXX: For debugging only
                // fs.writeFileSync('test.xlsx', newData, 'binary');

                done();
            });

        });

        it("moves columns left or right when filling lists", function(done) {

            fs.readFile(path.join(__dirname, 'templates', 'test-cols.xlsx'), function(err, data) {
                buster.expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                
                t.substitute(1, {
                    emptyCols: [],
                    multiCols: ["one", "two"],
                    singleCols: [10]
                });

                var newData = t.generate(),
                    archive = new zip(newData, {base64: false, checkCRC32: true});

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // C4 should have moved left, and the old B4 should now be deleted
                buster.expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text).toEqual("101");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C4']")).toBeNull();

                // C5 should have moved right, and the old B5 should now be expanded
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B5']/v").text, 10)
                    ].find("t").text
                ).toEqual("one");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C5']/v").text, 10)
                    ].find("t").text
                ).toEqual("two");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D5']/v").text).toEqual("102");
                
                // C6 should not have moved, and the old B6 should be replaced
                buster.expect(sheet1.find("./sheetData/row/c[@r='B6']/v").text).toEqual("10");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C6']/v").text).toEqual("103");

                // XXX: For debugging only
                // fs.writeFileSync('test.xlsx', newData, 'binary');

                done();
            });

        });

        it("moves rows down when filling tables", function(done) {

            fs.readFile(path.join(__dirname, 'templates', 'test-tables.xlsx'), function(err, data) {
                buster.expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                
                t.substitute("Tables", {
                    ages: [{name: "John", age: 10}, {name: "Bob", age: 2}],
                    scores: [{name: "John", score: 100}, {name: "Bob", score: 110}, {name: "Jim", score: 120}],
                    coords: [],
                    dates: [
                        {name: "John", dates: [new Date("2013-01-01"), new Date("2013-01-02")]},
                        {name: "Bob", dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")]},
                        {name: "Jim", dates: []},
                    ]
                });

                var newData = t.generate(),
                    archive = new zip(newData, {base64: false, checkCRC32: true});

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Marker above table hasn't moved
                buster.expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text).toEqual("101");

                // Headers on row 6 haven't moved
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Name");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Age");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Name");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='F6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Score");

                // Rows 7 contains table values for the two tables, plus the original marker in G7
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C7']/v").text).toEqual("10");
                
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("100");

                buster.expect(sheet1.find("./sheetData/row/c[@r='G7']/v").text).toEqual("102");

                // Row 8 contains table values, and no markers
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text, 10)
                    ].find("t").text
                ).toEqual("Bob");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C8']/v").text).toEqual("2");
                
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E8']/v").text, 10)
                    ].find("t").text
                ).toEqual("Bob");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("110");

                buster.expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

                // Row 9 contains no values for the first table, and again no markers
                buster.expect(sheet1.find("./sheetData/row/c[@r='B9']")).toBeNull();
                buster.expect(sheet1.find("./sheetData/row/c[@r='C9']")).toBeNull();
                
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E9']/v").text, 10)
                    ].find("t").text
                ).toEqual("Jim");
                buster.expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("120");

                buster.expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

                // Row 12 contains two blank cells and a marker
                buster.expect(sheet1.find("./sheetData/row/c[@r='B12']/v")).toBeNull();
                buster.expect(sheet1.find("./sheetData/row/c[@r='C12']/v")).toBeNull();
                buster.expect(sheet1.find("./sheetData/row/c[@r='D12']/v").text).toEqual("103");

                // Row 15 contains a name, two dates, and a placeholder that was shifted to the right
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B15']/v").text, 10)
                    ].find("t").text
                ).toEqual("John");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C15']/v").text).toEqual("41275");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D15']/v").text).toEqual("41276");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E15']/v").text).toEqual("104");

                // Row 16 contains a name and three dates
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B16']/v").text, 10)
                    ].find("t").text
                ).toEqual("Bob");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C16']/v").text).toEqual("41275");
                buster.expect(sheet1.find("./sheetData/row/c[@r='D16']/v").text).toEqual("41276");
                buster.expect(sheet1.find("./sheetData/row/c[@r='E16']/v").text).toEqual("41277");

                // Row 17 contains a name and no dates 
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B17']/v").text, 10)
                    ].find("t").text
                ).toEqual("Jim");
                buster.expect(sheet1.find("./sheetData/row/c[@r='C17']")).toBeNull();

                // XXX: For debugging only
                // fs.writeFileSync('test.xlsx', newData, 'binary');

                done();
            });

        });

        it("moves named tables, named cells and merged cells", function(done) {

            fs.readFile(path.join(__dirname, 'templates', 'test-named-tables.xlsx'), function(err, data) {
                buster.expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                
                t.substitute("Tables", {
                    ages: [
                        {name: "John", age: 10},
                        {name: "Bill", age: 12}
                    ],
                    days: ["Monday", "Tuesday", "Wednesday"],
                    hours: [
                        {name: "Bob", days: [10, 20, 30]},
                        {name: "Jim", days: [12, 24, 36]}
                    ],
                    progress: 100
                });

                var newData = t.generate(),
                    archive = new zip(newData, {base64: false, checkCRC32: true});

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot(),
                    workbook      = etree.parse(t.archive.file("xl/workbook.xml").asText()).getroot(),
                    table1        = etree.parse(t.archive.file("xl/tables/table1.xml").asText()).getroot(),
                    table2        = etree.parse(t.archive.file("xl/tables/table2.xml").asText()).getroot(),
                    table3        = etree.parse(t.archive.file("xl/tables/table3.xml").asText()).getroot();

                // Named ranges have moved
                buster.expect(workbook.find("./definedNames/definedName[@name='BelowTable']").text).toEqual("Tables!$B$18");
                buster.expect(workbook.find("./definedNames/definedName[@name='Moving']").text).toEqual("Tables!$G$8");
                buster.expect(workbook.find("./definedNames/definedName[@name='RangeBelowTable']").text).toEqual("Tables!$B$19:$C$19");
                buster.expect(workbook.find("./definedNames/definedName[@name='RangeRightOfTable']").text).toEqual("Tables!$E$14:$F$14");
                buster.expect(workbook.find("./definedNames/definedName[@name='RightOfTable']").text).toEqual("Tables!$F$8");

                // Merged cells have moved
                buster.expect(sheet1.find("./mergeCells/mergeCell[@ref='B2:C2']")).not.toBeNull(); // title - unchanged

                buster.expect(sheet1.find("./mergeCells/mergeCell[@ref='B10:C10']")).toBeNull(); // pushed down
                buster.expect(sheet1.find("./mergeCells/mergeCell[@ref='B12:C12']")).not.toBeNull(); // pushed down

                buster.expect(sheet1.find("./mergeCells/mergeCell[@ref='E7:F7']")).toBeNull(); // pushed down and accross
                buster.expect(sheet1.find("./mergeCells/mergeCell[@ref='G8:H8']")).not.toBeNull(); // pushed down and accross

                // Table ranges and autofilter definitions have moved
                buster.expect(table1.attrib.ref).toEqual("B4:C7"); // Grown
                buster.expect(table1.find("./autoFilter").attrib.ref).toEqual("B4:C6"); // Grown

                buster.expect(table2.attrib.ref).toEqual("B8:E10"); // Grown and pushed down
                buster.expect(table2.find("./autoFilter").attrib.ref).toEqual("B8:E10"); // Grown and pushed down

                buster.expect(table3.attrib.ref).toEqual("C14:D16"); // Grown and pushed down
                buster.expect(table3.find("./autoFilter").attrib.ref).toEqual("C14:D16"); // Grown and pushed down

                // XXX: For debugging only
                fs.writeFileSync('test.xlsx', newData, 'binary');

                done();
            });

        });

    


    });

});
