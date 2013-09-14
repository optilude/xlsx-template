/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, describe, before, after, it, expect */
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
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                expect(t.sharedStrings).toEqual([
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
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                
                t.replaceString("Plan table", "The plan");

                t.writeSharedStrings();
                
                var text = t.archive.file("xl/sharedStrings.xml").asText();
                expect(text).not.toMatch("<si><t>Plan table</t></si>");
                expect(text).toMatch("<si><t>The plan</t></si>");

                done();
            });

        });

        it("can substitute values and generate a file", function(done) {
            
            fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), function(err, data) {
                expect(err).toBeNull();

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
                expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
                    ].find("t").text
                ).toEqual("Extracted on 2013-01-02T00:00:00.000Z");

                // revision placeholder - cell C4 changed from string to number
                expect(sheet1.find("./sheetData/row/c[@r='C4']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

                // dates placeholder - added cells
                expect(sheet1.find("./sheetData/row/c[@r='D6']").attrib.t).toEqual("d");
                expect(sheet1.find("./sheetData/row/c[@r='E6']").attrib.t).toEqual("d");
                expect(sheet1.find("./sheetData/row/c[@r='F6']").attrib.t).toEqual("d");

                expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual(new Date("2013-01-01").toISOString());
                expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual(new Date("2013-01-02").toISOString());
                expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual(new Date("2013-01-03").toISOString());

                // planData placeholder - added rows and cells
                expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John Smith");
                expect(sheet1.find("./sheetData/row/c[@r='B8']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text, 10)
                    ].find("t").text
                ).toEqual("James Smith");
                expect(sheet1.find("./sheetData/row/c[@r='B9']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B9']/v").text, 10)
                    ].find("t").text
                ).toEqual("Jim Smith");

                expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
                    ].find("t").text
                ).toEqual("Developer");
                expect(sheet1.find("./sheetData/row/c[@r='C8']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C8']/v").text, 10)
                    ].find("t").text
                ).toEqual("Analyst");
                expect(sheet1.find("./sheetData/row/c[@r='C9']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C9']/v").text, 10)
                    ].find("t").text
                ).toEqual("Manager");

                expect(sheet1.find("./sheetData/row/c[@r='D7']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='D8']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='D9']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text).toEqual("4");

                expect(sheet1.find("./sheetData/row/c[@r='E7']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='E8']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='E9']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text).toEqual("4");

                expect(sheet1.find("./sheetData/row/c[@r='F7']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='F8']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='F9']").attrib.t).toEqual("n");
                expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("4");

                // XXX: For debugging only
                // fs.writeFileSync('test.xlsx', newData, 'binary');

                done();
            });

        });

        it("moves columns left or right when filling lists", function(done) {

            fs.readFile(path.join(__dirname, 'templates', 'test-cols.xlsx'), function(err, data) {
                expect(err).toBeNull();

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
                expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text).toEqual("101");
                expect(sheet1.find("./sheetData/row/c[@r='C4']")).toBeNull();

                // C5 should have moved right, and the old B5 should now be expanded
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B5']/v").text, 10)
                    ].find("t").text
                ).toEqual("one");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C5']/v").text, 10)
                    ].find("t").text
                ).toEqual("two");
                expect(sheet1.find("./sheetData/row/c[@r='D5']/v").text).toEqual("102");
                
                // C6 should not have moved, and the old B6 should be replaced
                expect(sheet1.find("./sheetData/row/c[@r='B6']/v").text).toEqual("10");
                expect(sheet1.find("./sheetData/row/c[@r='C6']/v").text).toEqual("103");

                // XXX: For debugging only
                // fs.writeFileSync('test.xlsx', newData, 'binary');

                done();
            });

        });

        it("moves rows down when filling tables", function(done) {

            fs.readFile(path.join(__dirname, 'templates', 'test-tables.xlsx'), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                
                t.substitute(1, {
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
                expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text).toEqual("101");

                // Headers on row 6 haven't moved
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Name");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Age");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Name");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='F6']/v").text, 10)
                    ].find("t").text
                ).toEqual("Score");

                // Rows 7 contains table values for the two tables, plus the original marker in G7
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John");
                expect(sheet1.find("./sheetData/row/c[@r='C7']/v").text).toEqual("10");
                
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John");
                expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("100");

                expect(sheet1.find("./sheetData/row/c[@r='G7']/v").text).toEqual("102");

                // Row 8 contains table values, and no markers
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text, 10)
                    ].find("t").text
                ).toEqual("Bob");
                expect(sheet1.find("./sheetData/row/c[@r='C8']/v").text).toEqual("2");
                
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E8']/v").text, 10)
                    ].find("t").text
                ).toEqual("Bob");
                expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("110");

                expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

                // Row 9 contains no values for the first table, and again no markers
                expect(sheet1.find("./sheetData/row/c[@r='B9']")).toBeNull();
                expect(sheet1.find("./sheetData/row/c[@r='C9']")).toBeNull();
                
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='E9']/v").text, 10)
                    ].find("t").text
                ).toEqual("Jim");
                expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("120");

                expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

                // Row 12 contains two blank cells and a marker
                expect(sheet1.find("./sheetData/row/c[@r='B12']/v")).toBeNull();
                expect(sheet1.find("./sheetData/row/c[@r='C12']/v")).toBeNull();
                expect(sheet1.find("./sheetData/row/c[@r='D12']/v").text).toEqual("103");

                // Row 15 contains a name, two dates, and a placeholder that was shifted to the right
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B15']/v").text, 10)
                    ].find("t").text
                ).toEqual("John");
                expect(sheet1.find("./sheetData/row/c[@r='C15']/v").text).toEqual(new Date("2013-01-01").toISOString());
                expect(sheet1.find("./sheetData/row/c[@r='D15']/v").text).toEqual(new Date("2013-01-02").toISOString());
                expect(sheet1.find("./sheetData/row/c[@r='E15']/v").text).toEqual("104");

                // Row 16 contains a name and three dates
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B16']/v").text, 10)
                    ].find("t").text
                ).toEqual("Bob");
                expect(sheet1.find("./sheetData/row/c[@r='C16']/v").text).toEqual(new Date("2013-01-01").toISOString());
                expect(sheet1.find("./sheetData/row/c[@r='D16']/v").text).toEqual(new Date("2013-01-02").toISOString());
                expect(sheet1.find("./sheetData/row/c[@r='E16']/v").text).toEqual(new Date("2013-01-03").toISOString());

                // Row 17 contains a name and no dates 
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B17']/v").text, 10)
                    ].find("t").text
                ).toEqual("Jim");
                expect(sheet1.find("./sheetData/row/c[@r='C17']")).toBeNull();

                // XXX: For debugging only
                // fs.writeFileSync('test.xlsx', newData, 'binary');

                done();
            });

        });

    


    });

});