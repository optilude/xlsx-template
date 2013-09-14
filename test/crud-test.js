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
                            days: [4, 4,4 ]
                        }, {
                            name: "Jim Smith",
                            role: "Manager",
                            days: [4, 4,4 ]
                        }
                    ]
                });

                var newData = t.generate(),
                    archive = new zip(newData, {base64: false, checkCRC32: true});

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // extract date placeholder - interpolated into string referenced at B4
                expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
                expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text).toEqual("8");
                expect(sharedStrings.findall("./si")[8].find("t").text).toEqual("Extracted on 2013-01-02T00:00:00.000Z");

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

    


    });

});