/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*tslint:disable:comment-format*/
/*tslint:disable:typedef*/
/*tslint:disable:max-line-length*/
/*global require, __dirname, describe, before, it */
"use strict";

var XlsxTemplate = require("../lib"),
    fs           = require("fs"),
    path         = require("path"),
    etree        = require("elementtree");

function getSharedString(sharedStrings, sheet1, index) {
    return sharedStrings.findall("./si")[
        parseInt(sheet1.find("./sheetData/row/c[@r='" + index + "']/v").text, 10)
    ].find("t").text;
}

describe("CRUD operations", function() {

    describe("XlsxTemplate", function() {

        it("can load data", function(done) {

            fs.readFile(path.join(__dirname, "templates", "t1.xlsx"), function(err, data) {
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

            fs.readFile(path.join(__dirname, "templates", "t1.xlsx"), function(err, data) {
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

            fs.readFile(path.join(__dirname, "templates", "t1.xlsx"), function(err, data) {
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

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Dimensions should be updated
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

                // extract date placeholder - interpolated into string referenced at B4
                expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
                    ].find("t").text
                ).toEqual("Extracted on 41276");

                // revision placeholder - cell C4 changed from string to number
                expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

                // dates placeholder - added cells
                expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual("41275");
                expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual("41276");
                expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual("41277");

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

                expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text).toEqual("4");

                expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text).toEqual("4");

                expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("4");

                // XXX: For debugging only
                fs.writeFileSync("test/output/test1.xlsx", newData, "binary");

                done();
            });

        });

        it("can substitute values with descendant properties and generate a file", function(done) {

            fs.readFile(path.join(__dirname, "templates", "t2.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);

                t.substitute(1, {
                    demo: { extractDate: new Date("2013-01-02") },
                    revision: 10,
                    dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
                    planData: [
                        {
                            name: "John Smith",
                            role: { name: "Developer" },
                            days: [8, 8, 4]
                        }, {
                            name: "James Smith",
                            role: { name: "Analyst" },
                            days: [4, 4, 4]
                        }, {
                            name: "Jim Smith",
                            role: { name: "Manager" },
                            days: [4, 4, 4]
                        }
                    ]
                });

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Dimensions should be updated
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

                // extract date placeholder - interpolated into string referenced at B4
                expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
                    ].find("t").text
                ).toEqual("Extracted on 41276");

                // revision placeholder - cell C4 changed from string to number
                expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

                // dates placeholder - added cells
                expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual("41275");
                expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual("41276");
                expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual("41277");

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

                expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text).toEqual("4");

                expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text).toEqual("4");

                expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("4");
                expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("4");

                // XXX: For debugging only
                fs.writeFileSync("test/output/test2.xlsx", newData, "binary");

                done();
            });

        });

		it("can substitute values when single item array contains an object and generate a file", function(done) {

            fs.readFile(path.join(__dirname, "templates", "t3.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);

                t.substitute(1, {
                    demo: { extractDate: new Date("2013-01-02") },
                    revision: 10,
                    planData: [
                        {
                            name: "John Smith",
                            role: { name: "Developer" }
                        }
                    ]
                });

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Dimensions should be updated
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:C7");

                // extract date placeholder - interpolated into string referenced at B4
                expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
                    ].find("t").text
                ).toEqual("Extracted on 41276");

                // revision placeholder - cell C4 changed from string to number
                expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

                // planData placeholder - added rows and cells
                expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John Smith");

                expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
                    ].find("t").text
                ).toEqual("Developer");

                // XXX: For debugging only
                fs.writeFileSync("test/output/test6.xlsx", newData, "binary");

                done();
            });

        });

		it("can substitute values when single item array contains an object with sub array containing primatives and generate a file", function(done) {

            fs.readFile(path.join(__dirname, "templates", "t2.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);

                t.substitute(1, {
                    demo: { extractDate: new Date("2013-01-02") },
                    revision: 10,
                    dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
                    planData: [
                        {
                            name: "John Smith",
                            role: { name: "Developer" },
                            days: [8, 8, 4]
                        }
                    ]
                });

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Dimensions should be updated
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F7");

                // extract date placeholder - interpolated into string referenced at B4
                expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
                    ].find("t").text
                ).toEqual("Extracted on 41276");

                // revision placeholder - cell C4 changed from string to number
                expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

                // dates placeholder - added cells
                expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual("41275");
                expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual("41276");
                expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual("41277");

                // planData placeholder - added rows and cells
                expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
                    ].find("t").text
                ).toEqual("John Smith");

                expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
                    ].find("t").text
                ).toEqual("Developer");


                expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
                expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");

                // XXX: For debugging only
                fs.writeFileSync("test/output/test7.xlsx", newData, "binary");

                done();
            });

        });

        it("moves columns left or right when filling lists", function(done) {

            fs.readFile(path.join(__dirname, "templates", "test-cols.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);

                t.substitute(1, {
                    emptyCols: [],
                    multiCols: ["one", "two"],
                    singleCols: [10]
                });

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Dimensions should be set
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:E6");

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
                fs.writeFileSync("test/output/test3.xlsx", newData, "binary");

                done();
            });

        });

        it("moves rows down when filling tables", function(done) {

            fs.readFile(path.join(__dirname, "templates", "test-tables.xlsx"), function(err, data) {
                expect(err).toBeNull();

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

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Dimensions should be updated
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:H17");

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
                expect(sheet1.find("./sheetData/row/c[@r='C15']/v").text).toEqual("41275");
                expect(sheet1.find("./sheetData/row/c[@r='D15']/v").text).toEqual("41276");
                expect(sheet1.find("./sheetData/row/c[@r='E15']/v").text).toEqual("104");

                // Row 16 contains a name and three dates
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B16']/v").text, 10)
                    ].find("t").text
                ).toEqual("Bob");
                expect(sheet1.find("./sheetData/row/c[@r='C16']/v").text).toEqual("41275");
                expect(sheet1.find("./sheetData/row/c[@r='D16']/v").text).toEqual("41276");
                expect(sheet1.find("./sheetData/row/c[@r='E16']/v").text).toEqual("41277");

                // Row 17 contains a name and no dates
                expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='B17']/v").text, 10)
                    ].find("t").text
                ).toEqual("Jim");
                expect(sheet1.find("./sheetData/row/c[@r='C17']")).toBeNull();

                // XXX: For debugging only
                fs.writeFileSync("test/output/test4.xlsx", newData, "binary");

                done();
            });

        });

        it("moves rows down when filling tables with merged line cell", function(done) {

            fs.readFile(path.join(__dirname, "templates", "test-tables-merged-line.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);

                t.substitute("Tables", {
                    score_first: {name: "Jason", score: 1},
                    scores: [
                        {name: "John", score: 100, extra:'O'},
                        {name: "Bob", score: 110, extra:'O'}, 
                        {name: "Jim", score: 120, extra:'O'}
                    ],
                    score_last: {name: "Fox", score: 99},

                    score2_first: {name: "Daddy", score: 1},
                    scores2: [
                        {name: "Son1", score: 100, extra:'O'},
                        {name: "Son2", score: 110, extra:'O'}, 
                        {name: "Son3", score: 120, extra:'O'},
                        {name: "Son4", score: 130, extra:'O'}
                    ],
                    score2_last: {name: "Mom", score: 99},
                });

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                expect(sheet1.find("./sheetData/row[@r='8']/c[@r='B8']").attrib.s).toEqual("3");
                expect(sheet1.find("./sheetData/row[@r='8']/c[@r='B8']/v").text).toEqual("23");
                expect(sheet1.find("./sheetData/row[@r='8']/c[@r='C8']").attrib.s).toEqual("4");

                expect(sheet1.find("./sheetData/row[@r='8']/c[@r='D8']").attrib.s).toEqual("5");
                expect(sheet1.find("./sheetData/row[@r='8']/c[@r='E8']").attrib.s).toEqual("6");
                expect(sheet1.find("./sheetData/row[@r='8']/c[@r='F8']").attrib.s).toEqual("7");

                expect(sheet1.find("./sheetData/row[@r='18']/c[@r='B18']").attrib.s).toEqual("3");
                expect(sheet1.find("./sheetData/row[@r='18']/c[@r='C18']").attrib.s).toEqual("4");
                expect(sheet1.find("./sheetData/row[@r='18']/c[@r='D18']").attrib.s).toEqual("5");

                // XXX: For debugging only
                fs.writeFileSync("test/output/test-tables-merged-line-out.xlsx", newData, "binary");

                done();
            });

        });
        
        it("replaces hyperlinks in sheet", function(done) {
          fs.readFile(path.join(__dirname, "templates", "test-hyperlinks.xlsx"), function(err, data) {
            expect(err).toBeNull();

            var t = new XlsxTemplate(data);

            t.substitute(1, {
              email: "john@bob.com",
              subject: "hello",
              url: "http://www.google.com",
              domain: "google"
            });

            var newData = t.generate();

            var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
              sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot(),
              rels          = etree.parse(t.archive.file("xl/worksheets/_rels/sheet1.xml.rels").asText()).getroot()
            ;

            //expect(sheet1.find("./hyperlinks/hyperlink/c[@r='C16']/v").text).toEqual("41275");
            expect(rels.find("./Relationship[@Id='rId2']").attrib.Target).toEqual("http://www.google.com");
            expect(rels.find("./Relationship[@Id='rId1']").attrib.Target).toEqual("mailto:john@bob.com?subject=Hello%20hello");

            // XXX: For debugging only
            fs.writeFileSync("test/output/test9.xlsx", newData, "binary");

            done();
          });
        });

        it("moved hyperlinks in sheet", function(done) {
          fs.readFile(path.join(__dirname, "templates", "test-moved-hyperlinks.xlsx"), function(err, data) {
            expect(err).toBeNull();

            var t = new XlsxTemplate(data);

            t.substitute(1, {
              email: "john@bob.com",
              subject: "hello",
              url: "http://www.google.com",
              domain: "google",
              rows: [{
                name: 'One',
                amount: 1,
              }, {
                name: 'Two',
                amount: 2,
              }, {
                name: 'Three',
                amount: 3,
              }],
              list: ['A', 'B', 'C', 'D']
            });

            var newData = t.generate();

            var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
              sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot(),
              rels          = etree.parse(t.archive.file("xl/worksheets/_rels/sheet1.xml.rels").asText()).getroot()
            ;

            // Every hyperlink has being substituted
            expect(rels.find("./Relationship[@Id='rId1']").attrib.Target).toEqual("mailto:john@bob.com?subject=Hello%20hello");
            expect(rels.find("./Relationship[@Id='rId2']").attrib.Target).toEqual("http://www.google.com");
            expect(rels.find("./Relationship[@Id='rId3']").attrib.Target).toEqual("mailto:john@bob.com?subject=Hello%20hello");
            expect(rels.find("./Relationship[@Id='rId4']").attrib.Target).toEqual("http://www.google.com");
            expect(rels.find("./Relationship[@Id='rId5']").attrib.Target).toEqual("mailto:john@bob.com?subject=Hello%20hello");
            expect(rels.find("./Relationship[@Id='rId6']").attrib.Target).toEqual("http://www.google.com");

            // Hyperlinks have moved
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='B7']")).not.toBeNull(); // before table and list - unchanged
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='C7']")).not.toBeNull(); // before table and list - unchanged

            expect(sheet1.find("./hyperlinks/hyperlink[@ref='B14']")).toBeNull(); // pushed down
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='B16']")).not.toBeNull(); // pushed down
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='C14']")).toBeNull(); // pushed down
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='C16']")).not.toBeNull(); // pushed down

            expect(sheet1.find("./hyperlinks/hyperlink[@ref='F14']")).toBeNull(); // pushed down and accross
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='I16']")).not.toBeNull(); // pushed down and accross
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='G14']")).toBeNull(); // pushed down and accross
            expect(sheet1.find("./hyperlinks/hyperlink[@ref='J16']")).not.toBeNull(); // pushed down and accross

            // XXX: For debugging only
            fs.writeFileSync("test/output/test-moved-hyperlinks.xlsx", newData, "binary");

            done();
          });
        });

        it("moves named tables, named cells and merged cells", function(done) {

            fs.readFile(path.join(__dirname, "templates", "test-named-tables.xlsx"), function(err, data) {
                expect(err).toBeNull();

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

                var newData = t.generate();

                var sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot(),
                    workbook      = etree.parse(t.archive.file("xl/workbook.xml").asText()).getroot(),
                    table1        = etree.parse(t.archive.file("xl/tables/table1.xml").asText()).getroot(),
                    table2        = etree.parse(t.archive.file("xl/tables/table2.xml").asText()).getroot(),
                    table3        = etree.parse(t.archive.file("xl/tables/table3.xml").asText()).getroot();

                // Dimensions should be updated
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:L29");

                // Named ranges have moved
                expect(workbook.find("./definedNames/definedName[@name='BelowTable']").text).toEqual("Tables!$B$18");
                expect(workbook.find("./definedNames/definedName[@name='Moving']").text).toEqual("Tables!$G$8");
                expect(workbook.find("./definedNames/definedName[@name='RangeBelowTable']").text).toEqual("Tables!$B$19:$C$19");
                expect(workbook.find("./definedNames/definedName[@name='RangeRightOfTable']").text).toEqual("Tables!$E$14:$F$14");
                expect(workbook.find("./definedNames/definedName[@name='RightOfTable']").text).toEqual("Tables!$F$8");
                expect(workbook.find("./definedNames/definedName[@name='_xlnm.Print_Titles']").text).toEqual("Tables!$1:$3");

                // Merged cells have moved
                expect(sheet1.find("./mergeCells/mergeCell[@ref='B2:C2']")).not.toBeNull(); // title - unchanged

                expect(sheet1.find("./mergeCells/mergeCell[@ref='B10:C10']")).toBeNull(); // pushed down
                expect(sheet1.find("./mergeCells/mergeCell[@ref='B12:C12']")).not.toBeNull(); // pushed down

                expect(sheet1.find("./mergeCells/mergeCell[@ref='E7:F7']")).toBeNull(); // pushed down and accross
                expect(sheet1.find("./mergeCells/mergeCell[@ref='G8:H8']")).not.toBeNull(); // pushed down and accross

                // Table ranges and autofilter definitions have moved
                expect(table1.attrib.ref).toEqual("B4:C7"); // Grown
                expect(table1.find("./autoFilter").attrib.ref).toEqual("B4:C6"); // Grown

                expect(table2.attrib.ref).toEqual("B8:E10"); // Grown and pushed down
                expect(table2.find("./autoFilter").attrib.ref).toEqual("B8:E10"); // Grown and pushed down

                expect(table3.attrib.ref).toEqual("C14:D16"); // Grown and pushed down
                expect(table3.find("./autoFilter").attrib.ref).toEqual("C14:D16"); // Grown and pushed down

                // XXX: For debugging only
                fs.writeFileSync("test/output/test5.xlsx", newData, "binary");

                done();
            });

        });

        it("Correctly parse when formula in the file", function(done) {

            fs.readFile(path.join(__dirname, "templates", "template.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                t.substitute(1, {
                    people: [
                        {
                            name: "John Smith",
                            age: 55,
                        },
                        {
                            name: "John Doe",
                            age: 35,
                        }
                    ]
                });

                done();
            });

        });

        it("Correctly recalculate formula", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-formula.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                t.substitute(1, {
                    data: [
                        { name: "A", quantity: 10, unitCost: 3 },
                        { name: "B", quantity: 15, unitCost: 5 },
                    ]
                });

                var newData = t.generate();
                var sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
                expect(sheet1).toBeDefined();

                expect(sheet1.find("./sheetData/row/c[@r='D2']/f").text).toEqual("Table3[Qty]*Table3[UnitCost]");
                expect(sheet1.find("./sheetData/row/c[@r='D2']/v")).toBeNull();

                // This part is not working
                // expect(sheet1.find("./sheetData/row/c[@r='D3']/f").text).toEqual("Table3[Qty]*Table3[UnitCost]");

                // fs.writeFileSync('test/output/test6.xlsx', newData, 'binary');
                done();
            });
        });

        it("File without dimensions works", function(done) {
            fs.readFile(path.join(__dirname, "templates", "gdocs.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                t.substitute(1, {
                    planData: [
                        { name: "A", role: "Role 1" },
                        { name: "B", role: "Role 2" },
                    ]
                });

                var newData = t.generate();
                var sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
                expect(sheet1).toBeDefined();

                // fs.writeFileSync('test/output/test7.xlsx', newData, 'binary');
                done();
            });
        });

        it("Array indexing", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-array.xlsx"), function(err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                t.substitute(1, {
                    data: [
                        "First row",
                        { name: "B" },
                    ]
                });

                var newData = t.generate();
                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
                expect(sheet1).toBeDefined();
                expect(sheet1.find("./sheetData/row/c[@r='A2']/v")).not.toBeNull();
                expect(getSharedString(sharedStrings, sheet1, "A2")).toEqual("First row");
                expect(sheet1.find("./sheetData/row/c[@r='B2']/v")).not.toBeNull();
                expect(getSharedString(sharedStrings, sheet1, "B2")).toEqual("B");

                // fs.writeFileSync('test/output/test8.xlsx', newData, 'binary');
                done();
            });
        });

        it("Arrays with single element", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-nested-arrays.xlsx"), function(err, buffer) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(buffer);
                var data = { "sales": [ { "payments": [123] } ] };
                t.substitute(1, data);

                var newData = t.generate();
                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
                expect(sheet1).toBeDefined();
                var a1 = sheet1.find("./sheetData/row/c[@r='A1']/v");
                var firstElement = sheet1.findall("./sheetData/row/c[@r='A1']");
                expect(a1).not.toBeNull();
                expect(a1.text).toEqual("123");
                expect(firstElement).not.toBeNull();
                expect(firstElement.length).toEqual(1);

                fs.writeFileSync("test/output/test-nested-arrays.xlsx", newData, "binary");
                done();
            });
        });

        it("will correctly fill cells on all rows where arrays are used to dynamically render multiple cells", function (done) {
            fs.readFile(path.join(__dirname, "templates", "t2.xlsx"), function (err, data) {
                expect(err).toBeNull();

                var t = new XlsxTemplate(data);

                t.substitute(1, {
                    demo: { extractDate: new Date("2013-01-02") },
                    revision: 10,
                    dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
                    planData: [{
                            name: "John Smith",
                            role: { name: "Developer" },
                            days: [1, 2, 3]
                        },
                        {
                            name: "James Smith",
                            role: { name: "Analyst" },
                            days: [1, 2, 3, 4, 5]
                        },
                        {
                            name: "Jim Smith",
                            role: { name: "Manager" },
                            days: [1, 2, 3, 4, 5, 6, 7]
                        }
                    ]
                });

                var newData = t.generate();

                // var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                var sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // Dimensions should be updated
                expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

                // Check length of all rows
                expect(sheet1.find("./sheetData/row[@r='7']")._children.length).toEqual(2 + 3);
                expect(sheet1.find("./sheetData/row[@r='8']")._children.length).toEqual(2 + 5);
                expect(sheet1.find("./sheetData/row[@r='9']")._children.length).toEqual(2 + 7);

                fs.writeFileSync("test/output/test8.xlsx", newData, "binary");

                done();
            });
        });

	it("do not move Images", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-move-images.xlsx"), function(err, data) {
                expect(err).toBeNull();
                var option = {
                    moveImages : false
                };
                var t = new XlsxTemplate(data, option);
                t.substitute(1, {
                    users: [
                        {
                            name: "John",
                            surname : "Smith"
                        },
                        {
                            name: "John",
                            surname : "Doe"
                        }
                    ]
                });
                var newData = t.generate();
                var drawingSheet = etree.parse(t.archive.file("xl/drawings/drawing1.xml").asText()).getroot();
                expect(drawingSheet).toBeDefined();
                drawingSheet.findall("xdr:twoCellAnchor").forEach(element => {
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='3']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("3");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("9");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='6']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("2");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("9");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='8']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("1");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("11");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='10']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("10");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("24");
                    }
                });
                fs.writeFileSync("test/output/test_donotmoveImages.xlsx", newData, "binary");
                done();
            });
        });

        it("Move Images", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-move-images.xlsx"), function(err, data) {
                expect(err).toBeNull();
                const option = {
                    moveImages : true
                };
                const t = new XlsxTemplate(data, option);
                t.substitute(1, {
                    users: [
                        {
                            name: "John",
                            surname : "Smith"
                        },
                        {
                            name: "John",
                            surname : "Doe"
                        }
                    ]
                });
                var newData = t.generate();
                var drawingSheet = etree.parse(t.archive.file("xl/drawings/drawing1.xml").asText()).getroot();
                expect(drawingSheet).toBeDefined();
                drawingSheet.findall("xdr:twoCellAnchor").forEach(element => {
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='3']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("4");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("10");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='6']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("2");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("9");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='8']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("1");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("11");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='10']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("11");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("25");
                    }
                });
                fs.writeFileSync("test/output/test_moveImages.xlsx", newData, "binary");
                done();
            });
        });

        it("Move Images with sameLine option", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-move-images.xlsx"), function(err, data) {
                expect(err).toBeNull();
                var option = {
                    moveImages : true,
                    moveSameLineImages : true,
                };
                var t = new XlsxTemplate(data, option);
                t.substitute(1, {
                    users: [
                        {
                            name: "John",
                            surname : "Smith"
                        },
                        {
                            name: "John",
                            surname : "Doe"
                        }
                    ]
                });
                var newData = t.generate();
                var drawingSheet = etree.parse(t.archive.file("xl/drawings/drawing1.xml").asText()).getroot();
                expect(drawingSheet).toBeDefined();
                drawingSheet.findall("xdr:twoCellAnchor").forEach(element => {
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='3']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("4");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("10");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='6']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("3");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("10");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='8']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("1");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("11");
                    }
                    if(element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='10']")) {
                        expect(element.find("xdr:from/xdr:row").text).toEqual("11");
                        expect(element.find("xdr:to/xdr:row").text).toEqual("25");
                    }
                });
                fs.writeFileSync("test/output/test_moveImages_withSameLineOption.xlsx", newData, "binary");
                done();
            });
        });

        it("Insert image and create rels", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-insert-images.xlsx"), function(err, data) {
                expect(err).toBeNull();
                var option = {
                    imageRootPath : path.join(__dirname, "templates", "dataset")
                };
                var t = new XlsxTemplate(data, option);
                var imgB64 = "iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC";

                t.substitute("init_rels", {
                    imgB64 : imgB64,
                });
                var newData = t.generate();
                var rels = etree.parse(t.archive.file("xl/worksheets/_rels/sheet1.xml.rels").asText()).getroot();
                expect(rels.findall("Relationship").length).toEqual(1);
                expect(rels.findall("Relationship")[0].attrib.Id).toEqual("rId1");
                expect(rels.findall("Relationship")[0].attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
                expect(rels.findall("Relationship")[0].attrib.Target).toEqual("../drawings/drawing2.xml");
                var drawing2 = etree.parse(t.archive.file("xl/drawings/drawing2.xml").asText()).getroot();
                expect(drawing2.findall("xdr:oneCellAnchor").length).toEqual(1);
                expect(drawing2.findall("xdr:oneCellAnchor")[0].findall("xdr:from")[0].findall("xdr:col")[0].text).toEqual("1");
                expect(drawing2.findall("xdr:oneCellAnchor")[0].findall("xdr:from")[0].findall("xdr:row")[0].text).toEqual("2");
                expect(drawing2.findall("xdr:oneCellAnchor")[0].findall("xdr:pic")[0].findall("xdr:blipFill")[0].findall("a:blip")[0].attrib["r:embed"]).toEqual("rId1");
                var relsdrawing2 = etree.parse(t.archive.file("xl/drawings/_rels/drawing2.xml.rels").asText()).getroot();
                expect(relsdrawing2.findall("Relationship").length).toEqual(1);
                expect(relsdrawing2.findall("Relationship")[0].attrib.Id).toEqual("rId1");
                expect(relsdrawing2.findall("Relationship")[0].attrib.Target).toEqual("../media/image1.jpg");
                expect(relsdrawing2.findall("Relationship")[0].attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                //TODO : How can i compare the jpg file in the archive with my imgB64 variable ?
                //var image = t.archive.file("xl/media/image1.jpg");
                fs.writeFileSync("test/output/insert_image.xlsx", newData, "binary");
                done();
            });
        });

        it("Insert imageincells and create rels", function(done) {
            fs.readFile(path.join(__dirname, 'templates', 'test-insert-images_in_cell.xlsx'), function(err, data) {
                expect(err).toBeNull();
                var option = {
                    imageRootPath : path.join(__dirname, 'templates', 'dataset')
                }
                var t = new XlsxTemplate(data, option);
                var imgB64 = 'iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC';

                t.substitute('init_rels', {
                    imgB64 : imgB64,
                });
                var newData = t.generate();
                var richDataFile = etree.parse(t.archive.file("xl/richData/_rels/richValueRel.xml.rels").asText()).getroot();
                expect(richDataFile.findall("Relationship").length).toEqual(4);
                var richDataFile = etree.parse(t.archive.file("xl/richData/rdrichvalue.xml").asText()).getroot();
                expect(richDataFile.findall("rv").length).toEqual(4);
                expect(parseInt(richDataFile.attrib.count)).toEqual(richDataFile.findall("rv").length);
                var richDataFile = etree.parse(t.archive.file("xl/richData/richValueRel.xml").asText()).getroot();
                expect(richDataFile.findall("rel").length).toEqual(4);
                var richDataFile = etree.parse(t.archive.file("xl/metadata.xml").asText()).getroot();
                expect(parseInt(richDataFile.find('futureMetadata').attrib.count)).toEqual(4);
                expect(parseInt(richDataFile.find('valueMetadata').attrib.count)).toEqual(4);
                expect(richDataFile.find('futureMetadata').findall("bk").length).toEqual(4);
                expect(richDataFile.find('valueMetadata').findall("bk").length).toEqual(4);
                var sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
                expect(sheet1.find("./sheetData/row[@r='3']/c[@r='B3']").attrib.vm).toEqual("1");
                expect(sheet1.find("./sheetData/row[@r='5']/c[@r='K5']").attrib.vm).toEqual("2");
                expect(sheet1.find("./sheetData/row[@r='6']/c[@r='D6']").attrib.vm).toEqual("3");
                expect(sheet1.find("./sheetData/row[@r='10']/c[@r='H10']").attrib.vm).toEqual("4");
                fs.writeFileSync('test/output/insert_imageincell.xlsx', newData, 'binary');
                done();
            });
        });

        it("Insert some format of image", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-insert-images.xlsx"), function(err, data) {
                expect(err).toBeNull();
                var option = {
                    imageRootPath : path.join(__dirname, "templates", "dataset")
                };
                var t = new XlsxTemplate(data, option);
                var imgB64 = "iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC";

                t.substitute("multi_test", {
                    imgB64 : imgB64,
                    imgBuffer : Buffer.from(imgB64, "base64"),
                    imgPath : "image1.png",
                    high : "high.png",
                    large : "large.png",
                    imgArray : [
                        {filename : "image1.png"},
                        {filename : "image2.png"},
                        {filename : "image3.png"},
                        {filename : "image4.png"},
                    ],
                    someText : "Hello Image",
                });
                var newData = t.generate();
                //TODO : make some test
                fs.writeFileSync("test/output/insert_images_format.xlsx", newData, "binary");
                done();
            });
        });

        it("Insert image does not exist", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-insert-images.xlsx"), function(err, data) {
                expect(err).toBeNull();
                var option = {
                    imageRootPath : path.join(__dirname, "templates", "dataset")
                };
                var t = new XlsxTemplate(data, option);
                var imgB64 = "iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC";

                expect(() => {t.substitute("breaking_image", {
                    imgBuffer : Buffer.from(imgB64, "base64"),
                    imgPath : "image_does_not_exist.png",
                });}).toThrow(TypeError);
                // TODO : @kant2002 if image is null or empty string in user substitution data, throw an error or not ?
                // expect(() => {t.substitute('breaking_image', {
                //     imgBuffer : Buffer.from(imgB64, 'base64'),
                //     imgPath : null,
                // })}).toThrow(TypeError);
                done();
            });
        });

        it("Catch insert image does not exist", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-insert-images.xlsx"), function(err, data) {
                expect(err).toBeNull();
                var substituteImgBuffer = Buffer.from("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", "base64");
                var handleImageFunction = function(substitution, error) {
                    return substituteImgBuffer;
                };
                var option = {
                    imageRootPath : path.join(__dirname, "templates", "dataset"),
                    handleImageError : handleImageFunction
                };
                var t = new XlsxTemplate(data, option);

                expect(() => {t.substitute("breaking_image", {
                    imgPath : "image_does_not_exist.png",
                });}).not.toThrow(TypeError);

                var newData = t.generate();
                expect(t.archive.file("xl/media/image1.jpg").asText()).not.toBeNull();
                // TODO : @kant2002 How can I test if "xl/media/image1.jpg" is equal to substituteImgBuffer ?
                done();
            });
        });

        it("Insert 100 image", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-insert-images.xlsx"), function(err, buffer) {
                expect(err).toBeNull();
                expect(err).toBeNull();
                var option = {
                    imageRootPath : path.join(__dirname, "templates", "dataset")
                };
                var t = new XlsxTemplate(buffer, option);
                // tslint:disable-next-line:max-line-length
                var imgB64 = "iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC";
                var data = {
                    test : "Bug : If remove me, there are an error on ref.match([regex])",
                    imgArray: [] as { filename: string }[]
                };
                for (let i = 0; i < 100; i++) {
                    data.imgArray.push({ filename : imgB64 });
                }

                t.substitute("more_than_100", data);
                var newData = t.generate();

                var drawing2 = etree.parse(t.archive.file("xl/drawings/drawing2.xml").asText()).getroot();
                expect(drawing2.findall("xdr:oneCellAnchor").length).toEqual(100);
                var relsdrawing2 = etree.parse(t.archive.file("xl/drawings/_rels/drawing2.xml.rels").asText()).getroot();
                expect(relsdrawing2.findall("Relationship").length).toEqual(100);
                for (let i = 1; i < 101; i++) {
                    var image = t.archive.file("xl/media/image" + i + ".jpg");
                    expect(image == null).toEqual(false);
                }
                fs.writeFileSync("test/output/insert_100_images.xlsx", newData, "binary");
                done();
            });
        });

        it("pushDownPageBreakOnTableSubstitution option", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-movePageBreakOption.xlsx"), function(err, buffer) {
                expect(err).toBeNull();
                expect(err).toBeNull();
                var option = {
                    pushDownPageBreakOnTableSubstitution : true
                };
                var t = new XlsxTemplate(buffer, option);
                var data = {
                    myarray : [
                        {name : "foo"},
                        {name : "john"},
                        {name : "doe"},
                    ]
                };
                t.substitute(1, data);
                var newData = t.generate();
                var workbook = etree.parse(t.archive.file("xl/workbook.xml").asText()).getroot();

                workbook.findall("definedNames/definedName").forEach(function(name) {
                    expect(name.text === "Feuil1!$A$1:$J$12");
                });

                fs.writeFileSync("test/output/movePageBreakOption.xlsx", newData, "binary");
                done();
            });
        });

        it("subsituteAllTableRow option", function(done) {
            fs.readFile(path.join(__dirname, "templates", "test-movePageBreakOption.xlsx"), function(err, buffer) {
                expect(err).toBeNull();
                expect(err).toBeNull();
                var option = {
                    subsituteAllTableRow : true
                };
                var t = new XlsxTemplate(buffer, option);
                var data = {
                    myarray : [
                        {name : "foo"},
                        {name : "john"},
                        {name : "doe"},
                    ]
                };
                t.substitute(1, data);
                var newData = t.generate();
                var sheet1      = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                expect(sheet1.find("./sheetData/row[@r='4']/c[@r='C4']").attrib.s).toEqual("2");
                expect(sheet1.find("./sheetData/row[@r='4']/c[@r='D4']").attrib.s).toEqual("3");
                expect(sheet1.find("./sheetData/row[@r='4']/c[@r='F4']/v").text).toEqual("5");
                expect(sheet1.find("./sheetData/row[@r='4']/c[@r='H4']").attrib.s).toEqual("4");

                expect(sheet1.find("./sheetData/row[@r='5']/c[@r='C5']").attrib.s).toEqual("2");
                expect(sheet1.find("./sheetData/row[@r='5']/c[@r='D5']").attrib.s).toEqual("3");
                expect(sheet1.find("./sheetData/row[@r='5']/c[@r='F5']/v").text).toEqual("5");
                expect(sheet1.find("./sheetData/row[@r='5']/c[@r='H5']").attrib.s).toEqual("4");

                expect(sheet1.find("./sheetData/row[@r='6']/c[@r='C6']").attrib.s).toEqual("2");
                expect(sheet1.find("./sheetData/row[@r='6']/c[@r='D6']").attrib.s).toEqual("3");
                expect(sheet1.find("./sheetData/row[@r='6']/c[@r='F6']/v").text).toEqual("5");
                expect(sheet1.find("./sheetData/row[@r='6']/c[@r='H6']").attrib.s).toEqual("4");

                fs.writeFileSync("test/output/substituteAllTheRow.xlsx", newData, "binary");
                done();
            });
        });
    });

    describe("Multiple sheets", function() {
        it("Each sheet should take each name", function (done) {
            fs.readFile(path.join(__dirname, "templates", "multple-sheets-arrays.xlsx"), function(err, data) {
                expect(err).toBeNull();

                // Create a template
                var t = new XlsxTemplate(data);
                for (let sheetNumber = 1; sheetNumber <= 2; sheetNumber++) {
                    // Set up some placeholder values matching the placeholders in the template
                    var values = {
                        page: "page: " + sheetNumber,
                        sheetNumber
                    };

                    // Perform substitution
                    t.substitute(sheetNumber, values);
                }

                // Get binary data
                var newData = t.generate();
                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
                var sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
                var sheet2        = etree.parse(t.archive.file("xl/worksheets/sheet2.xml").asText()).getroot();
                expect(sheet1).toBeDefined();
                expect(sheet2).toBeDefined();
                expect(getSharedString(sharedStrings, sheet1, "A1")).toEqual("page: 1");
                expect(getSharedString(sharedStrings, sheet2, "A1")).toEqual("page: 2");
                expect(getSharedString(sharedStrings, sheet1, "A2")).toEqual("Page 1");
                expect(getSharedString(sharedStrings, sheet2, "A2")).toEqual("Page 2");

                fs.writeFileSync("test/output/multple-sheets-arrays.xlsx", newData, "binary");
                done();
            });
        });

        it("Each sheet should be replaced", function (done) {
            fs.readFile(path.join(__dirname, "templates", "multple-sheet-substitution.xlsx"), function(err, data) {
                expect(err).toBeNull();

                // Create a template
                var t = new XlsxTemplate(data);
                t.substituteAll({
                    foo: "bar"
                });

                // Get binary data
                var newData = t.generate();
                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
                var sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
                var sheet2        = etree.parse(t.archive.file("xl/worksheets/sheet2.xml").asText()).getroot();
                expect(sheet1).toBeDefined();
                expect(sheet2).toBeDefined();
                expect(getSharedString(sharedStrings, sheet1, "A1")).toEqual("bar");
                expect(getSharedString(sharedStrings, sheet2, "A1")).toEqual("bar");

                fs.writeFileSync("test/output/multple-sheet-substitution.xlsx", newData, "binary");
                done();
            });
        });
    });

    describe("Copy sheets", function() {
        it("Standard-copy-sheets", function (done) {
            fs.readFile(path.join(__dirname, "templates", "test-copy-sheet.xlsx"), function(err, data) {
                expect(err).toBeNull();
                // Create a template
                var t = new XlsxTemplate(data);
                // Test copy simple sheet
                t.copySheet("simple_sheet", "simple_sheet_copy");
                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
                var copysheet4    = etree.parse(t.archive.file("xl/worksheets/sheet4.xml").asText()).getroot();
                expect(copysheet4).toBeDefined();
                expect(getSharedString(sharedStrings, copysheet4, "B4")).toEqual("Some data");
                expect(getSharedString(sharedStrings, copysheet4, "C6")).toEqual("${data.foo}");
                expect(getSharedString(sharedStrings, copysheet4, "C11")).toEqual("AZE²&é\"'(-è_çà)=~#{[|`\\^@]} ^$ù*,;:!¨£%µ?./§¤€<>Ôöôîïï");
                // TODO : How can I test the fill color of cell D9 ?
                // test copy sheet with image
                t.copySheet("image_sheet", "image_sheet_copy");
                var relscopysheet2    = etree.parse(t.archive.file("xl/worksheets/_rels/sheet5.xml.rels").asText()).getroot();
                expect(relscopysheet2.findall("Relationship").length).toEqual(1);
                expect(relscopysheet2.findall("Relationship")[0].attrib.Target).toEqual("../drawings/drawing1.xml");
                // test copy sheet with utf-8 header/footer
                // without binary mode (no ok) :
                // This test can be remove if you wish. This is just for show the issue if binary is set to true (default)
                t.copySheet("utf8_header_sheet", "utf8_header_sheet_copy");
                var copysheet6    = etree.parse(t.archive.file("xl/worksheets/sheet6.xml").asText()).getroot();
                expect(copysheet6).toBeDefined();
                expect(copysheet6.find("headerFooter/oddHeader").text).not.toEqual("&LTEST UTF8 éàè&ùï&C&A");
                // with binary mode (ok) :
                t.copySheet("utf8_header_sheet", "utf8_header_sheet_copy_v2", false);
                var copysheet7    = etree.parse(t.archive.file("xl/worksheets/sheet7.xml").asText()).getroot();
                expect(copysheet7).toBeDefined();
                expect(copysheet7.find("headerFooter/oddHeader").text).toEqual("&LTEST UTF8 éàè&ùï&C&A");
                // Get binary data
                var newData = t.generate();
                fs.writeFileSync("test/output/copy-sheet.xlsx", newData, "binary");
                done();
            });
        });
    });

    describe("Rebuild file", function() {
        it("Rebuild archive file with custom Relationship workbook rels after delete a sheet", function (done) {
            fs.readFile(path.join(__dirname, "templates", "custom-xml.xlsx"), function(err, buffer) {
                expect(err).toBeNull();
                // Create a template
                var t = new XlsxTemplate(buffer);

                t.deleteSheet("Sheet2");
                var newData     = t.generate();
                var wbrels      = etree.parse(t.archive.file("xl/_rels/workbook.xml.rels").asText()).getroot();
                const relationship1 = wbrels.find("Relationship[@Id='rId1']");
                expect(relationship1.attrib.Target).toEqual("worksheets/sheet1.xml");
                expect(relationship1.attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                const relationship2 = wbrels.find("Relationship[@Id='rId2']");
                expect(relationship2.attrib.Target).toEqual("worksheets/sheet3.xml");
                expect(relationship2.attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                const relationship3 = wbrels.find("Relationship[@Id='rId3']");
                expect(relationship3.attrib.Target).toEqual("theme/theme1.xml");
                expect(relationship3.attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
                const relationship4 = wbrels.find("Relationship[@Id='rId4']");
                expect(relationship4.attrib.Target).toEqual("styles.xml");
                expect(relationship4.attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
                const relationship5 = wbrels.find("Relationship[@Id='rId5']");
                expect(relationship5.attrib.Target).toEqual("sharedStrings.xml");
                expect(relationship5.attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
                const relationship6 = wbrels.find("Relationship[@Id='rId6']");
                expect(relationship6.attrib.Target).toEqual("xmlMaps.xml");
                expect(relationship6.attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps");
                const relationship7 = wbrels.find("Relationship[@Id='rId7']");
                expect(relationship7.attrib.Target).toEqual("connections.xml");
                expect(relationship7.attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections");
                fs.writeFileSync("test/output/custom-xml.xlsx", newData, "binary");
                done();
            });
        });
    });
});
