/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, describe, before, it */
"use strict";

var XlsxTemplate = require('../lib');

describe("Helpers", function() {

    describe('stringIndex', function() {
        it("adds new strings to the index if required", function() {
            var t = new XlsxTemplate();
            expect(t.stringIndex("foo")).toEqual(0);
            expect(t.stringIndex("bar")).toEqual(1);
            expect(t.stringIndex("foo")).toEqual(0);
            expect(t.stringIndex("baz")).toEqual(2);
        });

    });

    describe('replaceString', function() {

        it("adds new string if old string not found", function() {
            var t = new XlsxTemplate();

            expect(t.replaceString("foo", "bar")).toEqual(0);
            expect(t.sharedStrings).toEqual(["bar"]);
            expect(t.sharedStringsLookup).toEqual({"bar": 0});
        });

        it("replaces strings if found", function() {
            var t = new XlsxTemplate();

            t.addSharedString("foo");
            t.addSharedString("baz");

            expect(t.replaceString("foo", "bar")).toEqual(0);
            expect(t.sharedStrings).toEqual(["bar", "baz"]);
            expect(t.sharedStringsLookup).toEqual({"bar": 0, "baz": 1});
        });

    });

    describe('extractPlaceholders', function() {

        it("can extract simple placeholders", function() {
            var t = new XlsxTemplate();

            expect(t.extractPlaceholders("${foo}")).toEqual([{
                full: true,
                key: undefined,
                name: "foo",
                placeholder: "${foo}",
                subType: undefined,
                type: "normal"
            }]);
        });

        it("can extract simple placeholders inside strings", function() {
            var t = new XlsxTemplate();

            expect(t.extractPlaceholders("A string ${foo} bar")).toEqual([{
                full: false,
                key: undefined,
                name: "foo",
                placeholder: "${foo}",
                subType: undefined,
                type: "normal"
            }]);
        });

        it("can extract multiple placeholders from one string", function() {
            var t = new XlsxTemplate();

            expect(t.extractPlaceholders("${foo} ${bar}")).toEqual([{
                full: false,
                key: undefined,
                name: "foo",
                placeholder: "${foo}",
                subType: undefined,
                type: "normal"
            }, {
                full: false,
                key: undefined,
                name: "bar",
                placeholder: "${bar}",
                subType: undefined,
                type: "normal"
            }]);
        });

        it("can extract placeholders with keys", function() {
            var t = new XlsxTemplate();

            expect(t.extractPlaceholders("${foo.bar}")).toEqual([{
                full: true,
                key: "bar",
                name: "foo",
                placeholder: "${foo.bar}",
                subType: undefined,
                type: "normal"
            }]);
        });

        it("can extract placeholders with types", function() {
            var t = new XlsxTemplate();

            expect(t.extractPlaceholders("${table:foo}")).toEqual([{
                full: true,
                key: undefined,
                name: "foo",
                placeholder: "${table:foo}",
                subType: undefined,
                type: "table"
            }]);
        });

        it("can extract placeholders with types and keys", function() {
            var t = new XlsxTemplate();

            expect(t.extractPlaceholders("${table:foo.bar}")).toEqual([{
                full: true,
                key: "bar",
                name: "foo",
                placeholder: "${table:foo.bar}",
                subType: undefined,
                type: "table"
            }]);
        });

        it("can handle strings with no placeholders", function() {
            var t = new XlsxTemplate();

            expect(t.extractPlaceholders("A string")).toEqual([]);
        });

    });

    describe('isRange', function() {

        it("Returns true if there is a colon", function() {
            var t = new XlsxTemplate();
            expect(t.isRange("A1:A2")).toEqual(true);
            expect(t.isRange("$A$1:$A$2")).toEqual(true);
            expect(t.isRange("Table!$A$1:$A$2")).toEqual(true);
        });

        it("Returns false if there is not a colon", function() {
            var t = new XlsxTemplate();
            expect(t.isRange("A1")).toEqual(false);
            expect(t.isRange("$A$1")).toEqual(false);
            expect(t.isRange("Table!$A$1")).toEqual(false);
        });

    });

    describe('splitRef', function() {

        it("splits single digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.splitRef("A1")).toEqual({table: null, col: "A", colAbsolute: false, row: 1, rowAbsolute: false});
        });

        it("splits multiple digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.splitRef("AB12")).toEqual({table: null, col: "AB", colAbsolute: false, row: 12, rowAbsolute: false});
        });

        it("splits absolute references", function() {
            var t = new XlsxTemplate();
            expect(t.splitRef("$AB12")).toEqual({table: null, col: "AB", colAbsolute: true, row: 12, rowAbsolute: false});
            expect(t.splitRef("AB$12")).toEqual({table: null, col: "AB", colAbsolute: false, row: 12, rowAbsolute: true});
            expect(t.splitRef("$AB$12")).toEqual({table: null, col: "AB", colAbsolute: true, row: 12, rowAbsolute: true});
        });

        it("splits references with tables", function() {
            var t = new XlsxTemplate();
            expect(t.splitRef("Table one!AB12")).toEqual({table: "Table one", col: "AB", colAbsolute: false, row: 12, rowAbsolute: false});
            expect(t.splitRef("Table one!$AB12")).toEqual({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: false});
            expect(t.splitRef("Table one!$AB12")).toEqual({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: false});
            expect(t.splitRef("Table one!AB$12")).toEqual({table: "Table one", col: "AB", colAbsolute: false, row: 12, rowAbsolute: true});
            expect(t.splitRef("Table one!$AB$12")).toEqual({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: true});
        });

    });

    describe('splitRange', function() {

        it("splits single digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.splitRange("A1:B1")).toEqual({start: "A1", end: "B1"});
        });

        it("splits multiple digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.splitRange("AB12:CC13")).toEqual({start: "AB12", end: "CC13"});
        });

    });

    describe('joinRange', function() {

        it("join single digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.joinRange({start: "A1", end: "B1"})).toEqual("A1:B1");
        });

        it("join multiple digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.joinRange({start: "AB12", end: "CC13"})).toEqual("AB12:CC13");
        });

    });

    describe('joinRef', function() {

        it("joins single digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.joinRef({col: "A", row: 1})).toEqual("A1");
        });

        it("joins multiple digit and letter values", function() {
            var t = new XlsxTemplate();
            expect(t.joinRef({col: "AB", row: 12})).toEqual("AB12");
        });

        it("joins multiple digit and letter values and absolute references", function() {
            var t = new XlsxTemplate();
            expect(t.joinRef({col: "AB", colAbsolute: true, row: 12, rowAbsolute: false})).toEqual("$AB12");
            expect(t.joinRef({col: "AB", colAbsolute: true, row: 12, rowAbsolute: true})).toEqual("$AB$12");
            expect(t.joinRef({col: "AB", colAbsolute: false, row: 12, rowAbsolute: false})).toEqual("AB12");
        });

        it("joins multiple digit and letter values and tables", function() {
            var t = new XlsxTemplate();
            expect(t.joinRef({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: false})).toEqual("Table one!$AB12");
            expect(t.joinRef({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: true})).toEqual("Table one!$AB$12");
            expect(t.joinRef({table: "Table one", col: "AB", colAbsolute: false, row: 12, rowAbsolute: false})).toEqual("Table one!AB12");
        });

    });

    describe('nexCol', function() {

        it("increments single columns", function() {
            var t = new XlsxTemplate();

            expect(t.nextCol("A1")).toEqual("B1");
            expect(t.nextCol("B1")).toEqual("C1");
        });

        it("maintains row index", function() {
            var t = new XlsxTemplate();

            expect(t.nextCol("A99")).toEqual("B99");
            expect(t.nextCol("B11231")).toEqual("C11231");
        });

        it("captialises letters", function() {
            var t = new XlsxTemplate();

            expect(t.nextCol("a1")).toEqual("B1");
            expect(t.nextCol("b1")).toEqual("C1");
        });

        it("increments the last letter of double columns", function() {
            var t = new XlsxTemplate();

            expect(t.nextCol("AA12")).toEqual("AB12");
        });

        it("rolls over from Z to A and increments the preceding letter", function() {
            var t = new XlsxTemplate();

            expect(t.nextCol("AZ12")).toEqual("BA12");
        });

        it("rolls over from Z to A and adds a new letter if required", function() {
            var t = new XlsxTemplate();

            expect(t.nextCol("Z12")).toEqual("AA12");
            expect(t.nextCol("ZZ12")).toEqual("AAA12");
        });

    });

    describe('nexRow', function() {

        it("increments single digit rows", function() {
            var t = new XlsxTemplate();

            expect(t.nextRow("A1")).toEqual("A2");
            expect(t.nextRow("B1")).toEqual("B2");
            expect(t.nextRow("AZ2")).toEqual("AZ3");
        });

        it("captialises letters", function() {
            var t = new XlsxTemplate();

            expect(t.nextRow("a1")).toEqual("A2");
            expect(t.nextRow("b1")).toEqual("B2");
        });

        it("increments multi digit rows", function() {
            var t = new XlsxTemplate();

            expect(t.nextRow("A12")).toEqual("A13");
            expect(t.nextRow("AZ12")).toEqual("AZ13");
            expect(t.nextRow("A123")).toEqual("A124");
        });

    });

    describe('charToNum', function() {

        it("can return single letter numbers", function() {
            var t = new XlsxTemplate();

            expect(t.charToNum("A")).toEqual(1);
            expect(t.charToNum("B")).toEqual(2);
            expect(t.charToNum("Z")).toEqual(26);
        });

        it("can return double letter numbers", function() {
            var t = new XlsxTemplate();

            expect(t.charToNum("AA")).toEqual(27);
            expect(t.charToNum("AZ")).toEqual(52);
            expect(t.charToNum("BZ")).toEqual(78);
        });

        it("can return triple letter numbers", function() {
            var t = new XlsxTemplate();

            expect(t.charToNum("AAA")).toEqual(703);
            expect(t.charToNum("AAZ")).toEqual(728);
            expect(t.charToNum("ADI")).toEqual(789);
        });

    });

    describe('numToChar', function() {

        it("can convert single letter numbers", function() {
            var t = new XlsxTemplate();

            expect(t.numToChar(1)).toEqual("A");
            expect(t.numToChar(2)).toEqual("B");
            expect(t.numToChar(26)).toEqual("Z");
        });

        it("can convert double letter numbers", function() {
            var t = new XlsxTemplate();

            expect(t.numToChar(27)).toEqual("AA");
            expect(t.numToChar(52)).toEqual("AZ");
            expect(t.numToChar(78)).toEqual("BZ");
        });

        it("can convert triple letter numbers", function() {
            var t = new XlsxTemplate();

            expect(t.numToChar(703)).toEqual("AAA");
            expect(t.numToChar(728)).toEqual("AAZ");
            expect(t.numToChar(789)).toEqual("ADI");
        });

    });

    describe('isWithin', function() {

        it("can check 1x1 cells", function() {
            var t = new XlsxTemplate();

            expect(t.isWithin("A1", "A1", "A1")).toEqual(true);
            expect(t.isWithin("A2", "A1", "A1")).toEqual(false);
            expect(t.isWithin("B1", "A1", "A1")).toEqual(false);
        });

        it("can check 1xn cells", function() {
            var t = new XlsxTemplate();

            expect(t.isWithin("A1", "A1", "A3")).toEqual(true);
            expect(t.isWithin("A3", "A1", "A3")).toEqual(true);
            expect(t.isWithin("A4", "A1", "A3")).toEqual(false);
            expect(t.isWithin("A5", "A1", "A3")).toEqual(false);
            expect(t.isWithin("B1", "A1", "A3")).toEqual(false);
        });

        it("can check nxn cells", function() {
            var t = new XlsxTemplate();

            expect(t.isWithin("A1", "A2", "C3")).toEqual(false);
            expect(t.isWithin("A3", "A2", "C3")).toEqual(true);
            expect(t.isWithin("B2", "A2", "C3")).toEqual(true);
            expect(t.isWithin("A5", "A2", "C3")).toEqual(false);
            expect(t.isWithin("D2", "A2", "C3")).toEqual(false);
        });

        it("can check large nxn cells", function() {
            var t = new XlsxTemplate();

            expect(t.isWithin("AZ1", "AZ2", "CZ3")).toEqual(false);
            expect(t.isWithin("AZ3", "AZ2", "CZ3")).toEqual(true);
            expect(t.isWithin("BZ2", "AZ2", "CZ3")).toEqual(true);
            expect(t.isWithin("AZ5", "AZ2", "CZ3")).toEqual(false);
            expect(t.isWithin("DZ2", "AZ2", "CZ3")).toEqual(false);
        });

    });

    describe('stringify', function() {

        it("can stringify dates", function() {
            var t = new XlsxTemplate();

            expect(t.stringify(new Date("2013-01-01"))).toEqual(41275);
        });

        it("can stringify numbers", function() {
            var t = new XlsxTemplate();

            expect(t.stringify(12)).toEqual("12");
            expect(t.stringify(12.3)).toEqual("12.3");
        });

        it("can stringify booleans", function() {
            var t = new XlsxTemplate();

            expect(t.stringify(true)).toEqual("1");
            expect(t.stringify(false)).toEqual("0");
        });

        it("can stringify strings", function() {
            var t = new XlsxTemplate();

            expect(t.stringify("foo")).toEqual("foo");
        });

    });

    describe('substituteScalar', function() {

        it("can substitute simple string values", function() {
            var t = new XlsxTemplate(),
                string = "${foo}",
                substitution = "bar",
                placeholder = {
                    full: true,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("bar");
            expect(col.attrib.t).toEqual("s");
            expect(String(val.text)).toEqual("1");
            expect(t.sharedStrings).toEqual(["${foo}", "bar"]);
        });
        
        it("Substitution of shared simple string values", function() {
            var t = new XlsxTemplate(),
                string = "${foo}",
                substitution = "bar",
                placeholder = {
                    full: true,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("bar");
            
            // Explicitly share substritution strings if they could be reused.
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("bar");
            expect(col.attrib.t).toEqual("s");
            expect(String(val.text)).toEqual("1");
            expect(t.sharedStrings).toEqual(["${foo}", "bar"]);
        });

        it("can substitute simple numeric values", function() {
            var t = new XlsxTemplate(),
                string = "${foo}",
                substitution = 10,
                placeholder = {
                    full: true,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("10");
            expect(col.attrib.t).not.toBeDefined();
            expect(val.text).toEqual("10");
            expect(t.sharedStrings).toEqual(["${foo}"]);
        });

        it("can substitute simple boolean values (false)", function() {
            var t = new XlsxTemplate(),
                string = "${foo}",
                substitution = false,
                placeholder = {
                    full: true,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("0");
            expect(col.attrib.t).toEqual("b");
            expect(val.text).toEqual("0");
            expect(t.sharedStrings).toEqual(["${foo}"]);
        });

        it("can substitute simple boolean values (true)", function() {
            var t = new XlsxTemplate(),
                string = "${foo}",
                substitution = true,
                placeholder = {
                    full: true,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("1");
            expect(col.attrib.t).toEqual("b");
            expect(val.text).toEqual("1");
            expect(t.sharedStrings).toEqual(["${foo}"]);
        });

        it("can substitute dates", function() {
            var t = new XlsxTemplate(),
                string = "${foo}",
                substitution = new Date("2013-01-01"),
                placeholder = {
                    full: true,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual(41275);
            expect(col.attrib.t).not.toBeDefined();
            expect(val.text).toEqual(41275);
            expect(t.sharedStrings).toEqual(["${foo}"]);
        });

        it("can substitute parts of strings", function() {
            var t = new XlsxTemplate(),
                string = "foo: ${foo}",
                substitution = "bar",
                placeholder = {
                    full: false,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("foo: bar");
            expect(col.attrib.t).toEqual("s");
            expect(val.text).toEqual("1");
            expect(t.sharedStrings).toEqual(["foo: ${foo}", "foo: bar"]);
        });

        it("can substitute parts of strings with booleans", function() {
            var t = new XlsxTemplate(),
                string = "foo: ${foo}",
                substitution = false,
                placeholder = {
                    full: false,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("foo: 0");
            expect(col.attrib.t).toEqual("s");
            expect(val.text).toEqual("1");
            expect(t.sharedStrings).toEqual(["foo: ${foo}", "foo: 0"]);
        });

        it("can substitute parts of strings with numbers", function() {
            var t = new XlsxTemplate(),
                string = "foo: ${foo}",
                substitution = 10,
                placeholder = {
                    full: false,
                    key: undefined,
                    name: "foo",
                    placeholder: "${foo}",
                    type: "normal"
                },
                val = {
                    text: "0"
                },
                col = {
                    attrib: {t: "s"},
                    find: function() {
                        return val;
                    }
                };

            t.addSharedString(string);
            expect(t.substituteScalar(col, string, placeholder, substitution)).toEqual("foo: 10");
            expect(col.attrib.t).toEqual("s");
            expect(val.text).toEqual("1");
            expect(t.sharedStrings).toEqual(["foo: ${foo}", "foo: 10"]);
        });


    });

});
