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

    describe('imageUtils', function() {

        it("default handle image error", function() {
            var t = new XlsxTemplate()
            // TODO : @kant2002 how can I test if t.handleImageError is public ?
            expect(() => {t.handleImageError(null, new TypeError())}).toThrow(TypeError);
        });

        it("Override handle image error", function() {
            var t = new XlsxTemplate()
            t.handleImageError = function(substitution, error) {
                var imgB64 = 'data:image/png;base64,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'
                return Buffer.from(imgB64, 'base64');
            }
            expect(() => {t.handleImageError(null, new TypeError())}).not.toThrow(TypeError);
        });

    });

});
