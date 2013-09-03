/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, describe, before, after, it, expect */
"use strict";

var buster       = require('buster'),
    XlsxTemplate = require('../lib');

buster.spec.expose();
buster.testRunner.timeout = 500;

describe("Helpers", function() {
    
    before(function(done) {
        done();
    });

    describe('stringIndex', function() {

        it("adds new strings to the index if required", function() {
            var t = new XlsxTemplate();
            expect(t.stringIndex("foo")).toEqual(0);
            expect(t.stringIndex("bar")).toEqual(1);
            expect(t.stringIndex("foo")).toEqual(0);
            expect(t.stringIndex("baz")).toEqual(2);
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

});