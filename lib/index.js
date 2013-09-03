/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, Buffer */
"use strict";

var fs    = require('fs'),
    path  = require('path'),
    zip   = require('node-zip'),
    etree = require('elementtree');

module.exports = (function() {

    /**
     * Create a new workbook. Either pass the raw data of a .xlsx file,
     * or call `loadTemplate()` later.
     */
    var Workbook = function(data) {
        var self = this;

        self.archive = null;
        self.sharedStrings = [];
        self.sharedStringsLookup = {};

        if(data) {
            self.loadTemplate(data);
        }
    };

    /**
     * Load a .xlsx file from a byte array.
     */
    Workbook.prototype.loadTemplate = function(data) {
        var self = this;

        if(Buffer.isBuffer(data)) {
            data = data.toString('binary');
        }

        self.archive = new zip(data, {base64: false, checkCRC32: true});
        self.readSharedStrings();
    };

    /**
     *
     */
    Workbook.prototype.substitute = function(sheet, substitions) {
        
        // Get sheet and parse XML tree

        // Loop over /worksheet/sheetData/row
        // Count rows from 1
        // Update row@r (row reference). This is necessary because we may clone
        // rows

        // Loop over row/c (columns)
        // Update col@r (col reference) based on current row and cell. This is
        // necessary because we may clone rows and columns.

        // If c[@t="s"] (string column), look up /c/v@text as integer in
        // `this.sharedStrings`

        // If string contains `${token}`, invoke substituion on this element
        // using the using the following rules:

        //  - if `token` is not found in `substituions`, do nothing
        
        //  - if `substituions[token]` is a string, or if the shared string
        //    contains other characters than the substituion reference,
        //    replace the characters in the shared string, but leave the
        //    shared string reference intact.
        
        //  - if `substituions[token]` is a number (`n`), boolean (`b`), or
        //    date (`d`; ISO8601 format) insert it into the column
        //    reference and change the cell type.
        
        //  - if `substituions[token]` is a list, clone and repeat the column
        //    for each item in the list (converting strings and types)
        
        //  - if `token` starts with `table:`, assume it is in the format
        //    `${table:listName.varName}`. If `substituions[listName]` is a
        //    list of objects, clone the row and repeat for each item in the
        //    list, substituting each `${table:listName.varName}` with the
        //    property `varName` of each value in the list `listName`.

        // Write back the modified XML tree
    };

    /**
     * Generate a new binary .xlsx file
     */
    Workbook.prototype.generate = function() {
        var self = this;

        self.writeSharedStrings();
        return self.archive.generate();
    };

    // Helpers

    // Read the shared strings from the workbook
    Workbook.prototype.readSharedStrings = function() {
        var self = this;

        var root = etree.parse(self.archive.file('xl/sharedStrings.xml').asText());
        self.sharedStrings = [];
        root.findall('si/t').forEach(function(t) {
            self.sharedStrings.push(t.text);
            self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1;
        });
    };

    // Get the number of a shared string, adding a new one if necessary.
    Workbook.prototype.stringIndex = function(s) {
        var self = this;

        var idx = self.sharedStringsLookup[s];
        if(idx === undefined) {
            idx = self.sharedStrings.length;
            self.sharedStrings.push(s);
            self.sharedStringsLookup[s] = idx;
        }
        return idx;
    };

    // Write back the new shared strings list
    Workbook.prototype.writeSharedStrings = function() {
        // TODO
    };

    // Get the next column's cell reference given a reference like "B2".
    Workbook.prototype.nextCol = function(ref) {
        ref = ref.toUpperCase();
        return ref.replace(/[A-Z]+/, function(match) {
            var chars = match.split('');
            for(var i = chars.length - 1; i >= 0; --i) {

                // Increment char code
                var code = chars[i].charCodeAt(0) + 1;
                if(code > 90) { // Z
                    code = 65;  // A
                }
                chars[i] = String.fromCharCode(code);

                // Unless we rolled over to 'A', we don't need to increment
                // the preceding
                if(code != 65) {
                    return chars.join('');
                }
            }

            // If we got here without returning, then we need to add another
            // character
            return 'A' + chars.join('');
        });
    };

    // Get the next row's cell reference given a reference like "B2".
    Workbook.prototype.nextRow = function(ref) {
        ref = ref.toUpperCase();
        return ref.replace(/[0-9]+/, function(match) {
            return (parseInt(match, 10) + 1).toString();
        });
    };

    return Workbook;
})();