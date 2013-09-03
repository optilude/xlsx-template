/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, Buffer */
"use strict";

var fs    = require('fs'),
    path  = require('path'),
    zip   = require('node-zip'),
    etree = require('elementtree');

module.exports = (function() {

    var SHARED_STRINGS = "xl/sharedStrings.xml";

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
    Workbook.prototype.substitute = function(sheet, substitutions) {

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

        //  - if `token` is not found in `substitutions`, do nothing
        
        //  - if `substitutions[token]` is a string, or if the shared string
        //    contains other characters than the substituion reference,
        //    replace the characters in the shared string, but leave the
        //    shared string reference intact.
        
        //  - if `substitutions[token]` is a number (`n`), boolean (`b`), or
        //    date (`d`; ISO8601 format) insert it into the column
        //    reference and change the cell type.
        
        //  - if `substitutions[token]` is a list, clone and repeat the column
        //    for each item in the list (converting strings and types)
        
        //  - if `token` starts with `table:`, assume it is in the format
        //    `${table:listName.varName}`. If `substitutions[listName]` is a
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

        var tree = etree.parse(self.archive.file(SHARED_STRINGS).asText());
        self.sharedStrings = [];
        tree.findall('si/t').forEach(function(t) {
            self.sharedStrings.push(t.text);
            self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1;
        });
    };

    // Write back the new shared strings list
    Workbook.prototype.writeSharedStrings = function() {
        var self = this;

        var tree = etree.parse(self.archive.file(SHARED_STRINGS).asText()),
            root = tree.getroot(),
            children = root.getchildren();
        
        root.delSlice(0, children.length);


        self.sharedStrings.forEach(function(string) {
            var si = new etree.Element("si"),
                t  = new etree.Element("t");

            t.text = string;
            si.append(t);
            root.append(si);
        });

        root.attrib.count = self.sharedStrings.length;
        root.attrib.uniqueCount = self.sharedStrings.length;

        self.archive.file(SHARED_STRINGS, etree.tostring(root));
    };

    // Add a new shared string
    Workbook.prototype.addSharedString = function(s) {
        var self = this;

        var idx = self.sharedStrings.length;
        self.sharedStrings.push(s);
        self.sharedStringsLookup[s] = idx;

        return idx;
    };

    // Get the number of a shared string, adding a new one if necessary.
    Workbook.prototype.stringIndex = function(s) {
        var self = this;

        var idx = self.sharedStringsLookup[s];
        if(idx === undefined) {
            idx = self.addSharedString(s);
        }
        return idx;
    };

    // Replace a shared string with a new one at the same index. Return the
    // index.
    Workbook.prototype.replaceString = function(oldString, newString) {
        var self = this;

        var idx = self.sharedStringsLookup[oldString];
        if(idx === undefined) {
            idx = self.addSharedString(newString);
        } else {
            self.sharedStrings[idx] = newString;
            delete self.sharedStringsLookup[oldString];
            self.sharedStringsLookup[newString] = idx;
        }

        return idx;
    };
    
    // Return a list of tokens that may exist in the string.
    // Keys are: `placeholder` (the full placeholder, including the `${}`
    // delineators), `name` (the name part of the token), `key` (the object key
    // for `table` tokens), `full` (boolean indicating whether this placeholder
    // is the entirety of the string) and `type` (one of `table` or `cell`)
    Workbook.prototype.extractPlaceholders = function(string) {
        // Yes, that's right. It's a bunch of brackets and question marks and stuff.
        var re = /\${(?:(.+?):)?(.+?)(?:\.(.+?))?}/g;
        
        var match = null, matches = [];
        while((match = re.exec(string)) !== null) {
            matches.push({
                placeholder: match[0],
                type: match[1] || 'normal',
                name: match[2],
                key: match[3],
                full: match[0].length === string.length
            });
        }

        return matches;
    };

    // Split a reference into an object with keys `row` and `col`.
    Workbook.prototype.splitRef = function(ref) {
        var match = ref.match(/([A-Z]+)([0-9]+)/);
        return {
            col: match[1],
            row: parseInt(match[2], 10)
        };
    };

    // Join an object with keys `row` and `col` into a single reference string
    Workbook.prototype.joinRef = function(ref) {
        return ref.col.toUpperCase() + Number(ref.row).toString();
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