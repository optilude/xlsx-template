/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, Buffer */
"use strict";

var fs    = require('fs'),
    path  = require('path'),
    zip   = require('node-zip'),
    etree = require('elementtree'),
    _     = require('lodash');

module.exports = (function() {

    var SHARED_STRINGS = "xl/sharedStrings.xml",
        WORKSHEETS     = "xl/worksheets/";

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
     * Interpolate values for the sheet with the given number (1-based)
     * using the given substitutions (an object).
     */
    Workbook.prototype.substitute = function(sheet, substitutions) {
        var self = this;

        function getCurrentCell(currentCell, currentRow, col) {
            if(currentCell === null) {
                return self.joinRef({
                    row: currentRow,
                    col: self.splitRef(col.attrib.r).col
                });
            } else {
                return self.nextCol(currentCell);
            }
        }

        // Get sheet and parse XML tree
        var file = WORKSHEETS + "sheet" + sheet + ".xml",
            tree = etree.parse(self.archive.file(file).asText()),
            root = tree.getroot();

        // Loop over /worksheet/sheetData/row
        var sheetData = root.find("sheetData"),
            currentRow = null,
            rowsInserted = 0,
            rows = [];

        sheetData.findall("row").forEach(function(row) {

            // Update row@r (row reference) in case we have inserted additional rows.
            row.attrib.r = currentRow = parseInt(row.attrib.r, 10) + rowsInserted;;
            rows.push(row);
            
            // Loop over row/c (columns)
            var cols = [],
                currentCell = null,
                cellsInserted = 0;

            row.findall("c").forEach(function(col) {
                var appendCell = true;

                // If c[@t="s"] (string column), look up /c/v@text as integer in
                // `this.sharedStrings`
                if(col.attrib.t === "s") {
                    var cellValue   = col.find("v"),
                        stringIndex = parseInt(cellValue.text, 10),
                        string      = self.sharedStrings[stringIndex];

                    // Determine if we have a substitution
                    if(string !== undefined) {

                        // Replace each placeholder
                        var placeholders = self.extractPlaceholders(string);
                        placeholders.forEach(function(placeholder) {

                            var substitution = substitutions[placeholder.name];
                            if(substitution === undefined) {
                                return;
                            }

                            var substituted = self.substituteScalar(col, string, placeholder, substitution);
                            if(substituted !== undefined) {
                                string = substituted; // in case we have multiple placeholders
                            } else if(placeholder.full && substitution instanceof Array) {
                                // A column or table substitution
                                if(placeholder.type === "table") { // multiple rows
                                    // TODO
                                } else if(placeholder.type === "normal") { // columns on this row
                                    appendCell = false; // don't double-insert cells
                                    --cellsInserted; // we technically delete one before we start adding back
                                    substitution.forEach(function(substitutionElement, idx) {
                                        ++cellsInserted;
                                        var newCol = self.cloneElement(col);
                                        self.insertValue(newCol, substitutionElement);
                                        newCol.attrib.r = currentCell = getCurrentCell(currentCell, currentRow, newCol);
                                        cols.push(newCol);
                                    });
                                }
                            }
                        });
                    }
                }

                // if we are inserting rows or columns, we may not want to keep the original cell anymore
                if(appendCell) {
                    col.attrib.r = currentCell = getCurrentCell(currentCell, currentRow, col);
                    cols.push(col);
                }
                
            });

            // We may have inserted columns, so re-build the children of the row
            self.replaceChildren(row, cols);

            // Update row spans attribute
            if(cellsInserted !== 0 && row.attrib.spans) {
                var rowSpan = row.attrib.spans.split(':').map(function(f) { return parseInt(f, 10); });
                rowSpan[1] += cellsInserted;
                row.attrib.spans = rowSpan.join(":");
            }

        });

        // We may have inserted rows, so re-build the children of the sheetData
        self.replaceChildren(sheetData, rows);

        // TODO: Update <mergeCells />

        // Write back the modified XML tree
        self.archive.file(file, etree.tostring(root));
    };

    /**
     * Generate a new binary .xlsx file
     */
    Workbook.prototype.generate = function() {
        var self = this;

        self.writeSharedStrings();
        return self.archive.generate({base64:false,compression:'DEFLATE'});
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

    Workbook.prototype.stringify = function (value) {
        var self = this;

        if(value instanceof Date) {
            return value.toISOString();
        } else if(typeof(value) === "number" || typeof(value) === "boolean") {
            return Number(value).toString();
        } else {
            return String(value).toString();
        }
    };

    Workbook.prototype.insertCellValue = function(col, substitution) {
        var self = this;

        var cellValue = col.find("v"),
            stringified = self.stringify(substitution);

        if(typeof(substitution) === "number") {
            col.attrib.t = "n";
            cellValue.text = stringified;
        } else if(typeof(substitution) === "boolean" ) {
            col.attrib.t = "b";
            cellValue.text = stringified;
        } else if(substitution instanceof Date) {
            col.attrib.t = "d";
            cellValue.text = stringified;
        } else {
            col.attrib.t = "s";
            cellValue.text = self.stringIndex(stringified);
        }
    };

    Workbook.prototype.substituteScalar = function(col, string, placeholder, substitution) {
        var self = this;

        var cellValue = col.find("v"),
            stringified = self.stringify(substitution),
            newString = stringified;

        if(placeholder.full && typeof(substitution) === "number") {
            col.attrib.t = "n";
            cellValue.text = stringified;
        } else if(placeholder.full && typeof(substitution) === "boolean" ) {
            col.attrib.t = "b";
            cellValue.text = stringified;
        } else if(placeholder.full && substitution instanceof Date) {
            col.attrib.t = "d";
            cellValue.text = stringified;
        } else if(placeholder.full && typeof(substitution) === "string") {
            col.attrib.t = "s";
            self.replaceString(string, stringified);
        } else if(!placeholder.full) {
            col.attrib.t = "s";
            newString = string.replace(placeholder.placeholder, stringified);
            self.replaceString(string, newString);
        } else {
            return null;
        }

        return newString;
    };


    Workbook.prototype.cloneElement = function(element) {
        var self = this;

        var newElement = etree.Element(element.tag, element.attrib);
        newElement.text = element.text;
        newElement.tail = element.tail;
        newElement.getchildren().forEach(function(child) {
            newElement.append(self.cloneElement(child));
        });
        return newElement;
    };

    Workbook.prototype.replaceChildren = function(parent, children) {
        parent.delSlice(0, parent.len());
        children.forEach(function(child) {
            parent.append(child);
        });
    };

    return Workbook;
})();