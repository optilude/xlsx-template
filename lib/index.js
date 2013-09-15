/*jshint globalstrict:true, devel:true */
/*global require, module, exports, process, __dirname, Buffer */
"use strict";

var fs    = require('fs'),
    path  = require('path'),
    zip   = require('node-zip'),
    etree = require('elementtree');

module.exports = (function() {

    var DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        SHARED_STRINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

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

        // Load relationships
        var rels = etree.parse(self.archive.file("_rels/.rels").asText()).getroot(),
            workbookPath = rels.find("Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']").attrib.Target;

        self.prefix       = path.dirname(workbookPath);
        self.workbook     = etree.parse(self.archive.file(workbookPath).asText()).getroot();
        self.workbookRels = etree.parse(self.archive.file(path.join(self.prefix, '_rels', path.basename(workbookPath) + '.rels')).asText()).getroot();

        self.sharedStringsPath = path.join(self.prefix, self.workbookRels.find("Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']").attrib.Target);
        self.sharedStrings = [];
        etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot().findall('si/t').forEach(function(t) {
            self.sharedStrings.push(t.text);
            self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1;
        });
    };

    /**
     * Interpolate values for the sheet with the given number (1-based) or
     * name (if a string) using the given substitutions (an object).
     */
    Workbook.prototype.substitute = function(sheet, substitutions) {
        var self = this;

        var sheetFilename = self.getSheetFilename(sheet),
            tree = etree.parse(self.archive.file(sheetFilename).asText()),
            root = tree.getroot();

        // TODO: Find tables
        //      - Load refs from ./tableParts/tablePart@r:id
        //      - Find corresponding table file in worksheets/_rels/sheet${num}.xml.rels under Relationships/Relationship@Target
        //      - Load table XML file. Record table@ref (range) and table/autofilter@ref (range)

        var sheetData = root.find("sheetData"),
            currentRow = null,
            totalRowsInserted = 0,
            rows = [];

        sheetData.findall("row").forEach(function(row) {
            row.attrib.r = currentRow = self.getCurrentRow(row, totalRowsInserted);
            rows.push(row);
            
            var cells = [],
                cellsInserted = 0,
                newTableRows = [];

            row.findall("c").forEach(function(cell) {
                var appendCell = true;
                cell.attrib.r = self.getCurrentCell(cell, currentRow, cellsInserted);

                // If c[@t="s"] (string column), look up /c/v@text as integer in
                // `this.sharedStrings`
                if(cell.attrib.t === "s") {
                    
                    // Look for a shared string that may contain placeholders
                    var cellValue   = cell.find("v"),
                        stringIndex = parseInt(cellValue.text, 10),
                        string      = self.sharedStrings[stringIndex];

                    if(string === undefined) {
                        return;
                    }

                    // Loop over placeholders
                    self.extractPlaceholders(string).forEach(function(placeholder) {

                        // Only substitute things for which we have a substitution
                        var substitution = substitutions[placeholder.name];
                        if(substitution === undefined) {
                            return;
                        }

                        if(placeholder.full && placeholder.type === "table" && substitution instanceof Array) {
                            var newCellsInserted = self.substituteTable(
                                row, newTableRows,
                                cells, cell, substitution, placeholder.key
                            );

                            // Did we insert new columns (array values)?
                            if(newCellsInserted !== 0) {
                                appendCell = false; // don't double-insert cells
                                cellsInserted += newCellsInserted;

                                // TODO: Update table definitions if cell within a table

                            }
                        } else if(placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                            appendCell = false; // don't double-insert cells
                            cellsInserted += self.substituteArray(
                                cells, cell, substitution
                            );

                            // TODO: Update table definitions if cell within a table

                        } else {
                            string = self.substituteScalar(cell, string, placeholder, substitution);
                        }
                    });
                }

                // if we are inserting columns, we may not want to keep the original cell anymore
                if(appendCell) {
                    cells.push(cell);
                }
                
            }); // cells loop

            // We may have inserted columns, so re-build the children of the row
            self.replaceChildren(row, cells);

            // Update row spans attribute
            self.updateRowSpan(row, cellsInserted);

            // Add newly inserted rows
            newTableRows.forEach(function(row) {
                rows.push(row);
                ++totalRowsInserted;
            });

        }); // rows loop

        // We may have inserted rows, so re-build the children of the sheetData
        self.replaceChildren(sheetData, rows);

        // Write back the modified XML trees
        self.archive.file(sheetFilename, etree.tostring(root));
        self.writeSharedStrings();
    };

    /**
     * Generate a new binary .xlsx file
     */
    Workbook.prototype.generate = function() {
        var self = this;

        // XXX: Getting errors with compression DEFLATE
        return self.archive.generate({base64:false /*,compression:'DEFLATE'*/});
    };

    // Helpers

    // Write back the new shared strings list
    Workbook.prototype.writeSharedStrings = function() {
        var self = this;

        var root = etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot(),
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

        self.archive.file(self.sharedStringsPath, etree.tostring(root));
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

    // Get sheet filename
    Workbook.prototype.getSheetFilename = function(sheet) {
        var self = this;

        var sheetReference = (typeof(sheet) === "number") ?
                self.workbook.find("sheets/sheet[@sheetId='" + sheet + "']") :
                self.workbook.find("sheets/sheet[@name='" + sheet + "']");

        if(sheetReference === null) {
            throw new Error("Sheet " + sheet + "not found");
        }

        var sheetId = sheetReference.attrib['r:id'],
            relationship = self.workbookRels.find("Relationship[@Id='" + sheetId + "']");

        return path.join(self.prefix, relationship.attrib.Target);
    };

    // Load tables for a given sheet
    Workbook.prototype.loadTables = function(sheetTree, sheetFilename) {
        var self = this;

        var sheetDirectory = path.dirname(sheetFilename),
            sheetName      = path.basename(sheetFilename);

        var rels = etree.parse(self.archive.file(path.join(sheetDirectory, '_rels', sheetName + '.rels')).asText()).getroot(),
            tables = []; // [{filename: ..., tree: ....}]

        sheetTree.findAll("tableParts/tablePart").forEach(function(tablePart) {
            var relationshipId = tablePart.attrib['r:id'],
                target         = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target,
                tableFilename  = path.join(sheetDirectory, target),
                tableTree      = etree.parse(self.archive.file(tableFilename).asText());

            tables.push({
                filename: tableFilename,
                tree: tableTree
            });
        });

        return tables;
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
        var self = this;
        ref = ref.toUpperCase();
        return ref.replace(/[A-Z]+/, function(match) {
            return self.numToChar(self.charToNum(match) + 1);
        });
    };

    // Get the next row's cell reference given a reference like "B2".
    Workbook.prototype.nextRow = function(ref) {
        ref = ref.toUpperCase();
        return ref.replace(/[0-9]+/, function(match) {
            return (parseInt(match, 10) + 1).toString();
        });
    };

    // Turn a reference like "AA" into a number like 27
    Workbook.prototype.charToNum = function(str) {
        var num = 0;
        for(var idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration) {
            var thisChar = str.charCodeAt(idx) - 64, // A -> 1; B -> 2; ... Z->26
                multiplier = Math.pow(26, iteration);
            num += multiplier * thisChar;
        }
        return num;
    };

    // Turn a number like 27 into a reference like "AA"
    Workbook.prototype.numToChar = function(num) {
        var str = "";


        for(var i = 0; num > 0; ++i) {
            var remainder = num % 26,
                charCode = remainder + 64;
            num = (num - remainder) / 26;

            // Compensate for the fact that we don't represent zero, e.g. A = 1, Z = 26, but AA = 27
            if(remainder === 0) { // 26 -> Z
                charCode = 90;
                --num;
            }

            str = String.fromCharCode(charCode) + str;
        }

        return str;
    };

    // Is ref inside the table defined by startRef and endRef?
    Workbook.prototype.isWithin = function(ref, startRef, endRef) {
        var self = this;

        var start  = self.splitRef(startRef),
            end    = self.splitRef(endRef),
            target = self.splitRef(ref);

        start.col  = self.charToNum(start.col);
        end.col    = self.charToNum(end.col);
        target.col = self.charToNum(target.col);

        return (
            start.row <= target.row && target.row <= end.row &&
            start.col <= target.col && target.col <= end.col
        );

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

    Workbook.prototype.insertCellValue = function(cell, substitution) {
        var self = this;

        var cellValue = cell.find("v"),
            stringified = self.stringify(substitution);

        if(typeof(substitution) === "number") {
            cell.attrib.t = "n";
            cellValue.text = stringified;
        } else if(typeof(substitution) === "boolean" ) {
            cell.attrib.t = "b";
            cellValue.text = stringified;
        } else if(substitution instanceof Date) {
            cell.attrib.t = "d";
            cellValue.text = stringified;
        } else {
            cell.attrib.t = "s";
            cellValue.text = self.stringIndex(stringified);
        }

        return stringified;
    };

    // Perform substitution of a single value
    Workbook.prototype.substituteScalar = function(cell, string, placeholder, substitution) {
        var self = this;

        if(placeholder.full && typeof(substitution) === "string") {
            self.replaceString(string, substitution);
        }

        if(placeholder.full) {
            return self.insertCellValue(cell, substitution);
        } else {
            var newString = string.replace(placeholder.placeholder, self.stringify(substitution));
            cell.attrib.t = "s";
            self.replaceString(string, newString);
            return newString;
        }

    };

    // Perform a columns substitution from an array
    Workbook.prototype.substituteArray = function(cells, cell, substitution) {
        var self = this;

        var newCellsInserted = -1, // we technically delete one before we start adding back
            currentCell = cell.attrib.r;
                                
        // add a cell for each element in the list
        substitution.forEach(function(element) {
            ++newCellsInserted;
            
            if(newCellsInserted > 0) {
                currentCell = self.nextCol(currentCell);
            }

            var newCell = self.cloneElement(cell);
            self.insertCellValue(newCell, element);

            newCell.attrib.r = currentCell;
            cells.push(newCell);
        });

        return newCellsInserted;
    };

    // Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
    // Returns total number of new cells inserted on the original row.
    Workbook.prototype.substituteTable = function(row, newTableRows, cells, cell, substitution, key) {
        var self = this,
            newCellsInserted = 0; // on the original row

        // if no elements, blank the cell, but don't delete it
        if(substitution.length === 0) {
            delete cell.attrib.t;
            self.replaceChildren(cell, []);
        } else {
            substitution.forEach(function(element, idx) {
                var newRow, newCell,
                    newCellsInsertedOnNewRow = 0,
                    newCells = [],
                    value = element[key];

                if(idx === 0) { // insert in the row where the placeholders are
                    
                    if(value instanceof Array) {
                        newCellsInserted = self.substituteArray(cells, cell, value);
                    } else {
                        self.insertCellValue(cell, value);
                    }

                } else { // insert new rows (or reuse rows just inserted)
                    
                    // Do we have an existing row to use? If not, create one.
                    if((idx - 1) < newTableRows.length) {
                        newRow = newTableRows[idx - 1];
                    } else {
                        newRow = self.cloneElement(row, false);
                        newRow.attrib.r = self.getCurrentRow(row, newTableRows.length + 1);
                        newTableRows.push(newRow);
                    }

                    // Create a new cell
                    newCell = self.cloneElement(cell);
                    newCell.attrib.r = self.joinRef({
                        row: newRow.attrib.r,
                        col: self.splitRef(newCell.attrib.r).col
                    });

                    if(value instanceof Array) {
                        newCellsInsertedOnNewRow = self.substituteArray(newCells, newCell, value);
                        
                        // Add each of the new cells created by substituteArray()
                        newCells.forEach(function(newCell) {
                            newRow.append(newCell);
                        });

                        self.updateRowSpan(newRow, newCellsInsertedOnNewRow);
                    } else {
                        self.insertCellValue(newCell, value);

                        // Add the cell that previously held the placeholder
                        newRow.append(newCell);
                    }

                    // TODO: Update table definitions if cell within a table
                }
            });
        }

        return newCellsInserted;
    };

    Workbook.prototype.cloneElement = function(element, deep) {
        var self = this;

        var newElement = etree.Element(element.tag, element.attrib);
        newElement.text = element.text;
        newElement.tail = element.tail;

        if(deep !== false) {
            element.getchildren().forEach(function(child) {
                newElement.append(self.cloneElement(child));
            });
        }

        return newElement;
    };

    Workbook.prototype.replaceChildren = function(parent, children) {
        parent.delSlice(0, parent.len());
        children.forEach(function(child) {
            parent.append(child);
        });
    };

    Workbook.prototype.getCurrentRow = function(row, rowsInserted) {
        return parseInt(row.attrib.r, 10) + rowsInserted;
    };

    Workbook.prototype.getCurrentCell = function(cell, currentRow, cellsInserted) {
        var self = this;

        var colRef = self.splitRef(cell.attrib.r).col,
            colNum = self.charToNum(colRef);

        return self.joinRef({
            row: currentRow,
            col: self.numToChar(colNum + cellsInserted)
        });
    };

    Workbook.prototype.updateRowSpan = function(row, cellsInserted) {
        if(cellsInserted !== 0 && row.attrib.spans) {
            var rowSpan = row.attrib.spans.split(':').map(function(f) { return parseInt(f, 10); });
            rowSpan[1] += cellsInserted;
            row.attrib.spans = rowSpan.join(":");
        }
    };

    return Workbook;
})();