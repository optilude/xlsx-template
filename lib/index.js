/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, module, Buffer */
"use strict";

var path  = require('path'),
    JSZip   = require('jszip'),
    etree = require('elementtree');

module.exports = (function() {

    var DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        CALC_CHAIN_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
        SHARED_STRINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        HYPERLINK_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    /**
     * Create a new workbook. and call `loadTemplate` later.
     */
    var Workbook = function() {
        var self = this;
        self.archive = null;
        self.sharedStrings = [];
        self.sharedStringsLookup = {};
    };

    var _get_simple = function (obj, desc) {
        if (desc.indexOf("[") >=0 ) {
            var specification = desc.split(/[[[\]]/);
            var property = specification[0];
            var index = specification[1];
            return obj[property][index];
        }

        return obj[desc];
    }

    /**
     * Based on http://stackoverflow.com/questions/8051975
     * Mimic https://lodash.com/docs#get
     */
    var _get = function(obj, desc, defaultValue) {
        var arr = desc.split('.');
        try {
            while (arr.length) {
                obj = _get_simple(obj, arr.shift());
            }
        } catch(ex) {
            /* invalid chain */
            obj = undefined;
        }
        return obj === undefined ? defaultValue : obj;
    }

    /**
    * Delete unused sheets if needed
    */
    Workbook.prototype.deleteSheet = function(sheetName){
      var self = this;
      var sheet = self.loadSheet(sheetName);

      var sh = self.workbook.find("sheets/sheet[@sheetId='" + sheet.id + "']");
      self.workbook.find("sheets").remove(sh);

      var rel = self.workbookRels.find("Relationship[@Id='" + sh.attrib['r:id'] + "']");
      self.workbookRels.remove(rel);

      self._rebuild();
      return self
    };

    /**
    * Clone sheets in current workbook template
    */
    Workbook.prototype.copySheet = function(sheetName, copyName){
      var self = this;
      var sheet = self.loadSheet(sheetName); //filename, name , id, root
      var newSheetIndex = (self.workbook.findall("sheets/sheet").length+1).toString();
      var fileName = 'worksheets' + '/' + 'sheet' + newSheetIndex + '.xml';
      var arcName = self.prefix + '/' + fileName;

      self.archive.file(arcName, etree.tostring(sheet.root) );
      self.archive.files[arcName].options.binary = true;

      var newSheet = etree.SubElement( self.workbook.find('sheets'), 'sheet' );
      newSheet.attrib.name = copyName || 'Sheet' + newSheetIndex;
      newSheet.attrib.sheetId = newSheetIndex;
      newSheet.attrib['r:id'] = 'rId' + newSheetIndex;

      var newRel = etree.SubElement(self.workbookRels, 'Relationship');
      newRel.attrib.Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
      newRel.attrib.Target = fileName;

      self._rebuild();
//    TODO: work with "definedNames"
//    var defn = etree.SubElement(self.workbook.find('definedNames'), 'definedName');
//
      return self
    };


    /**
    *  Partially rebuild after copy/delete sheets
    */
    Workbook.prototype._rebuild = function(){
    //each <sheet> 'r:id' attribute in '\xl\workbook.xml'
    //must point to correct <Relationship> 'Id' in xl\_rels\workbook.xml.rels
      var self = this;
      var order = ['worksheet', 'theme', 'styles','sharedStrings'];

      self.workbookRels.findall("*")
      .sort(function(rel1, rel2){ //using order
        var index1 = order.indexOf( path.basename(rel1.attrib.Type) );
        var index2 = order.indexOf( path.basename(rel2.attrib.Type) );
        if ((index1 + index2) == 0) {
            if(rel1.attrib.Id && rel2.attrib.Id) return rel1.attrib.Id.substring(3) - rel2.attrib.Id.substring(3);
          return rel1._id - rel2._id;
        }
        return index1 - index2
      })
      .forEach(function(item, index) {
        item.attrib.Id = 'rId' + (index+1);
      })

      self.workbook.findall("sheets/sheet").forEach(function(item, index) {
        item.attrib['r:id'] = 'rId' + (index+1);
        item.attrib.sheetId = (index+1).toString();
      })

      self.archive.file(self.prefix + '/' + '_rels' + '/' + path.basename(self.workbookPath) + '.rels', etree.tostring(self.workbookRels));
      self.archive.file(self.workbookPath, etree.tostring(self.workbook));
      self.sheets = self.loadSheets(self.prefix, self.workbook, self.workbookRels);
    }


    /**
     * Load a .xlsx file from a byte array.
     */
    Workbook.prototype.loadTemplate = async function(data) {
        var self = this;

        if(Buffer.isBuffer(data)) {
            data = data.toString('binary');
        }

        self.archive = await new JSZip.loadAsync(data, {base64: false, checkCRC32: true});

        // Load relationships
        var rels = etree.parse(await self.archive.file("_rels/.rels").async("string")).getroot(),
            workbookPath = rels.find("Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']").attrib.Target;

        self.workbookPath = workbookPath;
        self.prefix       = path.dirname(workbookPath);
        self.workbook     = etree.parse(await self.archive.file(workbookPath).async("string")).getroot();
        self.workbookRels = etree.parse(await self.archive.file(self.prefix + "/" + '_rels' + "/" + path.basename(workbookPath) + '.rels').async("string")).getroot();
        self.sheets       = self.loadSheets(self.prefix, self.workbook, self.workbookRels);
        self.calChainRel  = self.workbookRels.find("Relationship[@Type='" + CALC_CHAIN_RELATIONSHIP + "']")

        if (self.calChainRel) {
          self.calcChainPath = self.prefix + "/" + self.calChainRel.attrib.Target;
        }

        self.sharedStringsPath = self.prefix + "/" + self.workbookRels.find("Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']").attrib.Target;
        self.sharedStrings = [];
        etree.parse(await self.archive.file(self.sharedStringsPath).async("string")).getroot().findall('si').forEach(function(si) {
            var t = {text:''};
            si.findall('t').forEach(function(tmp){
                t.text += tmp.text;
            });
            si.findall('r/t').forEach(function(tmp){
                t.text += tmp.text;
            });
            self.sharedStrings.push(t.text);
            self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1;
        });
    };

    /**
     * Interpolate values for the sheet with the given number (1-based) or
     * name (if a string) using the given substitutions (an object).
     */
    Workbook.prototype.substitute = async function(sheetName, substitutions) {
        var self = this;

        var sheet = await self.loadSheet(sheetName);

        var dimension = sheet.root.find("dimension"),
            sheetData = sheet.root.find("sheetData"),
            currentRow = null,
            totalRowsInserted = 0,
            totalColumnsInserted = 0,
            namedTables = await self.loadTables(sheet.root, sheet.filename),
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
                        var substitution = _get(substitutions, placeholder.name, ''),
                            newCellsInserted = 0;

                        if(placeholder.full && placeholder.type === "table" && substitution instanceof Array) {
                            newCellsInserted = self.substituteTable(
                                row, newTableRows,
                                cells, cell,
                                namedTables, substitution, placeholder.key
                            );

                            // don't double-insert cells
                            // this applies to arrays only, incorrectly applies to object arrays when there a single row, thus not rendering single row
                            if (newCellsInserted !== 0 || substitution.length) {
                                if (substitution.length === 1) {
                                    appendCell = true;
                                }
                                if (substitution[0][placeholder.key] instanceof Array) {
                                    appendCell = false;
                                }
                            }

                            // Did we insert new columns (array values)?
                            if(newCellsInserted !== 0) {
                                cellsInserted += newCellsInserted;
                                self.pushRight(self.workbook, sheet.root, cell.attrib.r, newCellsInserted);
                            }
                        } else if(placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                            appendCell = false; // don't double-insert cells
                            newCellsInserted = self.substituteArray(
                                cells, cell, substitution
                            );

                            if(newCellsInserted !== 0) {
                                cellsInserted += newCellsInserted;
                                self.pushRight(self.workbook, sheet.root, cell.attrib.r, newCellsInserted);
                            }
                        } else {
                            if (placeholder.key) {
                                substitution = _get(substitutions, placeholder.name + '.' + placeholder.key);
                            }
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
            if(cellsInserted !== 0) {
                self.updateRowSpan(row, cellsInserted);

                if(cellsInserted > totalColumnsInserted) {
                    totalColumnsInserted = cellsInserted;
                }

            }

            // Add newly inserted rows
            if(newTableRows.length > 0) {
                newTableRows.forEach(function(row) {
                    rows.push(row);
                    ++totalRowsInserted;
                });
                self.pushDown(self.workbook, sheet.root, namedTables, currentRow, newTableRows.length);
            }

        }); // rows loop

        // We may have inserted rows, so re-build the children of the sheetData
        self.replaceChildren(sheetData, rows);

        // Update placeholders in table column headers
        self.substituteTableColumnHeaders(namedTables, substitutions);

        // Update placeholders in hyperlinks
        await self.substituteHyperlinks(sheet.filename, substitutions);

        // Update <dimension /> if we added rows or columns
        if(dimension) {
            if(totalRowsInserted > 0 || totalColumnsInserted > 0) {
                var dimensionRange = self.splitRange(dimension.attrib.ref),
                    dimensionEndRef = self.splitRef(dimensionRange.end);

                dimensionEndRef.row += totalRowsInserted;
                dimensionEndRef.col = self.numToChar(self.charToNum(dimensionEndRef.col) + totalColumnsInserted);
                dimensionRange.end = self.joinRef(dimensionEndRef);

                dimension.attrib.ref = self.joinRange(dimensionRange);
            }
        }

       //Here we are forcing the values in formulas to be recalculated
      // existing as well as just substituted
        sheetData.findall("row").forEach(function(row) {
          row.findall("c").forEach(function(cell) {
            var formulas = cell.findall('f');
            if (formulas && formulas.length > 0) {
              cell.findall('v').forEach(function(v){
                cell.remove(v);
              });
            }
          })
        })

        // Write back the modified XML trees
        self.archive.file(sheet.filename, etree.tostring(sheet.root));
        self.archive.file(self.workbookPath, etree.tostring(self.workbook));

        // Remove calc chain - Excel will re-build, and we may have moved some formulae
        if(self.calcChainPath && self.archive.file(self.calcChainPath)) {
            self.archive.remove(self.calcChainPath);
        }

        await self.writeSharedStrings();
        self.writeTables(namedTables);
    };

    /**
     * Generate a new binary .xlsx file
     */
    Workbook.prototype.generate = async function(options) {
        var self = this;

        if(!options) {
            options = {
                type: "binarystring"
            }
        }

        return await self.archive.generateAsync(options);
    };

    // Helpers

    // Write back the new shared strings list
    Workbook.prototype.writeSharedStrings = async function() {
        var self = this;

        var root = etree.parse(await self.archive.file(self.sharedStringsPath).async("string")).getroot(),
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

    // Get a list of sheet ids, names and filenames
    Workbook.prototype.loadSheets = function(prefix, workbook, workbookRels) {
        var sheets = [];

        workbook.findall("sheets/sheet").forEach(function(sheet) {
            var sheetId      = sheet.attrib.sheetId,
                relId        = sheet.attrib['r:id'],
                relationship = workbookRels.find("Relationship[@Id='" + relId + "']"),
                filename     = prefix + "/" + relationship.attrib.Target;

            sheets.push({
                id: parseInt(sheetId, 10),
                name: sheet.attrib.name,
                filename: filename
            });
        });

        return sheets;
    };

    // Get sheet a sheet, including filename and name
    Workbook.prototype.loadSheet = async function(sheet) {
        var self = this;

        var info = null;

        for(var i = 0; i < self.sheets.length; ++i) {
            if((typeof(sheet) === "number" && self.sheets[i].id === sheet) || (self.sheets[i].name === sheet))  {
                info = self.sheets[i];
                break;
            }
        }

        if(info === null && (typeof(sheet) === "number")){
            //Get the sheet that corresponds to the 0 based index if the id does not work
            info = self.sheets[sheet - 1];
        }

        if(info === null) {
            throw new Error("Sheet " + sheet + " not found");
        }

        return {
            filename: info.filename,
            name: info.name,
            id: info.id,
            root: etree.parse(await self.archive.file(info.filename).async("string")).getroot()
        };
    };

    // Load tables for a given sheet
    Workbook.prototype.loadTables = async function(sheet, sheetFilename) {
        var self = this;

        var sheetDirectory = path.dirname(sheetFilename),
            sheetName      = path.basename(sheetFilename),
            relsFilename   = sheetDirectory + "/" + '_rels' + "/" + sheetName + '.rels',
            relsFile       = self.archive.file(relsFilename),
            tables         = []; // [{filename: ..., root: ....}]

        if(relsFile === null) {
            return tables;
        }

        var rels = etree.parse(await relsFile.async("string")).getroot();

        for(let tablePart of sheet.findall("tableParts/tablePart")) {
            var relationshipId = tablePart.attrib['r:id'],
                target         = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target,
                tableFilename  = target.replace('..', self.prefix),
                tableTree      = etree.parse(await self.archive.file(tableFilename).async("string"));

            tables.push({
                filename: tableFilename,
                root: tableTree.getroot()
            });
        }

        return tables;
    };

    // Write back possibly-modified tables
    Workbook.prototype.writeTables = function(tables) {
        var self = this;

        tables.forEach(function(namedTable) {
            self.archive.file(namedTable.filename, etree.tostring(namedTable.root));
        });
    };

    //Perform substitution in hyperlinks
    Workbook.prototype.substituteHyperlinks = async function(sheetFilename, substitutions) {
      let self = this;
      const sheetDirectory = path.dirname(sheetFilename),
        sheetName      = path.basename(sheetFilename),
        relsFilename   = sheetDirectory + "/" + '_rels' + "/" + sheetName + '.rels',
        relsFile       = self.archive.file(relsFilename);

      etree.parse(await self.archive.file(self.sharedStringsPath).async("string")).getroot();

      if(relsFile === null) {
        return;
      }

      const rels = etree.parse(await relsFile.async("string")).getroot();

      const relationships = rels._children;

      const newRelationships = [];

      relationships.forEach(function(relationship){
        newRelationships.push(relationship);

        if(relationship.attrib.Type === HYPERLINK_RELATIONSHIP) {

          let target = relationship.attrib.Target;

          //Double-decode due to excel double encoding url placeholders
          target = decodeURI(decodeURI(target));
          self.extractPlaceholders(target).forEach(function (placeholder) {
              const substitution = substitutions[placeholder.name];

              if (substitution === undefined) {
                return;
              }
              target = target.replace(placeholder.placeholder, self.stringify(substitution));

              relationship.attrib.Target = encodeURI(target);
            }
          );
        }
      });

      self.replaceChildren(rels, newRelationships);

      self.archive.file(relsFilename, etree.tostring(rels));
    };

    // Perform substitution in table headers
    Workbook.prototype.substituteTableColumnHeaders = function(tables, substitutions) {
        var self = this;

        tables.forEach(function(table) {
            var root = table.root,
                columns = root.find("tableColumns"),
                autoFilter = root.find("autoFilter"),
                tableRange = self.splitRange(root.attrib.ref),
                idx = 0,
                inserted = 0,
                newColumns = [];

            columns.findall("tableColumn").forEach(function(col) {
                ++idx;
                col.attrib.id = Number(idx).toString();
                newColumns.push(col);

                var name = col.attrib.name;

                self.extractPlaceholders(name).forEach(function(placeholder) {
                    var substitution = substitutions[placeholder.name];
                    if(substitution === undefined) {
                        return;
                    }

                    // Array -> new columns
                    if(placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                        substitution.forEach(function(element, i) {
                            var newCol = col;
                            if(i > 0) {
                                newCol = self.cloneElement(newCol);
                                newCol.attrib.id = Number(++idx).toString();
                                newColumns.push(newCol);
                                ++inserted;
                                tableRange.end = self.nextCol(tableRange.end);
                            }
                            newCol.attrib.name = self.stringify(element);
                        });
                    // Normal placeholder
                    } else {
                        name = name.replace(placeholder.placeholder, self.stringify(substitution));
                        col.attrib.name = name;
                    }
                });
            });

            self.replaceChildren(columns, newColumns);

            // Update range if we inserted columns
            if(inserted > 0) {
                columns.attrib.count = Number(idx).toString();
                root.attrib.ref = self.joinRange(tableRange);
                if(autoFilter !== null) {
                    // XXX: This is a simplification that may stomp on some configurations
                    autoFilter.attrib.ref = self.joinRange(tableRange);
                }
            }

            //update ranges for totalsRowCount
            var tableRoot  = table.root,
                tableRange = self.splitRange(tableRoot.attrib.ref),
                tableStart = self.splitRef(tableRange.start),
                tableEnd   = self.splitRef(tableRange.end);

            if (tableRoot.attrib.totalsRowCount) {
                var autoFilter = tableRoot.find("autoFilter");
                if(autoFilter !== null) {
                    autoFilter.attrib.ref = self.joinRange({
                        start: self.joinRef(tableStart),
                        end: self.joinRef(tableEnd),
                    });
                }

                ++tableEnd.row;
                tableRoot.attrib.ref = self.joinRange({
                    start: self.joinRef(tableStart),
                    end: self.joinRef(tableEnd),
                });

            }
        });
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

    // Split a reference into an object with keys `row` and `col` and,
    // optionally, `table`, `rowAbsolute` and `colAbsolute`.
    Workbook.prototype.splitRef = function(ref) {
        var match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)(\$)?([0-9]+)/);
        return {
            table: match && match[1] || null,
            colAbsolute: Boolean(match && match[2]),
            col: match && match[3],
            rowAbsolute: Boolean(match && match[4]),
            row: parseInt(match && match[5], 10)
        };
    };

    // Join an object with keys `row` and `col` into a single reference string
    Workbook.prototype.joinRef = function(ref) {
        return (ref.table?       ref.table + "!" : "") +
               (ref.colAbsolute?             "$" : "") +
                ref.col.toUpperCase()                 +
               (ref.rowAbsolute?             "$" : "" )+
               Number(ref.row).toString();
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

    // Is ref a range?
    Workbook.prototype.isRange = function(ref) {
        return ref.indexOf(':') !== -1;
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

    // Turn a value of any type into a string
    Workbook.prototype.stringify = function (value) {
        if(value instanceof Date) {
            //In Excel date is a number of days since 01/01/1900
            //           timestamp in ms    to days      + number of days from 1900 to 1970
            return Number( (value.getTime()/(1000*60*60*24)) + 25569);
        } else if(typeof(value) === "number" || typeof(value) === "boolean") {
            return Number(value).toString();
        } else if(typeof(value) === "string") {
            return String(value).toString();
        }

        return "";
    };

    // Insert a substitution value into a cell (c tag)
    Workbook.prototype.insertCellValue = function(cell, substitution) {
        var self = this;

        var cellValue = cell.find("v"),
            stringified = self.stringify(substitution);

        if (typeof substitution ==='string' && substitution[0] === '='){
          //substitution, started with '=' is a formula substitution
          var formula = new etree.Element("f");
          formula.text = substitution.substr(1);
          cell.insert(1, formula);
          delete cell.attrib.t;  //cellValue will be deleted later
          return formula.text
        }

        if(typeof(substitution) === "number" || substitution instanceof Date) {
            delete cell.attrib.t;
            cellValue.text = stringified;
        } else if(typeof(substitution) === "boolean" ) {
            cell.attrib.t = "b";
            cellValue.text = stringified;
        } else {
            cell.attrib.t = "s";
            cellValue.text = Number(self.stringIndex(stringified)).toString();
        }

        return stringified;
    };

    // Perform substitution of a single value
    Workbook.prototype.substituteScalar = function(cell, string, placeholder, substitution) {
        var self = this;

        if(placeholder.full) {
            return self.insertCellValue(cell, substitution);
        } else {
            var newString = string.replace(placeholder.placeholder, self.stringify(substitution));
            cell.attrib.t = "s";
            return self.insertCellValue(cell, newString)
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
    Workbook.prototype.substituteTable = function(row, newTableRows, cells, cell, namedTables, substitution, key) {
        var self = this,
            newCellsInserted = 0; // on the original row

        // if no elements, blank the cell, but don't delete it
        if(substitution.length === 0) {
            delete cell.attrib.t;
            self.replaceChildren(cell, []);
        } else {

            var parentTables = namedTables.filter(function(namedTable) {
                var range = self.splitRange(namedTable.root.attrib.ref);
                return self.isWithin(cell.attrib.r, range.start, range.end);
            });

            substitution.forEach(function(element, idx) {
                var newRow, newCell,
                    newCellsInsertedOnNewRow = 0,
                    newCells = [],
                    value = _get(element, key, '');

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

                    // expand named table range if necessary
                    parentTables.forEach(function(namedTable) {
                        var tableRoot = namedTable.root,
                            autoFilter = tableRoot.find("autoFilter"),
                            range = self.splitRange(tableRoot.attrib.ref);

                        if(!self.isWithin(newCell.attrib.r, range.start, range.end)) {
                            range.end = self.nextRow(range.end);
                            tableRoot.attrib.ref = self.joinRange(range);
                            if(autoFilter !== null) {
                                // XXX: This is a simplification that may stomp on some configurations
                                autoFilter.attrib.ref = tableRoot.attrib.ref;
                            }
                        }
                    });
                }
            });
        }

        return newCellsInserted;
    };

    // Clone an element. If `deep` is true, recursively clone children
    Workbook.prototype.cloneElement = function(element, deep) {
        var self = this;

        var newElement = etree.Element(element.tag, element.attrib);
        newElement.text = element.text;
        newElement.tail = element.tail;

        if(deep !== false) {
            element.getchildren().forEach(function(child) {
                newElement.append(self.cloneElement(child, deep));
            });
        }

        return newElement;
    };

    // Replace all children of `parent` with the nodes in the list `children`
    Workbook.prototype.replaceChildren = function(parent, children) {
        parent.delSlice(0, parent.len());
        children.forEach(function(child) {
            parent.append(child);
        });
    };

    // Calculate the current row based on a source row and a number of new rows
    // that have been inserted above
    Workbook.prototype.getCurrentRow = function(row, rowsInserted) {
        return parseInt(row.attrib.r, 10) + rowsInserted;
    };

    // Calculate the current cell based on asource cell, the current row index,
    // and a number of new cells that have been inserted so far
    Workbook.prototype.getCurrentCell = function(cell, currentRow, cellsInserted) {
        var self = this;

        var colRef = self.splitRef(cell.attrib.r).col,
            colNum = self.charToNum(colRef);

        return self.joinRef({
            row: currentRow,
            col: self.numToChar(colNum + cellsInserted)
        });
    };

    // Adjust the row `spans` attribute by `cellsInserted`
    Workbook.prototype.updateRowSpan = function(row, cellsInserted) {
        if(cellsInserted !== 0 && row.attrib.spans) {
            var rowSpan = row.attrib.spans.split(':').map(function(f) { return parseInt(f, 10); });
            rowSpan[1] += cellsInserted;
            row.attrib.spans = rowSpan.join(":");
        }
    };

    // Split a range like "A1:B1" into {start: "A1", end: "B1"}
    Workbook.prototype.splitRange = function(range) {
        var split = range.split(":");
        return {
            start: split[0],
            end: split[1]
        };
    };

    // Join into a a range like "A1:B1" an object like {start: "A1", end: "B1"}
    Workbook.prototype.joinRange = function(range) {
        return range.start + ":" + range.end;
    };

    // Look for any merged cell or named range definitions to the right of
    // `currentCell` and push right by `numCols`.
    Workbook.prototype.pushRight = function(workbook, sheet, currentCell, numCols) {
        var self = this;

        var cellRef = self.splitRef(currentCell),
            currentRow = cellRef.row,
            currentCol = self.charToNum(cellRef.col);

        // Update merged cells on the same row, at a higher column
        sheet.findall("mergeCells/mergeCell").forEach(function(mergeCell) {
            var mergeRange    = self.splitRange(mergeCell.attrib.ref),
                mergeStart    = self.splitRef(mergeRange.start),
                mergeStartCol = self.charToNum(mergeStart.col),
                mergeEnd      = self.splitRef(mergeRange.end),
                mergeEndCol   = self.charToNum(mergeEnd.col);

            if(mergeStart.row === currentRow && currentCol < mergeStartCol) {
                mergeStart.col = self.numToChar(mergeStartCol + numCols);
                mergeEnd.col = self.numToChar(mergeEndCol + numCols);

                mergeCell.attrib.ref = self.joinRange({
                    start: self.joinRef(mergeStart),
                    end: self.joinRef(mergeEnd),
                });
            }
        });

        // Named cells/ranges
        workbook.findall("definedNames/definedName").forEach(function(name) {
            var ref = name.text;

            if(self.isRange(ref)) {
                var namedRange    = self.splitRange(ref),
                    namedStart    = self.splitRef(namedRange.start),
                    namedStartCol = self.charToNum(namedStart.col),
                    namedEnd      = self.splitRef(namedRange.end),
                    namedEndCol   = self.charToNum(namedEnd.col);

                if(namedStart.row === currentRow && currentCol < namedStartCol) {
                    namedStart.col = self.numToChar(namedStartCol + numCols);
                    namedEnd.col = self.numToChar(namedEndCol + numCols);

                    name.text = self.joinRange({
                        start: self.joinRef(namedStart),
                        end: self.joinRef(namedEnd),
                    });
                }
            } else {
                var namedRef = self.splitRef(ref),
                    namedCol = self.charToNum(namedRef.col);

                if(namedRef.row === currentRow && currentCol < namedCol) {
                    namedRef.col = self.numToChar(namedCol + numCols);

                    name.text = self.joinRef(namedRef);
                }
            }

        });
    };

    // Look for any merged cell, named table or named range definitions below
    // `currentRow` and push down by `numRows` (used when rows are inserted).
    Workbook.prototype.pushDown = function(workbook, sheet, tables, currentRow, numRows) {
        var self = this;

	var mergeCells = sheet.find("mergeCells");

        // Update merged cells below this row
        sheet.findall("mergeCells/mergeCell").forEach(function(mergeCell) {
            var mergeRange    = self.splitRange(mergeCell.attrib.ref),
                mergeStart    = self.splitRef(mergeRange.start),
                mergeEnd      = self.splitRef(mergeRange.end);

            if(mergeStart.row > currentRow) {
                mergeStart.row += numRows;
                mergeEnd.row += numRows;

                mergeCell.attrib.ref = self.joinRange({
                    start: self.joinRef(mergeStart),
                    end: self.joinRef(mergeEnd),
                });

            }


		//add new merge cell
            if (mergeStart.row == currentRow) {
              for (var i = 1; i <= numRows; i++) {
                var newMergeCell = self.cloneElement(mergeCell);
                mergeStart.row += 1;
                mergeEnd.row += 1;
                newMergeCell.attrib.ref = self.joinRange({
                  start: self.joinRef(mergeStart),
                  end: self.joinRef(mergeEnd)
                });
                mergeCells.attrib.count += 1;
                mergeCells._children.push(newMergeCell);
              }
            }
        });

        // Update named tables below this row
        tables.forEach(function(table) {
            var tableRoot  = table.root,
                tableRange = self.splitRange(tableRoot.attrib.ref),
                tableStart = self.splitRef(tableRange.start),
                tableEnd   = self.splitRef(tableRange.end);


            if(tableStart.row > currentRow) {
                tableStart.row += numRows;
                tableEnd.row += numRows;

                tableRoot.attrib.ref = self.joinRange({
                    start: self.joinRef(tableStart),
                    end: self.joinRef(tableEnd),
                });

                var autoFilter = tableRoot.find("autoFilter");
                if(autoFilter !== null) {
                    // XXX: This is a simplification that may stomp on some configurations
                    autoFilter.attrib.ref = tableRoot.attrib.ref;
                }
            }
        });

        // Named cells/ranges
        workbook.findall("definedNames/definedName").forEach(function(name) {
            var ref = name.text;

            if(self.isRange(ref)) {
                var namedRange    = self.splitRange(ref),
                    namedStart    = self.splitRef(namedRange.start),
                    namedEnd      = self.splitRef(namedRange.end);

                if(namedStart){
                    if(namedStart.row > currentRow) {
                        namedStart.row += numRows;
                        namedEnd.row += numRows;

                        name.text = self.joinRange({
                            start: self.joinRef(namedStart),
                            end: self.joinRef(namedEnd),
                        });

                    }
                }
            } else {
                var namedRef = self.splitRef(ref);

                if(namedRef.row > currentRow) {
                    namedRef.row += numRows;
                    name.text = self.joinRef(namedRef);
                }
            }

        });
    };

    return Workbook;
})();
