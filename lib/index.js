/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, module, Buffer */
"use strict";

var path  = require('path'),
    sizeOf = require('image-size'),
    fs = require('fs'),
    etree = require('elementtree');
import zip from "jszip";

module.exports = (function() {

    var DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        CALC_CHAIN_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
        SHARED_STRINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        HYPERLINK_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    /**
     * Create a new workbook. Either pass the raw data of a .xlsx file,
     * or call `loadTemplate()` later.
     */
    var Workbook = function(data, option = { imageRootPath: undefined }) {
        this.archive = null;
        this.sharedStrings = [];
        this.sharedStringsLookup = {};
        this.option = option;
        this.sharedStringsPath = "";
        this.sheets = [];
        this.sheet = null;
        this.workbook = null;
        this.workbookPath = null;
        this.contentTypes = null;
        this.prefix = null;
        this.workbookRels = null;
        this.calChainRel = null;
        this.calcChainPath = "";

        if(data) {
            this.loadTemplate(data);
        }
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
    Workbook.prototype.loadTemplate = function(data) {
        var self = this;

        if(Buffer.isBuffer(data)) {
            data = data.toString('binary');
        }

        self.archive = new zip(data, {base64: false, checkCRC32: true});

        // Load relationships
        var rels = etree.parse(self.archive.file("_rels/.rels").asText()).getroot(),
            workbookPath = rels.find("Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']").attrib.Target;

        self.workbookPath = workbookPath;
        self.prefix       = path.dirname(workbookPath);
        self.workbook     = etree.parse(self.archive.file(workbookPath).asText()).getroot();
        self.workbookRels = etree.parse(self.archive.file(self.prefix + "/" + '_rels' + "/" + path.basename(workbookPath) + '.rels').asText()).getroot();
        self.sheets       = self.loadSheets(self.prefix, self.workbook, self.workbookRels);
        self.calChainRel  = self.workbookRels.find("Relationship[@Type='" + CALC_CHAIN_RELATIONSHIP + "']")

        if (self.calChainRel) {
          self.calcChainPath = self.prefix + "/" + self.calChainRel.attrib.Target;
        }

        self.sharedStringsPath = self.prefix + "/" + self.workbookRels.find("Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']").attrib.Target;
        self.sharedStrings = [];
        etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot().findall('si').forEach(function(si) {
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

        self.contentTypes = etree.parse(self.archive.file('[Content_Types].xml').asText()).getroot();
        var jpgType = self.contentTypes.find('Default[@Extension="jpg"]');
        if(jpgType===null){
            etree.SubElement(self.contentTypes, 'Default', {'ContentType':'image/png', 'Extension':'jpg'});
        }
    };

    /**
     * Interpolate values for the sheet with the given number (1-based) or
     * name (if a string) using the given substitutions (an object).
     */
    Workbook.prototype.substitute = function(sheetName, substitutions) {
        var self = this;

        var sheet = self.loadSheet(sheetName);
        self.sheet = sheet;

        var dimension = sheet.root.find("dimension"),
            sheetData = sheet.root.find("sheetData"),
            currentRow = null,
            totalRowsInserted = 0,
            totalColumnsInserted = 0,
            namedTables = self.loadTables(sheet.root, sheet.filename),
            rows = [],
            drawing = null;

        var rels = self.loadSheetRels(sheet.filename);			   
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
                            if(placeholder.subType==='image' && drawing == null){
                                if(rels){
                                    drawing = self.loadDrawing(sheet.root, sheet.filename, rels.root);
                                }else{
                                    console.log("Need to implement initRels. Or init this with Excel")
                                }
                            }
                            newCellsInserted = self.substituteTable(
                                row, newTableRows,
                                cells, cell,
                                namedTables, substitution, placeholder.key,
                                placeholder, drawing
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
                        } else if(placeholder.type === "image" && placeholder.full) {
                            if(rels != null){
                                if (drawing==null){
                                    drawing = self.loadDrawing(sheet.root, sheet.filename, rels.root);
                                }
                                string = self.substituteImage(cell, string, placeholder, substitution, drawing);
                            }else{
                                console.log("Need to implement initRels. Or init this with Excel")
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
                //Move images for each subsitute array if option is active
                if(self.option["moveImages"] && rels){
                    if (drawing == null){
                        //Maybe we can load drawing at the begining of function and remove all the self.loadDrawing() along the function ?
                        //If we make this, we create all the time the drawing file (like rels file at this moment)
                        drawing = self.loadDrawing(sheet.root, sheet.filename, rels.root);
                    }
                    if(drawing != null){
                        self.moveAllImages(drawing, row.attrib.r, newTableRows.length)
                    }
                }				
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
        self.substituteHyperlinks(rels, substitutions);

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
        if(rels){
            self.archive.file(rels.filename, etree.tostring(rels.root));
        }
        self.archive.file('[Content_Types].xml', etree.tostring(self.contentTypes));
        // Remove calc chain - Excel will re-build, and we may have moved some formulae
        if(self.calcChainPath && self.archive.file(self.calcChainPath)) {
            self.archive.remove(self.calcChainPath);
        }

        self.writeSharedStrings();
        self.writeTables(namedTables);
        self.writeDrawing(drawing);																
    };

    /**
     * Generate a new binary .xlsx file
     */
    Workbook.prototype.generate = function(options) {
        var self = this;

        if(!options) {
            options = {
                base64: false
            }
        }

        return self.archive.generate(options);
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
    Workbook.prototype.loadSheet = function(sheet) {
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
            root: etree.parse(self.archive.file(info.filename).asText()).getroot()
        };
    };

    //Load rels for a sheetName
    Workbook.prototype.loadSheetRels = function (sheetFilename) {
        var self = this;
        var sheetDirectory = path.dirname(sheetFilename),
            sheetName = path.basename(sheetFilename),
            relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/'),
            relsFile = self.archive.file(relsFilename);
        if (relsFile === null) {
            return self.initSheetRels(sheetFilename);
        }
        var rels = {filename: relsFilename, root: etree.parse(relsFile.asText()).getroot()}
        return rels;
    }
    
    Workbook.prototype.initSheetRels = function (sheetFilename) {
        var sheetDirectory = path.dirname(sheetFilename),
            sheetName = path.basename(sheetFilename),
            relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/');
        var element = etree.Element;
        var ElementTree = etree.ElementTree;
        var root = element('Relationships');
        root.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        var relsEtree = new ElementTree(root);
        var rels = {filename: relsFilename, root: relsEtree.getroot()}
        return rels;
    }
    //Load Drawing file
    Workbook.prototype.loadDrawing = function (sheet, sheetFilename, rels) {
        var self = this;
        var sheetDirectory = path.dirname(sheetFilename),
            sheetName = path.basename(sheetFilename),
            drawing = {filename: '', root: null};
        var drawingPart = sheet.find("drawing");
        if (drawingPart === null) {
            drawing = self.initDrawing(sheet, rels);										
            return drawing;
        }
        var relationshipId = drawingPart.attrib['r:id'],
            target = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target,
            drawingFilename = path.join(sheetDirectory, target).replace(/\\/g, '/'),
            drawingTree = etree.parse(self.archive.file(drawingFilename).asText());
        drawing.filename = drawingFilename;
        drawing.root = drawingTree.getroot();
        drawing.relFilename = path.dirname(drawingFilename) + '/_rels/' + path.basename(drawingFilename) + '.rels';
        drawing.relRoot = etree.parse(self.archive.file(drawing.relFilename).asText()).getroot();
        return drawing;
    };
    
    Workbook.prototype.addContentType = function(partName, contentType){
        var self = this;
        etree.SubElement(self.contentTypes, 'Override', { 'ContentType':contentType, 'PartName':partName});																		
    }
                                        
    Workbook.prototype.initDrawing = function (sheet, rels) {
        var self = this;
        var maxId = self.findMaxId(rels, 'Relationship', 'Id', /rId(\d*)/);
        var rel = etree.SubElement(rels, 'Relationship');
        sheet.insert(sheet._children.length, etree.Element('drawing', {'r:id':'rId'+maxId}));
        rel.set('Id', 'rId' + maxId);
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing');
        var drawing = {};
        var drawingFilename = 'drawing' + self.findMaxFileId(/xl\/drawings\/drawing\d*\.xml/,/drawing(\d*)\.xml/ ) + '.xml';
        rel.set('Target', '../drawings/' + drawingFilename);
        drawing.root = etree.Element('xdr:wsDr');
        drawing.root.set('xmlns:xdr', "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        drawing.root.set('xmlns:a', "http://schemas.openxmlformats.org/drawingml/2006/main");
        drawing.filename = 'xl/drawings/' + drawingFilename;
        drawing.relFilename = 'xl/drawings/_rels/' + drawingFilename + '.rels'
        drawing.relRoot = etree.Element('Relationships');
        drawing.relRoot.set('xmlns', "http://schemas.openxmlformats.org/package/2006/relationships");
        self.addContentType('/'+drawing.filename, 'application/vnd.openxmlformats-officedocument.drawing+xml');
        return drawing;
    };

    //Write Drawing file
    Workbook.prototype.writeDrawing = function (drawing) {
        var self = this;
        if (drawing!==null){
            self.archive.file(drawing.filename, etree.tostring(drawing.root));
            self.archive.file(drawing.relFilename, etree.tostring(drawing.relRoot));
        }
    };

    //Move all images after fromRow of nbRow row
    Workbook.prototype.moveAllImages = function(drawing, fromRow, nbRow){
        var self = this;
        drawing.root.getchildren().forEach(function(drawElement){
            if(drawElement.tag == "xdr:twoCellAnchor"){
                self._moveTwoCellAnchor(drawElement, fromRow, nbRow)
            }
            //TODO : make the other tags image
        })
    };
    
    //Move TwoCellAnchor tag images after fromRow of nbRow row
    Workbook.prototype._moveTwoCellAnchor = function(drawingElement, fromRow, nbRow){
        var self = this;
        var _moveImage = function(drawingElement, fromRow, nbRow){
            var from = Number.parseInt(drawingElement.find('xdr:from').find('xdr:row').text, 10) + Number.parseInt(nbRow, 10)
            drawingElement.find('xdr:from').find('xdr:row').text = from
            var to = Number.parseInt(drawingElement.find('xdr:to').find('xdr:row').text, 10) + Number.parseInt(nbRow, 10)
            drawingElement.find('xdr:to').find('xdr:row').text = to
        }
        if(self.option["moveSameLineImages"]){
            if(parseInt(drawingElement.find('xdr:from').find('xdr:row').text) + 1 >= fromRow){
                _moveImage(drawingElement, fromRow, nbRow)
            }
        }else{
            if(parseInt(drawingElement.find('xdr:from').find('xdr:row').text) + 1 > fromRow){
                _moveImage(drawingElement, fromRow, nbRow)
            }
        }
    };																	
    

    // Load tables for a given sheet
    Workbook.prototype.loadTables = function(sheet, sheetFilename) {
        var self = this;

        var sheetDirectory = path.dirname(sheetFilename),
            sheetName      = path.basename(sheetFilename),
            relsFilename   = sheetDirectory + "/" + '_rels' + "/" + sheetName + '.rels',
            relsFile       = self.archive.file(relsFilename),
            tables         = []; // [{filename: ..., root: ....}]

        if(relsFile === null) {
            return tables;
        }

        var rels = etree.parse(relsFile.asText()).getroot();

        sheet.findall("tableParts/tablePart").forEach(function(tablePart) {
            var relationshipId = tablePart.attrib['r:id'],
                target         = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target,
                tableFilename  = target.replace('..', self.prefix),
                tableTree      = etree.parse(self.archive.file(tableFilename).asText());

            tables.push({
                filename: tableFilename,
                root: tableTree.getroot()
            });
        });

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
    Workbook.prototype.substituteHyperlinks = function(rels, substitutions) {
      let self = this;
      etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot();
      if(rels === null) {
        return;
      }
      const relationships = rels.root._children;
      relationships.forEach(function(relationship){
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
        var re = /\${(?:(.+?):)?(.+?)(?:\.(.+?))?(?::(.+?))??}/g;

        var match = null, matches = [];
        while((match = re.exec(string)) !== null) {
            matches.push({
                placeholder: match[0],
                type: match[1] || 'normal',
                name: match[2],
                key: match[3],
                subType: match[4],
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
    Workbook.prototype.substituteTable = function(row, newTableRows, cells, cell, namedTables, substitution, key, placeholder, drawing) {
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
                    } else if(placeholder.subType=='image' && value!=""){
                        self.substituteImage(cell, placeholder.placeholder, placeholder, value, drawing);
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
                    } else if(placeholder.subType=='image' && value!=''){
                        self.substituteImage(newCell, placeholder.placeholder, placeholder, value, drawing);
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

    Workbook.prototype.substituteImage = function (cell, string, placeholder, substitution, drawing) {
        var self = this;
        var self = this;
        self.substituteScalar(cell, string, placeholder, '');
        if (substitution==null || substitution==""){
            return true;
        }
        //get max refid
        //update rel file.
        var maxId = self.findMaxId(drawing.relRoot, 'Relationship', 'Id', /rId(\d*)/);
        var maxFildId = self.findMaxFileId(/xl\/media\/image\d*.jpg/, /image(\d*)\.jpg/);
        var rel = etree.SubElement(drawing.relRoot, 'Relationship');
        rel.set('Id', 'rId' + maxId);
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
        
        rel.set('Target', '../media/image' + maxFildId + '.jpg');
        function toArrayBuffer(buffer) {
            var ab = new ArrayBuffer(buffer.length);
            var view = new Uint8Array(ab);
            for (var i = 0; i < buffer.length; ++i) {
                view[i] = buffer[i];
            }
            return ab;
        };
        substitution = self.imageToBuffer(substitution)
        //put image to media.
        self.archive.file('xl/media/image' + maxFildId + '.jpg', toArrayBuffer(substitution), {binary:true, base64:false});
        var dimension = sizeOf(substitution);
        var imageWidth = self.pixelsToEMUs(dimension.width);
        var imageHeight = self.pixelsToEMUs(dimension.height);
        //var sheet = self.loadSheet(self.substitueSheetName);
        var imageInMergeCell = false;
        self.sheet.root.findall("mergeCells/mergeCell").forEach(function(mergeCell) {
            //If image is in merge cell, fit the image
            if(self.cellInMergeCells(cell, mergeCell)){
                var mergeCellWidth = self.getWidthMergeCell(mergeCell, self.sheet)
                var mergeCellHeight = self.getHeightMergeCell(mergeCell, self.sheet)
                var mergeWidthEmus = self.columnWidthToEMUs(mergeCellWidth);
                var mergeHeightEmus = self.rowHeightToEMUs(mergeCellHeight);
                /*if(imageWidth <= mergeWidthEmus && imageHeight <= mergeHeightEmus){
                    //Image as more little than the merge cell
                    imageWidth = mergeWidthEmus;
                    imageHeight = mergeHeightEmus;
                }*/
                var widthRate = imageWidth / mergeWidthEmus;
                var heightRate = imageHeight / mergeHeightEmus;
                if(widthRate > heightRate){
                    imageWidth = Math.floor(imageWidth / widthRate);
                    imageHeight = Math.floor(imageHeight / widthRate);
                }else{
                    imageWidth = Math.floor(imageWidth / heightRate);
                    imageHeight = Math.floor(imageHeight / heightRate);
                }
                imageInMergeCell = true;
            }
        })
        if(imageInMergeCell == false){
            var ratio = 100;
            if(self.option && self.option.imageRatio){
                ratio = self.option.imageRatio;
            }
            if(ratio <= 0){
                ratio = 100;
            }
            imageWidth = Math.floor(imageWidth * ratio / 100);
            imageHeight = Math.floor(imageHeight * ratio / 100);
        }
        var imagePart = etree.SubElement(drawing.root, 'xdr:oneCellAnchor');
        var fromPart = etree.SubElement(imagePart, 'xdr:from');
        var fromCol = etree.SubElement(fromPart, 'xdr:col');
        fromCol.text = (self.charToNum(self.splitRef(cell.attrib.r).col)-1).toString();
        var fromColOff = etree.SubElement(fromPart, 'xdr:colOff');
        fromColOff.text = '0';
        var fromRow = etree.SubElement(fromPart, 'xdr:row');
        fromRow.text = (self.splitRef(cell.attrib.r).row-1).toString();
        var fromRowOff = etree.SubElement(fromPart, 'xdr:rowOff');
        fromRowOff.text = '0';
        var extImagePart = etree.SubElement(imagePart, 'xdr:ext', { cx:imageWidth, cy:imageHeight });
        var picNode = etree.SubElement(imagePart, 'xdr:pic');
        var nvPicPr = etree.SubElement(picNode, 'xdr:nvPicPr');
        var cNvPr = etree.SubElement(nvPicPr, 'xdr:cNvPr', {id: maxId, name: 'image_' + maxId, descr: ''});
        var cNvPicPr = etree.SubElement(nvPicPr, 'xdr:cNvPicPr');
        var picLocks = etree.SubElement(cNvPicPr, 'a:picLocks', {noChangeAspect: '1'})
        var blipFill = etree.SubElement(picNode, 'xdr:blipFill');
        var blip = etree.SubElement(blipFill, 'a:blip', {
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "r:embed": "rId" + maxId
        });
        var stretch = etree.SubElement(blipFill, 'a:stretch');
        var fillRect = etree.SubElement(stretch, 'a:fillRect');
        var spPr = etree.SubElement(picNode, 'xdr:spPr');
        var xfrm = etree.SubElement(spPr, 'a:xfrm');
        var off = etree.SubElement(xfrm, 'a:off', {x:"0", y:"0" });
        var ext = etree.SubElement(xfrm, 'a:ext', { cx:imageWidth, cy:imageHeight });
        var prstGeom = etree.SubElement(spPr, 'a:prstGeom', {'prst': 'rect'});
        var avLst = etree.SubElement(prstGeom, 'a:avLst');
        var clientData = etree.SubElement(imagePart, 'xdr:clientData');
        return true;
    }

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

    Workbook.prototype.getWidthCell = function(numCol, sheet){
        var defaultWidth = sheet.root.find("sheetFormatPr").attrib["defaultColWidth"]
        if(!defaultWidth){
            //TODO : Check why defaultColWidth is not set ? 
            defaultWidth = 11.42578125
        }
        var finalWidth = defaultWidth;
        sheet.root.findall("cols/col").forEach(function(col) {
            if(numCol >= col.attrib["min"] && numCol <= col.attrib["max"]){
                if(col.attrib["width"] != undefined){
                    finalWidth = col.attrib["width"]
                }
            }
        })
        return Number.parseFloat(finalWidth);
    }
    Workbook.prototype.getWidthMergeCell = function(mergeCell, sheet){
        var self = this;
        var mergeWidth = 0;
        var mergeRange    = self.splitRange(mergeCell.attrib.ref),
            mergeStartCol = self.charToNum(self.splitRef(mergeRange.start).col),
            mergeEndCol   = self.charToNum(self.splitRef(mergeRange.end).col);
        for(let i = mergeStartCol; i < mergeEndCol + 1; i++){
            mergeWidth += self.getWidthCell(i, sheet);
        }
        return mergeWidth;
    }
    Workbook.prototype.getHeightCell = function(numRow, sheet){
        var defaultHight = sheet.root.find("sheetFormatPr").attrib["defaultRowHeight"]
        var finalHeight = defaultHight;
        sheet.root.findall("sheetData/row").forEach(function(row) {
            if(numRow == row.attrib["r"]){
                if(row.attrib["ht"] != undefined){
                    finalHeight = row.attrib["ht"]
                }
            }
        })
        return Number.parseFloat(finalHeight);
    }
    Workbook.prototype.getHeightMergeCell = function(mergeCell, sheet){
        var self = this;
        var mergeHeight = 0;
        var mergeRange    = self.splitRange(mergeCell.attrib.ref),
            mergeStartRow = self.splitRef(mergeRange.start).row,
            mergeEndRow   = self.splitRef(mergeRange.end).row;
        for(let i = mergeStartRow; i < mergeEndRow + 1; i++){
            mergeHeight += self.getHeightCell(i, sheet);
        }
        return mergeHeight;
    }

    Workbook.prototype.getNbRowOfMergeCell = function(mergeCell){
        var self = this;
        var mergeRange    = self.splitRange(mergeCell.attrib.ref),
            mergeStartRow = self.splitRef(mergeRange.start).row,
            mergeEndRow   = self.splitRef(mergeRange.end).row;
        return mergeEndRow - mergeStartRow +1 ;
    }

    Workbook.prototype.pixelsToEMUs = function (pixels) {
        return Math.round(pixels * 914400 / 96);
    }

    Workbook.prototype.columnWidthToEMUs = function (width) {
        // TODO : This is not the true. Change with true calcul
        // can find help here : 
        // https://docs.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths
        // https://stackoverflow.com/questions/58021996/how-to-set-the-fixed-column-width-values-in-inches-apache-poi
        // https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Sheet.html#setColumnWidth-int-int-
        // https://poi.apache.org/apidocs/dev/org/apache/poi/util/Units.html
        // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
        // http://lcorneliussen.de/raw/dashboards/ooxml/
        return this.pixelsToEMUs(width * 7.625579987895905);
    }

    Workbook.prototype.rowHeightToEMUs = function (height) {
        //TODO : need to be verify
        return Math.round(height / 72 * 914400);
    }

    Workbook.prototype.findMaxFileId = function(fileNameRegex, idRegex){
        var self = this;
        var files = self.archive.file(fileNameRegex);
        var maxFile = files.reduce(function(p, c){
            if(p==null){
                return c.name;
            }
            return p.name>c.name? p.name : c.name;
        }, null);
        var maxid = 0;
        if(maxFile!=null){
            maxid = idRegex.exec(maxFile)[1];
        }
        maxid++;
        return maxid;
    }

    Workbook.prototype.cellInMergeCells = function(cell, mergeCell){
        var self = this;
        var cellCol = self.charToNum(self.splitRef(cell.attrib.r).col);
        var cellRow = self.splitRef(cell.attrib.r).row;
        var mergeRange    = self.splitRange(mergeCell.attrib.ref),
            mergeStartCol = self.charToNum(self.splitRef(mergeRange.start).col),
            mergeEndCol   = self.charToNum(self.splitRef(mergeRange.end).col),
            mergeStartRow = self.splitRef(mergeRange.start).row,
            mergeEndRow   = self.splitRef(mergeRange.end).row;
        if(cellCol >= mergeStartCol && cellCol <= mergeEndCol ){
            if(cellRow >= mergeStartRow && cellRow <= mergeEndRow){
                return true;
            }
        }
        return false;
    }

    Workbook.prototype.isUrl = function(str) {
        var pattern = new RegExp('^(https?:\\/\\/)?'+ // protocol
          '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|'+ // domain name
          '((\\d{1,3}\\.){3}\\d{1,3}))'+ // OR ip (v4) address
          '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*'+ // port and path
          '(\\?[;&a-z\\d%_.~+=-]*)?'+ // query string
          '(\\#[-a-z\\d_]*)?$','i'); // fragment locator
        return !!pattern.test(str);
    }

    Workbook.prototype.toArrayBuffer = function(buffer) {
        var ab = new ArrayBuffer(buffer.length);
        var view = new Uint8Array(ab);
        for (var i = 0; i < buffer.length; ++i) {
            view[i] = buffer[i];
        }
        return ab;
    };

    Workbook.prototype.imageToBuffer = function(imageObj){
        //TODO : I think I can make this function more graceful
        if(!imageObj){
            return null;
        }
        if(imageObj instanceof Buffer){
            return imageObj
        }
        else{
            if(typeof(imageObj) === 'string'  || imageObj instanceof String){
                imageObj = imageObj.toString();
                //if(this.isUrl(imageObj)){
                    // TODO
                //}else{
                    if("imageRootPath" in this.option && fs.existsSync(this.option.imageRootPath + "/" + imageObj)){
                        //get the Absolute path file
                        return Buffer.from(fs.readFileSync(this.option.imageRootPath + "/" + imageObj, { encoding: 'base64' }), 'base64');
                    }else{
                        if(fs.existsSync(imageObj)){
                            //get the relatif path file
                            return Buffer.from(fs.readFileSync(imageObj, { encoding: 'base64' }), 'base64');
                        }
                    }
                //}
                try {
                    var buff = Buffer.from(imageObj, 'base64')
                    return buff;
                } catch (error) {
                    console.log("this is NOT a base64 string")
                    return null;
                }
            }
        }
    }
    
    Workbook.prototype.findMaxId = function (element, tag, attr, idRegex) {
        var maxId = 0;
        element.findall(tag).forEach((element)=> {
            var match = idRegex.exec(element.attrib[attr]);
            if (match == null) {
                throw new Error("Can not find the id!");
            }
            var cid = parseInt(match[1]);
            if(cid > maxId){
                maxId = cid;
            }
        })
        return ++maxId;
    }

    return Workbook;
})();
