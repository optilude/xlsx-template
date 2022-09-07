import * as etree from "elementtree";
import * as JSZip from "jszip";

export interface TemplatePlaceholder{
    type: string;
    string?: string;
    full: boolean;
    name: string;
    key: string;
    placeholder?: string
}

export interface NamedTable{
    filename: string;
    root: etree.Element;
}

namespace XlsxTemplate {

}

interface OutputByType {
    base64: string;
    uint8array: Uint8Array;
    arraybuffer: ArrayBuffer;
    blob: Blob;
    nodebuffer: Buffer;
}

export type GenerateOptions = {
    type: keyof OutputByType
};

interface RangeSplit
{
    start: string;
    end: string;
}

interface ReferenceAddress
{
    table?: string;
    colAbsolute?: boolean;
    col : string;
    rowAbsolute?: boolean;
    row : number;
}

class XlsxTemplate
{
    public readonly sharedStrings: string[];
    protected readonly workbook: etree.ElementTree;
    public readonly archive: JSZip | null;
    protected readonly workbookPath: string;
    protected readonly calcChainPath?: string;
    public readonly sharedStringsLookup: { [key: string]: number };

    constructor(data? : Buffer, option? : object);
    public deleteSheet(sheetName : string) : this;
    public copySheet(sheetName : string, copyName : string, binary? : boolean) : this;
    public loadTemplate(data : Buffer) : void;
    public substitute(sheetName : string | number, substitutions : Object) : void;
    public generate<T extends GenerateOptions>(options : T) : OutputByType[OutputByType];
    public generate() : any;

    public replaceString(oldString : string, newString : string) : number; // returns idx
    public stringIndex(s : string) : number; // returns idx
    public writeSharedStrings() : void;
    public splitRef(ref : string) : ReferenceAddress;
    public joinRef(ref : ReferenceAddress) : string;
    public addSharedString(s : string) : any; // I think s is a string? Not sure what its return "idx" is though, I think it's a number? Is "idx" short for "index"?
    public substituteScalar(cell : any, string: string, placeholder: TemplatePlaceholder, substitution: any): string;
    public charToNum(str : string) : number;
    public numToChar(num : number) : string;
    public nextCol(ref : string) : string; 
    public nextRow(ref : string) : string;
    public stringify(value : Date | number | boolean | string) : string;
    public isWithin(ref : string, startRef : string, endRef : string) : boolean;
    public isRange(ref : string) : boolean;
    public splitRange(range : string) : RangeSplit;
    public joinRange(range : RangeSplit) : string;
    public extractPlaceholders(templateString : string) : TemplatePlaceholder[];
    
    // need typing properly
    protected _rebuild() : void;
    protected loadSheets(prefix : any, workbook : etree.ElementTree, workbookRels : any) : any[];
    protected loadSheet(sheet : any) : { filename : any, name : any, id : any, root : any }; // this could definitely return a "Sheet" interface/class
    protected loadSheetRels(sheetFilename : string) : { rels : any};
    protected loadDrawing(sheet : any, sheetFilename : string, rels : any) : { drawing : any};
    protected writeDrawing(drawing : any);
    protected moveAllImages(drawing : any, fromRow : number, nbRow : number);
    protected loadTables(sheet : any, sheetFilename : any) : any;
    protected writeTables(tables : any) : void;
    protected substituteHyperlinks(sheetFilename : any, substitutions : any) : void;
    protected substituteTableColumnHeaders(tables : any, substitutions : any) : void;
    protected insertCellValue(cell : any, substitution : any) : string;
    protected substituteArray(cells : any[], cell : any, substitution : any);
    protected substituteTable(row : any, newTableRows : any, cells : any[], cell : any, namedTables : any, substitution : any, key : any, placeholder : TemplatePlaceholder, drawing : etree.ElementTree) : any;
    protected substituteImage(cell : any, string : string, placeholder: TemplatePlaceholder, substitution : any, drawing : etree.ElementTree) : boolean
    protected cloneElement(element : any, deep? : any) : any;
    protected replaceChildren(parent : any, children : any) : void;
    protected getCurrentRow(row : any, rowsInserted : any) : number;
    protected getCurrentCell(cell : any, currentRow : any, cellsInserted : any) : string;
    protected updateRowSpan(row : any, cellsInserted : any) : any;
    protected pushRight(workbook : etree.ElementTree, sheet : any, currentCell : any, numCols : any) : any;
    protected pushDown(workbook : etree.ElementTree, sheets : any, tables : any, currentRow : any, numRows : any) : any;

    //Insert for insert_image
    protected getWidthCell(numCol : number, sheet : etree.ElementTree) : number;
    protected getWidthMergeCell(mergeCell : etree.ElementTree, sheet : etree.ElementTree) : Float32Array
    protected getHeightCell(numRow : number, sheet : etree.ElementTree) : number;
    protected getHeightMergeCell(mergeCell : etree.ElementTree, sheet : etree.ElementTree) : Float32Array
    protected getNbRowOfMergeCell(mergeCell : etree.ElementTree) : Int16Array;
    protected pixels(pixels : number) : Int16Array;
    protected columnWidthToEMUs(width : number) : Int16Array;
    protected rowHeightToEMUs(height : number) : Int16Array;
    protected findMaxFileId(fileNameRegex : string, idRegex : string) : Int16Array;
    protected cellInMergeCells(cell : etree.ElementTree, mergeCell : etree.ElementTree) : boolean;
    protected isUrl(str : string) : boolean;
    protected toArrayBuffer(buffer : Buffer) : ArrayBuffer;
    protected imageToBuffer(imageObj : any) : Buffer;
    protected findMaxId(element : etree.ElementTree, tag : string, attr : string, idRegex : string) : number;
}

export as namespace XlsxTemplate;
export = XlsxTemplate;
