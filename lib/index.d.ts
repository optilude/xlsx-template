export default class Workbook
{
    constructor(data? : Buffer);
    public deleteSheet(sheetName : string) : this;
    public copySheet(sheetName : string, copyName : string) : this;
    public loadTemplate(data : Buffer) : void;
    public substitute(sheetName : string | number, substitutions : Object) : void;
    public generate(options? : GenerateOptions) : Buffer | Uint8Array | Blob | string | ArrayBuffer;

    // need typing properly
    protected _rebuild() : void;
    protected writeSharedStrings() : void;
    protected addSharedString(s : any) : any; // I think s is a string? Not sure what its return "idx" is though, I think it's a number? Is "idx" short for "index"?
    protected stringIndex(s : any) : any; // returns idx
    protected replaceString(oldString : string, newString : string) : any; // returns idx
    protected loadSheets(prefix : any, workbook : any, workbookRels : any) : any[];
    protected loadSheet(sheet : any) : { filename : any, name : any, id : any, root : any }; // this could definitely return a "Sheet" interface/class
    protected loadTables(sheet : any, sheetFilename : any) : any;
    protected writeTables(tables : any) : void;
    protected substituteHyperlinks(sheetFilename : any, substitutions : any) : void;
    protected substituteTableColumnHeaders(tables : any, substitutions : any) : void;
    protected extractPlaceholders(string : any) : any[];
    protected splitRef(ref : any) : { table : any, colAbsolute : any, col : any, rowAbsolute : any, row : any }
    protected joinRef(ref : any) : string;
    protected nextCol(ref : any) : string; 
    protected nextRow(ref : any) : string
    protected charToNum(str : string) : number;
    protected numToChar(num : number) : string;
    protected isRange(ref : any) : boolean;
    protected isWithin(ref : any, startRef : any, endRef : any) : boolean;
    protected stringify(value : any) : string;
    protected insertCellValue(cell : any, substitution : any) : string;
    protected substituteScalar(cell : any);
    protected substituteArray(cells : any[], cell : any, substitution : any);
    protected substituteTable(row : any, newTableRows : any, cells : any[], cell : any, namedTables : any, substitution : any, key : any) : any;
    protected cloneElement(element : any, deep : any) : any;
    protected replaceChildren(parent : any, children : any) : void;
    protected getCurrentRow(row : any, rowsInserted : any) : number;
    protected getCurrentCell(cell : any, currentRow : any, cellsInserted : any) : string;
    protected updateRowSpan(row : any, cellsInserted : any) : any;
    protected splitRange(range : string) : any;
    protected joinRange(range : any) : string
    protected pushRight(workbook : any, sheet : any, currentCell : any, numCols : any) : any;
    protected pushDown(workbook : any, sheets : any, tables : any, currentRow : any, numRows : any) : any;
}

export interface GenerateOptions
{
    type : "uint8array" | "arraybuffer" | "blob" | "nodebuffer" | "base64";
}