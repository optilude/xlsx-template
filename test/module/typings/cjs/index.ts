import XlsxTemplate = require("xlsx-template");

const generateType: XlsxTemplate.GenerateType = "base64";

const placeholder: XlsxTemplate.TemplatePlaceholder = {
  type: "normal",
  full: true,
  name: "foo",
  key: "bar",
};

const namedTable: XlsxTemplate.NamedTable = {
  filename: "table1.xml",
  root: {} as any,
};

const outputByType: XlsxTemplate.OutputByType = {
  base64: "base64string",
  uint8array: new Uint8Array(),
  arraybuffer: new ArrayBuffer(0),
  blob: new Blob(),
  nodebuffer: Buffer.from([]),
};

const rangeSplit: XlsxTemplate.RangeSplit = {
  start: "A1",
  end: "C3",
};

const referenceAddress: XlsxTemplate.ReferenceAddress = {
  col: "A",
  row: 1,
};