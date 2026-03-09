import XlsxTemplate, {
  GenerateOptions,
  GenerateType,
  TemplatePlaceholder,
  NamedTable,
  OutputByType,
  RangeSplit,
  ReferenceAddress,
} from "xlsx-template";

const generateType1: XlsxTemplate.GenerateType = "base64";
const generateType2: GenerateType = "base64";

const generateOptions1: XlsxTemplate.GenerateOptions = {
  type: "base64",
};
const generateOptions2: GenerateOptions = {
  type: "base64",
};

const placeholder1: XlsxTemplate.TemplatePlaceholder = {
  type: "normal",
  full: true,
  name: "foo",
  key: "bar",
};
const placeholder2: TemplatePlaceholder = {
  type: "normal",
  full: true,
  name: "foo",
  key: "bar",
};

const namedTable1: XlsxTemplate.NamedTable = {
  filename: "table1.xml",
  root: {} as any,
};
const namedTable2: NamedTable = {
  filename: "table1.xml",
  root: {} as any,
};

const outputByType1: XlsxTemplate.OutputByType = {
  base64: "base64string",
  uint8array: new Uint8Array(),
  arraybuffer: new ArrayBuffer(0),
  blob: new Blob(),
  nodebuffer: Buffer.from([]),
};
const outputByType2: OutputByType = {
  base64: "base64string",
  uint8array: new Uint8Array(),
  arraybuffer: new ArrayBuffer(0),
  blob: new Blob(),
  nodebuffer: Buffer.from([]),
};

const rangeSplit1: XlsxTemplate.RangeSplit = {
  start: "A1",
  end: "C3",
};
const rangeSplit2: RangeSplit = {
  start: "A1",
  end: "C3",
};

const referenceAddress1: XlsxTemplate.ReferenceAddress = {
  col: "A",
  row: 1,
};
const referenceAddress2: ReferenceAddress = {
  col: "A",
  row: 1,
};
