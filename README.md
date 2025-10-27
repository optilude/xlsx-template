# XLSX Template

[![CI validation](https://github.com/optilude/xlsx-template/actions/workflows/npm.yaml/badge.svg)](https://github.com/optilude/xlsx-template/actions/workflows/npm.yaml)

Generate Excel (.xlsx) reports from templates with dynamic data substitution. Create beautifully formatted Excel files using Excel as your template designer, then populate them with data from your Node.js application.

## Features

- **Real Excel files** - Generate actual `.xlsx` files
- **Template-based** - Design templates in Excel with all formatting, formulas, and styles
- **Dynamic data** - Replace placeholders with values, arrays, tables, and images
- **Preserve formatting** - Cell formatting, merged cells, formulas, named tables are maintained
- **Image support** - Insert images from file paths or Base64
- **Multiple sheets** - Work with multiple sheets, copy/delete sheets dynamically
- **Lightweight & cross-platform** - Only 3 dependencies ([jszip](https://www.npmjs.com/package/@kant2002/jszip), [elementtree](https://www.npmjs.com/package/elementtree), [image-size](https://www.npmjs.com/package/image-size)). Direct XML DOM manipulation with no superfluous overhead. Works seamlessly on Windows, Linux, and macOS.

## Installation

```bash
npm install xlsx-template
```

## Quick Start

**1. Create an Excel template** (`template.xlsx`) with placeholders:

|   | A | B |
|---|---|---|
| 1 | Report Date | `${reportDate}` |
| 2 | Company | `${companyName}` |

**2. Use the template in your code:**

```javascript
const XlsxTemplate = require('xlsx-template');
const fs = require('fs');

// Load the template
fs.readFile('template.xlsx', (err, data) => {
    const template = new XlsxTemplate(data);
    
    // Define replacement values
    const values = {
        reportDate: new Date(),
        companyName: 'Acme Corporation'
    };
    
    // Perform substitution on sheet 1 (can also use sheet name as string: 'Sheet1')
    template.substitute(1, values);
    
    // Generate the output file
    const output = template.generate();
    
    // Save to disk
    fs.writeFileSync('output.xlsx', output);
});
```

## Placeholder Types

### 1. Simple Values (Scalars)

Replace a placeholder with a single value.

**Excel template:**

|   | A | B |
|---|---|---|
| 1 | Extracted on: | `${extractDate}` |

**Code:**
```javascript
const values = {
    extractDate: new Date('2024-01-15')
};
template.substitute(1, values);
```

**Result:** 

|   | A | B |
|---|---|---|
| 1 | Extracted on: | Jan-15-2024 |

**Notes:**
- Placeholders can be standalone in a cell or part of text: `"Total: ${amount}"`
- Excel cell formatting (date, number, currency) is preserved

### 2. Array Indexing

Access specific array elements directly in templates.

**Excel template:**

|   | A | B |
|---|---|---|
| 1 | First date: | `${dates[0]}` |
| 2 | Second date: | `${dates[1]}` |

**Code:**
```javascript
const values = {
    dates: [new Date('2024-01-01'), new Date('2024-02-01')]
};
template.substitute(1, values);
```

**Result:**

|   | A | B |
|---|---|---|
| 1 | First date: | Jan-01-2024 |
| 2 | Second date: | Feb-01-2024 |

### 3. Column Arrays

Expand an array horizontally across columns.

**Excel template:**

|   | A |
|---|---|
| 1 | `${dates}` |

**Code:**
```javascript
const values = {
    dates: [
        new Date('2024-01-01'),
        new Date('2024-02-01'),
        new Date('2024-03-01')
    ]
};
template.substitute(1, values);
```

**Result:**

|   | A | B | C |
|---|---|---|---|
| 1 | Jan-01-2024 | Feb-01-2024 | Mar-01-2024 |

**Notes:**
- The placeholder must be the **only content** in its cell

### 4. Table Rows

Generate multiple rows from an array of objects.

**Excel template:**

|   | A | B | C |
|---|---|---|---|
| 1 | Name | Age | Department |
| 2 | `${table:team.name}` | `${table:team.age}` | `${table:team.dept}` |

**Code:**
```javascript
const values = {
    team: [
        { name: 'Alice Johnson', age: 28, dept: 'Engineering' },
        { name: 'Bob Smith', age: 34, dept: 'Marketing' },
        { name: 'Carol White', age: 25, dept: 'Sales' }
    ]
};
template.substitute(1, values);
```

**Result:**

|   | A | B | C |
|---|---|---|---|
| 1 | Name | Age | Department |
| 2 | Alice Johnson | 28 | Engineering |
| 3 | Bob Smith | 34 | Marketing |
| 4 | Carol White | 25 | Sales |

**Notes:**
- Syntax: `${table:arrayName.propertyName}`
- Each object in the array creates a new row
- If a property is an array, it expands horizontally

### 5. Images

Insert images into cells.

**Excel template:**

|   | A | B |
|---|---|---|
| 1 | Logo: | `${image:companyLogo}` |

**Code:**
```javascript
const values = {
    companyLogo: '/path/to/logo.png'  // or Base64, Buffer
};
template.substitute(1, values);
```

**Result:**

|   | A | B |
|---|---|---|
| 1 | Logo: | 🖼️ |

**Supported image formats:**
- File path (absolute or relative): `'/path/to/image.png'`
- Base64 string: `'data:image/png;base64,iVBORw0KG...'`
- Buffer: `fs.readFileSync('image.png')`
- **URL: Not supported** - This library is synchronous and cannot fetch remote images. Fetch the image in your own code first, then pass it as one of the supported formats above.

**Image options:**
```javascript
const template = new XlsxTemplate(data, {
    imageRootPath: '/absolute/path/to/images',  // Base path for relative image paths
    imageRatio: 75                               // Scale images to 75% (only for non-merged cells)
});
```

**Table images:**

|   | A | B |
|---|---|---|
| 1 | Product | Photo |
| 2 | `${table:products.name}` | `${table:products.photo:image}` |

```javascript
const values = {
    products: [
        { name: 'Product 1', photo: 'product1.jpg' },
        { name: 'Product 2', photo: 'product2.jpg' }
    ]
};
```

**Result:**

|   | A | B |
|---|---|---|
| 1 | Product | Photo |
| 2 | Product 1 | 🖼️ |
| 3 | Product 2 | 🖼️ |

#### Images in Merged Cells

Images automatically fit the size of merged cells.

**Excel template with merged cells B1:C2:**

|   | A | B-C (merged) |
|---|---|---|
| 1-2 (merged) | Large Image: | `${image:banner}` |

```javascript
const values = {
    banner: 'banner-image.png'
};
template.substitute(1, values);
```

**Result:** The image will be automatically resized to fit the merged cell area (B1:C2).

### 6. Images in Cell

Insert images that automatically fit cell size (requires Excel 2308+).

> **⚠️ Warning:** This feature requires Excel version 2308 or later. Excel 2302 and earlier do not support this.

This is the equivalent of Excel's "Place in Cell" feature (right-click on image → "Place in Cell").

**Excel template:**

|   | A |
|---|---|
| 1 | `${imageincell:profilePicture}` |

**Code:**
```javascript
const values = {
    profilePicture: 'avatar.jpg'
};
template.substitute(1, values);
```

**Result:**

|   | A |
|---|---|
| 1 | 🖼️ |

**Features:**
- Images automatically match cell/merged cell size
- Respects cell alignment and formatting
- Perfect for profile pictures, thumbnails, etc.

**Table usage:**

|   | A | B |
|---|---|---|
| 1 | Employee | Avatar |
| 2 | `${table:employees.name}` | `${table:employees.avatar:imageincell}` |

```javascript
const values = {
    employees: [
        { name: 'Alice', avatar: 'alice.jpg' },
        { name: 'Bob', avatar: 'bob.jpg' }
    ]
};
```

**Result:**

|   | A | B |
|---|---|---|
| 1 | Employee | Avatar |
| 2 | Alice | 🖼️ |
| 3 | Bob | 🖼️ |

## API Reference

### `new XlsxTemplate(data, options)`

Create a new template instance.

**Parameters:**
- `data` (Buffer|String) - The `.xlsx` file content (binary data)
- `options` (Object) - Optional configuration:
  - `imageRootPath` (String) - Root directory for relative image paths
  - `imageRatio` (Number) - Image scaling percentage (default: 100)
  - `moveImages` (Boolean) - Move images when inserting table rows (default: false)
  - `moveSameLineImages` (Boolean) - Move images on the same line as inserted rows (default: false)
  - `subsituteAllTableRow` (Boolean) - Apply substitutions to all cells in table rows (default: false)
  - `pushDownPageBreakOnTableSubstitution` (Boolean) - Adjust page breaks when tables grow (default: false)

**Example:**
```javascript
const template = new XlsxTemplate(data, {
    imageRootPath: __dirname + '/images',
    imageRatio: 80,
    moveImages: true
});
```

### `template.substitute(sheetNumber, values)`

Replace placeholders with values on a specific sheet.

**Parameters:**
- `sheetNumber` (Number|String) - Sheet index (1-based) or sheet name
- `values` (Object) - Key-value pairs for placeholder substitution

**Example:**
```javascript
template.substitute(1, { name: 'John', age: 30 });
template.substitute('Sales Report', { quarter: 'Q1', revenue: 50000 });
```

### `template.substituteAll(values)`

Replace placeholders on **all sheets** with the same values.

**Parameters:**
- `values` (Object) - Key-value pairs for placeholder substitution

**Example:**
```javascript
template.substituteAll({ 
    companyName: 'Acme Corp',
    reportDate: new Date() 
});
```

### `template.generate(options)`

Generate the final Excel file.

**Parameters:**
- `options` (Object) - JSZip generation options:
  - `type` (String) - Output format:
    - `'nodebuffer'` - Node.js Buffer (recommended for file I/O)
    - `'base64'` - Base64 string
    - `'uint8array'` - Uint8Array
    - `'arraybuffer'` - ArrayBuffer
    - `'blob'` - Blob (browser only)

**Returns:** The generated file in the specified format

**Example:**
```javascript
const buffer = template.generate({ type: 'nodebuffer' });
fs.writeFileSync('output.xlsx', buffer);

// Or for web download
const base64 = template.generate({ type: 'base64' });
const downloadLink = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64;
```

### `template.copySheet(sheetName, newSheetName)`

Copy an existing sheet to a new sheet in the same workbook.

**Parameters:**
- `sheetName` (String|Number) - Source sheet name or index (1-based)
- `newSheetName` (String) - Name for the new sheet (optional, defaults to "SheetN")

**Returns:** `this` (for chaining)

> **Note:** The optional `binary` parameter (third parameter) is deprecated and should not be used. It will be removed in a future version. Always use the default behavior which preserves UTF-8 encoding correctly.

**Example:**
```javascript
template.copySheet('Template', 'January Report');
template.copySheet(1, 'Q1 Data');

// Chain operations
template.copySheet('Template', 'Report1')
        .substitute('Report1', { month: 'January' })
        .copySheet('Template', 'Report2')
        .substitute('Report2', { month: 'February' });
```

**Notes:**
- Copies all content/Relation: data, formatting, formulas, comments, images, print settings
- Merged cells and named ranges are preserved
- Comments (including threaded comments) are copied with unique IDs

### `template.deleteSheet(sheetName)`

Delete a sheet from the workbook.

**Parameters:**
- `sheetName` (String|Number) - Sheet name or index (1-based) to delete

**Returns:** `this` (for chaining)

**Example:**
```javascript
template.deleteSheet('Sheet2');
template.deleteSheet(3);
```

## Complete Example

```javascript
const XlsxTemplate = require('xlsx-template');
const fs = require('fs');

// Load template
const templateData = fs.readFileSync('sales-template.xlsx');
const template = new XlsxTemplate(templateData, {
    imageRootPath: __dirname + '/assets',
    moveImages: true
});

// Prepare data
const data = {
    reportDate: new Date(),
    companyName: 'Acme Corporation',
    region: 'North America',
    
    // Table data
    salesData: [
        { product: 'Widget A', qty: 150, price: 25.50, photo: 'widget-a.jpg' },
        { product: 'Widget B', qty: 200, price: 30.00, photo: 'widget-b.jpg' },
        { product: 'Gadget X', qty: 80, price: 55.75, photo: 'gadget-x.jpg' }
    ],
    
    // Chart data (arrays)
    months: ['Jan', 'Feb', 'Mar', 'Apr'],
    revenues: [45000, 52000, 48000, 61000],
    
    // Company logo
    logo: 'company-logo.png',
    
    // Formula
    totalFormula: '=SUM(D2:D100)'
};

// Apply substitutions
template.substitute(1, data);

// Generate output
const output = template.generate({ type: 'nodebuffer' });
fs.writeFileSync('sales-report.xlsx', output);

console.log('Report generated successfully!');
```

## Important Notes & Limitations

### File Format
- ✅ Only `.xlsx` format is supported
- ❌ `.xls`, `.xlsb`, `.xlsm` formats are **not supported**


### Merged Cells & Named Ranges
- Merged cells are automatically adjusted when rows/columns are inserted
- Named ranges and tables are moved correctly

### Table Substitutions
When using `${table:...}` placeholders:
- Rows below the table are pushed down automatically
- Columns to the right are shifted if arrays expand horizontally
- Use **Excel Named Tables** for best formula compatibility
- Page breaks can be automatically adjusted with the `pushDownPageBreakOnTableSubstitution` option

### Images
- Images in merged cells automatically fit the merged area
- Standard cells: use `imageRatio` option to scale
- The `moveImages` option shifts images when rows are inserted (move the anchor)

### Performance
- Large templates with many placeholders may take time to process
- Consider splitting very large reports across multiple sheets
- Image processing (especially Base64) can be memory-intensive

---

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

**Pull Request Requirements:**
- All PRs **must include unit tests** for new features or bug fixes
- Ensure all existing tests pass (`npm test`)
- Follow the existing code style and conventions

## License

MIT License - see [LICENSE](LICENSE) file for details

## Authors

- **Martin Aspeli** - Original author
- **Andrii Kurdiumov** (@kant2002) - Maintainer
- And many [contributors](https://github.com/optilude/xlsx-template/graphs/contributors)

---

## Changelog History

### Version 1.4.5
* Fixed UTF-8 encoding in `copySheet()` - sheet content now properly preserved in binary mode
* `copySheet()` now properly copies comments (including threaded comments)
* **Note**: Do not use `binary=false` parameter as it corrupts UTF-8 characters. May be deprecated in future versions.
* Added comprehensive sheet copying tests

### Version 1.4.4
* Move hyperlinks references on added rows and columns. (#184). Thanks @IagoSRL
* Fix line issue under table with merged cell. (#188). Thanks @muyoungko

### Version 1.4.3
* Fix potential issue when template has lot of images.
* Update image-size to 1.0.2

### Version 1.4.2
* Move to @kant2002/jszip which fix https://github.com/advisories/GHSA-36fh-84j7-cv5h
* Fix previously broken release.

### Version 1.4.1
* Move to @kant2002/jszip which fix https://github.com/advisories/GHSA-36fh-84j7-cv5h
* Also broke everything. **DONT USE THIS VERSION**

### Version 1.4.0
* substituteAll: Interpolate values for all the sheets using the given substitutions (#173) Thanks @jonathankeebler
* int` and `float` don't exist in Typescript, both are of type `number`. This fixes it. (#169) Thanks @EHadoux
* Insert images. (#126). Thanks @jdugh
* Add customXml in the order list for rebuild. (#154). Thanks @jdugh
* Adding 2 options affect table substitution : subsituteAllTableRow and pushDownPageBreakOnTableSubstitution. (#124). Thanks @jdugh

### Version 1.3.2
* Fix import statement for jszip

### Version 1.3.1
* Added the imageRatio parameter like a percent ratio when insert images. (#121)
* Add new substitution for images. (#110)
* Fixing Defined Range Name with Sheet Name. (#150)
* Add binary option for copySheet : for header/footer in UTF-8 (#130)

### Version 1.3.0
* Added support for optional moving of the images together with table. (#109)

### Version 1.2.0
* Specify license field in addition to licenses field in the package.json (#102)

### Version 1.1.0
* Added TypeScript definitions. #101
* NodeJS 12, 14 support

### Version 1.0.0
Nothing to see here. Just I'm being brave and make version 1.0.0

### Version 0.5.0
* Placeholder in hyperlinks. #87
* NodeJS 10 support

### Version 0.4.0
* Fix wrongly replacing text in shared strings #81

### Version 0.2.0

* Add ability copy and delete sheets.

### Version 0.0.7

* Fix bug with calculating <dimensions /> when adding columns

### Version 0.0.6

* You can now pass `options` to `generate()`, which are passed to JSZip
* Fix setting of sheet <dimensions /> when growing the sheet
* Fix corruption of sheet when writing dates
* Fix corruption of sheet when calculating calcChain

### Version 0.0.5

* Mysterious

### Version 0.0.4

Merged pending pull requests

* Deletion of the sheets.

### Version 0.0.3

Merged a number of overdue pull requests, including:

* Windows support
* Support for table footers
* Documentation improvements

### Version 0.0.2

* Fix a potential issue with the typing of string indices that could cause the
  first string to not render correctly if it contained a substitution.

### Version 0.0.1

* Initial release
