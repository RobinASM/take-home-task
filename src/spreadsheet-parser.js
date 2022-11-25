const Excel = require('exceljs');
const fs = require('fs');

var structure = {
    "name": "sheet2",
    "freeze": "A1",
    "styles": [],
    "merges": [],
    "rows": {},
    "cols": {},
    "validations": []
};

// get the current style inside the structure and returns the index.
const getStyleIndex = (cell) => {
    const bgColor = cell?.fill?.fgColor?.argb || null;

    let styleIndex = null;

    // if the cell is a number and has a background color, add it to the structure as an object.
    if (cell?.text?.match(/^[0-9]+$/) && cell?.text !== "0") {
        styleIndex = structure.styles.findIndex(style => style.bgcolor === `#${bgColor.toLocaleLowerCase()}` && style.format === "numberNoDecimal");

        if (styleIndex === -1) {
            structure.styles.push({
                "format": "numberNoDecimal",
                "bgcolor": `#${bgColor.toLocaleLowerCase()}`,
            });
            styleIndex = structure.styles.length - 1;
        }
    }

    if ((cell?.value === null || cell?.text === "0") && bgColor !== null) {
        styleIndex = structure.styles.findIndex(style => {
            const length = Object.keys(style).length;
            if (style.bgcolor === `#${bgColor.toLocaleLowerCase()}` && length === 1) {
                return style;
            }
        });

        if (styleIndex === -1) {
            structure.styles.push({
                "bgcolor": `#${bgColor.toLocaleLowerCase()}`,
            });
            styleIndex = structure.styles.length - 1;
        }
    }

    // if the cell is a percentage and has a background color, add it to the structure as an object.
    if (cell.numFmt === '0.00%') {
        styleIndex = structure.styles.findIndex(style => style.bgcolor === `#${bgColor.toLocaleLowerCase()}` && style.format === "percentNoDecimal");

        if (styleIndex === -1) {
            structure.styles.push({
                "format": "percentNoDecimal",
                "bgcolor": `#${bgColor.toLocaleLowerCase()}`,
            });
            styleIndex = structure.styles.length - 1;

        }
    }

    // if the cell is bold, add it to the structure as an object.
    if (cell.font?.bold) {
        styleIndex = structure.styles.findIndex(style => style?.font?.bold === true);
        
        if (styleIndex === -1) {
            structure.styles.push({
                "font": {
                    "bold": true,
                }
            });
            styleIndex = structure.styles.length - 1;
        }
    }

    return styleIndex;
}

const parser = async (source, output) => {

    // Read the source file.
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(source);

    // Get the first worksheet.
    const worksheet = workbook.getWorksheet(1);

    // Iterate over all rows (including empty rows) in a worksheet
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {

            structure.rows[rowNumber - 1] = structure.rows[rowNumber - 1] || { cells: {} };

            // if the cell has a formula, add it to the structure as an object.
            if (cell.formula) {
                structure.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": `=${cell.formula.toLocaleLowerCase()}`
                }

                const index = getStyleIndex(cell);
                if (index !== null) {
                    structure.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // if the cell is a number, add it to the structure as an object.
            if (typeof cell.value === 'number') {
                structure.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": String(cell.value)
                }

                const index = getStyleIndex(cell);
                if (index !== null) {
                    structure.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // if the cell is a string, add it to the structure as an object.
            if (typeof cell.value === 'string') {
                structure.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": cell.value
                }

                const index = getStyleIndex(cell);
                if (index !== null) {
                    structure.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // if the cell is null, add it to the structure as an object.
            if (cell.value === null) {
                structure.rows[rowNumber - 1].cells[colNumber - 1] = {
                    "text": ""
                }

                const index = getStyleIndex(cell);
                if (index !== null) {
                    structure.rows[rowNumber - 1].cells[colNumber - 1].style = index;
                }
            }

            // fill the cols object with the column width of every column.
            const colWidth = worksheet.getColumn(colNumber)?.width;
            if (colWidth) {
                structure.cols[colNumber - 1] = {
                    "width": colWidth
                }
            }

        });
    })

    fs.writeFileSync(output, JSON.stringify(structure, null, 2));
    return structure;
}

// parser("./data/sheet.xlsx", "./data/out.json").then((data) => {
//     console.log("Parser data: ", data);
// });

module.exports = parser