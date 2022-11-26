
const { parser, getStyleIndex } = require('../src/spreadsheet-parser');
const target = require('./target.json');

describe('Spreadsheet Parser', () => {

    it('parser function exists', () => {
        expect(parser).toBeDefined();
    });

    test('parser returns a json object', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");
        expect(typeof res).toBe('object');
    });

    test('parser returns the correct keys', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");
        expect(Object.keys(res)).toEqual(Object.keys(target));
    });

    test('parser returns the correct styles length', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");
        expect(res.styles.length).toEqual(target.styles.length);
    })

    test('parser returns the correct styles', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");

        res.styles.forEach((style, index) => {
            target.styles.forEach((targetStyle, targetIndex) => {
                if (style === targetStyle) {
                    expect(style).toEqual(targetStyle);
                }
            })
        })
    })

    test('parser returns the correct rows length', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");
        expect(res.rows.length).toEqual(target.rows.length);
    })

    test('parser returns the correct columns length', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");
        expect(res.cols.length).toEqual(target.cols.length);
    })

    test('parser returns the correct cells keys', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");

        Object.values(res.rows).forEach((row, indexRow) => {
            const cells = Object.values(row.cells);
            const targetKeys = Object.keys(target.rows[indexRow].cells);

            targetKeys.forEach((key, indexKey) => {
                expect(Object.keys(cells[key])).toEqual(Object.keys(target.rows[indexRow].cells[key]));
            })
        })
    })

    test('parser returns the correct cells text', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");

        Object.values(res.rows).forEach((row, indexRow) => {
            const cells = Object.values(row.cells);
            const targetKeys = Object.keys(target.rows[indexRow].cells);
            
            targetKeys.forEach((key, indexKey) => {
                expect(cells[indexKey].text).toEqual(target.rows[indexRow].cells[key].text);
            })
        })
    })

    test('parser returns the correct cells style', async () => {
        const res = await parser("./data/sheet.xlsx", "./data/out.json");

        Object.values(res.rows).forEach((row, indexRow) => {
            const cells = Object.values(row.cells);
            const targetKeys = Object.keys(target.rows[indexRow].cells);
            
            targetKeys.forEach((key, indexKey) => {
                const targetStyleId = target.rows[indexRow].cells[key].style
                const targetStyle = target.styles[targetStyleId];

                const sourceStyleId = cells[key].style;
                const sourceStyle = res.styles[sourceStyleId];

                expect(sourceStyle).toEqual(targetStyle);
            })
        })
    })
    
})
