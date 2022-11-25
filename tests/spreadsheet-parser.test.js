
const parser = require('../src/spreadsheet-parser');
const target = require('./target.json');

describe('Spreadsheet Parser', () => {
    
    // check if parser exists.
    it('parser function exists', () => {
        expect(parser).toBeDefined();
    });

    // check if parser returns a json object.
    it('parser returns a json object', () => {
        expect(parser("../data/sheet.xlsx", 'out.json')).toBeInstanceOf(Object);
    });

    
})
