const { Command } = require('commander');
const program = new Command();
const parser = require('./src/spreadsheet-parser');

program
    .name('spreadsheet-parser')
    .description('Spreadsheet Parser of Excel files.')
    .version('0.8.0')
    .option('-s, --source <string>', 'Specify the relative path of the source file. (.xlsx)')
    .option('-o, --output <string>', 'Specify the relative path of the output file. (.json)')

program.parse();

const options = program.opts();

if (!options.source) {
    console.log('Please specify the source file.');
} 

if (!options.output) {
    console.log('Please specify the output path.');
}

if (options.source && options.output) {
    parser(options.source, options.output);
}

if (Object.keys(options).length === 0) {
    program.help();
}