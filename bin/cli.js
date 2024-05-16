#!/usr/bin/env node

const { program } = require('commander');
const convertToMarkdown = require('../lib/convert');
const fs = require('fs');

program
  .version('1.0.0')
  .description('Convert XLSX to Markdown table')
  .requiredOption('-f, --file <path>', 'XLSX file path')
  .requiredOption('-s, --sheet <name>', 'Sheet name')
  .requiredOption('-c, --columns <columns>', 'Columns to include (comma separated)', val => val.split(','))
  .option('-o, --output <path>', 'Output file path', 'output.md');

program.parse(process.argv);

const options = program.opts();

try {
  const markdownTable = convertToMarkdown(options.file, options.sheet, options.columns);
  fs.writeFileSync(options.output, markdownTable, 'utf8');
  console.log(`Markdown table written to ${options.output}`);
} catch (error) {
  console.error(`Error: ${error.message}`);
}
