'use strict';

var assert = require('assert');
var xlsxConverter = require('./../index');
var testFile = 'xlsx-converter-test.xlsx';
var outFile = 'out.json';

describe('xlsx-converter', function () {
  it('Should convert the testFile to json!', function () {
    xlsxConverter.convert(testFile).then(result => {
      try {
        JSON.stringify(result);
      } catch (err) {
        console.log(err);
      }
      done();
    });
  });

  it('Should convert the test file and write to out.json!', function () {
    xlsxConverter.convert(testFile, {write: outFile}).then(result => {
      try {
        JSON.stringify(result);
      } catch (err) {
        console.log(err);
      }
      done();
    });
  });
});
