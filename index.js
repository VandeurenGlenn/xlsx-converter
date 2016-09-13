'use strict';
// const xlsxTojson = require('./libs/xlsx-to-json');
const {readFile} = require('xlsx');
const {writeFile} = require('fs');

class XlsxConverter {
  constructor() {
    this.options = {};
  }
  
  /**
   * Array containing what to return ...
   *
   * ### Options are
   * - 'all'
   * - 'v': raw value
   * - 'w': formatted text
   * - 't': type
   * - 'f': formula
   * - 'r': rich text encoding
   * - 'h': HTML
   * - 'c': comments
   * - 'z': format
   * - 'l': hyperlink
   * - 's': style/theme
   *
   * checkout the [xlsx](https://www.npmjs.com/package/xlsx#cell-object) for more info
   */

  get byOptions() {
    return this.options.by || undefined;
  }

  /**
   * path for writing the file to
   */
  get write() {
    return this.options.write || undefined;
  }

  /**
   * The extension to convert to
   *
   * @default {string} json
   */
  get to() {
    return this.options.to || 'json';
  }

  /**
   * The extension to convert to
   *
   * @default {string} json
   */
  get sheets() {
    return this.options.sheets || undefined;
  }

  /**
   * @method reads a file
   * @param {string} path to the file
   * @param {array} sheets an array of sheetnumbers to return
   */
   // TODO: read fails silently when file is not found
  read(path, sheets) {
    return new Promise((resolve, reject) => {
      var file = readFile(path);
      var result = [];
      try {
        sheets = sheets || file.SheetNames;
        for (let sheet of sheets) {
          // be sure sheet is an number
          let _sheet = file.Sheets[sheet];
          result.push(_sheet);
        }
      } catch (err) {
        reject(err);
      }
      resolve(result);
    });
  }

  convert(path, opts) {
    this.options = opts || {};
    return new Promise((resolve, reject) => {
      this.read(path, this.sheets).then(result => {
        switch (this.to) {
          case 'json':
            result = this.toJSON(result);
            break;
          case 'csv':

            break;
        }
        if (this.write) {
          var JSONString = JSON.stringify(result, null, 2);
          writeFile(`${this.write}.${this.to}`, JSONString, (err, succes) => {
            if (err) {
              reject(err);
            }
            resolve(result);
          });
        } else {
          resolve(result);
        }
      }).catch(err => {
        reject(err);
      });
    });
  }

  toJSON(input) {
    // TODO: return ref (range) or only when defined as byOption ?
    let output = {};
    for (let obj of input) {
      for (let i of Object.keys(obj)) {
        if (i !== '!ref') {
          var key = i.replace(/[A-Z]/g, '');
          if (!output[key]) {
            output[key] = [];
          }
          output[key].push(this.constructByOptions(obj[i], this.byOptions));
        }
      }
    }
    return output;
  }

  constructByOptions(input, options) {
    if (options) {
      var result = [];
      for (let option of options) {
        if (option === 'all') {
          result.push(input);
        } else {
          result.push(input[option]);
        }
      }
    }
    return result || input['v'];
  }
}

module.exports = new XlsxConverter();
