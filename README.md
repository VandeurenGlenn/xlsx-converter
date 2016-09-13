# xlsx-converter
> convert xlsx to json, csv, etc (more will follow)

## Install
```sh
$ npm install --save xlsx-converter
```
## Usage

```js
const {convert} = require('xlsx-converter');

convert('file.xlsx').then(result => {
  // do something with the result
  console.log(result);
});
```

## Options

### To
Convert to

> ** type: ** *String* the file extension to convert to
>
> ** default: ** json
>
** options **
* json
* csv

```js
convert('file.xlsx', { to: 'csv' });
```

### Sheets
An array containing the sheetPages to convert (as an number or by name).

> ** type: ** *Array* sheetPages to convert
>
> ** default: ** converts all
>
** example **
['1', '2']
>
** note: ** * ['1', '2'] and ['Sheet1', 'Sheet2'] have the same result *

```js
convert('file.xlsx', { to: 'csv' });
```

### Write
Write the result
> ** type: ** *String* destination

```js
convert('file.xlsx', { write: 'out.json' });
```

## Status
Work in progress
> *Documented options are available except converting to csv*

#### TODO
*soon*

## Under the hood
- [xlsx](https://www.npmjs.com/package/xlsx)

## Contributing
- suggestions welcome! [create suggestion](https://github.com/vandeurenglenn/xlsx-converter/issues/new)

## Issues
When finding an issue consider creating a [gh-issue](https://github.com/vandeurenglenn/xlsx-converter/issues/new) for easy follow up.

You can find open issues [here](https://github.com/vandeurenglenn/xlsx-converter/issues).
<br>
<br>
<br>
*That wraps it up*
