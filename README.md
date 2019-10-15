# Mini Sheets
Minified and simplified Google Sheets with [Promises](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)

## Example Usage
```javascript
const {Drive, Spreadsheets} = require('minisheets');
const minisheets = new Spreadsheets(CLIENT_ID, TOKEN);
minisheets.createSpreadsheet(spreadsheetTitle, {sheet1: [[1]]}).then(worksheet=>console.log(worksheet)).catch(e=>console.log(e));
```

## Table of Contents
+ [gAPI](#gapi)
    + [constructor(clientId, token)](#constructorclientid-token)
    + [Spreadsheets](#spreadsheets)
        + [Worksheet Object](#worksheet-object)
            + [Grid Data Format](#grid-data-format)
            + [Metadata Format](#metadata-format)
        + [createSpreadsheet(title, gridData, metadata)](#createspreadsheet)
        + [getSpreadsheet(spreadsheetId, options)](#getspreadsheetspreadsheetid-options)
        + [setSpreadsheet(spreadsheetId, gridData, metadata)](#setspreadsheetspreadsheetid-griddata-metadata)
    + [Drive](#drive)
        + [getFile(fileId)](#getfilefileid)
        + [setFile(fileId, properties)](#setfilefileid-properties)
        + [deleteFile(fileId)](#deletefilefildid)

## gAPI
---
### constructor(clientId, token)
> Initialize gAPI (parent class of Drive & Spreadsheets)

Arguments:
+ `clientId`<**String**>: string of client id
+ `token`<**Object**>: authentication token object


## Spreadsheets
---
```javascript
const {Spreadsheets} = require('minisheets');
let worksheets = new Spreadsheets(clientId, token);
```

---
### Worksheet Object
> Object returned from using Spreadsheets methods

```javascript
{
    id: spreadsheetId,
    title: spreadsheetTitle,
    sheets: {
        Sheet1: [[String, Number],
                 [...]],
        Sheet2: ...,
        ...
    },
    metadata: {
        Sheet1: {
            key1: String,
            key2: Number,
            ...
        },
        ...
    }
}
```

#### Grid Data Format
```javascript
{
    sheetTitle: [[String, Number],
                 [..., ...]],
    ...
}
```

#### Metadata Format
```javascript
{
    sheetTitle: {
        key: value,
        ...
    },
    ...
}
```

---
### createSpreadsheet(title, gridData, metadata)
> Creates a new Google Sheets spreadsheet

Arguments:
+ `title` <**String**>: ID of Google Sheets spreadsheet
+ `gridData` <[**Grid Data Format**](#grid-data-format)>: Object of sheets
+ `metadata` <[**Metadata Format**](#metadata-format)>: Object of metadata

Returns:
+ <[**Worksheet Object**](#worksheet-object)>

Usage:
```javascript
worksheets.createSpreadsheet(newTitle, {sheet1: [['Hello', 'World', 1]]}, {sheet1: {key1: 'Hello World'}}).then(console.log);
/*{
    title: newTitle,
    id: generatedId,
    sheets: {
        sheet1: [['Hello', 'World', 1]]
    },
    metadata: {
        sheet1: {
            key1: 'Hello World'
        }
    }
}*/
```

---
### getSpreadsheet(spreadsheetId, \_options)
> Gets the data of a Google Sheets spreadsheet

Arguments:
+ `spreadsheetId` <**String**>: ID of Google Sheets spreadsheet
+ (OPTIONAL) <**Object**>: options Object
    + `include` <**String**\|**Array**>: only include sheets with specified titles

Returns:
+ <[**Worksheet Object**](#worksheet-object)>

Usage:
```javascript
worksheets.getSpreadsheet(spreadsheetId, {include: ['sheet1']}).then(console.log);
/*{
    title: spreadsheetTitle,
    id: spreadsheetId,
    sheets: {
        sheet1: [['Hello', 'World', 1]]
    },
    metadata: {
        sheet1: {
            key1: 'Hello World'
        }
    }
}*/
```

---
### setSpreadsheet(spreadsheetId, gridData, metadata)
> Changes a Google Sheets spreadsheet

Arguments:
+ `spreadsheetId` <**String**>: ID of Google Sheets spreadsheet
+ `gridData` <[**Grid Data Format**](#grid-data-format)>: Object of sheets
+ `metadata` <[**Metadata Format**](#metadata-format)>: Object of metadata

Returns:
+ <[**Worksheet Object**](#worksheet-object)>

Usage:
```javascript
worksheets.setSpreadsheet(spreadsheetId, {sheet1: [['Hello', 'World', 2]]}, {sheet1: {key1: 'Hello There'}}).then(console.log);
/*{
    title: spreadsheetTitle,
    id: spreadsheetId,
    sheets: {
        sheet1: [['Hello', 'World', 2]]
    },
    metadata: {
        sheet1: {
            key1: 'Hello There'
        }
    }
}*/
```

## Drive
---
```javascript
const {Drive} = require('minisheets');
let files = new Drive(clientId, token);
```

---
### getFile(fileId)
> Fetches the properties of a file within Google Drive

Arguments:
+ `fileId` <**String**>: ID of file

Return:
+ <**Object**>: [metadata](https://developers.google.com/drive/api/v3/reference/files)
    + `null` if file does not exist

Usage:
```javascript
files.getFile(fileId).then(console.log);
/*{
    "kind": "drive#file",
    "id": String,
    "name": String,
    "mimeType": String
}*/
```

---
### setFile(fileId, properties)
> Alters properties of a file within Google Drive

Arguments:
+ `fileId` <**String**>: ID of file

Return:
+ <**Object**>: [metadata](https://developers.google.com/drive/api/v3/reference/files)
    + `null` if file does not exist

Usage:
```javascript
files.deleteFile(fileId).then(console.log); // true
```

---
### deleteFile(fileId)
> Deletes a file within Google Drive

Arguments:
+ `fileId` <**String**>: ID of file

Return:
+ <**Boolean**>: `false` if the file does not exist 

Usage:
```javascript
files.setFile(fileId, {name: newName}).then(console.log);
/*{
    "kind": "drive#file",
    "id": String,
    "name": String,
    "mimeType": String
}*/
```