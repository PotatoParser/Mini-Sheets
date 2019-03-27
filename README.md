# SimpleSheets
Simplifying Google Sheets

## Example Usage
```javascript
const sSheets = require("simplesheets");
const SimpleSheets = sSheets(PATH_TO_OAUTH_OBJECT, PATH_TO_TOKEN_OBJECT);
SimpleSheets.create({
    spreadsheetTitle:{
        sheetTitle:[["hello"]]
    }
}).then(simplesheet=>console.log(simplesheet)).catch(e=>console.log(e));
```

## Table of Contents
+ [SimpleSheets](#SimpleSheets)
    + [Initialization](#Initialization)
+ [Data format (worksheetData)](#Data-format-worksheetData)
+ [SimpleSheet Object](#SimpleSheet-Object)
+ [SimpleSheets Methods [Asynchronous]](#SimpleSheets-Methods-[Asynchronous])
    + [create(worksheetData)](#createworksheetData)
    + [createFromCSV(title, [csv,...])](#createFromCSVtitle-csv)
    + [get(id)](#getid)
    + [exists(id)](#existsid)    
    + [update(id, worksheetData)](#updateid-worksheetData)
    + [remove(id)](#removeid-folderId)
    + [move(id, folderId)](#moveid-folderId)

## SimpleSheets
---
### Initialization
> Initialize the SimpleSheets object by authenticating the user

Arguments:
+ `OAUTH`<**String**\|**Object**>: contains Google's Oauth2.0 object
    + <**String**> takes the location of an OAUTH JSON file
```javascript
// OAUTH Object (Minimal Required Properties)
{
    client_id: "...",
    client_secret: "...",
    redirect_uris: ["...", ...]
}
```
+ `Token`<**String**\|**Object**>: contains authentication token object
    + <**String**> takes the location of a token JSON file
    + <**Object**> matches [Google API Access Token](https://github.com/googleapis/google-api-nodejs-client#retrieve-access-token)
```javascript
// Token object (can be obtained from Google API)
{
    access_token: "...",
    refresh_token: "...", // Optional
    scope: "...",
    token_type: "...",
    expiry_date: ...
}
```

Return:
+ <[**SimpleSheets**](#SimpleSheets)>

```javascript
const SimpleSheets = simplesheets(<OAUTH Object>, <Access Token>);
```

---

## Data format (worksheetData)
> Dynamic rows and columns

Accepted cell values:
+ <**Date**>: converts into a String on spreadsheet
+ <**Number**>
+ <**String**>
+ `null`: empty cell
```javascript
{
    titleOfSpreadSheet: {
        titleOfSheet1: [[value, value,...],
                        [value, null,...,...],
                        [...]],
        titleOfSheet2: ...,
        ...
    }
}
```

---

## SimpleSheet Object
```javascript
{
    id: stringID,
    worksheet: worksheetData,
    folder: null // MyDrive
}
```

---

## SimpleSheets Methods *[Asynchronous]*
---
### create(worksheetData)
> Creates a spreadsheet based on the data given

Arguments:
+ `worksheetData` <**Object**>: an object in the form of [**worksheetData**](#Data-format-%28worksheetData%29)

Return:
+ <[**SimpleSheet Object**](#SimpleSheet-Object)>

---
### createFromCSV(title, [csv,...])
> Converts csv file(s) to a spreadsheet

Arguments:
+ `title` <**String**>: title of the spreadsheet
+ `[csv,...]` <**String**\|**Array**(**String**)>: single location of CSV file **OR** array of the locations of CSV files

Return:
+ <[**SimpleSheet Object**](#SimpleSheet-Object)>

---
### get(id)
> Gets the data of a spreadsheet

Arguments:
+ `id` <**String**>: the string id of the spreadsheet

Return:
+ <[**SimpleSheet Object**](#SimpleSheet-Object)>

---
### exists(id)
> Check to see if spreadsheet exists and is able to be read/written

Arguments:
+ `id` <**String**>: the string id of the spreadsheet

Return:
+ <**Boolean**>

---
### update(id, worksheetData)
> Updates the spreadsheet with new data

Arguments:
+ `id` <**String**>: the string id of the spreadsheet
+ `worksheetData` <**Object**>: an object in the form of [**worksheetData**](#Data-format-%28worksheetData%29)

Return:
+ <[**SimpleSheet Object**](#SimpleSheet-Object)>: contains updated values

---
### remove(id)
> Deletes the spreadsheet

Arguments:
+ `id` <**String**>: the string id of the spreadsheet

Return:
+ <**Boolean**>: `true` if the spreadsheet has been successfully deleted

---
### move(id, folderId)
> Moves the spreadsheet to a specified folder

Arguments:
+ `id` <**String**>: the string id of the spreadsheet
+ `folderId` <**String**>: the string id of the folder
    + Use `null` to move to *MyDrive* instead of a folder

Return:
+ `folderId` <**String**>: the string id of the folder