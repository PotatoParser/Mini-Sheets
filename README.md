# Mini Sheets
Minified and simplified Google Sheets

## Example Usage
```javascript
const mSheets = require("minisheets");
const MiniSheets = mSheets(PATH_TO_OAUTH_OBJECT, PATH_TO_TOKEN_OBJECT);
MiniSheets.create({
    spreadsheetTitle:{
        sheetTitle:[["hello"]]
    }
}).then(minisheet=>console.log(minisheet)).catch(e=>console.log(e));
```

## Table of Contents
+ [MiniSheets](#minisheets)
    + [Initialization](#initialization)
+ [Data format (worksheetData)](#data-format-worksheetdata)
+ [MiniSheet Object](#miniSheet-object)
+ [MiniSheets Methods [Asynchronous]](#minisheets-methods-asynchronous)
    + [create(worksheetData)](#createworksheetdata)
    + [createFromCSV(title, [csv,...])](#createfromcsvtitle-csv)
    + [get(id, options)](#getid-options)
    + [exists(id)](#existsid)    
    + [update(id, worksheetData, options)](#updateid-worksheetdata-options)
    + [remove(id)](#removeid-folderid)
    + [move(id, folderId)](#moveid-folderid)

## MiniSheets
---
### Initialization
> Initialize the MiniSheets object by authenticating the user

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
+ <[**MiniSheets**](#MiniSheets)>

```javascript
const MiniSheets = minisheets(<OAUTH Object>, <Access Token>);
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

## MiniSheet Object
```javascript
{
    id: stringID,
    worksheet: worksheetData,
    folder: null, // MyDrive
    details: "",
    title: "",
    trashed: false,
}
```

## MiniSheet Methods
---
### sheet(sheetTitle)
> Fetches a sheet based off of title

Arguments:
+ (OPTIONAL) `sheetTitle` <**String**>: a sheet title OR leave blank for first sheet

Return:
+ <**Object**>

---

## MiniSheets Methods *[Asynchronous]*
---
### create(worksheetData)
> Creates a spreadsheet based on the data given

Arguments:
+ `worksheetData` <**Object**>: an object in the form of [**worksheetData**](#Data-format-%28worksheetData%29)

Return:
+ <[**MiniSheet Object**](#MiniSheet-Object)>

---
### createFromCSV(title, [csv,...])
> Converts csv file(s) to a spreadsheet

Arguments:
+ `title` <**String**>: title of the spreadsheet
+ `[csv,...]` <**String**\|**Array**(**String**)>: single location of CSV file **OR** array of the locations of CSV files

Return:
+ <[**MiniSheet Object**](#MiniSheet-Object)>

---
### get(id, options)
> Gets the data of a spreadsheet

Arguments:
+ `id` <**String**>: the string id of the spreadsheet
+ (OPTIONAL) `options` <**Object**>: options object
    + `include` <**String**|**Array**>: specifies only certain sheets

Return:
+ <[**MiniSheet Object**](#MiniSheet-Object)>

---
### exists(id)
> Check to see if spreadsheet exists and is able to be read/written

Arguments:
+ `id` <**String**>: the string id of the spreadsheet

Return:
+ <**Boolean**>

---
### update(id, worksheetData, options)
> Updates the spreadsheet with new data

Arguments:
+ `id` <**String**>: the string id of the spreadsheet
+ `worksheetData` <**Object**>: an object in the form of [**worksheetData**](#Data-format-%28worksheetData%29)
+ (OPTIONAL) `options` <**Object**>: options object
    + `flex` <**Boolean**>: enables exact update from worksheetData **[Default: false]**
    + `include` <**String**|**Array**>: specifies only certain sheets

Return:
+ <[**MiniSheet Object**](#MiniSheet-Object)>: contains updated values

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