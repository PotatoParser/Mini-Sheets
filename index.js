const {google} = require("googleapis");
const fs = require('fs');
const path = require('path');

Object.defineProperty(Object.prototype, "firstKey", {
	enumerable: false,
	value: function(){
		return Object.keys(this)[0];
	}
});

class Sheet {
	constructor(sheetObj){
		let arr = [];
		let data = sheetObj.data[0].rowData;
		for (let k = 0; k < sheetObj.properties.gridProperties.rowCount; k++) {
			let row = [];
			if (data[k] === undefined) {
				for (let i = 0; i < sheetObj.properties.gridProperties.columnCount; i++) row.push(null);
				arr.push(row);
				continue;				
			}
			let insideTemp = data[k].values;
			if (!insideTemp) {
				for (let i = 0; i < sheetObj.properties.gridProperties.columnCount; i++) row.push(null);
				arr.push(row);
				continue;
			}
			for (let j = 0; j < insideTemp.length; j++) {
				let other = insideTemp[j].formattedValue;
				let cell = (typeof other === 'object') ? other[Object.keys(other)[0]] : ((isNaN(Number(other)) ? ((isNaN(Date.parse(other))) ? other : new Date(other)) : Number(other)));
				row.push(cell);
			}
			arr.push(row);
		}
		this.data = arr;
		this.title = sheetObj.properties.title;
		this.id = sheetObj.properties.sheetId;
		this.dimensions = {row: sheetObj.properties.gridProperties.rowCount, col: sheetObj.properties.gridProperties.columnCount};
	}
}

function OAUTH2(clientJSON, token) {
	let client_id, client_secret, redirect_uris;
	if (typeof clientJSON === 'string') {
		clientJSON = JSON.parse(fs.readFileSync(clientJSON, {encoding: "utf8"}));
	}
	client_id = clientJSON.client_id;
	client_secret = clientJSON.client_secret;
	redirect_uris = clientJSON.redirect_uris;
	if (typeof token === 'string') {
		token = JSON.parse(fs.readFileSync(token, {encoding: "utf8"}));
	}
	let oauthCode = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
	oauthCode.generateAuthUrl({access_type:"offline",scope: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file"]});
	oauthCode.setCredentials(token);
	return oauthCode;
}

class Format {
	static fullRandId(arr) {
		let rand = ()=>{
			let temp = "";
			for (let i = 0; i < 9; i++) temp+=Math.floor(Math.random()*(9) + 1);
			return Number(temp);	
		}
		let num = rand();
		while(arr.indexOf(num) !== -1) num = rand();
		return num;
	}
	static toArrayFull(sheetObj){
		let temp = sheetObj.sheets;
		let sheets = {
			[sheetObj.properties.title]: {}
		};
		for (let i = 0; i < temp.length; i++) {
			let tempSheet = new Sheet(temp[i]);
			sheets[sheetObj.properties.title][temp[i].properties.title] = tempSheet;
		}
		return sheets;
	}
	static fromFullToArray(fullArray) {
		let temp = fullArray.firstKey();
		let sheets = {
			[temp]: {}
		}
		for (let key in fullArray[temp]) {
			sheets[temp][key] = fullArray[temp][key].data;
		}
		return sheets;
	}
	static toArray(sheetObj){
		let temp = sheetObj.sheets;
		let sheets = {
			[sheetObj.properties.title]: {}
		};
		for (let i = 0; i < temp.length; i++) {
			let arr = [];
			let data = temp[i].data[0].rowData;
			for (let k = 0; k < data.length; k++) {
				let row = [];
				let insideTemp = data[k].values;
				if (!insideTemp) continue;
				for (let j = 0; j < insideTemp.length; j++) {
					let other = insideTemp[j].formattedValue;
					let cell = (typeof other === 'object') ? other[Object.keys(other)[0]] : ((isNaN(Number(other)) ? ((isNaN(Date.parse(other))) ? other : new Date(other)) : Number(other)));
					row.push(cell);
				}
				arr.push(row);
			}
			sheets[sheetObj.properties.title][temp[i].properties.title] = arr;
		}
		return sheets;
	}		
	static value(val){
		let variable;
		if (val instanceof Date) val = val.toString();
		switch (typeof val) {
			case "string": variable = "stringValue"; break;
			case "boolean": variable = "boolValue"; break;
			case "number": variable = "numberValue"; break;
			default: throw new Error(`Wrong cell type detected: ${JSON.stringify(val)}`);
		}
		return {[variable]: (val !== '') ? val : null}
	}
	static createSheet(title, data) {
		let maxRow = data.length, maxColumn = data[0].length;
		data.forEach((val)=>{maxColumn = Math.max(maxColumn, val.length)});
		let sheet = {
			properties: {
				title: title,
				gridProperties: {
					rowCount: maxRow,
					columnCount: maxColumn
				}
			},
			data: []
		}
		let tempRow = {
			startRow: 0,
			startColumn: 0,
			rowData: []
		};		
		for (let i = 0; i < data.length; i++) {
			let temp = {values: []};
			for (let k = 0; k < data[i].length; k++) {
				let val = {
					userEnteredValue: Format.value(data[i][k])
				};
				temp.values.push(val);
			}
			tempRow.rowData.push(temp);
		}
		sheet.data.push(tempRow);
		return sheet;
	}	
}

class gSheets {
	constructor(data, auth) {
		this.authenticated = google.sheets({version: "v4", auth});
		this.data = data;
		this.id = null;		
	}
	constructWorksheet(){
		let title = this.data.firstKey();
		if (this.data[title] instanceof Array) throw new Error("Spreadsheet does not have any title!");
		let worksheet = {
			properties: {
				title: title
			},
			sheets: []
		}
		for (let key in this.data[title]) {
			worksheet.sheets.push(Format.createSheet(key, this.data[title][key]));
		}
		return worksheet;
	}
	static async get(id, auth, decoder) {
		let temp = new gSheets(null, auth);
		temp.id = id;
		return new Promise((resolve, reject)=>{
			temp.authenticated.spreadsheets.get({spreadsheetId: id, includeGridData: true}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);					
					reject(new Error(err.errors[0].message));
				}
				else {
					temp.data = decoder(res.data);	
					temp.id = res.data.spreadsheetId;
					resolve(temp);
				}
			});
		});	
	}
	create() {
		return new Promise((resolve, reject)=>{
			this.authenticated.spreadsheets.create({resource: this.constructWorksheet()}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);
					reject(new Error(err.errors[0].message));
				}
				else {
					this.id = res.data.spreadsheetId;
					resolve(this);
				}
			});
		});
	}
	static async update(id, auth, newData) {
		let sheetObj = await this.get(id, auth, Format.toArrayFull);
		let requests = [];
		let obj = sheetObj.data;
		let title = obj.firstKey();
		let title2 = newData.firstKey();
		let currentIds = [];
		for (let key in obj[title]) currentIds.push(obj[title][key].id);
		for (let key in obj[title]) {
			if (newData[title2][key] === undefined) requests.push({deleteSheet: {sheetId: obj[title][key].id}});
		}
		for (let key in newData[title]) {
			let _old = obj[title][key];
			let _new = newData[title][key];
			if (_old === undefined) {
				let temp = Format.createSheet(key, _new);
				let id = Format.fullRandId(currentIds);
				currentIds.push(id);
				requests.push({addSheet: {properties: {title: temp.properties.title, sheetId: id, gridProperties: temp.properties.gridProperties}}});	
				requests.push({updateCells: {fields: "*", rows: temp.data[0].rowData, range: {sheetId: id, startRowIndex: 0, startColumnIndex: 0}}})		
				continue;
			}
			let sheetUp = this.compare(_old, _new, title2);
			if (sheetUp) sheetUp.forEach(d=>requests.push(d));
		}
		if (title !== title2) {
			requests.push({updateSpreadsheetProperties: {
					properties: {
						title: title2,
					},
					fields: "*"
				}
			});	
		}		
		if (requests.length === 0) {
			sheetObj.data = Format.fromFullToArray(sheetObj.data);
			return sheetObj;		
		}
		return new Promise((resolve, reject)=>{
			sheetObj.authenticated.spreadsheets.batchUpdate({spreadsheetId: id, resource: {requests: requests, includeSpreadsheetInResponse: true, responseIncludeGridData: true}}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);
					reject(new Error(err.errors[0].message));
				}
				else {
					sheetObj.data = Format.toArray(res.data.updatedSpreadsheet);
					sheetObj.id = res.data.updatedSpreadsheet.spreadsheetId;
					resolve(sheetObj);
				}
			});
		});		
	}
	static compare(oldSheet, newData, newTitle) {
		let sheetUpdate = [];
		let oldData = oldSheet.data;
		let newDimensions = {
			row: newData.length,
			col: (()=>{
				let maxTemp = newData[0].length;
				newData.forEach(val=>{maxTemp = Math.max(maxTemp, val.length)});
				return maxTemp;
			})()
		}
		if (oldSheet.dimensions.row !== newDimensions.row || oldSheet.dimensions.col !== newDimensions.col) {
			sheetUpdate.push({updateSheetProperties: {
					properties: {
						title: newTitle,
						sheetId: oldSheet.id,
						gridProperties: {
							rowCount: newDimensions.row,
							columnCount: newDimensions.col,
						}
					},
					fields: "*"
				}
			});
		}
		for (let i = 0; i < newData.length; i++) {
			for (let j = 0; j < newData[i].length; j++) {
				if (oldData[i] === undefined) {
					sheetUpdate.push({updateCells: {
						rows: [{values: [{userEnteredValue: Format.value(newData[i][j] || oldData[i][j])}]}], 
						fields: "*",
						start: {sheetId: oldSheet.id, rowIndex: i, columnIndex: j}}
					});	
					continue;				
				}
				if (oldData[i][j] !== newData[i][j]) 
					sheetUpdate.push({updateCells: {
						rows: [{values: [{userEnteredValue: Format.value(newData[i][j] || oldData[i][j])}]}], 
						fields: "*",
						start: {sheetId: oldSheet.id, rowIndex: i, columnIndex: j}}
					});
			}
		}
		return sheetUpdate.length === 0 ? undefined : sheetUpdate;		
	}
	static async remove(id, auth){
		let tempAuth = google.drive({version: "v2", auth});
		return new Promise((resolve, reject)=>{
			tempAuth.files.delete({fileId: id}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);					
					reject(new Error(err.errors[0].message));
				}
				else resolve(res);
			});
		});			
	}
	static async getFolder(id, auth) {
		let tempAuth = google.drive({version: "v2", auth});
		return await new Promise((resolve, reject)=>{
			tempAuth.files.get({fileId: id, scope: "parents"}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);
					if (err.errors[0].reason === "notFound") return resolve(false);
					else return reject(new Error(err.errors[0].message));
				}
				if (res.data.labels.trashed) {
					console.warn('\x1b[33m%s\x1b[0m',`WARNING: Spreadsheet is in trash bin: ${id}`);
					return resolve(false);
				}
				let parents = [];
				res.data.parents.forEach((val)=>{
					if (!val.isRoot) parents.push(val.id);
				});
				resolve((parents.length === 1) ? parents.join(",") : null);
			});
		});
	}
	static async move(id, folderId, auth) {
		let tempAuth = google.drive({version: "v2", auth});
		let resource = {
			fileId: id,
			addParents: folderId
		}
		let parents = await this.getFolder(id, auth);
		if (parents === false) throw new Error("Cannot find spreadsheet");
		if (parents) resource["removeParents"] = parents;
		if (!parents && !folderId) return id;
		if (folderId === null) {
			delete resource.addParents;
		}
		return new Promise((resolve, reject)=>{
			tempAuth.files.update(resource, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);					
					reject(new Error(err.errors[0].message));
				}
				else resolve((folderId) ? folderId : id);
			});
		});
	}
}

function csvToJSON(){
	let obj = {};
	for (let k = 0; k < arguments.length; k++) {
		let filePath = arguments[k];
		let temp = fs.readFileSync(filePath, {encoding: "utf8"});
		let rows = [];
		temp = temp.split("\r\n");
		for (let i = 0; i < temp.length; i++) {
			temp[i] = temp[i].split(",");
			for (let j = 0; j < temp[i].length; j++) {
				if (!isNaN(Number(temp[i][j]))) temp[i][j] = Number(temp[i][j]);
			}
		}		
		obj[path.parse(filePath).name] = temp;
	}
	return obj;
}

class MiniSheet {
	constructor(id, data, folder) {
		this.id = id;
		this.worksheet = data;
		this.folder = folder || null;
	}
}

class MiniSheets {
	constructor(auth, token) {
		this.auth = OAUTH2(auth, token);
	}
	async create(data) {
		if (!(data instanceof Object)) throw new Error("Invalid worksheet data");		
		let temp = new gSheets(data, this.auth);
		return filter(await temp.create(), null);
	}
	async createFromCSV(title, data) {
		return this.create({[title]: csvToJSON.apply(null, files)});		
	}
	async get(id){
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");
		let folder = await gSheets.getFolder(id, this.auth);
		if (!folder) return false;		
		let sh = await gSheets.get(id, this.auth, Format.toArray);
		return filter(sh, folder);
	}
	async exists(id) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");		
		return ((await gSheets.getFolder(id, this.auth) === false)) ? false : true;
	}
	async update(id, newData) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");
		if (!(newData instanceof Object)) throw new Error("Invalid worksheet data");
		let folder = await gSheets.getFolder(id, this.auth);		
		let sh = await gSheets.update(id, this.auth, newData);
		return filter(sh, folder);
	}
	async remove(id) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");		
		return gSheets.remove(id, this.auth);
	}
	async move(id, folderId) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");		
		if (typeof folderId !== "string" && folderId !== null) throw new Error("Invalid folder ID, ID must be a string");
		return gSheets.move(id, folderId, this.auth);
	}
}

function filter(simpleObj, folder){
	if (simpleObj instanceof gSheets) {
		return new MiniSheet(simpleObj.id, simpleObj.data, folder);
	} else throw new Error("Cannot convert MiniSheets into safe object to be modified");
}

function MiniSheets_Global(auth, token){
	return new MiniSheets(auth, token);
}
module.exports = MiniSheets_Global;