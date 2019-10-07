const {google} = require("googleapis");
const fs = require('fs');
const path = require('path');
const OPTIONS = {include: [], flex: false};

function firstKey(obj){
	return Object.keys(obj)[0];
}
function empty(obj){
	return Object.keys(obj).length === 0;
}
function cpy(duplicateObj, originalObj){
	for (let key in originalObj) {
		if (originalObj[key] !== undefined) {
			if (duplicateObj[key] === undefined) duplicateObj[key] = originalObj[key];
			if (originalObj[key] instanceof Array) {
				if (typeof duplicateObj[key] === "string") {
					duplicateObj[key] = [duplicateObj[key]];
				}
			}
			if (typeof originalObj[key] !== typeof duplicateObj[key]) throw new TypeError(`Type Mismatch: ${key}`);
		}
	}
	for (let key in duplicateObj) {
		if (originalObj[key] === undefined) delete duplicateObj[key];
	}
	return duplicateObj;	
}

function validate(originalObj, verifyObj) {
	for (let key in verifyObj) {
		if (originalObj[key] !== undefined) {
			verifyObj[key] = originalObj[key];
		}
	}
	return verifyObj;
}

class Sheet {
	constructor(sheetObj) {
		this.sheetId = sheetObj.properties.sheetId;
		this.sheetTitle = sheetObj.properties.title;
		this.height = sheetObj.properties.gridProperties.rowCount;
		this.width = sheetObj.properties.gridProperties.columnCount;
		this.data = [];
		let sheetData = sheetObj.data[0].rowData;
		for (let i = 0; i < sheetData.length; i++) {
			let sheetRow = sheetData[i].values;
			let row = [];
			for (let k = 0; k < sheetRow.length; k++) {
				let type = firstKey(sheetRow[k]);
				let value = sheetRow[k].effectiveValue[type];				
				switch(type) {
					case 'numberValue': row.push(Number(value)); break;
					case 'boolValue': row.push(!!value); break;
					default: row.push(value); break;
				}
			}
			this.data.push(row);
		}
	}
}

class Metadata {
	constructor(sheetMeta) {
		this.id = sheetMeta.metadataId;
		this.value = sheetMeta.metadataValue;
		this.path = null;
		if (sheetMeta.location.locationType === 'SHEET') this.path = sheetMeta.location.sheetId;
	}
}

class Worksheet {
	constructor(worksheetObj) {
		this.worksheetId = worksheetObj.spreadsheetId;
		this.worksheetTitle = worksheetObj.properties.title;
		this.sheets = {};
		this.metadata = [];
		for (let i = 0; i < worksheetObj.sheets.length; i++) {
			let title = worksheetObj.sheets[i].properties.title;
			this.sheets[title] = new Sheet(worksheetObj.sheets[i]);
			let sheetMeta = sheetObj.developerMetadata;
			this.metadata[title] = {};			
			for (let k = 0; k < sheetMeta.length; k++) {
				this.metadata[title][sheetMeta[i].metadataKey] = new Metadata(sheetMeta[i]);
			}	
		}
	}
	simplify(){
		this.title = this.worksheetTitle;
		for (let sheetTitle in this.sheets) {
			this.sheets[sheetTitle] = this.sheets[sheetTitle].data;
			let meta = this.metadata;
			for (let metaKey in this.metadata[sheetTitle]) {
				this.metadata[sheetTitle][metaKey] = this.metadata[sheetTitle][metaKey].value;
			}
		}
		delete this.worksheetId;
		delete this.worksheetTitle;
		return this;
	}
}

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
				let cell = (typeof other === 'object') ? other[Object.keys(other)[0]] : (!isNaN(Number(other)) ? Number(other) : other);
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
	client_secret = clientJSON["client_secret"];
	if (typeof token === 'string') {
		token = JSON.parse(fs.readFileSync(token, {encoding: "utf8"}));
	}
	let oauthCode = new google.auth.OAuth2(client_id, client_secret);
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
		if (!sheetObj) return undefined;
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
		let temp = firstKey(fullArray);
		let sheets = {
			[temp]: {}
		}
		for (let key in fullArray[temp]) {
			sheets[temp][key] = fullArray[temp][key].data;
		}
		return sheets;
	}
	static toArray(sheetObj){
		if (!sheetObj) return undefined;
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
					let cell = (typeof other === 'object') ? other[Object.keys(other)[0]] : (!isNaN(Number(other)) ? Number(other) : other);
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
		if (val === null || val === undefined) val = '';		
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
		let title = firstKey(this.data);
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
	static async fetch(id, auth) {
		let temp = new gSheets(null, auth);
		temp.id = id;
		return new Promise((resolve, reject)=>{
			temp.authenticated.spreadsheets.get({spreadsheetId: id, includeGridData: false}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);					
					reject(new Error(err.errors[0].message));
				}
				else {
					let sheetObj = {[res.data.properties.title]: {}};
					res.data.sheets.forEach(v=>{
						sheetObj[res.data.properties.title][v.properties.title] = v.properties.sheetId;
					});
					resolve(sheetObj);
				}
			});
		});	
	}
	static async get(id, auth, decoder, options) {
		options = cpy(options || {}, OPTIONS);
		let temp = new gSheets(null, auth);
		temp.id = id;
		if (options.include.length > 0) {
			let all = await this.fetch(id, auth);
			let requests = [];
			for (let key in all[firstKey(all)]) {
				if (options.include.indexOf(key) !== -1) requests.push({gridRange: {sheetId: all[firstKey(all)][key]}});
			}
			return new Promise((resolve, reject)=>{
				temp.authenticated.spreadsheets.getByDataFilter({spreadsheetId: id, resource: {dataFilters: requests, includeGridData: true}}, (err, res)=>{
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
	static async update(id, auth, newData, options) {
		options = cpy(options || {}, OPTIONS);
		let sheetObj = await this.get(id, auth, Format.toArrayFull, options);
		let requests = [];
		let obj = sheetObj.data;
		let title = firstKey(obj);
		let currentIds = [];
		for (let key in obj[title]) currentIds.push(obj[title][key].id);
		if (!options.flex) {
			for (let key in obj[title]) {
				if (newData[key] === undefined) requests.push({deleteSheet: {sheetId: obj[title][key].id}});
			}
		}
		for (let key in newData) {
			let _old = obj[title][key];
			let _new = newData[key];
			if (_new === undefined) continue;
			if (_old === undefined) {
				let temp = Format.createSheet(key, _new);
				let id = Format.fullRandId(currentIds);
				currentIds.push(id);
				requests.push({addSheet: {properties: {title: temp.properties.title, sheetId: id, gridProperties: temp.properties.gridProperties}}});	
				requests.push({updateCells: {fields: "*", rows: temp.data[0].rowData, range: {sheetId: id, startRowIndex: 0, startColumnIndex: 0}}})		
				continue;
			}
			let sheetUp = this.compare(_old, _new, key);
			if (sheetUp) sheetUp.forEach(d=>requests.push(d));
		}
		/*if (title !== title2) {
			requests.push({updateSpreadsheetProperties: {
					properties: {
						title: title2,
					},
					fields: "*"
				}
			});	
		}	*/	
		if (requests.length === 0) {
			sheetObj.data = Format.fromFullToArray(sheetObj.data);
			return sheetObj;		
		}
		return new Promise((resolve, reject)=>{
			let REQ = {spreadsheetId: id, resource: {requests: requests, includeSpreadsheetInResponse: true, responseIncludeGridData: true}};
			if (options.include) REQ.resource.responseRanges = options.include;
			sheetObj.authenticated.spreadsheets.batchUpdate(REQ, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);
					reject(new Error(err.errors[0].message));
				}
				else {
					sheetObj.data = Format.toArray(res.data.updatedSpreadsheet);
					sheetObj.id = res.data.spreadsheetId;
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
						rows: [{values: [{userEnteredValue: Format.value((newData[i][j]!== undefined) ? newData[i][j] : oldData[i][j])}]}], 
						fields: "*",
						start: {sheetId: oldSheet.id, rowIndex: i, columnIndex: j}}
					});	
					continue;				
				}
				if (oldData[i][j] !== newData[i][j]) 
					sheetUpdate.push({updateCells: {
						rows: [{values: [{userEnteredValue: Format.value((newData[i][j]!== undefined) ? newData[i][j] : oldData[i][j])}]}], 
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
	static async getProp(id, auth) {
		let prop = {folder: null, trashed: false, title: null, description: "", open: false};
		let tempAuth = google.drive({version: "v2", auth});
		return await new Promise((resolve, reject)=>{
			tempAuth.files.get({fileId: id, scope: "parents"}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);
					if (err.errors[0].reason === "notFound") {
						console.warn('\x1b[33m%s\x1b[0m', `WARNING: Spreadsheet is not found: ${id}`);
						return resolve(prop);
					}
					else return reject(new Error(err.errors[0].message));
				}
				if (res.data.labels.trashed) {
					console.warn('\x1b[33m%s\x1b[0m',`WARNING: Spreadsheet is in trash bin: ${id}`);
					prop.trashed = true;
				} else prop.open = true;
				prop.title = res.data.title;
				prop.description = res.data.description;
				let parents = [];
				res.data.parents.forEach((val)=>{
					if (!val.isRoot) parents.push(val.id);
				});
				if (parents.length === 1)
					prop.folder = parents.join(",");
				resolve(prop);
			});
		});
	}
	static async move(id, folderId, auth) {
		let tempAuth = google.drive({version: "v2", auth});
		let resource = {
			fileId: id,
			addParents: folderId
		}
		let parentsProp = await this.getProp(id, auth);
		if (!parentsProp.open) throw new Error("Cannot find spreadsheet");
		if (parentsProp.folder) resource["removeParents"] = parents;
		if (!parentsProp.folder && !folderId) return id;
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
	static async setProp(id, newProp, auth) {
		newProp = newProp || {}
		let tempAuth = google.drive({version: "v2", auth});
		let resource = {};
		if (newProp.details) resource.description = newProp.details;
		if (newProp.title) resource.title = newProp.title;	
		if (empty(resource)) return null;
		return new Promise((resolve, reject)=>{
			tempAuth.files.update({fileId: id, resource: resource}, (err, res)=>{
				if (err) {
					if (!err.errors) return reject(err);					
					reject(new Error(err.errors[0].message));
				}
				else resolve((resource) ? resource : '');
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
	constructor(id, data, prop) {
		this.id = id;
		this.worksheet = data ? data[firstKey(data)] : undefined;
		this.folder = prop.folder;
		this.details = prop.description || "";
		this.title = prop.title;
		this.trashed = prop.trashed;
	}
	sheet(name){
		if (!name) {
			let first = firstKey(this.worksheet);
			return this.worksheet[first];
		} else if (name){
			return this.worksheet[name];
		}
	}
}

class MiniSheets {
	constructor(auth, token) {
		this.auth = OAUTH2(auth, token);
	}
	async create(data) {
		if (!(data instanceof Object)) throw new Error("Invalid worksheet data");		
		let temp = new gSheets(data, this.auth);
		return filter(await temp.create(), {folder: null, trashed: false, title: firstKey(data), description: "", open: true});
	}
	async createFromCSV(title, data) {
		return this.create({[title]: csvToJSON.apply(null, files)});		
	}
	async get(id, options){
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");
		let fileProp = await gSheets.getProp(id, this.auth);
		if (!fileProp.open) return false;
		let sh = await gSheets.get(id, this.auth, Format.toArray, options);
		return filter(sh, fileProp);
	}
	async exists(id) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");		
		return ((await gSheets.getProp(id, this.auth).open)) ? true : false;
	}
	async update(id, newData, options) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");
		if (!(newData instanceof Object)) throw new Error("Invalid worksheet data");
		let fileProp = await gSheets.getProp(id, this.auth);		
		let sh = await gSheets.update(id, this.auth, newData, options);
		return filter(sh, fileProp);
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
	async setProp(id, props) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");		
		if (typeof props !== 'object') throw new Error("Invalid properties, properties must be an object");	
		return gSheets.setProp(id, props, this.auth);
	}
	async getProp(id) {
		if (typeof id !== "string") throw new Error("Invalid ID, ID must be a string");		
		return gSheets.getProp(id, this.auth);
	}	
}

function filter(simpleObj, prop){
	if (simpleObj instanceof gSheets) {
		return new MiniSheet(simpleObj.id, simpleObj.data, prop);
	} else throw new Error("Cannot convert MiniSheets into safe object to be modified");
}

function MiniSheets_Global(auth, token){
	return new MiniSheets(auth, token);
}
module.exports = MiniSheets_Global;