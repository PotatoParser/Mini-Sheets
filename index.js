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
function validate(originalObj, verifyObj) {
	for (let key in verifyObj) {
		if (originalObj[key] !== undefined) {
			verifyObj[key] = originalObj[key];
		}
	}
	return verifyObj;
}

class SingleMetadata {
	constructor(key, value, id, parentMetadata) {
		this.key = key;
		this.value = value;
		this.id = id;
		this.parent = parentMetadata;
	}
	toSingleMetadataObject(){
		let temp = {
			metadataKey: this.key,
			metadataValue: this.value,
			visibility: 'DOCUMENT',
			location: {
				sheetId: this.parent.path
			}
		};
		if (this.id) temp.metadataId = this.id;		
		return temp;
	}
	equals(singleMetadata){
		for (let key in this) {
			if (key === 'parent') continue;
			if (this[key] !== singleMetadata[key]) return false;
		}
		return true;
	}
}

class Metadata {
	constructor(sheetMetaPath = null, sheetMetaPairs = {}) {
		this.path = sheetMetaPath;
		this.data = {};
		for (let metaKey in sheetMetaPairs) {
			this.data[metaKey] = new SingleMetadata(metaKey, sheetMetaPairs[metaKey], null, this);
		}
	}
	static parse(sheetObj) {
		let parsedSheetMeta = new Metadata();
		parsedSheetMeta.path = sheetObj.properties.sheetId;
		if (!sheetObj.developerMetadata) return parsedSheetMeta;
		for (let i = 0; i < sheetObj.developerMetadata.length; i++) {
			let meta = sheetObj.developerMetadata[i];
			parsedSheetMeta.data[meta.metadataKey] = new SingleMetadata(meta.metadataKey, meta.metadataValue, meta.metadataId, parsedSheetMeta);
		}
		return parsedSheetMeta;
	}
	fillEmpty(oldMetadata){
		if (!this.path) this.path = oldMetadata.path;
		for (let metaKey in this.data) {
			if (oldMetadata.data[metaKey]) {
				this.data[metaKey].id = this.data[metaKey].id || oldMetadata.data[metaKey].id;			
			}
		}
		for (let metaKey in oldMetadata.data) {
			if (!this.data.hasOwnProperty(metaKey)) this.data[metaKey] = new SingleMetadata(metaKey, oldMetadata.data[metaKey].value, oldMetadata.data[metaKey].id, this);
		}
	}
	toMetadataObject(){
		let obj = [];
		for (let metaKey in this.data) {
			obj.push(this.data[metaKey].toSingleMetadataObject());
		}
		return obj;
	}
}

class Sheet {
	constructor(sheetTitle, sheetGridData = [[]]) {
		this.sheetId = null;
		this.sheetTitle = sheetTitle;
		let columnCount = sheetGridData[0].length;
		for (let i = 1; i < sheetGridData.length; i++) {
			if (sheetGridData[i].length !== columnCount) throw new Error('Column count does not match per row');
		}
		this.rows = sheetGridData.length;
		this.columns = sheetGridData[0].length;
		this.data = sheetGridData;
	}
	static parse(sheetObj) {
		let parsedSheet = new Sheet();
		parsedSheet.sheetId = sheetObj.properties.sheetId;
		parsedSheet.sheetTitle = sheetObj.properties.title;
		parsedSheet.rows = sheetObj.properties.gridProperties.rowCount;
		parsedSheet.columns = sheetObj.properties.gridProperties.columnCount;
		parsedSheet.data = [];
		if (sheetObj.data) {
			let sheetData = sheetObj.data[0].rowData;
			for (let i = 0; i < sheetData.length; i++) {
				let sheetRow = sheetData[i].values;
				let row = [];
				for (let k = 0; k < sheetRow.length; k++) {
					if (empty(sheetRow[k])) {
						row.push(null);
						continue;
					}
					let type = firstKey(sheetRow[k].effectiveValue || sheetRow[k].userEnteredValue);
					let value = (sheetRow[k].effectiveValue) ? sheetRow[k].effectiveValue[type] : sheetRow[k].userEnteredValue[type];				
					switch(type) {
						case 'numberValue': row.push(Number(value)); break;
						case 'boolValue': row.push(!!value); break;
						default: row.push(value); break;
					}
				}
				parsedSheet.data.push(row);
			}		
		}
		return parsedSheet;
	}
	fillEmpty(oldSheet = {}){
		if (oldSheet.sheetId) this.sheetId = oldSheet.sheetId;
	}
	toSheetObject(metadata = {data: {}}){
		let obj = {
			properties: {
				sheetId: this.sheetId,
				title: this.sheetTitle,
				gridProperties: {
					rowCount: this.rows,
					columnCount: this.columns
				}
			},
			data: [{
				rowData: [],
				startRow: 0,
				startColumn: 0,
			}],
			developerMetadata: [],
		}
		for (let i = 0; i < this.data.length; i++) {
			let singleRow = {
				values: []
			}
			for (let k = 0; k < this.data[i].length; k++) {
				let gridValue = this.data[i][k];
				switch (typeof gridValue) {
					case 'string': singleRow.values.push({userEnteredValue: {stringValue: gridValue}}); break;
					case 'number': singleRow.values.push({userEnteredValue: {numberValue: gridValue}}); break;
					case 'boolean': singleRow.values.push({userEnteredValue: {boolValue: gridValue}}); break;
					default: singleRow.values.push({}); break;
				}
			}
			obj.data[0].rowData.push(singleRow);
		}
		if (empty(metadata.data)) delete obj.developerMetadata;
		else obj.developerMetadata = metadata.toMetadataObject();
		return obj;
	}
}

class Worksheet {
	constructor(worksheetObj = {spreadsheetId: null, properties: {title: null}, sheets: []}) {
		this.worksheetId = worksheetObj.spreadsheetId;
		this.worksheetTitle = worksheetObj.properties.title;
		this.maxId = 0;
		this.sheets = {};
		this.metadata = {};
		for (let i = 0; i < worksheetObj.sheets.length; i++) {
			let title = worksheetObj.sheets[i].properties.title;
			this.sheets[title] = Sheet.parse(worksheetObj.sheets[i]);
			this.metadata[title] = Metadata.parse(worksheetObj.sheets[i]);
		}
	}
	simplify(){
		this.title = this.worksheetTitle;
		for (let sheetTitle in this.sheets) {
			this.sheets[sheetTitle] = this.sheets[sheetTitle].data;
			let meta = this.metadata;
			this.metadata = {};
			for (let sheetMetaTitle in meta) {
				this.metadata[sheetMetaTitle] = null;
				for (let metaKey in meta[sheetMetaTitle].data) {
					if (!this.metadata[sheetMetaTitle]) this.metadata[sheetMetaTitle] = {};
					this.metadata[sheetMetaTitle][metaKey] = meta[sheetMetaTitle].data[metaKey].value;
				}
			}
		}
		delete this.worksheetId;
		delete this.worksheetTitle;
		return this;
	}
	static create(title, gridData = {}, metadata = {}){
		if (!title) throw new Error('Missing Spreadsheet Title');
		let generatedWorksheet = new Worksheet();
		generatedWorksheet.worksheetTitle = title;
		for (let sheetTitle in gridData) {
			generatedWorksheet.sheets[sheetTitle] = new Sheet(sheetTitle, gridData[sheetTitle]);
			this.maxId = Math.max(this.maxId, generatedWorksheet.sheets[sheetTitle].sheetId);
		}
		for (let sheetTitle in metadata) {
			//if (!gridData[sheetTitle]) throw new Error(`${sheetTitle} does not exist as a sheet title`);
			generatedWorksheet.metadata[sheetTitle] = {};
			for (let metaTitle in metadata[sheetTitle]) {
				generatedWorksheet.metadata[sheetTitle] = new Metadata(sheetTitle, metadata[sheetTitle]);
			}
		}
		return generatedWorksheet;
	}
	toSpreadsheetObject(){
		let obj = {
			properties: {
				title: this.worksheetTitle
			},
		}
		if (this.worksheetId) obj.spreadsheetId = this.worksheetId;
		if (!empty(this.sheets)) obj.sheets = [];
		for (let sheetTitle in this.sheets) {
			obj.sheets.push(this.sheets[sheetTitle].toSheetObject(this.metadata[sheetTitle]));
		}
		return obj
	}
	createSheetId(){
		this.maxId++;
		return this.maxId;
	}
}

class Google {
	constructor(client_id, token){
		let client_secret;
		if (typeof client_id === 'object') {
			client_secret = client_id.client_secret;
			client_id = client_id.client_id;
		}
		this.oauth = new google.auth.OAuth2(client_id, client_secret);
		this.oauth.setCredentials(token);
	}
}

class Drive extends Google {
}

class Spreadsheets extends Google {
	constructor(client_id, token) {
		super(client_id, token);
		this.spreadsheets = google.sheets({version: "v4", auth: this.oauth}).spreadsheets;
	}
	createSpreadsheet(title, gridData, metadata) {
		return new Promise((resolve, reject)=>{
			this.spreadsheets.create({resource: Worksheet.create(title, gridData, metadata).toSpreadsheetObject()}, (err, res)=>{
				if (err) return reject(err);
				resolve(new Worksheet(res.data));
			});
		});
	}
	getRawSpreadsheet(spreadsheetId) {
		return new Promise((resolve, reject)=>{
			this.spreadsheets.get({spreadsheetId: spreadsheetId, includeGridData: false}, (err, res)=>{
				if (err) reject(err);
				else {
					resolve(new Worksheet(res.data));
				}
			});
		});
	}
	async getSpreadsheet(spreadsheetId, _options = {}){
		_options = validate(_options, {include: []});
		if (!(_options.include instanceof Array)) _options.include = [_options.include];
		if (_options.include.length > 0) {
			let preSpreadsheet = await this.getRawSpreadsheet(spreadsheetId);
			return await new Promise((resolve, reject)=>{
				let dataFilters = [];
				for (let i = 0; i < _options.include.length; i++) {
					if (!preSpreadsheet.sheets[_options.include]) return reject(new Error('Sheet name does not exist'));
					dataFilters.push({gridRange: {sheetId: preSpreadsheet.sheets[_options.include].sheetId}});
				}
				this.spreadsheets.getByDataFilter({
					spreadsheetId: spreadsheetId, 
					resource: {
						dataFilters: dataFilters, 
						includeGridData: true
					}
				}, (err, res)=>{
					if (err) return reject(err);
					resolve(new Worksheet(res.data));
				});
			});
		} else {
			return await new Promise((resolve, reject)=>{
				this.spreadsheets.get({
					spreadsheetId: spreadsheetId,
					includeGridData: true
				}, (err, res)=>{
					if (err) return reject(err);
					resolve(new Worksheet(res.data));
				});
			});
		}
	}
	async setSpreadsheet(spreadsheetId, gridData, metadata){
		//for (let sheetTitle in gridData) if (_options.include.indexOf(sheetTitle) === -1) throw new Error(`${sheetTitle} not included in options`);
		let preSpreadsheet = await this.getRawSpreadsheet(spreadsheetId),
			newSpreadsheet = Worksheet.create(preSpreadsheet.worksheetTitle, gridData, metadata);
		let requests = [];
		let sheetTitles = {};
		for (let key in newSpreadsheet.sheets) sheetTitles[key] = true;
		for (let key in newSpreadsheet.metadata) sheetTitles[key] = true;
		for (let sheetTitle in sheetTitles) {
			let newConvertedSheet, oldConvertedSheet = {developerMetadata: []}, 
				newSheet = newSpreadsheet.sheets[sheetTitle], 
				preSheet = preSpreadsheet.sheets[sheetTitle],
				newMeta = newSpreadsheet.metadata[sheetTitle],
				preMeta = preSpreadsheet.metadata[sheetTitle] || {data: {}};
			if(newSheet) newSheet.fillEmpty(preSheet);
			if(newMeta) newMeta.fillEmpty(preMeta);
			console.log(newMeta);
			if (newSheet && !preSheet) {
				let tempId = preSpreadsheet.generateId();
				newSheet.sheetId = tempId;			
				requests.push({
					addSheet: {
						properties: newSheet.toSheetObject().properties
					}
				});				
			}
			if (newMeta) {
				for (let metaKey in newMeta.data) {
					if (!preMeta.data[metaKey]) {
						requests.push({
							createDeveloperMetadata: {
								developerMetadata: newMeta.data[metaKey].toSingleMetadataObject()
							}
						});
					} else if (!newMeta.data[metaKey].equals(preMeta.data[metaKey])) {
						requests.push({
							updateDeveloperMetadata: {
								fields: '*',
								dataFilters: [{
									gridRange: {
										sheedId: newMeta.data[metaKey].parent.path
									}
								}],
								developerMetadata: newMeta.data[metaKey].toSingleMetadataObject()
							}
						});						
					}
				}
			}
			if ((newSheet && preSheet) && (newSheet.rows !== preSheet.rows || tempSheet.columns !== preSheet.columns)) {
				requests.push({
					updateSheetProperties: {
						fields: '*',
						properties: newSheet.toSheetObject().properties
					}
				});				
			}
			if (!empty(gridData)) {
				requests.push({
					updateCells: {
						fields: '*',
						rows: newSheet.toSheetObject().data[0].rowData,
						range: {
							sheetId: newSheet.toSheetObject().properties.sheetId,
							startRowIndex: 0,
							startColumnIndex: 0,
						}
					}
				});
			}
			return requests;
		}
	}
}

/*class Sheet {
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
}*/

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