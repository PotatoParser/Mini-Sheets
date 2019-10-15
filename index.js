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
		if (this.value !== undefined && typeof this.value === 'string') {
			if (typeof JSON.parse(this.value) === 'object') this.value = JSON.parse(this.value);
			else if (this.value === 'true' || this.value === 'false') this.value = JSON.parse(this.value);
			else if (!isNaN(Number(this.value))) this.value = Number(this.value);
		}
		this.id = id;
		this.parent = parentMetadata;
	}
	toSingleMetadataObject(){
		let temp = {
			metadataKey: this.key,
			metadataValue: typeof this.value === 'object' ? JSON.stringify(this.value) : String(this.value),
			visibility: 'DOCUMENT',
		};
		if (this.parent.path) temp.location = {sheetId: this.parent.path};
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
			if (this.data[metaKey].value !== undefined) obj.push(this.data[metaKey].toSingleMetadataObject());
		}
		return obj;
	}
}

class Sheet {
	constructor(sheetTitle, sheetGridData) {
		this.sheetId = null;
		this.sheetTitle = sheetTitle;
		if (sheetGridData === null) {
			this.data = null;
			return this;
		}
		if (sheetGridData === true) sheetGridData = [[]];
		sheetGridData = sheetGridData || [[]];
		if (!sheetGridData[0]) throw new Error('Incomplete grid data');
		let columnCount = sheetGridData[0].length;
		for (let i = 1; i < sheetGridData.length; i++) {
			if (sheetGridData[i].length !== columnCount) throw new Error('Column count does not match per row');
		}
		this.rows = sheetGridData.length || 1;
		this.columns = sheetGridData[0].length || 1;
		this.data = sheetGridData;
	}
	static parse(sheetObj) {
		let parsedSheet = new Sheet();
		parsedSheet.sheetId = sheetObj.properties.sheetId;
		parsedSheet.sheetTitle = sheetObj.properties.title;
		parsedSheet.rows = sheetObj.properties.gridProperties.rowCount;
		parsedSheet.columns = sheetObj.properties.gridProperties.columnCount;
		parsedSheet.data = [];
		if (sheetObj.data && sheetObj.data[0].rowData) {
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
		if (this.sheetId) {
			obj.properties.sheetId = this.sheetId;
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
		let obj = {
			title: this.worksheetTitle,
			sheets: {},
			metadata: {},
			id: this.worksheetId
		}
		for (let sheetTitle in this.sheets) {
			obj.sheets[sheetTitle] = this.sheets[sheetTitle].data;
		}
		for (let sheetTitle in this.metadata) {
			if (!this.metadata[sheetTitle] || empty(this.metadata[sheetTitle].data)) continue;
			obj.metadata[sheetTitle] = {};
			for (let metaKey in this.metadata[sheetTitle].data) {
				obj.metadata[sheetTitle][metaKey] = this.metadata[sheetTitle].data[metaKey].value;
			}
		}
		return obj;
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
			generatedWorksheet.metadata[sheetTitle] = {};
			for (let metaTitle in metadata[sheetTitle]) {
				generatedWorksheet.metadata[sheetTitle] = new Metadata(null, metadata[sheetTitle]);
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

class gAPI {
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

class Drive extends gAPI {
	constructor(client_id, token) {
		super(client_id, token);
		this.drive = google.drive({version: "v3", auth: this.oauth}).files;
	}
	getFile(fileId) {
		return new Promise((resolve, reject)=>{
			this.drive.get({fileId: fileId, fields: '*'}, (err, res)=>{
				if (err) {
					if (err.errors[0].reason === 'notFound') {
						console.warn('\x1b[33m%s\x1b[0m', `File: '${fileId}' Not Found`);
						resolve(null);
					} else reject(err);
				}
				else resolve(res.data);
			});
		});
	}
	setFile(fileId, properties) {
		let request = {
			fileId: fileId, 
			resource: properties,
			fields: '*'
		};
		return new Promise((resolve, reject)=>{
			this.drive.update(request, (err, res)=>{
				if (err) {
					if (err.errors[0].reason === 'notFound') {
						console.warn('\x1b[33m%s\x1b[0m', `File: '${fileId}' Not Found`);
						resolve(null);
					} else reject(err);
				}
				else resolve(res.data);
			});
		});		
	}
	async deleteFile(fileId) {
		return await new Promise((resolve, reject)=>{
			this.drive.delete({fileId: fileId}, (err, res)=>{
				if (err) {
					if (err.errors[0].reason === 'notFound') {
						console.warn('\x1b[33m%s\x1b[0m', `File: '${fileId}' Not Found`);
						resolve(false);
					} else reject(err);
				}
				else resolve(true);
			});
		});		
	}
}

class Spreadsheets extends gAPI {
	constructor(client_id, token) {
		super(client_id, token);
		this.spreadsheets = google.sheets({version: "v4", auth: this.oauth}).spreadsheets;
	}
	createSpreadsheet(title, gridData, metadata) {
		return new Promise((resolve, reject)=>{
			this.spreadsheets.create({resource: Worksheet.create(title, gridData, metadata).toSpreadsheetObject()}, (err, res)=>{
				if (err) reject(err);
				else resolve(new Worksheet(res.data).simplify());
			});
		});
	}
	getRawSpreadsheet(spreadsheetId) {
		return new Promise((resolve, reject)=>{
			this.spreadsheets.get({spreadsheetId: spreadsheetId, includeGridData: false}, (err, res)=>{
				if (err) reject(err);
				else resolve(new Worksheet(res.data));
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
					if (err) reject(err);
					else resolve(new Worksheet(res.data).simplify());
				});
			});
		} else {
			return await new Promise((resolve, reject)=>{
				this.spreadsheets.get({
					spreadsheetId: spreadsheetId,
					includeGridData: true
				}, (err, res)=>{
					if (err) reject(err);
					else resolve(new Worksheet(res.data).simplify());
				});
			});
		}
	}
	async setSpreadsheet(spreadsheetId, gridData, metadata){
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
			if (newSheet.data === null) {
				if (newSheet.sheetId) {
					requests.push({
						deleteSheet: {
							sheetId: newSheet.sheetId
						}
					});
				}
				continue;
			}
			if(newMeta) newMeta.fillEmpty(preMeta);
			if (newSheet && !preSheet) {
				let tempId = preSpreadsheet.createSheetId();
				newSheet.sheetId = tempId;			
				requests.push({
					addSheet: {
						properties: newSheet.toSheetObject().properties
					}
				});				
			}
			if (newMeta) {
				for (let metaKey in newMeta.data) {
					if (newMeta.data[metaKey].value === undefined && newMeta.data[metaKey].id) {
						requests.push({
							deleteDeveloperMetadata: {
								dataFilter: {
									developerMetadataLookup: {
										metadataId: newMeta.data[metaKey].id
									}
								}
							}
						})
					}
					else if (!preMeta.data[metaKey]) {
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
									developerMetadataLookup: {
										metadataId: newMeta.data[metaKey].id
									}
								}],
								developerMetadata: newMeta.data[metaKey].toSingleMetadataObject()
							}
						});						
					}
				}
			}
			if ((newSheet && preSheet) && (newSheet.rows !== preSheet.rows || newSheet.columns !== preSheet.columns)) {
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
		}
		if (requests.length === 0) return null;
		return await new Promise((resolve, reject)=>{
			this.spreadsheets.batchUpdate({
				spreadsheetId: spreadsheetId, 
				resource: {
					requests: requests, 
					includeSpreadsheetInResponse: true, 
					responseIncludeGridData: true
				}
			}, (err, res)=>{
				if (err) reject(err);
				else resolve(new Worksheet(res.data.updatedSpreadsheet).simplify());
			});
		});
		return requests;
	}
}
module.exports = {
	Drive: Drive,
	Spreadsheets: Spreadsheets
};