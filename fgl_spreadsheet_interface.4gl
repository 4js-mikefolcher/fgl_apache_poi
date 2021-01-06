IMPORT util

IMPORT FGL fgl_spreadsheet_helper

PUBLIC TYPE ISpreadsheet INTERFACE
	getRecordDefinition() RETURNS DYNAMIC ARRAY OF TFields,
	getJSONObject() RETURNS util.JSONObject,
	getColumnTitles() RETURNS DYNAMIC ARRAY OF STRING
END INTERFACE