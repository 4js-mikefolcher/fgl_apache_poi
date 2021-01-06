IMPORT util
IMPORT FGL fgl_spreadsheet_api

TYPE TData RECORD
	stringField STRING,
	moneyField MONEY(12,2),
	decimalField DECIMAL(8,4),
	integerField INTEGER,
	smallintField SMALLINT,
	dateField DATE,
	datetimeField DATETIME YEAR TO SECOND,
	charField CHAR(35),
	varcharField VARCHAR(100),
	floatField FLOAT,
	smallfloatField SMALLFLOAT,
	booleanField BOOLEAN
END RECORD

MAIN
	DEFINE dataList DYNAMIC ARRAY OF TData
	DEFINE dataRec TData
	DEFINE idx INTEGER
	DEFINE excelHandler fgl_spreadsheet_api.TSpreadsheet

	CALL STARTLOG("fgl_excel_api_test.log")

	FOR idx = 1 TO 100
		LET dataRec.stringField = SFMT("Row #%1", idx)
		LET dataRec.charField = SFMT("%1 x %1 y %1 z %1", idx)
		LET dataRec.varcharField = SFMT("'%1' '%1' '%1' '%1'", ASCII(idx + 32))
		LET dataRec.booleanField = idx MOD 2
		LET dataRec.dateField = TODAY - idx UNITS DAY
		LET dataRec.datetimeField = CURRENT YEAR TO SECOND
		IF dataRec.booleanField THEN
			#odd
			LET dataRec.smallintField = -1 * idx
			LET dataRec.integerField = -2 * idx
			LET dataRec.floatField = -7.17 * idx
			LET dataRec.smallfloatField = -1.1 * idx
			LET dataRec.moneyField = -11.69 * idx
			LET dataRec.decimalField = -13.3954 * idx
		ELSE
			#even
			LET dataRec.smallintField = idx
			LET dataRec.integerField = 2 * idx
			LET dataRec.floatField = 7.153 * idx
			LET dataRec.smallfloatField = 1.1 * idx
			LET dataRec.moneyField = 11.9 * idx
			LET dataRec.decimalField = 13.3954 * idx
		END IF
		LET dataList[idx] = dataRec
	END FOR

	#Steps to create an excel spreadsheet from a array of records
	CALL excelHandler.init()
	CALL excelHandler.setHeaders(excelHeader())
	CALL excelHandler.setRecordDefinition(base.TypeInfo.create(dataRec))
	CALL excelHandler.setTitle("Test Spreadsheet API")
	IF excelHandler.createSpreadsheet(util.JSONArray.fromFGL(dataList)) THEN
		DISPLAY SFMT("Excel file path: %1", excelHandler.getFilename())
	END IF

END MAIN

PRIVATE FUNCTION excelHeader() RETURNS DYNAMIC ARRAY OF STRING
	DEFINE headerList DYNAMIC ARRAY OF STRING = [
		"String",
		"Money",
		"Decimal",
		"Integer",
		"Small Integer",
		"Date",
		"Datetime",
		"Char",
		"Varchar",
		"Float",
		"Small Float",
		"Boolean"
	]

	RETURN headerList

END FUNCTION