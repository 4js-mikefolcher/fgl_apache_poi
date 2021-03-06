IMPORT JAVA java.io.FileOutputStream
IMPORT JAVA java.io.FileInputStream


IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFWorkbook
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFSheet
IMPORT JAVA org.apache.poi.ss.usermodel.Row
IMPORT JAVA org.apache.poi.ss.usermodel.Cell
IMPORT JAVA org.apache.poi.ss.usermodel.CellStyle
IMPORT JAVA org.apache.poi.ss.usermodel.HorizontalAlignment
IMPORT JAVA org.apache.poi.ss.usermodel.Font
IMPORT JAVA org.apache.poi.ss.usermodel.PrintSetup
IMPORT JAVA org.apache.poi.ss.usermodel.DataFormat
IMPORT JAVA java.util.Date
IMPORT JAVA java.util.Calendar
IMPORT JAVA java.text.SimpleDateFormat
IMPORT JAVA java.time.LocalDateTime
IMPORT JAVA java.time.ZoneId




PUBLIC TYPE workbookType XSSFWorkbook
PUBLIC TYPE sheetType XSSFSheet
PUBLIC TYPE rowType Row
PUBLIC TYPE cellType Cell
PUBLIC TYPE cellStyleType CellStyle
PUBLIC TYPE fontType Font
PUBLIC TYPE cellFormat DataFormat


#Constants the corralate to the Built-in Formats
PUBLIC CONSTANT cDecimalFormat = 40
PUBLIC CONSTANT cMoneyFormat = 8
PUBLIC CONSTANT cIntegerFormat = 38
PUBLIC CONSTANT cDateFormat = 14
PUBLIC CONSTANT cDatetimeFormat = 22


FUNCTION workbook_create()
    RETURN XSSFWorkbook.create()
END FUNCTION



FUNCTION workbook_writeToFile(w, filename)
DEFINE w workbookType
DEFINE filename STRING
DEFINE fo FileOutputStream

    LET fo = FileOutputStream.create(filename)
    CALL w.write(fo)
    CALL fo.close()
END FUNCTION


FUNCTION workbook_open(filename)
DEFINE filename STRING
DEFINE fi FileInputStream
DEFINE w workbookType

    LET fi = FileInputStream.create(filename)
    LET w = XSSFWorkbook.create(fi)
    RETURN w
END FUNCTION



FUNCTION workbook_createsheet(w)
DEFINE w workbookType
DEFINE s sheetType

    LET s= w.createSheet()
    RETURN s
END FUNCTION



FUNCTION sheet_createrow(s,idx)
DEFINE s sheetType
DEFINE idx INTEGER
DEFINE r rowType
    LET r = s.createRow(idx)
    RETURN r
END FUNCTION



FUNCTION sheet_autosizecolumn(s, c)
DEFINE s sheetType
DEFINE c INTEGER

    CALL s.autoSizeColumn(c)
END FUNCTION

FUNCTION sheet_columnwidth_set(s, c, wf)
DEFINE s sheetType
DEFINE c INTEGER
DEFINE wf FLOAT
DEFINE wi INTEGER

    LET wi = 256.0*wf
    CALL s.setColumnWidth(c,wi)
END FUNCTION



FUNCTION row_createcell(r,idx)
DEFINE r rowType
DEFINE idx INTEGER
DEFINE c cellType
    LET c = r.createCell(idx)
    RETURN c
END FUNCTION




FUNCTION cell_value_set(c, v)
DEFINE c cellType
DEFINE v STRING
    CALL c.setCellValue(v)
END FUNCTION



FUNCTION cell_number_set(c, v)
DEFINE c cellType
DEFINE v FLOAT
  
    CALL c.setCellValue(v)
END FUNCTION


FUNCTION cell_date_set(c, v)
DEFINE c cellType
DEFINE v STRING
DEFINE javaDate java.util.Date
DEFINE dateFormatter SimpleDateFormat

	LET dateFormatter = java.text.SimpleDateFormat.create("yyyy-MM-dd")
	LET javaDate = dateFormatter.parse(v)
	CALL c.setCellValue(javaDate)
END FUNCTION

FUNCTION cell_datetime_set(c, v)
DEFINE c cellType
DEFINE v STRING
DEFINE javaDatetime Calendar
DEFINE dateFormatter SimpleDateFormat

	LET dateFormatter = java.text.SimpleDateFormat.create("yyyy-MM-dd HH:mm:ss")
	LET javaDatetime = Calendar.getInstance()
	CALL javaDatetime.setTime(dateFormatter.parse(v))
	CALL c.setCellValue(LocalDateTime.ofInstant(javaDatetime.toInstant(), ZoneId.systemDefault()))

END FUNCTION

FUNCTION cell_formula_set(c, v)
DEFINE c cellType
DEFINE v STRING
    CALL c.setCellFormula(v)
END FUNCTION



-- Map A to 0 B to 1, Z to 25, AA to 26, AZ to 51
FUNCTION column2row(col)
DEFINE col STRING

    CASE
        WHEN col MATCHES "[A-Z]" 
            RETURN ORD(col) - 65
        WHEN col MATCHES "[A-Z][A-Z]"
            RETURN ((ORD(col.subString(1,1))-65)*26) + (ORD(col.subString(2,2)) - 65)
    END CASE
    RETURN -1
END FUNCTION
    



FUNCTION cell_style_set(c, s)
DEFINE c cellType
DEFINE s cellStyleType
    CALL c.setCellStyle(s)
END FUNCTION

FUNCTION cell_style_format_create(w workBookType, fmt STRING) RETURNS (cellStyleType)
DEFINE c cellFormat
DEFINE s cellStyleType

	LET s = w.createCellStyle()
	LET c = w.createDataFormat()
	CALL s.setDataFormat(c.getFormat(fmt))

	RETURN s

END FUNCTION

FUNCTION cell_style_builtin_create(w workBookType, fmt SMALLINT) RETURNS (cellStyleType)
DEFINE s cellStyleType

	LET s = w.createCellStyle()
	CALL s.setDataFormat(fmt)

	RETURN s
END FUNCTION

FUNCTION font_create(w)
DEFINE w workBookType
DEFINE f fontType
    LET f = w.createFont()
    RETURN f
END FUNCTION

FUNCTION font_set(f, a, v)
DEFINE f fontType
DEFINE a STRING
DEFINE v STRING

    CASE 
        WHEN a="weight" AND v="bold"
            CALL f.setBold(true)
        WHEN a="weight" AND v="normal"
            CALL f.setBold(false)
        -- add more as required
    END CASE
END FUNCTION



FUNCTION style_create(w)
DEFINE w workBookType
DEFINE s cellStyleType
    LET s = w.createCellStyle()
    RETURN s
END FUNCTION



FUNCTION style_set(s, a, v)
DEFINE s cellStyleType
DEFINE a STRING
DEFINE v STRING

    CASE 
        WHEN a="alignment" AND v="center"
            CALL s.setAlignment(HorizontalAlignment.CENTER)
            
        WHEN a="alignment" AND v="left"
            CALL s.setAlignment(HorizontalAlignment.LEFT)

        WHEN a="alignment" AND v="right"
            CALL s.setAlignment(HorizontalAlignment.RIGHT)

        WHEN a="alignment" AND v="justify"
            CALL s.setAlignment(HorizontalAlignment.JUSTIFY)

        WHEN a="alignment" AND v="general"
            CALL s.setAlignment(HorizontalAlignment.GENERAL)

        -- add more as required
    END CASE
END FUNCTION



FUNCTION style_font_set(s,f)
DEFINE s cellStyleType
DEFINE f fontType
    CALL s.setFont(f)
END FUNCTION


FUNCTION workbook_fit_to_page(old_filename STRING, new_filename STRING)
DEFINE w workbookType
DEFINE s sheetType

DEFINE ps PrintSetup
CONSTANT SHORT_ONE SMALLINT = 1

    LET w = workbook_open(old_filename)
    LET s = w.getSheetAt(0)
    CALL s.setFitToPage(true)
    CALL s.setAutobreaks(true)
    LET ps = s.getPrintSetup()
    CALL ps.setFitWidth(SHORT_ONE)
    CALL ps.setFitHeight(SHORT_ONE)

    CALL workbook_writeToFile(w, new_filename)
END FUNCTION