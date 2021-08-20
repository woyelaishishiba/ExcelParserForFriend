package com.helper.excel

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.*


object ExcelUtils {

    fun parseExcel() {
        val filePath = "E:\\excel\\source_data.xlsx"
        val book = readExcel(filePath) ?: return
        if (book.numberOfSheets != 3) throw IllegalStateException("表单为什么不是3个呢")
        val map1 = parseKeyValue(book.getSheetAt(0))
        val map2 = parseProperty(book.getSheetAt(1))
        val finalMap = merge(map1, map2)
        println(finalMap.size)
        writeSheet(finalMap, book.getSheetAt(2), book, filePath)
        println("完成")
    }

    private fun writeSheet(map: Map<String, NodeInfo>, sheet: Sheet, book: Workbook, filePath: String) {
        val firstRow = sheet.getRow(0)
        for (i in 1 until sheet.physicalNumberOfRows) {
            val row = sheet.getRow(i)
            val which = getCellFormatValue(row.getCell(0))
            for (j in 1 until firstRow.physicalNumberOfCells) {
                val key = getCellFormatValue(firstRow.getCell(j))
                val value = map[which]?.map?.get(key) ?: continue
                row.createCell(j, value.cellType)?.let { putCellValue(value, it) }
            }
        }
        var output: FileOutputStream? = null
        try {
            output = FileOutputStream(filePath)
            book.write(output)
        } finally {
            output?.close()
        }
    }

    private fun putCellValue(value: Cell, target: Cell) {
        when (value.cellType) {
            CellType.NUMERIC -> target.setCellValue(value.numericCellValue)
            CellType.STRING -> target.setCellValue(value.richStringCellValue)
            CellType.FORMULA -> target.setCellValue(value.cellFormula)
            CellType.BOOLEAN -> target.setCellValue(value.booleanCellValue)
            else -> Unit
        }
    }

    private fun parseProperty(sheet: Sheet): Map<String, NodeInfo> {
        val result = HashMap<String, NodeInfo>()
        if (sheet.getRow(0).physicalNumberOfCells != 3) throw IllegalStateException("输入为什么不是三列呢")
        for (i in 1 until sheet.physicalNumberOfRows) {
            val row = sheet.getRow(i)
            val key = getCellFormatValue(row.getCell(0))
            result.putIfAbsent(key, NodeInfo())
            result[key]?.addEntry(getCellFormatValue(row.getCell(1)), row.getCell(2))
        }
        return result
    }

    private fun parseKeyValue(sheet: Sheet): Map<String, NodeInfo> {
        val result = HashMap<String, NodeInfo>()
        val firstRow = sheet.getRow(0)
        for (i in 1 until sheet.physicalNumberOfRows) {
            val row = sheet.getRow(i)
            val which = getCellFormatValue(row.getCell(0))
            result.putIfAbsent(which, NodeInfo())
            for (j in 1 until row.physicalNumberOfCells) {
                val key = getCellFormatValue(firstRow.getCell(j))
                result[which]?.addEntry(key, row.getCell(j))
            }
        }
        return result
    }

    private fun readExcel(filePath: String): Workbook? {
        val extString = filePath.substring(filePath.lastIndexOf("."))
        var inputStream: InputStream? = null
        try {
            inputStream = FileInputStream(filePath)
            return when (extString) {
                ".xls" -> HSSFWorkbook(inputStream)
                ".xlsx" -> XSSFWorkbook(inputStream)
                else -> return null
            }
        } catch (e: FileNotFoundException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        } finally {
            inputStream?.close()
        }
        return null
    }

    private fun getCellFormatValue(cell: Cell?): String {
        cell ?: return ""
        return when (cell.cellType) {
            CellType.NUMERIC -> cell.numericCellValue.toString()
            CellType.FORMULA -> if (DateUtil.isCellDateFormatted(cell)) cell.dateCellValue.toString() else cell.numericCellValue.toString()
            CellType.STRING -> cell.richStringCellValue.string
            else -> ""
        }
    }
}
