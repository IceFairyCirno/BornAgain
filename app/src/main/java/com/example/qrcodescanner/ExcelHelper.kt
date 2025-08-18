package com.example.qrcodescanner

import android.content.Context
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

class ExcelHelper {
    fun createExcel(context: Context, fileName: String): File? {
        val file = File(context.filesDir, fileName)
        if (file.exists()) {
            return null
        }
        val workbook = XSSFWorkbook()
        workbook.createSheet("Sheet1")
        try {
            FileOutputStream(file).use { fileOut ->
                workbook.write(fileOut)
            }
            println("Excel file created at: ${file.absolutePath}")
            return file
        } catch (e: IOException) {
            println("Error creating Excel file: ${e.message}")
            return null
        } finally {
            workbook.close()
        }
    }
    fun modifyExcel(context: Context, fileName: String, rowData: List<String>) {
        val file = File(context.filesDir, fileName)
        if (!file.exists()) {
            println("File does not exist: ${file.absolutePath}")
            return
        }
        val workbook = try {
            FileInputStream(file).use { WorkbookFactory.create(it) }
        } catch (e: Exception) {
            println("Error: Excel file is corrupted or unreadable: ${e.message}")
            return
        }
        val tempFile = File(context.filesDir, "$fileName.temp")
        try {
            val sheet = workbook.getSheetAt(0)
            val row = sheet.createRow(sheet.lastRowNum + 1)
            rowData.forEachIndexed { index, data ->
                row.createCell(index).setCellValue(data)
            }
            FileOutputStream(tempFile).use { fileOut ->
                workbook.write(fileOut)
            }
            workbook.close()
            if (tempFile.exists() && file.delete()) {
                tempFile.renameTo(file)
                println("Row appended to Excel file: ${file.absolutePath}")
            } else {
                println("Error: Failed to replace original file")
            }
        } catch (e: IOException) {
            println("Error writing to Excel file: ${e.message}")
        } catch (e: Exception) {
            println("Unexpected error during modification: ${e.message}")
        } finally {
            if (tempFile.exists()) {
                tempFile.delete()
            }
        }
    }
    fun modifyExcelCells(context: Context, cellRef1: String, newValue1: String, cellRef2: String, newValue2: String){
        try {
            val ref1 = CellReference(cellRef1)
            val ref2 = CellReference(cellRef2)
            val row1 = ref1.row
            val col1 = ref1.col.toInt()
            val row2 = ref2.row
            val col2 = ref2.col.toInt()

            val file = File(context.filesDir, "record.xlsx")
            if (!file.exists()) {
                throw Exception("record.xlsx not found in internal storage")
            }

            FileInputStream(file).use { inputStream ->
                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheet("Sheet3")
                    ?: throw Exception("Sheet3 not found in record.xlsx")

                val cell1 = sheet.getRow(row1)?.getCell(col1, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                val cell2 = sheet.getRow(row2)?.getCell(col2, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                cell1?.setCellValue(newValue1)
                cell2?.setCellValue(newValue2)

                FileOutputStream(file).use { outputStream ->
                    workbook.write(outputStream)
                }

                workbook.close()

            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }
    fun getExcelCells(context: Context, cellRef1: String, cellRef2: String): Pair<String, String> {
        try {
            val ref1 = CellReference(cellRef1)
            val ref2 = CellReference(cellRef2)
            val row1 = ref1.row
            val col1 = ref1.col.toInt()
            val row2 = ref2.row
            val col2 = ref2.col.toInt()

            // Access record.xlsx
            val file = File(context.filesDir, "record.xlsx")
            if (!file.exists()) {
                throw Exception("record.xlsx not found in internal storage")
            }

            // Read the Excel file
            FileInputStream(file).use { inputStream ->
                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheet("Sheet3")
                    ?: throw Exception("Sheet3 not found in record.xlsx")

                // Read cell values
                val cell1 = sheet.getRow(row1)?.getCell(col1, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                val cell2 = sheet.getRow(row2)?.getCell(col2, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                val value1 = cell1.toString()
                val value2 = cell2.toString()

                workbook.close()
                return Pair(value1, value2)
            }
        } catch (e: Exception) {
            e.printStackTrace()
            return Pair("null", "null")
        }
    }
}