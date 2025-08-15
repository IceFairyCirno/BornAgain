package com.example.qrcodescanner

import android.content.Context
import org.apache.poi.ss.usermodel.WorkbookFactory
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
    fun searchExcel(context: Context, fileName: String, columnIndex: Int, searchValue: String): List<List<String>> {
        val file = File(context.filesDir, fileName)
        if (!file.exists()) {
            println("File does not exist: ${file.absolutePath}")
            return emptyList()
        }

        val workbook = try {
            WorkbookFactory.create(file)
        } catch (e: IOException) {
            println("Error opening Excel file: ${e.message}")
            return emptyList()
        }

        val sheet = workbook.getSheetAt(0)
        val result = mutableListOf<List<String>>()
        try {
            for (row in sheet) {
                val cell = row.getCell(columnIndex, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                if (cell.toString() == searchValue) {
                    val rowData = mutableListOf<String>()
                    row.forEach { rowData.add(it.toString()) }
                    result.add(rowData)
                }
            }
            println("Search completed in: ${file.absolutePath}")
            return result
        } catch (e: Exception) {
            println("Error searching Excel file: ${e.message}")
            return emptyList()
        } finally {
            workbook.close()
        }
    }
}