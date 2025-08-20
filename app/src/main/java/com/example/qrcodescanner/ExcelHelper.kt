package com.example.qrcodescanner

import android.content.Context
import android.util.Log
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.text.SimpleDateFormat
import java.util.Locale

class ExcelHelper(private val context: Context) {
    companion object {
        private const val TAG = "ExcelHelper"
    }

    fun initExcel(filename: String) {
        try {
            val file = File(context.filesDir, filename)
            if (file.exists()) {
                Log.i(TAG, "Excel file already exists: $filename, skipping creation")
                return
            }

            Log.d(TAG, "Initializing Excel file: $filename")
            val workbook = XSSFWorkbook()
            workbook.createSheet("record")
            workbook.createSheet("bodydata")
            workbook.createSheet("settings")

            FileOutputStream(file).use { outputStream ->
                workbook.write(outputStream)
            }
            workbook.close()
            Log.i(TAG, "Successfully created Excel file: $filename with sheets: record, bodydata, settings")
        } catch (e: Exception) {
            Log.e(TAG, "Failed to create Excel file: $filename, error: ${e.message}", e)
            throw RuntimeException("Failed to create Excel file: ${e.message}")
        }
    }

    fun deleteExcel(filename: String): Boolean {
        try {
            Log.d(TAG, "Attempting to delete Excel file: $filename")
            val file = File(context.filesDir, filename)
            return if (file.exists()) {
                val deleted = file.delete()
                if (deleted) {
                    Log.i(TAG, "Successfully deleted Excel file: $filename")
                } else {
                    Log.w(TAG, "Failed to delete Excel file: $filename (exists but deletion failed)")
                }
                deleted
            } else {
                Log.w(TAG, "Excel file does not exist: $filename")
                false
            }
        } catch (e: Exception) {
            Log.e(TAG, "Failed to delete Excel file: $filename, error: ${e.message}", e)
            throw RuntimeException("Failed to delete Excel file: ${e.message}")
        }
    }

    fun searchFromBottomExcel(filename: String, sheetName: String, col: Int, colValue: String, n: Int): List<List<String>> {
        try {
            Log.d(TAG, "Searching from bottom in $filename, sheet: $sheetName, column: $col, value: $colValue, limit: $n")
            val file = File(context.filesDir, filename)
            if (!file.exists()) {
                Log.w(TAG, "Excel file does not exist: $filename")
                return emptyList()
            }

            FileInputStream(file).use { inputStream ->
                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    Log.w(TAG, "Sheet does not exist: $sheetName in $filename")
                    workbook.close()
                    return emptyList()
                }

                val result = mutableListOf<List<String>>()
                val lastRow = sheet.lastRowNum

                for (rowIndex in lastRow downTo 0) {
                    if (result.size >= n) break
                    val row = sheet.getRow(rowIndex) ?: continue
                    val cell = row.getCell(col - 1) ?: continue
                    if (cell.toString() == colValue) {
                        val rowData = mutableListOf<String>()
                        row.forEach { cell ->
                            rowData.add(cell.toString())
                        }
                        result.add(rowData)
                        Log.d(TAG, "Found matching row at index $rowIndex: $rowData")
                    }
                }

                workbook.close()
                Log.i(TAG, "Search completed, found ${result.size} matching rows")
                return result
            }
        } catch (e: Exception) {
            Log.e(TAG, "Failed to search Excel file: $filename, sheet: $sheetName, error: ${e.message}", e)
            throw RuntimeException("Failed to search Excel file: ${e.message}")
        }
    }

    fun searchUniqueFromBottomExcel(filename: String, sheetName: String, col: Int, n: Int): List<List<String>> {
        try {
            Log.d(TAG, "Searching unique values from bottom in $filename, sheet: $sheetName, column: $col, limit: $n")
            val file = File(context.filesDir, filename)
            if (!file.exists()) {
                Log.w(TAG, "Excel file does not exist: $filename")
                return emptyList()
            }

            FileInputStream(file).use { inputStream ->
                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    Log.w(TAG, "Sheet does not exist: $sheetName in $filename")
                    workbook.close()
                    return emptyList()
                }

                val result = mutableListOf<List<String>>()
                val seenValues = mutableSetOf<String>()
                val lastRow = sheet.lastRowNum

                for (rowIndex in lastRow downTo 0) {
                    if (result.size >= n) break
                    val row = sheet.getRow(rowIndex) ?: continue
                    val cell = row.getCell(col - 1) ?: continue
                    val cellValue = cell.toString()
                    if (cellValue.isNotEmpty() && seenValues.add(cellValue)) {
                        val rowData = mutableListOf<String>()
                        row.forEach { cell ->
                            rowData.add(cell.toString())
                        }
                        result.add(rowData)
                        Log.d(TAG, "Found unique row at index $rowIndex with value $cellValue: $rowData")
                    }
                }

                workbook.close()
                Log.i(TAG, "Unique search completed, found ${result.size} unique rows")
                return result
            }
        } catch (e: Exception) {
            Log.e(TAG, "Failed to search unique values in $filename, sheet: $sheetName, error: ${e.message}", e)
            throw RuntimeException("Failed to search unique values in Excel file: ${e.message}")
        }
    }

    fun modifyExcelFromBottom(filename: String, sheetName: String, data: List<String>) {
        try {
            Log.d(TAG, "Appending row to $filename, sheet: $sheetName, data: $data")
            val file = File(context.filesDir, filename)
            val workbook = if (file.exists()) {
                FileInputStream(file).use { inputStream ->
                    XSSFWorkbook(inputStream)
                }
            } else {
                Log.w(TAG, "Excel file does not exist, creating new: $filename")
                XSSFWorkbook()
            }

            val sheet = workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName)
            val newRow = sheet.createRow(sheet.lastRowNum + 1)

            data.forEachIndexed { index, value ->
                val cell = newRow.createCell(index)
                cell.setCellValue(value)
            }

            FileOutputStream(file).use { outputStream ->
                workbook.write(outputStream)
            }
            workbook.close()
            Log.i(TAG, "Successfully appended row to $filename, sheet: $sheetName")
        } catch (e: Exception) {
            Log.e(TAG, "Failed to append row to $filename, sheet: $sheetName, error: ${e.message}", e)
            throw RuntimeException("Failed to modify Excel file: ${e.message}")
        }
    }

    fun getCellExcel(filename: String, sheetName: String, cell: String): String {
        try {
            Log.d(TAG, "Reading cell $cell from $filename, sheet: $sheetName")
            val file = File(context.filesDir, filename)
            if (!file.exists()) {
                Log.w(TAG, "Excel file does not exist: $filename")
                return ""
            }

            FileInputStream(file).use { inputStream ->
                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    Log.w(TAG, "Sheet does not exist: $sheetName in $filename")
                    workbook.close()
                    return ""
                }

                val cellReference = CellReference(cell)
                val row = sheet.getRow(cellReference.row)
                val cellValue = row?.getCell(cellReference.col.toInt())?.toString() ?: ""

                workbook.close()
                Log.d(TAG, "Cell $cell value: $cellValue")
                return cellValue
            }
        } catch (e: Exception) {
            Log.e(TAG, "Failed to read cell $cell from $filename, sheet: $sheetName, error: ${e.message}", e)
            throw RuntimeException("Failed to read cell from Excel file: ${e.message}")
        }
    }

    fun modifyCellExcel(filename: String, sheetName: String, cell: String, value: String) {
        try {
            Log.d(TAG, "Modifying cell $cell in $filename, sheet: $sheetName, new value: $value")
            val file = File(context.filesDir, filename)
            val workbook = if (file.exists()) {
                FileInputStream(file).use { inputStream ->
                    XSSFWorkbook(inputStream)
                }
            } else {
                Log.w(TAG, "Excel file does not exist, creating new: $filename")
                XSSFWorkbook()
            }

            val sheet = workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName)
            val cellReference = CellReference(cell)
            val row = sheet.getRow(cellReference.row) ?: sheet.createRow(cellReference.row)
            val targetCell = row.createCell(cellReference.col.toInt())
            targetCell.setCellValue(value)

            FileOutputStream(file).use { outputStream ->
                workbook.write(outputStream)
            }
            workbook.close()
            Log.i(TAG, "Successfully modified cell $cell in $filename, sheet: $sheetName")
        } catch (e: Exception) {
            Log.e(TAG, "Failed to modify cell $cell in $filename, sheet: $sheetName, error: ${e.message}", e)
            throw RuntimeException("Failed to modify cell in Excel file: ${e.message}")
        }
    }
    fun copyExcel(srcfile: String, srcsheet: String, targetfile: String, targetsheet: String) {
        try {
            Log.d(TAG, "Copying sheet $srcsheet from $srcfile to $targetsheet in $targetfile")
            val srcFile = File(context.filesDir, srcfile)
            if (!srcFile.exists()) {
                Log.w(TAG, "Source Excel file does not exist: $srcfile")
                throw RuntimeException("Source Excel file does not exist: $srcfile")
            }

            val targetFile = File(context.filesDir, targetfile)
            val targetWorkbook = if (targetFile.exists()) {
                FileInputStream(targetFile).use { inputStream ->
                    XSSFWorkbook(inputStream)
                }
            } else {
                Log.w(TAG, "Target Excel file does not exist, creating new: $targetfile")
                XSSFWorkbook()
            }

            FileInputStream(srcFile).use { inputStream ->
                val srcWorkbook = XSSFWorkbook(inputStream)
                val srcSheet = srcWorkbook.getSheet(srcsheet)
                if (srcSheet == null) {
                    Log.w(TAG, "Source sheet does not exist: $srcsheet in $srcfile")
                    srcWorkbook.close()
                    throw RuntimeException("Source sheet does not exist: $srcsheet")
                }

                val targetSheet = targetWorkbook.getSheet(targetsheet) ?: targetWorkbook.createSheet(targetsheet)
                val lastRowNum = srcSheet.lastRowNum
                val startRow = if (targetSheet.lastRowNum >= 0) targetSheet.lastRowNum + 1 else 0

                for (i in 0..lastRowNum) {
                    val srcRow = srcSheet.getRow(i)
                    if (srcRow != null) {
                        val newRow = targetSheet.createRow(startRow + i)
                        for (j in 0 until srcRow.lastCellNum) {
                            val srcCell = srcRow.getCell(j)
                            if (srcCell != null) {
                                val newCell = newRow.createCell(j)
                                when (srcCell.cellType) {
                                    CellType.NUMERIC -> newCell.setCellValue(srcCell.numericCellValue)
                                    CellType.STRING -> newCell.setCellValue(srcCell.stringCellValue)
                                    CellType.BOOLEAN -> newCell.setCellValue(srcCell.booleanCellValue)
                                    else -> newCell.setCellValue(srcCell.toString())
                                }
                            }
                        }
                    }
                }
            }

            FileOutputStream(targetFile).use { outputStream ->
                targetWorkbook.write(outputStream)
            }
            targetWorkbook.close()
            Log.i(TAG, "Successfully appended sheet $srcsheet to $targetsheet in $targetfile")
        } catch (e: Exception) {
            Log.e(TAG, "Failed to copy sheet from $srcfile to $targetfile, error: ${e.message}", e)
            throw RuntimeException("Failed to copy sheet: ${e.message}")
        }
    }

    fun searchFromBottomN(filename: String, sheetName: String, n: Int): List<List<String>> {
        try {
            Log.d(TAG, "Searching last $n rows from bottom in $filename, sheet: $sheetName")
            val file = File(context.filesDir, filename)
            if (!file.exists()) {
                Log.w(TAG, "Excel file does not exist: $filename")
                return emptyList()
            }

            FileInputStream(file).use { inputStream ->
                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    Log.w(TAG, "Sheet does not exist: $sheetName in $filename")
                    workbook.close()
                    return emptyList()
                }

                val result = mutableListOf<List<String>>()
                val lastRow = sheet.lastRowNum
                val rowsToFetch = minOf(n, lastRow + 1) // Ensure n doesn't exceed total rows

                for (rowIndex in lastRow downTo maxOf(0, lastRow - rowsToFetch + 1)) {
                    val row = sheet.getRow(rowIndex) ?: continue
                    val rowData = mutableListOf<String>()
                    row.forEach { cell ->
                        rowData.add(when (cell.cellType) {
                            CellType.NUMERIC -> cell.numericCellValue.toInt().toString()
                            else -> cell.toString()
                        })
                    }
                    // Only add the row if it contains at least one non-empty value
                    if (rowData.any { it.isNotEmpty() }) {
                        result.add(rowData)
                        Log.d(TAG, "Added row at index $rowIndex: $rowData")
                    }
                }

                workbook.close()
                Log.i(TAG, "Search completed, retrieved ${result.size} rows")
                return result.reversed() // Reverse to return rows in top-to-bottom order
            }
        } catch (e: Exception) {
            Log.e(TAG, "Failed to search last $n rows in $filename, sheet: $sheetName, error: ${e.message}", e)
            throw RuntimeException("Failed to search last $n rows: ${e.message}")
        }
    }

    fun processExcelFile(filename: String, sheetName: String): List<List<String>> {
        try {
            // Access the file from internal storage
            Log.d(TAG, "Attempting to access file: $filename")
            val file = File(context.filesDir, filename)
            if (!file.exists()) {
                Log.e(TAG, "File $filename does not exist in internal storage")
                return emptyList()
            }

            // Read the Excel file
            Log.d(TAG, "Reading Excel file: $filename")
            FileInputStream(file).use { inputStream ->
                val workbook = XSSFWorkbook(inputStream)
                Log.d(TAG, "Workbook loaded successfully")
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    Log.e(TAG, "Sheet $sheetName not found in $filename")
                    workbook.close()
                    return emptyList()
                }
                Log.d(TAG, "Sheet $sheetName loaded successfully")

                // Get all rows
                val rows = sheet.iterator().asSequence().toList()
                if (rows.isEmpty()) {
                    Log.w(TAG, "Sheet $sheetName is empty")
                    workbook.close()
                    return emptyList()
                }
                Log.i(TAG, "Found ${rows.size} rows in sheet $sheetName")

                // Map to store unique col2 values and their last row (with col1)
                val uniqueCol2Rows = mutableMapOf<String, Pair<String, Int>>()
                rows.forEachIndexed { index, row ->
                    val col2 = row.getCell(1)?.toString() ?: ""
                    val col1 = row.getCell(0)?.toString() ?: ""
                    if (col2.isNotEmpty() && col1.isNotEmpty()) {
                        uniqueCol2Rows[col2] = Pair(col1, index)
                    }
                }
                Log.d(TAG, "Unique col2 values with col1: ${uniqueCol2Rows.map { "${it.key}(${it.value.first})" }}")

                // Get the last 3 unique col2 rows based on their last row index
                val lastThreeCol2Rows = uniqueCol2Rows.entries
                    .sortedByDescending { it.value.second }
                    .take(3)
                    .map { Pair(it.value.first, it.key) } // Pair(col1, col2)
                Log.i(TAG, "Last 3 unique col2 rows (col1, col2): $lastThreeCol2Rows")

                val result = mutableListOf<List<String>>()

                // Process each (col1, col2) pair from the last 3 unique col2 rows
                for ((col1, col2) in lastThreeCol2Rows) {
                    Log.d(TAG, "Processing col1: $col1, col2: $col2")
                    // Find rows matching both col1 and col2
                    val matchingRows = rows.filter { row ->
                        val rowCol1 = row.getCell(0)?.toString() ?: ""
                        val rowCol2 = row.getCell(1)?.toString() ?: ""
                        rowCol1 == col1 && rowCol2 == col2
                    }
                    Log.d(TAG, "Found ${matchingRows.size} rows for col1: $col1, col2: $col2")

                    // Calculate average of col3 and count rows
                    val col3Values = matchingRows.mapNotNull { row ->
                        val cell = row.getCell(2)
                        when (cell?.cellType) {
                            CellType.NUMERIC -> cell.numericCellValue
                            CellType.STRING -> try {
                                cell.stringCellValue.toDoubleOrNull().also { value ->
                                    if (value == null) {
                                        Log.w(TAG, "Cannot parse string cell as double in row ${row.rowNum}, col3: ${cell.stringCellValue}")
                                    }
                                }
                            } catch (e: NumberFormatException) {
                                Log.w(TAG, "Invalid number format in row ${row.rowNum}, col3: ${cell.stringCellValue}")
                                null
                            }
                            else -> {
                                Log.w(TAG, "Non-numeric cell in row ${row.rowNum}, col3: ${cell?.toString()}")
                                null
                            }
                        }
                    }
                    val avgCol3 = if (col3Values.isNotEmpty()) {
                        String.format("%.2f", col3Values.average())
                    } else {
                        Log.w(TAG, "No valid double col3 values for col1: $col1, col2: $col2")
                        "0.0"
                    }
                    val rowCount = matchingRows.size.toString()
                    Log.d(TAG, "col1: $col1, col2: $col2, avgCol3: $avgCol3, rowCount: $rowCount")

                    // Add [col1, col2, avgCol3, rowCount] to result
                    result.add(listOf(col1, col2, avgCol3, rowCount))
                }

                workbook.close()
                Log.d(TAG, "Workbook closed")
                Log.i(TAG, "Processing complete. Result size: ${result.size}")
                return result
            }
        } catch (e: Exception) {
            Log.e(TAG, "Error processing file $filename: ${e.message}", e)
            return emptyList()
        }
    }

    fun getLastCellWithContentAsString(filename: String, sheetName: String): String {
        val fileInputStream = context.openFileInput(filename)
        val workbook: Workbook = XSSFWorkbook(fileInputStream)
        val sheet: Sheet = workbook.getSheet(sheetName) ?: return "0.0"

        // Iterate from the last row to the first row
        for (rowIndex in sheet.lastRowNum downTo 0) {
            val row: Row? = sheet.getRow(rowIndex)
            if (row != null) {
                // Iterate through the cells of the row from last to first
                for (cellIndex in row.lastCellNum - 1 downTo 0) {
                    val cell: Cell? = row.getCell(cellIndex)
                    if (cell != null && cell.cellType != CellType.BLANK) {
                        // Convert cell content to String
                        return when (cell.cellType) {
                            CellType.STRING -> cell.stringCellValue
                            CellType.NUMERIC -> cell.numericCellValue.toString()
                            CellType.BOOLEAN -> cell.booleanCellValue.toString()
                            CellType.FORMULA -> cell.cellFormula
                            else -> "0.0"
                        }
                    }
                }
            }
        }

        workbook.close()
        fileInputStream.close()
        return "0.0"
    }

    fun getLastNUniqueFirstColumnTimeDifferences(filename: String, sheetName: String, n: Int): List<Double> {
        Log.d(TAG, "Processing file: $filename, sheet: $sheetName, n: $n")
        val differences = mutableListOf<Double>()

        try {
            // Access internal storage
            val file = context.getFileStreamPath(filename)
            if (!file.exists()) {
                Log.e(TAG, "File not found: $filename")
                return differences
            }
            Log.d(TAG, "File found: $filename")

            FileInputStream(file).use { fis ->
                // Open workbook
                val workbook = WorkbookFactory.create(fis)
                val sheet = workbook.getSheet(sheetName)
                if (sheet == null) {
                    Log.e(TAG, "Sheet not found: $sheetName")
                    return differences
                }
                Log.d(TAG, "Sheet found: $sheetName")

                // Handle empty sheet
                if (sheet.lastRowNum < 0) {
                    Log.w(TAG, "Sheet is empty (no rows)")
                    return differences
                }

                // Collect last n unique values in first column
                val uniqueValues = mutableListOf<String>()
                val seenValues = mutableSetOf<String>()
                for (rowNum in sheet.lastRowNum downTo 0) {
                    val row = sheet.getRow(rowNum) ?: continue
                    val firstCell = row.getCell(0) ?: continue
                    val firstValue = when (firstCell.cellType) {
                        CellType.STRING -> firstCell.stringCellValue
                        CellType.NUMERIC -> firstCell.numericCellValue.toString()
                        CellType.BOOLEAN -> firstCell.booleanCellValue.toString()
                        CellType.FORMULA -> firstCell.cachedFormulaResultType.let { cachedType ->
                            when (cachedType) {
                                CellType.STRING -> firstCell.stringCellValue
                                CellType.NUMERIC -> firstCell.numericCellValue.toString()
                                CellType.BOOLEAN -> firstCell.booleanCellValue.toString()
                                else -> firstCell.toString()
                            }
                        }
                        else -> firstCell.toString()
                    }
                    if (firstValue.isNotEmpty() && seenValues.add(firstValue)) {
                        uniqueValues.add(firstValue)
                        Log.d(TAG, "Found unique first column value: '$firstValue' at row $rowNum")
                    }
                    if (uniqueValues.size >= n) break
                }
                Log.d(TAG, "Found ${uniqueValues.size} unique values: $uniqueValues")

                // For each unique value, find first and last row, get last non-blank cell, and calculate time difference
                val timeFormat = SimpleDateFormat("HH:mm", Locale.US).apply { isLenient = false }
                for (uniqueValue in uniqueValues) {
                    var firstRow: Row? = null
                    var lastRow: Row? = null
                    var firstRowNum = -1
                    var lastRowNum = -1

                    // Find first and last row with the unique value in first column
                    for (rowNum in 0..sheet.lastRowNum) {
                        val row = sheet.getRow(rowNum) ?: continue
                        val firstCell = row.getCell(0) ?: continue
                        val firstValue = when (firstCell.cellType) {
                            CellType.STRING -> firstCell.stringCellValue
                            CellType.NUMERIC -> firstCell.numericCellValue.toString()
                            CellType.BOOLEAN -> firstCell.booleanCellValue.toString()
                            CellType.FORMULA -> firstCell.cachedFormulaResultType.let { cachedType ->
                                when (cachedType) {
                                    CellType.STRING -> firstCell.stringCellValue
                                    CellType.NUMERIC -> firstCell.numericCellValue.toString()
                                    CellType.BOOLEAN -> firstCell.booleanCellValue.toString()
                                    else -> firstCell.toString()
                                }
                            }
                            else -> firstCell.toString()
                        }
                        if (firstValue == uniqueValue) {
                            if (firstRow == null) {
                                firstRow = row
                                firstRowNum = rowNum
                            }
                            lastRow = row
                            lastRowNum = rowNum
                        }
                    }

                    if (firstRow == null || lastRow == null) {
                        Log.w(TAG, "Could not find rows for value: '$uniqueValue'")
                        differences.add(0.0)
                        continue
                    }
                    Log.d(TAG, "Value '$uniqueValue': first row $firstRowNum, last row $lastRowNum")

                    // Get last non-blank cell with content for first and last row
                    fun getLastNonBlankCell(row: Row): String? {
                        for (colNum in (row.lastCellNum - 1) downTo 0) {
                            val cell = row.getCell(colNum) ?: continue
                            if (cell.cellType != CellType.BLANK) {
                                val value = when (cell.cellType) {
                                    CellType.STRING -> cell.stringCellValue
                                    CellType.NUMERIC -> if (DateUtil.isCellDateFormatted(cell)) cell.dateCellValue.toString() else cell.numericCellValue.toString()
                                    CellType.BOOLEAN -> cell.booleanCellValue.toString()
                                    CellType.FORMULA -> cell.cachedFormulaResultType.let { cachedType ->
                                        when (cachedType) {
                                            CellType.STRING -> cell.stringCellValue
                                            CellType.NUMERIC -> cell.numericCellValue.toString()
                                            CellType.BOOLEAN -> cell.booleanCellValue.toString()
                                            else -> cell.toString()
                                        }
                                    }
                                    CellType.ERROR -> cell.errorCellValue.toString()
                                    else -> cell.toString()
                                }
                                if (value.isNotEmpty()) {
                                    Log.d(TAG, "Found non-blank cell at column ${colNum} in row ${row.rowNum}: '$value'")
                                    return value
                                }
                                Log.d(TAG, "Skipping empty cell at column $colNum in row ${row.rowNum}")
                            } else {
                                Log.d(TAG, "Skipping blank cell at column $colNum in row ${row.rowNum}")
                            }
                        }
                        Log.w(TAG, "No non-blank cells with content in row ${row.rowNum}")
                        return null
                    }

                    val firstRowLastValue = getLastNonBlankCell(firstRow)
                    val lastRowLastValue = getLastNonBlankCell(lastRow)

                    if (firstRowLastValue == null || lastRowLastValue == null) {
                        Log.w(TAG, "No valid last cell for value '$uniqueValue' (first: $firstRowLastValue, last: $lastRowLastValue)")
                        differences.add(0.0)
                        continue
                    }

                    // Calculate time difference if both values are in HH:mm format
                    try {
                        val firstTime = timeFormat.parse(firstRowLastValue)?.time ?: throw IllegalArgumentException("Invalid first time format")
                        val lastTime = timeFormat.parse(lastRowLastValue)?.time ?: throw IllegalArgumentException("Invalid last time format")
                        val diffMs = lastTime - firstTime
                        val diffHours = diffMs / (1000.0 * 60 * 60)
                        Log.d(TAG, "Time difference for '$uniqueValue': $lastRowLastValue - $firstRowLastValue = $diffHours hours")
                        differences.add(diffHours)
                    } catch (e: Exception) {
                        Log.w(TAG, "Invalid time format for '$uniqueValue' (first: $firstRowLastValue, last: $lastRowLastValue), setting difference to 0.0")
                        differences.add(0.0)
                    }
                }

                Log.d(TAG, "Returning time differences: $differences")
                return differences
            }
        } catch (e: IOException) {
            Log.e(TAG, "IO Exception while reading file: ${e.message}", e)
            return differences
        } catch (e: Exception) {
            Log.e(TAG, "Unexpected error: ${e.message}", e)
            return differences
        }
    }
}