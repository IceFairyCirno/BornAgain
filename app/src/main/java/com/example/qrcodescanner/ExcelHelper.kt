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

    fun processExcelTimesWithKeys(filename: String, sheetName: String, n: Int): Pair<List<String>, List<Double>> {
        val TAG = "ExcelTimeProcessor"
        val keys = mutableListOf<String>()
        val results = mutableListOf<Double>()
        val timeFormat = SimpleDateFormat("HH:mm", Locale.US)

        Log.d(TAG, "Opening file: $filename, sheet: $sheetName")

        FileInputStream(File(context.filesDir, filename)).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                Log.e(TAG, "Sheet '$sheetName' not found in $filename")
                return Pair(emptyList(), emptyList())
            }

            Log.d(TAG, "Sheet '$sheetName' found. Reading first column values.")

            val col1Values = mutableListOf<String>()
            for (row in sheet) {
                val cellVal = row.getCell(0)?.toString()?.trim()
                if (!cellVal.isNullOrEmpty()) {
                    col1Values.add(cellVal)
                }
            }

            Log.d(TAG, "Total rows read: ${col1Values.size}")

            // Get last n unique values in reverse order
            val lastNUnique = col1Values.asReversed().distinct().take(n)
            Log.d(TAG, "Last $n unique values in first column: $lastNUnique")

            for (value in lastNUnique) {
                val matchingRows = sheet.filter { row ->
                    row.getCell(0)?.toString()?.trim() == value
                }

                if (matchingRows.isEmpty()) {
                    Log.w(TAG, "No matching rows found for value '$value'")
                    keys.add(value)
                    results.add(0.0)
                    continue
                }

                val firstRow = matchingRows.first()
                val lastRow = matchingRows.last()

                val firstTimeStr = firstRow.getCell(4)?.toString()?.trim() ?: ""
                val lastTimeStr = lastRow.getCell(4)?.toString()?.trim() ?: ""

                Log.d(TAG, "Processing value '$value': firstTime='$firstTimeStr', lastTime='$lastTimeStr'")

                val diffHours = try {
                    if (firstTimeStr.matches(Regex("^\\d{2}:\\d{2}$")) &&
                        lastTimeStr.matches(Regex("^\\d{2}:\\d{2}$"))
                    ) {
                        val firstDate = timeFormat.parse(firstTimeStr)
                        val lastDate = timeFormat.parse(lastTimeStr)
                        if (firstDate != null && lastDate != null) {
                            val diffMillis = lastDate.time - firstDate.time
                            val diff = diffMillis / (1000.0 * 60.0 * 60.0)
                            Log.d(TAG, "Time difference for '$value': $diff hours")
                            diff
                        } else {
                            Log.w(TAG, "Failed to parse times for '$value'")
                            0.0
                        }
                    } else {
                        Log.w(TAG, "Time format mismatch for '$value': firstTime='$firstTimeStr', lastTime='$lastTimeStr'")
                        0.0
                    }
                } catch (e: Exception) {
                    Log.e(TAG, "Exception parsing times for '$value': ${e.message}")
                    0.0
                }

                keys.add(value)
                results.add(diffHours)
            }

            workbook.close()
        }
        Log.d(TAG, "Processing complete. Keys: $keys, Results: $results")
        return Pair(keys, results)
    }

    fun updateLastRowValueIfEmpty(
        filename: String,
        sheetName: String,
        colValue: String,
        result: Double
    ) {
        val TAG = "ExcelUpdater"

        Log.d(TAG, "Opening file: $filename, sheet: $sheetName")

        FileInputStream(File(context.filesDir, filename)).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                Log.e(TAG, "Sheet '$sheetName' not found in $filename")
                workbook.close()
                return
            }

            Log.d(TAG, "Searching for last row containing '$colValue' in column 1")

            val lastMatchingRow = (sheet.lastRowNum downTo 0).mapNotNull { rowIndex -> sheet.getRow(rowIndex) }
                .firstOrNull { row ->
                    row.getCell(0)?.toString()?.trim() == colValue
                }

            if (lastMatchingRow == null) {
                Log.w(TAG, "No row found with value '$colValue' in column 1")
                workbook.close()
                return
            }

            Log.d(TAG, "Last matching row number: ${lastMatchingRow.rowNum}")

            val cell = lastMatchingRow.getCell(5) ?: lastMatchingRow.createCell(5)
            val cellValueString = cell.toString()

            if (cell.cellType == org.apache.poi.ss.usermodel.CellType.BLANK
                || cellValueString.isEmpty()
                || (cell.cellType == org.apache.poi.ss.usermodel.CellType.STRING && cellValueString.isBlank())
            ) {
                Log.d(TAG, "Cell at column 6 is empty. Writing result: $result")
                cell.setCellValue(result)

                FileOutputStream(File(context.filesDir, filename)).use { fos ->
                    workbook.write(fos)
                    Log.d(TAG, "Saved updates to file: $filename")
                }
            } else {
                Log.d(TAG, "Cell at column 6 is not empty ('${cellValueString}'). No change made.")
            }

            workbook.close()
        }
    }
    fun updateLastRowValueAddIfNotEmpty(
        filename: String,
        sheetName: String,
        colValue: String,
        result: Double
    ) {
        val TAG = "ExcelUpdater"

        Log.d(TAG, "Opening file: $filename, sheet: $sheetName")

        FileInputStream(File(context.filesDir, filename)).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                Log.e(TAG, "Sheet '$sheetName' not found in $filename")
                workbook.close()
                return
            }

            Log.d(TAG, "Searching for last row containing '$colValue' in column 1")

            val lastMatchingRow = (sheet.lastRowNum downTo 0).mapNotNull { rowIndex -> sheet.getRow(rowIndex) }
                .firstOrNull { row ->
                    row.getCell(0)?.toString()?.trim() == colValue
                }

            if (lastMatchingRow == null) {
                Log.w(TAG, "No row found with value '$colValue' in column 1")
                workbook.close()
                return
            }

            Log.d(TAG, "Last matching row number: ${lastMatchingRow.rowNum}")

            val cell = lastMatchingRow.getCell(5) ?: lastMatchingRow.createCell(5)
            val cellValueString = cell.toString()

            if (cell.cellType == org.apache.poi.ss.usermodel.CellType.BLANK
                || cellValueString.isEmpty()
                || (cell.cellType == org.apache.poi.ss.usermodel.CellType.STRING && cellValueString.isBlank())
            ) {
                Log.d(TAG, "Cell at column 6 is empty. Writing result: $result")
                cell.setCellValue(result)
            } else {
                val existingValue = try {
                    cell.numericCellValue
                } catch (e: Exception) {
                    cellValueString.toDoubleOrNull() ?: run {
                        Log.w(TAG, "Failed to parse cell value '$cellValueString' as double, defaulting to 0.0")
                        0.0
                    }
                }
                val newValue = existingValue + result
                Log.d(TAG, "Cell at column 6 has existing value $existingValue. Adding $result = $newValue")
                cell.setCellValue(newValue)
            }

            FileOutputStream(File(context.filesDir, filename)).use { fos ->
                workbook.write(fos)
                Log.d(TAG, "Saved updates to file: $filename")
            }
            workbook.close()
        }
    }

    fun getLastRowCol6Value(
        filename: String,
        sheetName: String,
        colValue: String
    ): String? {
        val TAG = "ExcelReader"

        Log.d(TAG, "Opening file: $filename, sheet: $sheetName")

        FileInputStream(File(context.filesDir, filename)).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                Log.e(TAG, "Sheet '$sheetName' not found in $filename")
                workbook.close()
                return null
            }

            Log.d(TAG, "Searching for last row containing '$colValue' in column 1")

            val lastMatchingRow = (sheet.lastRowNum downTo 0).mapNotNull { sheet.getRow(it) }
                .firstOrNull { row ->
                    row.getCell(0)?.toString()?.trim() == colValue
                }

            if (lastMatchingRow == null) {
                Log.w(TAG, "No row found with value '$colValue' in column 1")
                workbook.close()
                return null
            }

            Log.d(TAG, "Last matching row number: ${lastMatchingRow.rowNum}")

            val cell = lastMatchingRow.getCell(5)  // column 6 index is 5
            val cellValue = cell?.toString()?.takeIf { it.isNotBlank() }

            Log.d(TAG, "Value in column 6: ${cellValue ?: "null"}")

            workbook.close()
            return cellValue ?: "null"
        }
    }

    fun getUniqueColumnValuesWithLogging(filename: String, sheetName: String, colNum: Int): List<String> {
        val TAG = "ExcelUniqueValues"
        val uniqueValues = mutableSetOf<String>()

        Log.d(TAG, "Opening file: $filename, sheet: $sheetName, column: $colNum")

        val fileName = File(context.filesDir, filename)

        FileInputStream(fileName).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                Log.e(TAG, "Sheet '$sheetName' not found in $filename")
                workbook.close()
                return emptyList()
            }

            Log.d(TAG, "Iterating rows from 0 to ${sheet.lastRowNum}")

            for (rowIndex in 0..sheet.lastRowNum) {
                val row = sheet.getRow(rowIndex)
                if (row == null) {
                    Log.w(TAG, "Row $rowIndex is null, skipping")
                    continue
                }

                val cell = row.getCell(colNum - 1) // Adjust for 1-based colNum
                if (cell == null) {
                    Log.d(TAG, "Row $rowIndex, cell at column ${colNum} is null, skipping")
                    continue
                }

                val cellValue = when (cell.cellType) {
                    CellType.STRING -> cell.stringCellValue
                    CellType.NUMERIC -> cell.numericCellValue.toString()
                    CellType.BOOLEAN -> cell.booleanCellValue.toString()
                    CellType.FORMULA -> {
                        try {
                            val evaluator = workbook.creationHelper.createFormulaEvaluator()
                            val evaluatedCell = evaluator.evaluate(cell)
                            when (evaluatedCell.cellType) {
                                CellType.STRING -> evaluatedCell.stringValue
                                CellType.NUMERIC -> evaluatedCell.numberValue.toString()
                                CellType.BOOLEAN -> evaluatedCell.booleanValue.toString()
                                else -> {
                                    Log.w(TAG, "Row $rowIndex formula cell evaluated to unsupported type")
                                    null
                                }
                            }
                        } catch (e: Exception) {
                            Log.e(TAG, "Error evaluating formula in row $rowIndex: ${e.message}")
                            null
                        }
                    }
                    else -> {
                        Log.d(TAG, "Row $rowIndex, cell type ${cell.cellType} not handled")
                        null
                    }
                }?.trim()

                if (!cellValue.isNullOrEmpty()) {
                    val added = uniqueValues.add(cellValue)
                    if (added) {
                        Log.d(TAG, "Added unique value: '$cellValue' from row $rowIndex")
                    } else {
                        Log.d(TAG, "Duplicate value encountered: '$cellValue' from row $rowIndex")
                    }
                } else {
                    Log.d(TAG, "Empty or null cell value at row $rowIndex")
                }
            }

            workbook.close()
            Log.d(TAG, "Found total unique values: ${uniqueValues.size}")
        }

        return uniqueValues.toList()
    }

}