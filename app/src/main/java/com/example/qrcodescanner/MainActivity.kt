package com.example.qrcodescanner

import android.Manifest
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.util.Log
import android.view.View
import android.widget.ImageButton
import android.widget.TextView
import android.widget.Toast
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.cardview.widget.CardView
import androidx.core.content.ContextCompat
import com.example.qrcodescanner.databinding.ActivityMain2Binding
import com.example.qrcodescanner.databinding.ActivityMainBinding
import com.journeyapps.barcodescanner.CaptureActivity
import com.journeyapps.barcodescanner.ScanContract
import com.journeyapps.barcodescanner.ScanOptions
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.temporal.ChronoUnit

class MainActivity : AppCompatActivity() {
    private lateinit var excelHelper: ExcelHelper
    private lateinit var binding: ActivityMainBinding


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main)
        excelHelper = ExcelHelper()
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        excelHelper.createExcel(this, "record.xlsx")

        binding.ScanButton.setOnClickListener {
            val options = ScanOptions()
            options.setPrompt("Scan a QR Code")
            options.setBeepEnabled(true)
            options.setOrientationLocked(true)
            options.setCaptureActivity(CaptureActivity::class.java)
            barcodeLauncher.launch(options)
        }

        binding.MusicButton.setOnClickListener{
            val intent = Intent(this, MusicPlayer::class.java)
            startActivity(intent)
            finish()
        }

        binding.ProgressButton.setOnClickListener{
            val intent = Intent(this, Progress::class.java)
            startActivity(intent)
            finish()
        }

        val recentRecord = searchRecent(this, "record.xlsx", 1, true)
        setupRecentExercise(recentRecord.reversed())

        val week = listOf(binding.SundayContainer, binding.MondayContainer, binding.TuesdayContainer, binding.WednesdayContainer, binding.ThursdayContainer, binding.FridayContainer, binding.SaturdayContainer,)
        week[getTodayWeekday()].setCardBackgroundColor(ContextCompat.getColor(this, R.color.blue))
        val weektext = week[getTodayWeekday()].getChildAt(0) as? TextView
        weektext?.setTextColor(android.graphics.Color.WHITE)
    }

    private val barcodeLauncher = registerForActivityResult(ScanContract()){result ->
        if (result.contents!=null){
            if ("gym" in result.contents){
                val exerciseName = (result.contents).split("/").last()
                val intent: Intent
                if ("cable" in result.contents){
                    intent = Intent(this, SubMachines::class.java)
                } else{
                    intent = Intent(this, MainActivity2::class.java).apply{
                        putExtra("exerciseName", exerciseName)
                    }
                }
                startActivity(intent)
                finish()
            } else if ("cable" in result.contents){
                val intent = Intent(this, SubMachines::class.java)
                startActivity(intent)
                finish()
            }
        } else{
            Toast.makeText(this, "Invalid QR Code", Toast.LENGTH_SHORT).show()
        }
    }

    private fun searchRecent(context: Context, fileName: String, columnIndex: Int, isReverse: Boolean): List<List<String>> {
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
        val seenValues = mutableSetOf<String>()

        try {
            if (isReverse) {
                for (rowIndex in sheet.lastRowNum downTo 0) {
                    val row = sheet.getRow(rowIndex)
                    if (row != null) {
                        val cell = row.getCell(columnIndex, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        val cellValue = cell.toString()
                        if (seenValues.add(cellValue)) {
                            val rowData = mutableListOf<String>()
                            row.forEach { rowData.add(it.toString()) }
                            result.add(0, rowData)
                        }
                        if (seenValues.size >= 3) break
                    }
                }
            } else {
                for (row in sheet) {
                    val cell = row.getCell(columnIndex, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                    val cellValue = cell.toString()
                    if (seenValues.add(cellValue)) {
                        val rowData = mutableListOf<String>()
                        row.forEach { rowData.add(it.toString()) }
                        result.add(rowData)
                    }
                    if (seenValues.size >= 3) break
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

    private fun applyToCard(textViews: List<TextView>, values: List<String>){
        if (textViews.size != 4 || values.size != 4) {
            throw IllegalArgumentException("Both lists must contain exactly 4 elements.")
        }
        for (i in textViews.indices) {
            textViews[i].text = values[i]
        }

    }

    private fun getPassedDate(inputDate: String): String {
        val formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd")

        val date = LocalDate.parse(inputDate, formatter)
        val today = LocalDate.now()

        val daysDifference = ChronoUnit.DAYS.between(date, today).toInt()

        return when (daysDifference) {
            1 -> "Yesterday"
            in 2..Int.MAX_VALUE -> "$daysDifference Days Ago"
            else -> "Today"
        }
    }

    private fun getInputList(record: List<String>): List<String>{
        val rightSideData = countMatchingRows("record.xlsx", record[0], record[1])
        val avgWeight = "%.1f".format(rightSideData.second)
        val setNum = rightSideData.first.toString()

        return listOf(
            getPassedDate(record[0]),
            record[1],
            "$avgWeight kg",
            "$setNum Sets"
        )
    }

    private fun countMatchingRows(fileName: String, col0: String, col1: String): Pair<Int, Float> {
        var matchCount = 0
        var avgWeight = 0f
        try {
            val file = File(this.filesDir, fileName)
            if (!file.exists()) {
                return Pair(-1, -1f)
            }
            FileInputStream(file).use { fis ->
                val workbook = WorkbookFactory.create(fis)
                val sheet = workbook.getSheetAt(0)
                for (row in sheet) {
                    val cell0 = row.getCell(0)
                    val cell1 = row.getCell(1)
                    val cell0Value = cell0?.toString()?.trim()
                    val cell1Value = cell1?.toString()?.trim()

                    if (cell0Value == col0 && cell1Value == col1) {
                        val cell2 = row.getCell(2)
                        val cell2Value = cell2?.toString()?.trim()
                        if (cell2Value != null) {
                            avgWeight = avgWeight + cell2Value.toFloat()
                        }
                        matchCount++
                    }
                }
                workbook.close()
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return Pair(matchCount, avgWeight/matchCount)
    }

    private fun setupRecentExercise(recentRecord: List<List<String>>){
        binding.Card1.visibility = View.GONE
        binding.Card2.visibility = View.GONE
        binding.Card3.visibility = View.GONE

        if (recentRecord.size >= 1){
            binding.Card1.visibility = View.VISIBLE
            val inputData = getInputList(recentRecord[0])
            applyToCard(listOf(binding.Card1Date, binding.Card1Name, binding.Card1Weight, binding.Card1Sets), inputData)
        }
        if (recentRecord.size >= 2){
            binding.Card2.visibility = View.VISIBLE
            val inputData = getInputList(recentRecord[0])
            applyToCard(listOf(binding.Card1Date, binding.Card1Name, binding.Card1Weight, binding.Card1Sets), inputData)
            val inputData1 = getInputList(recentRecord[1])
            applyToCard(listOf(binding.Card2Date, binding.Card2Name, binding.Card2Weight, binding.Card2Sets), inputData1)
        }
        if (recentRecord.size >= 3){
            binding.Card3.visibility = View.VISIBLE
            val inputData = getInputList(recentRecord[0])
            applyToCard(listOf(binding.Card1Date, binding.Card1Name, binding.Card1Weight, binding.Card1Sets), inputData)
            val inputData1 = getInputList(recentRecord[1])
            applyToCard(listOf(binding.Card2Date, binding.Card2Name, binding.Card2Weight, binding.Card2Sets), inputData1)
            val inputData2 = getInputList(recentRecord[2])
            applyToCard(listOf(binding.Card3Date, binding.Card3Name, binding.Card3Weight, binding.Card3Sets), inputData2)
        }
    }

    private fun getTodayWeekday(): Int {
        val today = LocalDate.now()
        val dayOfWeek = today.dayOfWeek.value
        return if (dayOfWeek == 7) 0 else dayOfWeek
    }
}