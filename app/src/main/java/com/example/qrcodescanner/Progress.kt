package com.example.qrcodescanner

import android.content.Intent
import android.graphics.Color
import android.os.Bundle
import android.util.Log
import android.widget.EditText
import android.widget.LinearLayout
import android.widget.TextView
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import com.example.qrcodescanner.databinding.ActivityProgressBinding
import com.github.mikephil.charting.animation.Easing
import com.github.mikephil.charting.charts.BarChart
import com.github.mikephil.charting.charts.LineChart
import com.github.mikephil.charting.data.BarData
import com.github.mikephil.charting.data.BarDataSet
import com.github.mikephil.charting.data.BarEntry
import com.github.mikephil.charting.data.Entry
import com.github.mikephil.charting.data.LineData
import com.github.mikephil.charting.data.LineDataSet
import com.github.mikephil.charting.formatter.IndexAxisValueFormatter
import java.text.SimpleDateFormat
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.util.Date
import java.util.Locale

class Progress : AppCompatActivity() {

    private lateinit var binding: ActivityProgressBinding
    private lateinit var excelHelper: ExcelHelper

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_progress)

        binding = ActivityProgressBinding.inflate(layoutInflater)
        setContentView(binding.root)
        excelHelper = ExcelHelper(this)

        val bodydata = excelHelper.searchFromBottomN("born_again-db.xlsx", "bodydata", 7)

        val dates = convertDateList(getElementsAtIndex(bodydata, 0))
        val weights = getElementsAtIndex(bodydata, 1).map { it.toDouble() }

        val calories : MutableList<Double> = mutableListOf()
        val lastDateTime = excelHelper.processExcelTimesWithKeys("born_again-db.xlsx", "record", 7)
        val dates_calories = lastDateTime.first
        val durations = lastDateTime.second
        val met = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "C1").toDouble()
        val latestBodyWeights = excelHelper.getLastCellWithContentAsString("born_again-db.xlsx", "bodydata").toDouble()
        for (i in durations.indices) {
            val col6 = excelHelper.getLastRowCol6Value("born_again-db.xlsx", "record", dates_calories[i])
            if (col6 != "null"){
                val col6Double = col6?.toDoubleOrNull()
                calories.add(String.format("%.2f", col6Double).toDouble())
            } else{
                val cal = met * latestBodyWeights * durations[i]
                excelHelper.updateLastRowValueIfEmpty("born_again-db.xlsx", "record", dates_calories[i], cal)
                calories.add(String.format("%.2f", cal).toDouble())
            }
        }

        var input : String = "0"

        binding.EnterYourCardioButton.setOnClickListener {
            val editText = EditText(this).apply {
                layoutParams = LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT

                )
            }

            val container = LinearLayout(this).apply {
                orientation = LinearLayout.VERTICAL
                setPadding(50, 40, 50, 10)
                addView(TextView(this@Progress).apply { text = "Your Cardio Today:" })
                addView(editText)
            }

            AlertDialog.Builder(this)
                .setView(container)
                .setPositiveButton("OK") { _, _ ->
                    excelHelper.updateLastRowValueAddIfNotEmpty("born_again-db.xlsx", "record", getTodayDate(), editText.text.toString().toDouble())
                }
                .setNegativeButton("Cancel") { dialog, _ -> dialog.dismiss() }
                .show()
        }

        setupBodyWeightChangeChart(binding.Chart, weights, dates)
        setupWeightLiftedChart(binding.Chart2, calories.reversed(), convertDateList(dates_calories).reversed())
        setupEditBodyWeight(bodydata)

        setupHomeButton()
        setupSaveButton()
    }

    fun getTodayDate(): String {
        val today = LocalDate.now()
        val formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd")
        return today.format(formatter)
    }

    private fun getElementsAtIndex(lists: List<List<String>>, index: Int): List<String> {
        return lists.map { it[index] }
    }

    private fun convertDateList(dateList: List<String>): List<String> {
        return dateList.map { date ->
            val parts = date.split("-")
            if (parts.size == 3) {
                val day = parts[2]
                val month = parts[1]
                "$day/$month"
            } else {
                throw IllegalArgumentException("Invalid date format. Expected yyyy-mm-dd.")
            }
        }
    }

    private fun setupEditBodyWeight(nestedlist: List<List<String>>) {
        val latest = nestedlist.last()[0]
        if (LocalDate.parse(latest, DateTimeFormatter.ofPattern("yyyy-MM-dd")) == LocalDate.now()){
            binding.EditBodyWeight.isEnabled = false
            binding.EditBodyWeight.setText((nestedlist.last()[1]) + " kg")
            binding.SaveBodyWeight.isEnabled = false
        } else{
            binding.EditBodyWeight.isEnabled = true
            binding.SaveBodyWeight.isEnabled = true
        }
    }

    private fun setupSaveButton(){
        binding.SaveBodyWeight.setOnClickListener{
            excelHelper.modifyExcelFromBottom("born_again-db.xlsx", "bodydata", listOf(SimpleDateFormat("yyyy-MM-dd", Locale.getDefault()).format(Date()), binding.EditBodyWeight.text.toString()))
            binding.EditBodyWeight.isEnabled = false
            binding.SaveBodyWeight.isEnabled = false
        }
    }
    private fun setupHomeButton() {
        binding.ReturnHomeButton.setOnClickListener {
            val intent = Intent(this, MainActivity::class.java)
            startActivity(intent)
            finish()
        }
    }

    private fun setupBodyWeightChangeChart(lineChart: LineChart, doubleValues: List<Double>, xLabels: List<String>) {
        if (doubleValues.size != xLabels.size) {
            throw IllegalArgumentException("xLabels must have the same length as doubleValues")
        }

        lineChart.clear()

        val entries = doubleValues.mapIndexed { index, value ->
            Entry(index.toFloat(), value.toFloat())
        }

        val dataSet = LineDataSet(entries, "Weight (kg)").apply {
            color = Color.rgb(30, 144, 255)
            fillColor = Color.argb(50, 30, 144, 255)
            setDrawFilled(true)
            lineWidth = 2f
            setDrawCircles(true)
            circleRadius = 4f
        }

        val lineData = LineData(dataSet)

        lineChart.data = lineData

        lineChart.apply {
            description.isEnabled = false
            isDragEnabled = true
            setScaleEnabled(true)
            setPinchZoom(true)
            setDrawGridBackground(false)
            xAxis.apply {
                valueFormatter = IndexAxisValueFormatter(xLabels)
                position = com.github.mikephil.charting.components.XAxis.XAxisPosition.BOTTOM
                setDrawGridLines(false)
                textSize = 12f
                labelCount = doubleValues.size
                granularity = 1f
                isGranularityEnabled = true
            }
            axisLeft.apply {
                axisMinimum = 0f
                setDrawGridLines(false)
                textSize = 12f
            }
            axisRight.isEnabled = false
            legend.isEnabled = true
            legend.textSize = 12f
            animateXY(1000, 1000, Easing.EaseInOutQuad)
        }

        lineChart.invalidate()
    }

    private fun setupWeightLiftedChart(barChart: BarChart, doubleValues: List<Double>, xLabels: List<String>) {
        if (doubleValues.size != xLabels.size) {
            throw IllegalArgumentException("xLabels must have the same length as doubleValues")
        }

        barChart.clear()

        // Convert doubles to BarEntry objects for MPAndroidChart
        // Each BarEntry maps an x-index (0 to size-1) to a y-value
        val entries = doubleValues.mapIndexed { index, value ->
            BarEntry(index.toFloat(), value.toFloat())
        }

        // Create a BarDataSet to define the bars' appearance and data
        val dataSet = BarDataSet(entries, "Calories (kcal)").apply {
            // Set bar color (dark blue for visibility)
            color = Color.rgb(30, 144, 255)
            // Set bar value text color
            valueTextColor = Color.BLACK
            // Set bar value text size
            valueTextSize = 12f
            // Enable displaying values on top of bars
            setDrawValues(true)
        }

        // Create BarData object to hold the dataset
        val barData = BarData(dataSet)

        // Set the data to the chart, adjusting bar width for better spacing
        barData.barWidth = 0.5f // Width of each bar (adjustable, max 1.0f for no overlap)
        barChart.data = barData

        barChart.apply {
            description.isEnabled = false
            isDragEnabled = true
            setScaleEnabled(true)
            setPinchZoom(true)
            setDrawGridBackground(false)
            xAxis.apply {
                valueFormatter = IndexAxisValueFormatter(xLabels)
                position = com.github.mikephil.charting.components.XAxis.XAxisPosition.BOTTOM
                setDrawGridLines(false)
                textSize = 12f
                labelCount = doubleValues.size
                granularity = 1f
                isGranularityEnabled = true
            }
            axisLeft.apply {
                axisMinimum = 0f
                setDrawGridLines(false)
                textSize = 12f
            }
            axisRight.isEnabled = false
            legend.isEnabled = true
            legend.textSize = 12f
            // Animate both x and y axes for a smoother transition
            animateXY(2000, 2000, Easing.EaseInOutQuad)
        }

        barChart.invalidate()
    }

    fun checkNestedList(nestedList: List<List<String>>) {
        if (nestedList.isEmpty()) {
            Log.d("CheckNestedList", "The nested list is empty.")
        } else {
            Log.d("CheckNestedList", "The nested list contains ${nestedList.size} inner lists.")
            for ((index, innerList) in nestedList.withIndex()) {
                Log.d("CheckNestedList", "Inner list $index: ${innerList.size} elements")
            }
        }
    }

    fun checkStringList(stringList: List<String>) {
        if (stringList.isEmpty()) {
            Log.d("CheckStringList", "The string list is empty.")
        } else {
            Log.d("CheckStringList", "The string list contains ${stringList.size} elements.")
            Log.d("CheckStringList", "Elements: $stringList")
        }
    }
}