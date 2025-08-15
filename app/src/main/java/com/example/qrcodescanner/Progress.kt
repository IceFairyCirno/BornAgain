package com.example.qrcodescanner

import android.content.Context
import android.content.Intent
import android.os.Bundle
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import com.example.qrcodescanner.databinding.ActivityMainBinding
import com.example.qrcodescanner.databinding.ActivityProgressBinding
import com.github.mikephil.charting.charts.BarChart
import com.github.mikephil.charting.charts.LineChart
import com.github.mikephil.charting.components.XAxis
import com.github.mikephil.charting.data.BarData
import com.github.mikephil.charting.data.BarDataSet
import com.github.mikephil.charting.data.BarEntry
import com.github.mikephil.charting.data.Entry
import com.github.mikephil.charting.data.LineData
import com.github.mikephil.charting.data.LineDataSet
import com.github.mikephil.charting.formatter.IndexAxisValueFormatter
import com.github.mikephil.charting.formatter.ValueFormatter
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.text.SimpleDateFormat
import java.util.Calendar
import java.util.Locale

class Progress : AppCompatActivity() {

    private lateinit var binding: ActivityProgressBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_progress)

        binding = ActivityProgressBinding.inflate(layoutInflater)
        setContentView(binding.root)

        val weights = listOf(50, 55, 60, 58, 62)

        val weightLiftedData = getSumOfLastFiveUniqueCol1()
        val bodyweightData = getBodyWeight()

        setupWeightChart(binding.Chart, bodyweightData.second, bodyweightData.first)
        setupBarChart(binding.Chart2, weightLiftedData.first, weightLiftedData.second)
        setupHomeButton()
    }

    private fun setupWeightChart(lineChart: LineChart, weights: List<Double>, labels: List<String>) {
        // Exit if data or labels are empty or mismatched
        if (weights.isEmpty() || labels.size != weights.size) {
            lineChart.clear()
            return
        }

        // Create entries from the list of doubles
        val entries = weights.mapIndexed { index, weight ->
            Entry(index.toFloat(), weight.toFloat())
        }

        // Create and customize the dataset
        val dataSet = LineDataSet(entries, "Weight Progress").apply {
            color = lineChart.context.resources.getColor(R.color.blue)
            valueTextColor = lineChart.context.resources.getColor(android.R.color.black)
            valueTextSize = 10f // Not used since values are hidden
            lineWidth = 3f // Thicker line for visibility
            setCircleColor(android.R.color.holo_blue_dark)
            circleRadius = 5f // Larger points
            setDrawCircleHole(false)
            setDrawValues(false) // Disable values on data points
        }

        // Create LineData object
        val lineData = LineData(dataSet)

        // Configure the chart
        lineChart.apply {
            data = lineData
            description.isEnabled = false
            setTouchEnabled(true) // Enable touch gestures
            isDragEnabled = true
            setScaleEnabled(true)
            setPinchZoom(true)

            // Ensure chart fills the view
            setExtraOffsets(10f, 10f, 10f, 20f) // Extra bottom offset for x-axis labels
            setMinimumHeight(400) // Ensure minimum height
            setMinimumWidth(300) // Ensure minimum width

            // Customize x-axis
            xAxis.apply {
                setDrawGridLines(false) // Disable vertical grid lines
                valueFormatter = IndexAxisValueFormatter(labels) // Set custom labels
                labelCount = labels.size
                textSize = 12f // Larger x-axis labels
                granularity = 1f // One label per data point
                position = XAxis.XAxisPosition.BOTTOM // X-axis at bottom
            }

            // Customize y-axis
            axisLeft.apply {
                setDrawGridLines(false) // Disable horizontal grid lines
                textSize = 12f // Larger y-axis labels
                axisMinimum = (weights.minOrNull()?.toFloat() ?: 0f) * 0.9f // Slightly below min weight
                axisMaximum = (weights.maxOrNull()?.toFloat() ?: 100f) * 1.1f // Slightly above max weight
            }
            axisRight.isEnabled = false // Disable right y-axis

            // Fix scaling and viewport
            setVisibleXRangeMaximum(10f) // Show up to 10 data points at a time
            moveViewToX(0f) // Start at the first data point
            setScaleMinima(1f, 1f) // Prevent excessive zooming out

            // Animate for better visuals
            animateY(1000)

            // Refresh the chart
            invalidate()
        }
    }

    private fun getSumOfLastFiveUniqueCol1(): Pair<List<Double>, List<String>> {
        try {
            // Open the file from internal storage
            val file = File(this.filesDir, "record.xlsx")
            if (!file.exists()) return Pair(emptyList(), emptyList())

            // Date formatters for parsing and formatting
            val inputDateFormat = SimpleDateFormat("yyyy-MM-dd", Locale.getDefault())
            val outputDateFormat = SimpleDateFormat("dd/MM", Locale.getDefault())

            // Read the Excel file
            FileInputStream(file).use { fis ->
                val workbook = XSSFWorkbook(fis)
                val sheet = workbook.getSheetAt(0) // Assuming first sheet

                // Map to store sums for each unique Column 1 value
                val sumsByCol1 = mutableMapOf<String, Double>()
                val col1Order = mutableListOf<String>() // To track order of unique Column 1 values

                // Iterate through rows to collect Column 1 values and sum Column 3
                for (row in sheet) {
                    val col1Cell = row.getCell(0) // First column (index 0)
                    val col3Cell = row.getCell(2) // Third column (index 2)

                    if (col1Cell != null && col3Cell != null) {
                        // Get Column 1 value as string (handles strings or numbers)
                        val col1Value = when (col1Cell.cellType) {
                            CellType.STRING -> col1Cell.stringCellValue
                            CellType.NUMERIC -> col1Cell.numericCellValue.toString()
                            else -> continue
                        }

                        // Get Column 3 value as double
                        val col3Value = when (col3Cell.cellType) {
                            CellType.NUMERIC -> col3Cell.numericCellValue
                            CellType.STRING -> col3Cell.stringCellValue.toDoubleOrNull() ?: continue
                            else -> continue
                        }

                        // Add to sums and track order if valid date
                        if (col1Value.isNotBlank()) {
                            try {
                                // Parse and reformat date to ensure it's valid
                                val date = inputDateFormat.parse(col1Value)
                                val formattedDate = outputDateFormat.format(date)
                                if (!col1Order.contains(formattedDate)) {
                                    col1Order.add(formattedDate)
                                }
                                sumsByCol1[formattedDate] = sumsByCol1.getOrDefault(formattedDate, 0.0) + col3Value
                            } catch (e: Exception) {
                                continue // Skip if not a valid date
                            }
                        }
                    }
                }

                // Get the last 5 unique Column 1 values (or fewer if less than 5)
                val lastFiveCol1 = col1Order.takeLast(5)

                // Get the sums for these values
                val sums = lastFiveCol1.map { sumsByCol1[it] ?: 0.0 }.take(5)

                // Return both sums and formatted Column 1 values
                return Pair(sums, lastFiveCol1)
            }
        } catch (e: Exception) {
            e.printStackTrace()
            return Pair(emptyList(), emptyList())
        }
    }

    private fun setupBarChart(barChart: BarChart, data: List<Double>, labels: List<String>) {
        if (data.isEmpty() || labels.size != data.size) {
            barChart.clear()
            return
        }

        // Create entries for the bar chart
        val entries = data.mapIndexed { index, value ->
            BarEntry(index.toFloat(), value.toFloat())
        }

        // Create and customize the dataset
        val dataSet = BarDataSet(entries, "Session Progress").apply {
            color = barChart.context.resources.getColor(R.color.blue)
            valueTextColor = barChart.context.resources.getColor(android.R.color.black)
            valueTextSize = 12f // Not used since values are hidden
            setDrawValues(false) // Disable values on bars
        }

        // Create BarData object
        val barData = BarData(dataSet).apply {
            barWidth = 0.4f // Set bar width (adjust as needed)
        }

        // Configure the chart
        barChart.apply {
            barChart.data = barData
            description.isEnabled = false
            setTouchEnabled(true) // Enable touch gestures
            isDragEnabled = true
            setScaleEnabled(true)
            setPinchZoom(true)

            // Ensure chart fills the view
            setExtraOffsets(10f, 10f, 10f, 20f) // Extra bottom offset for x-axis labels
            setMinimumHeight(400) // Ensure minimum height
            setMinimumWidth(300) // Ensure minimum width

            // Customize x-axis
            xAxis.apply {
                position = XAxis.XAxisPosition.BOTTOM // X-axis at bottom
                setDrawGridLines(false) // Disable vertical grid lines
                valueFormatter = IndexAxisValueFormatter(labels) // Set custom labels
                labelCount = labels.size
                textSize = 12f // Larger x-axis labels
                granularity = 1f // One label per bar
            }

            // Customize y-axis
            axisLeft.apply {
                setDrawGridLines(false) // Disable horizontal grid lines
                textSize = 12f // Larger y-axis labels
                axisMinimum = (data.minOrNull()?.toFloat() ?: 0f) * 0.9f // Slightly below min
                axisMaximum = (data.maxOrNull()?.toFloat() ?: 100f) * 1.1f // Slightly above max
            }
            axisRight.isEnabled = false // Disable right y-axis

            // Fix scaling and viewport
            setVisibleXRangeMaximum(10f) // Show up to 10 bars at a time
            moveViewToX(0f) // Start at the first bar
            setScaleMinima(1f, 1f) // Prevent excessive zooming out

            // Animate for better visuals
            animateY(1000)

            // Refresh the chart
            invalidate()
        }
    }

    private fun setupHomeButton() {
        binding.ReturnHomeButton.setOnClickListener {
            val intent = Intent(this, MainActivity::class.java)
            startActivity(intent)
            finish()
        }
    }

    private fun getBodyWeight(): Pair<List<String>, List<Double>> {
        try {
            // Open the file from internal storage
            val file = File(this.filesDir, "record.xlsx")
            if (!file.exists()) return Pair(emptyList(), emptyList())

            // Date formatters for parsing and formatting
            val inputDateFormat = SimpleDateFormat("yyyy-MM-dd", Locale.getDefault())
            val outputDateFormat = SimpleDateFormat("dd/MM", Locale.getDefault())

            // Read the Excel file
            FileInputStream(file).use { fis ->
                val workbook = XSSFWorkbook(fis)
                val sheet = workbook.getSheetAt(1) // Access second sheet (index 1)

                // Map to store the last Column 2 value for each unique Column 1 value
                val col2ByCol1 = mutableMapOf<String, Double>()
                val col1Order = mutableListOf<String>() // Track order of unique Column 1 values

                // Iterate through rows to collect Column 1 and Column 2 values
                for (row in sheet) {
                    val col1Cell = row.getCell(0) // First column (index 0)
                    val col2Cell = row.getCell(1) // Second column (index 1)

                    if (col1Cell != null && col2Cell != null) {
                        // Get Column 1 value as string (handles strings or numbers)
                        val col1Value = when (col1Cell.cellType) {
                            CellType.STRING -> col1Cell.stringCellValue
                            CellType.NUMERIC -> col1Cell.numericCellValue.toString()
                            else -> continue
                        }

                        // Get Column 2 value as double
                        val col2Value = when (col2Cell.cellType) {
                            CellType.NUMERIC -> col2Cell.numericCellValue
                            CellType.STRING -> col2Cell.stringCellValue.toDoubleOrNull() ?: continue
                            else -> continue
                        }

                        // Parse and reformat date, store with Column 2 value
                        if (col1Value.isNotBlank()) {
                            try {
                                val date = inputDateFormat.parse(col1Value)
                                val formattedDate = outputDateFormat.format(date)
                                if (!col1Order.contains(formattedDate)) {
                                    col1Order.add(formattedDate)
                                }
                                col2ByCol1[formattedDate] = col2Value // Store last Column 2 value
                            } catch (e: Exception) {
                                continue // Skip if not a valid date
                            }
                        }
                    }
                }

                // Get the last 5 unique Column 1 values (or fewer if less than 5)
                val lastFiveCol1 = col1Order.takeLast(5)

                // Get corresponding Column 2 values
                val col2Values = lastFiveCol1.map { col2ByCol1[it] ?: 0.0 }.take(5)

                // Return formatted dates and Column 2 values
                return Pair(lastFiveCol1, col2Values)
            }
        } catch (e: Exception) {
            e.printStackTrace()
            return Pair(emptyList(), emptyList())
        }
    }
}