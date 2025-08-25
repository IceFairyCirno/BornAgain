package com.example.qrcodescanner

import android.app.NotificationChannel
import android.app.NotificationManager
import android.content.Context
import android.content.Intent
import android.os.Build
import android.os.Bundle
import android.os.CountDownTimer
import android.widget.Toast
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.NotificationCompat
import com.example.qrcodescanner.databinding.ActivityMain2Binding
import com.journeyapps.barcodescanner.CaptureActivity
import com.journeyapps.barcodescanner.ScanContract
import com.journeyapps.barcodescanner.ScanOptions
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.time.LocalDate
import java.time.LocalTime
import java.time.format.DateTimeFormatter

class MainActivity2 : AppCompatActivity() {

    private lateinit var binding: ActivityMain2Binding
    private lateinit var excelHelper: ExcelHelper

    var countDownTimer: CountDownTimer? = null
    private var currentTime: Long = 0
    private var currentSet = 1

    private val channelId = "default_channel"
    private val notificationId = 1

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        tempDeleteData()
        enableEdgeToEdge()
        setContentView(R.layout.activity_main2)
        binding = ActivityMain2Binding.inflate(layoutInflater)
        setContentView(binding.root)
        excelHelper = ExcelHelper(this)
        val exerciseName = intent.getStringExtra("exerciseName").toString()
        currentTime = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "A1").toLong()
        createNotificationChannel()
        binding.SetNumber.text = "Set $currentSet/${excelHelper.getCellExcel("born_again-db.xlsx", "settings", "B1")}"

        val formattedName = convertExerciseName(exerciseName)

        setupExerciseName(formattedName)
        setupTimer()
        setupCompeleteSetButton(formattedName)
        setupWeightRepsButtons()
        setupHomeButton()
        setupScanNextButton()
        setupSaveButton()
        setupHighestRecord(formattedName)
    }

    private fun setupExerciseName(input: String){
        binding.ExerciseName.text = input
    }

    private fun convertExerciseName(input: String): String{
        val formatted_exerciseName = input
            .split("_")
            .joinToString(" ") { word ->
                word.replaceFirstChar {
                    if (it.isLowerCase()) it.titlecase() else it.toString()
                }
            }
        return formatted_exerciseName
    }

    private fun setupWeightRepsButtons(){
        binding.WeightIncreaseButton.setOnClickListener{
            val currentWeight = binding.WeightDisplay.text.toString()
            val currentNumber = currentWeight.toFloatOrNull() ?: 0.0f
            val newNumber = currentNumber + 2.5f
            binding.WeightDisplay.text = String.format("%.1f", newNumber)
        }
        binding.WeightDecreaseButton.setOnClickListener{
            val currentWeight = binding.WeightDisplay.text.toString()
            val currentNumber = currentWeight.toFloatOrNull() ?: 0.0f
            val newNumber = currentNumber - 2.5f
            binding.WeightDisplay.text = String.format("%.1f", newNumber)
        }
        binding.RepsIncreaseButton.setOnClickListener{
            val currentReps = binding.RepsDisplay.text.toString()
            val currentNumber = currentReps.toIntOrNull() ?: 0
            val newNumber = currentNumber + 1
            binding.RepsDisplay.text = String.format("%02d", newNumber)
        }
        binding.RepsDecreaseButton.setOnClickListener{
            val currentReps = binding.RepsDisplay.text.toString()
            val currentNumber = currentReps.toIntOrNull() ?: 0
            val newNumber = currentNumber - 1
            binding.RepsDisplay.text = String.format("%02d", newNumber)
        }
    }

    private fun updateCompeletedSet(SetNum: String, Data: String){
        if (binding.Set1Text.text.toString().isEmpty()) {
            binding.Set1Text.text = "Set $SetNum"
            binding.Set1Load.text= Data
            binding.Set1Tick.text = "✓"

            binding.Set2Load.text = "" //For removing the "Nothing Here Yet" text
        } else if (binding.Set2Text.text.toString().isEmpty()) {
            binding.Set2Text.text = "Set $SetNum"
            binding.Set2Load.text= Data
            binding.Set2Tick.text = "✓"
        } else if (binding.Set3Text.text.toString().isEmpty()) {
            binding.Set3Text.text = "Set $SetNum"
            binding.Set3Load.text= Data
            binding.Set3Tick.text = "✓"
        } else {
            binding.Set1Text.text = binding.Set2Text.text
            binding.Set1Load.text = binding.Set2Load.text

            binding.Set2Text.text = binding.Set3Text.text
            binding.Set2Load.text = binding.Set3Load.text

            binding.Set3Text.text = "Set $SetNum"
            binding.Set3Load.text = Data
        }
    }

    private fun setupHomeButton() {
        binding.ReturnHomeButton.setOnClickListener {
            val intent = Intent(this, MainActivity::class.java)
            startActivity(intent)
            finish()
        }
    }

    private fun setupScanNextButton(){
        binding.NextExerciseButton.setOnClickListener{
            excelHelper.copyExcel("temp.xlsx", "record", "born_again-db.xlsx", "record")
            showMsg("Exercise Saved!")
            val options = ScanOptions()
            options.setPrompt("Scan a QR Code")
            options.setBeepEnabled(true)
            options.setOrientationLocked(true)
            options.setCaptureActivity(CaptureActivity::class.java)
            barcodeLauncher.launch(options)
        }
    }

    private val barcodeLauncher = registerForActivityResult(ScanContract()){result ->
        if (result.contents!=null){
            if ("gym" in result.contents) {
                val exerciseName = (result.contents).split("/").last()
                val intent: Intent
                if ("cable" in result.contents || "multi" in result.contents) {
                    intent = Intent(this, SubMachines::class.java).apply {
                        putExtra("machineName", exerciseName)
                    }
                } else {
                    intent = Intent(this, MainActivity2::class.java).apply {
                        putExtra("exerciseName", exerciseName)
                    }
                }
                startActivity(intent)
                finish()
            }
        } else{
            Toast.makeText(this, "Invalid QR Code", Toast.LENGTH_SHORT).show()
        }
    }

    private fun tempSaveData(rowData: List<String>){
        excelHelper.initExcel("temp.xlsx")
        excelHelper.modifyExcelFromBottom("temp.xlsx", "record", rowData)
    }

    private fun getCurrentDateFormatted(): String {
        val currentDate = LocalDate.now()
        val formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd")
        return currentDate.format(formatter)
    }
    private fun getCurrentTimeInHHMM(): String {
        val currentTime = LocalTime.now()
        val formatter = DateTimeFormatter.ofPattern("HH:mm")
        return currentTime.format(formatter)
    }


    private fun setupSaveButton(){
        binding.SaveExerciseButton.setOnClickListener{
            excelHelper.copyExcel("temp.xlsx", "record", "born_again-db.xlsx", "record")
            showMsg("Exercise Saved!")
            val intent = Intent(this, MainActivity::class.java)
            startActivity(intent)
            finish()
        }
    }

    private fun showMsg(message: String) {
        Toast.makeText(this, message, Toast.LENGTH_SHORT).show()
    }

    private fun tempDeleteData(){
        val file = File(this.filesDir, "temp.xlsx")
        if (file.exists()) {
            file.delete()
            println("Deleted temp")
        } else {
            println("No temp need to be delete")
        }
    }

    private fun createNotificationChannel() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
            val name = "Default Channel"
            val descriptionText = "Channel for default notifications"
            val importance = NotificationManager.IMPORTANCE_DEFAULT
            val channel = NotificationChannel(channelId, name, importance).apply {
                description = descriptionText
            }

            val notificationManager: NotificationManager =
                getSystemService(Context.NOTIFICATION_SERVICE) as NotificationManager
            notificationManager.createNotificationChannel(channel)
        }
    }

    private fun showNotification() {
        val builder = NotificationCompat.Builder(this, channelId)
            .setSmallIcon(R.drawable.tick)
            .setContentTitle("Time's up!")
            .setContentText("Continue your work")
            .setPriority(NotificationCompat.PRIORITY_DEFAULT)
            .setAutoCancel(true)

        val notificationManager: NotificationManager =
            getSystemService(Context.NOTIFICATION_SERVICE) as NotificationManager
        notificationManager.notify(notificationId, builder.build())
    }

    private fun renewCompletedSet(){
        val displayData = collectSetData()
        val setNum = currentSet.toString()
        updateCompeletedSet(setNum, displayData)
        currentSet += 1
        binding.SetNumber.text = "Set $currentSet/${excelHelper.getCellExcel("born_again-db.xlsx", "settings", "B1")}"
    }

    private fun collectSetData(): String{
        val currentWeight = binding.WeightDisplay.text.toString()
        val currentReps = (binding.RepsDisplay.text.toString()).toInt().toString()
        val outputText = "$currentWeight kg × $currentReps"
        return outputText
    }

    private fun setupTimer(){
        val totalTime: Long = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "A1").toLong()
        binding.Timer.text = String.format("%02d:%02d", totalTime / 1000 / 60, totalTime / 1000 % 60)
        binding.IncreaseTimerButton.setOnClickListener {
            currentTime = adjustTimer(currentTime, 5000){ newTime -> currentTime = newTime }
        }
        binding.DecreaseTimerButton.setOnClickListener {
            currentTime = adjustTimer(currentTime, -5000) { newTime -> currentTime = newTime }
        }
    }

    private fun startTimer(time: Long, onFinish: (Long) -> Unit) {
        countDownTimer?.cancel()

        countDownTimer = object : CountDownTimer(time, 1000) {
            override fun onTick(millisUntilFinished: Long) {
                val seconds = millisUntilFinished / 1000
                binding.Timer.text = String.format("%02d:%02d", seconds / 60, seconds % 60)
                onFinish(millisUntilFinished)
            }

            override fun onFinish() {
                showNotification()
                binding.Timer.text = "00:00"
            }
        }.start()
    }

    private fun adjustTimer(currentTime: Long, adjustment: Long, onFinish: (Long) -> Unit): Long {
        var newTime = currentTime + adjustment
        if (newTime < 1000) {
            newTime = 1 // Minimum 1 second
        }
        startTimer(newTime) { updatedTime -> onFinish(updatedTime) }
        return newTime
    }

    private fun setupCompeleteSetButton(formattedName: String){
        binding.SetConfirmButton.setOnClickListener {
            if (binding.RepsDisplay.text.toString().toInt() == 0){
                showMsg("You can't have 0 reps")
            } else{
                startTimer(excelHelper.getCellExcel("born_again-db.xlsx", "settings", "A1").toLong()) { newTime -> currentTime = newTime }
                val toTempSaveRow = listOf(
                    getCurrentDateFormatted(),
                    formattedName,
                    binding.WeightDisplay.text.toString(),
                    binding.RepsDisplay.text.toString().toInt().toString(),
                    getCurrentTimeInHHMM()
                )
                tempSaveData(toTempSaveRow)
                renewCompletedSet()
            }
        }
    }

    private fun setupHighestRecord(name: String){
        try {
            val file = File(this.filesDir, "born_again-db.xlsx")

            FileInputStream(file).use { fis ->
                val workbook = XSSFWorkbook(fis)
                val sheet = workbook.getSheetAt(0)

                val numbers = mutableListOf<Double>()
                for (row in sheet) {
                    val nameCell = row.getCell(1)
                    val numberCell = row.getCell(2)

                    if (nameCell?.stringCellValue == name && numberCell != null) {
                        val number = when (numberCell.cellType) {
                            CellType.NUMERIC -> numberCell.numericCellValue
                            CellType.STRING -> numberCell.stringCellValue.toDoubleOrNull() ?: continue
                            else -> continue
                        }
                        numbers.add(number)
                    } else{
                        numbers.add(0.0)
                    }
                }
                binding.HighestRecord.text = ("Highest Record: "+(numbers.maxOrNull()).toString()+"kg")
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }
}