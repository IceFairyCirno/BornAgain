package com.example.qrcodescanner

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Build
import android.os.Bundle
import android.view.View
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.ContextCompat
import com.example.qrcodescanner.databinding.ActivityMainBinding
import com.journeyapps.barcodescanner.CaptureActivity
import com.journeyapps.barcodescanner.ScanContract
import com.journeyapps.barcodescanner.ScanOptions
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.temporal.ChronoUnit

class MainActivity : AppCompatActivity() {
    private lateinit var excelHelper: ExcelHelper
    private lateinit var binding: ActivityMainBinding

    private val requestPermissionsLauncher = registerForActivityResult(
        ActivityResultContracts.RequestMultiplePermissions()
    ) { permissions ->
        permissions.entries.forEach { (permission, granted) ->
            if (!granted) {
                Toast.makeText(this, "$permission denied", Toast.LENGTH_SHORT).show()
            }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main)
        excelHelper = ExcelHelper(this)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        requestRequiredPermissions()
        excelHelper.initExcel("born_again-db.xlsx")

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
        }
        binding.ProgressButton.setOnClickListener{
            val intent = Intent(this, Progress::class.java)
            startActivity(intent)
        }
        binding.SettingButton.setOnClickListener{
            val intent = Intent(this, SystemConfig::class.java)
            startActivity(intent)
        }

        setupRecentExercise()

        val week = listOf(binding.SundayContainer, binding.MondayContainer, binding.TuesdayContainer, binding.WednesdayContainer, binding.ThursdayContainer, binding.FridayContainer, binding.SaturdayContainer,)
        week[getTodayWeekday()].setCardBackgroundColor(ContextCompat.getColor(this, R.color.blue))
        val weektext = week[getTodayWeekday()].getChildAt(0) as? TextView
        weektext?.setTextColor(android.graphics.Color.WHITE)

        binding.CardioButton.setOnClickListener {
            val builder = AlertDialog.Builder(this)
            builder.setTitle("Cardio")
            builder.setMessage("Enter your calories burned:")

            // Add an edit text for user input
            val input = EditText(this)
            input.width = 10
            builder.setView(input)

            builder.setPositiveButton("Submit") { dialog, which ->
                val userInput = input.text.toString()
                Toast.makeText(this, "Saved Calories Burned: $userInput", Toast.LENGTH_SHORT).show()
            }
            builder.setNegativeButton("Cancel") { dialog, which ->
                dialog.dismiss()
            }

            builder.show()
        }
    }

    private fun requestRequiredPermissions() {
        val permissions = mutableListOf(
            Manifest.permission.CAMERA,
            Manifest.permission.POST_NOTIFICATIONS
        )

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
            permissions.add(Manifest.permission.READ_MEDIA_AUDIO)
        } else {
            permissions.add(Manifest.permission.READ_EXTERNAL_STORAGE)
        }

        val permissionsToRequest = permissions.filter {
            ContextCompat.checkSelfPermission(this, it) != PackageManager.PERMISSION_GRANTED
        }

        if (permissionsToRequest.isNotEmpty()) {
            requestPermissionsLauncher.launch(permissionsToRequest.toTypedArray())
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
    private fun getRecents(): MutableList<List<String>>{
        val recentRecords = excelHelper.processExcelFile("born_again-db.xlsx", "record")
        val recordList: MutableList<List<String>> = mutableListOf()
        for (record in recentRecords) {
            recordList.add(listOf(
                getPassedDate(record[0]),
                record[1],
                "${record[2]} kg",
                "${record[3].trim().toDoubleOrNull()?.toInt() ?: 0} Sets"
            ))
        }
        return recordList
    }
    private fun setupRecentExercise(){

        val recents = getRecents()

        binding.Card1.visibility = View.GONE
        binding.Card2.visibility = View.GONE
        binding.Card3.visibility = View.GONE

        if (recents.size >= 1){
            binding.Card1.visibility = View.VISIBLE
            val inputData = recents[0]
            applyToCard(listOf(binding.Card1Date, binding.Card1Name, binding.Card1Weight, binding.Card1Sets), inputData)
        }
        if (recents.size >= 2){
            binding.Card2.visibility = View.VISIBLE
            val inputData = recents[0]
            applyToCard(listOf(binding.Card1Date, binding.Card1Name, binding.Card1Weight, binding.Card1Sets), inputData)
            val inputData1 = recents[1]
            applyToCard(listOf(binding.Card2Date, binding.Card2Name, binding.Card2Weight, binding.Card2Sets), inputData1)
        }
        if (recents.size >= 3){
            binding.Card3.visibility = View.VISIBLE
            val inputData = recents[0]
            applyToCard(listOf(binding.Card1Date, binding.Card1Name, binding.Card1Weight, binding.Card1Sets), inputData)
            val inputData1 = recents[1]
            applyToCard(listOf(binding.Card2Date, binding.Card2Name, binding.Card2Weight, binding.Card2Sets), inputData1)
            val inputData2 = recents[2]
            applyToCard(listOf(binding.Card3Date, binding.Card3Name, binding.Card3Weight, binding.Card3Sets), inputData2)
        }
    }
    private fun getTodayWeekday(): Int {
        val today = LocalDate.now()
        val dayOfWeek = today.dayOfWeek.value
        return if (dayOfWeek == 7) 0 else dayOfWeek
    }
}