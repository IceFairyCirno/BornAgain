package com.example.qrcodescanner

import android.content.Intent
import android.os.Bundle
import android.util.Log
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import com.example.qrcodescanner.databinding.ActivitySystemConfigBinding

class SystemConfig : AppCompatActivity() {
    private lateinit var excelHelper: ExcelHelper
    private lateinit var binding: ActivitySystemConfigBinding

    private var restTime: String = ""
    private var totalSets: String = ""

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_system_config)
        binding = ActivitySystemConfigBinding.inflate(layoutInflater)
        setContentView(binding.root)
        excelHelper = ExcelHelper(this)
        setupHomeButton()

        // Rest Time
        var restTime = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "A1")
        if (restTime == ""){
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "A1", "60000")
            restTime = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "A1")
            binding.SettingRestTime.text = (restTime.toInt()/60000.0).toString()+" min"
        } else{
            binding.SettingRestTime.text = (restTime.toInt()/60000.0).toString()+" min"
        }

        binding.IncreaseRestTimeButton.setOnClickListener{
            restTime = (restTime.toInt() + 30000).toString()
            binding.SettingRestTime.text = (restTime.toInt()/60000.0).toString()+" min"
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "A1", restTime)
        }
        binding.DecreaseRestTimeButton.setOnClickListener{
            if (restTime.toInt() - 30000 <= 0){
                restTime = "0"
            } else{
                restTime = (restTime.toInt()-30000).toString()
            }
            binding.SettingRestTime.text = (restTime.toInt()/60000.0).toString()+" min"
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "A1", restTime)
        }

        // Total Sets
        var totalSets = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "B1")
        if (totalSets == ""){
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "B1", "4")
            totalSets = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "B1")
            binding.SettingTotalSet.text = totalSets+" Sets"
        } else{
            binding.SettingTotalSet.text = totalSets+" Sets"
        }
        binding.IncreaseTotalSetButton.setOnClickListener{
            totalSets = (totalSets.toInt()+1).toString()
            binding.SettingTotalSet.text = totalSets+" Sets"
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "B1", totalSets)
        }
        binding.DecreaseTotalSetButton.setOnClickListener{
            if (totalSets.toInt()-1 < 0){
                totalSets = "0"
            } else{
                totalSets = (totalSets.toInt()-1).toString()
            }
            binding.SettingTotalSet.text = totalSets+" Sets"
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "B1", totalSets)
        }

        //MET Values
        var met = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "C1")
        if (met == ""){
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "C1", "3")
            met = excelHelper.getCellExcel("born_again-db.xlsx", "settings", "C1")
            binding.SettingMET.text = met
        } else{
            binding.SettingMET.text = met
        }
        binding.IncreaseMETButton.setOnClickListener{
            met = (met.toInt()+1).toString()
            binding.SettingMET.text = met
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "C1", met)
        }
        binding.DecreaseMETButton.setOnClickListener{
            if (met.toInt()-1 <= 1){
                met = "1"
            } else{
                met = (met.toInt()-1).toString()
            }
            binding.SettingMET.text = met
            excelHelper.modifyCellExcel("born_again-db.xlsx", "settings", "C1", met)
        }

    }
    private fun setupHomeButton() {
        binding.ReturnHomeButton.setOnClickListener {
            val intent = Intent(this, MainActivity::class.java)
            startActivity(intent)
            finish()
        }
    }

}