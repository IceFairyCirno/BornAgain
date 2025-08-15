package com.example.qrcodescanner

import android.content.Intent
import android.os.Bundle
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import com.example.qrcodescanner.databinding.ActivityMain3Binding

class SubMachines : AppCompatActivity() {

    private lateinit var binding: ActivityMain3Binding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main3)
        binding = ActivityMain3Binding.inflate(layoutInflater)
        setContentView(binding.root)

        binding.PullDown.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", "Pull Down")
            }
            startActivity(intent)
            finish()
        }
        binding.PushDown.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", "Push Down")
            }
            startActivity(intent)
            finish()
        }
        binding.ArmCurlBar.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", "Arm Curl Bar")
            }
            startActivity(intent)
            finish()
        }
        binding.RowWithExtensionTriangle.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", "Row With Extension")
            }
            startActivity(intent)
            finish()
        }
    }
}