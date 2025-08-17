package com.example.qrcodescanner

import android.content.Intent
import android.os.Bundle
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import com.example.qrcodescanner.databinding.ActivitySubMachinesBinding

class SubMachines : AppCompatActivity() {

    private lateinit var binding: ActivitySubMachinesBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_sub_machines)
        binding = ActivitySubMachinesBinding.inflate(layoutInflater)
        setContentView(binding.root)

        val machineName = intent.getStringExtra("machineName").toString()
        if ("cable" in machineName){
            binding.Sub1Image.setImageResource(R.drawable.armcurlbar)
            binding.Sub2Image.setImageResource(R.drawable.pulldown)
            binding.Sub3Image.setImageResource(R.drawable.pushdown)
            binding.Sub4Image.setImageResource(R.drawable.rowwithbackextensiontriangle)
            binding.Sub1Text.text = "Arm Curl Bar"
            binding.Sub2Text.text = "Pull Down"
            binding.Sub3Text.text = "Push Down"
            binding.Sub4Text.text = "Row With Extension"
        } else{
            binding.Sub1Image.setImageResource(R.drawable.smithshrug)
            binding.Sub2Image.setImageResource(R.drawable.smithsquat)
            binding.Sub3Image.setImageResource(R.drawable.smithchesspress)
            binding.Sub4Image.setImageResource(R.drawable.smithinclinepress)
            binding.Sub1Text.text = "Shrug (Smith)"
            binding.Sub2Text.text = "Squat (Smith)"
            binding.Sub3Text.text = "Bench Press (Smith)"
            binding.Sub4Text.text = "Incline Press (Smith)"
        }

        binding.Sub1Image.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", binding.Sub1Text.text.toString())
            }
            startActivity(intent)
            finish()
        }
        binding.Sub2Image.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", binding.Sub2Text.text.toString())
            }
            startActivity(intent)
            finish()
        }
        binding.Sub3Image.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", binding.Sub3Text.text.toString())
            }
            startActivity(intent)
            finish()
        }
        binding.Sub4Image.setOnClickListener{
            val intent = Intent(this, MainActivity2::class.java).apply{
                putExtra("exerciseName", binding.Sub4Text.text.toString())
            }
            startActivity(intent)
            finish()
        }

        setupHomeButton()
    }

    private fun setupHomeButton() {
        binding.ReturnHomeButton.setOnClickListener {
            val intent = Intent(this, MainActivity::class.java)
            startActivity(intent)
            finish()
        }
    }
}