package com.example.qrcodescanner

import android.content.Intent
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.media.MediaMetadataRetriever
import android.media.MediaPlayer
import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.widget.SeekBar
import android.widget.Toast
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import com.example.qrcodescanner.databinding.ActivityMusicPlayerBinding
import java.io.File
import androidx.core.net.toUri

class MusicPlayer : AppCompatActivity() {

    private lateinit var binding: ActivityMusicPlayerBinding

    private var mediaPlayer: MediaPlayer? = null
    private lateinit var songs: List<File>
    private var currentSongIndex = 0
    private var isPlaying = false
    private var isPaused = false

    override fun onCreate(savedInstanceState: Bundle?) {
        Log.d("Music", "OnCreate Called")
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_music_player)

        binding = ActivityMusicPlayerBinding.inflate(layoutInflater)
        setContentView(binding.root)

        loadSongs()
        updateSongDisplay()
        binding.PauseButton.setOnClickListener{
            playPauseSong()
        }
        binding.SkipButton.setOnClickListener{
            nextSong()
        }
        binding.RewindButton.setOnClickListener{
            lastSong()
        }

        binding.ProgressBar.setOnSeekBarChangeListener(object : SeekBar.OnSeekBarChangeListener {
            override fun onProgressChanged(seekBar: SeekBar?, progress: Int, fromUser: Boolean) {
                if (fromUser) {
                    mediaPlayer?.seekTo(progress)
                }
            }
            override fun onStartTrackingTouch(seekBar: SeekBar?) {
                mediaPlayer?.pause()
                binding.PauseButton.setImageResource(R.drawable.pauseicon)
            }
            override fun onStopTrackingTouch(seekBar: SeekBar?) {
                mediaPlayer?.start()
                binding.PauseButton.setImageResource(R.drawable.playicon)
                updateSeekBar()
            }
        })

        setupHomeButton()
    }

    private fun loadSongs() {
        Log.d("Music Player", "Loading songs")
        val musicDir = File("/storage/emulated/0/Music_App")
        songs = musicDir.listFiles { _, name -> name.endsWith(".mp3") }?.toList() ?: emptyList()

        if (songs.isNotEmpty()) {
            Log.d("Music Player", "val songs isn't empty!")
            Toast.makeText(this, "Loaded ${songs.size} songs", Toast.LENGTH_SHORT).show()
        }
    }

    private fun playSongFromStart() {
        if (songs.isNotEmpty()) {
            mediaPlayer?.release()
            mediaPlayer = MediaPlayer.create(this, Uri.fromFile(songs[currentSongIndex]))
            binding.ProgressBar.max = mediaPlayer?.duration!!
            mediaPlayer!!.setOnPreparedListener {
                updateSeekBar()
            }
            binding.TotalTime.text = formatDuration(mediaPlayer?.duration!!)
            mediaPlayer?.start()
            updateSongDisplay()
            mediaPlayer?.setOnCompletionListener {
                nextSong()
            }
        }
    }

    private fun displayCoverArt(path: String) {
        val retriever = MediaMetadataRetriever()
        retriever.setDataSource(this, path.toUri())
        val art: ByteArray? = retriever.embeddedPicture
        if (art != null) {
            val coverArt: Bitmap = BitmapFactory.decodeByteArray(art, 0, art.size)
            binding.MusicCoverArt.setImageBitmap(coverArt)
        } else {
            binding.MusicCoverArt.setImageResource(R.drawable.musicicon)
        }
        retriever.release()
    }

    private fun updateSongDisplay() {
        val songValues = extractArtistAndSongName((songs[currentSongIndex].name).toString())
        binding.SongName.text = songValues?.second
        binding.ArtistName.text = songValues?.first
        displayCoverArt("/storage/emulated/0/Music_App/${songs[currentSongIndex].name}")
    }

    private fun nextSong() {
        if (songs.isNotEmpty()) {
            binding.PauseButton.setImageResource(R.drawable.playicon)
            currentSongIndex = (currentSongIndex + 1) % songs.size
            isPaused = false
            playSongFromStart()
        }
    }
    private fun lastSong() {
        if (songs.isNotEmpty()) {
            binding.PauseButton.setImageResource(R.drawable.playicon)

            currentSongIndex = if (currentSongIndex > 0) {
                currentSongIndex - 1
            } else {
                songs.size - 1
            }
            isPaused = false
            playSongFromStart()
        }
    }

    private fun playPauseSong() {
        isPlaying = !isPlaying
        if (!isPlaying) {
            mediaPlayer?.pause()
            binding.PauseButton.setImageResource(R.drawable.pauseicon)
            isPaused = !isPaused
        } else{
            binding.PauseButton.setImageResource(R.drawable.playicon)
            if (isPaused){
                mediaPlayer?.start()
                isPaused = !isPaused
                updateSeekBar()
            }else{
                playSongFromStart()
            }
        }
    }

    private fun extractArtistAndSongName(fileName: String): Pair<String, String>? {
        val nameWithoutExtension = fileName.removeSuffix(".mp3")

        val regex = Regex("""\[(.+?)\]\s(.+)""")
        val matchResult = regex.find(nameWithoutExtension)

        return if (matchResult != null) {
            val artistName = matchResult.groups[1]?.value ?: ""
            val songName = matchResult.groups[2]?.value ?: ""
            Pair(artistName, songName)
        } else {
            null
        }
    }

    private fun updateSeekBar() {
        binding.ProgressBar.progress = mediaPlayer?.currentPosition!!
        binding.CurrentProgress.text = formatDuration(mediaPlayer?.currentPosition!!)
        if (mediaPlayer?.isPlaying!!) {
            binding.ProgressBar.postDelayed({ updateSeekBar() }, 1000) // Update every second
        }
    }

    private fun formatDuration(duration: Int): String {
        val minutes = (duration / 1000) / 60
        val seconds = (duration / 1000) % 60
        return String.format("%02d:%02d", minutes, seconds)
    }

    private fun setupHomeButton() {
        binding.ReturnHomeButton.setOnClickListener {
            mediaPlayer?.let {
                if (isPlaying) {
                    it.pause()
                    it.release()
                }else{
                    it.stop()
                    it.release()
                }
            }
            mediaPlayer = null // Prevent further calls on released MediaPlayer

            // Use flags to clear the activity stack if needed
            val intent = Intent(this, MainActivity::class.java).apply {
                flags = Intent.FLAG_ACTIVITY_CLEAR_TOP or Intent.FLAG_ACTIVITY_NEW_TASK
            }
            startActivity(intent)
            finish()
        }
    }

    private fun printDirectoryContents() {
        // Specify the directory path
        val musicDir = File("/storage/emulated/0/Music_App")

        // Check if the directory exists
        if (musicDir.exists() && musicDir.isDirectory) {
            // List all files in the directory
            val files = musicDir.listFiles()

            if (files != null && files.isNotEmpty()) {
                Log.i("DirectoryContents", "Contents of ${musicDir.absolutePath}:")
                for (file in files) {
                    Log.i("DirectoryContents", file.name) // Print file name
                }
            } else {
                Log.i("DirectoryContents", "The directory is empty.")
            }
        } else {
            Log.e("DirectoryContents", "Directory does not exist: ${musicDir.absolutePath}")
        }
    }
}