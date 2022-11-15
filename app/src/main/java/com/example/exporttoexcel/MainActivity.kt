package com.example.exporttoexcel

import android.Manifest
import android.content.pm.PackageManager
import android.os.Bundle
import android.os.Environment
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.example.exporttoexcel.databinding.ActivityMainBinding
import com.example.exporttoexcel.model.User
import jxl.Workbook
import jxl.WorkbookSettings
import jxl.write.Label
import jxl.write.WritableWorkbook
import java.io.File
import java.util.*
import kotlin.collections.ArrayList


class MainActivity : AppCompatActivity() {


    private lateinit var binding: ActivityMainBinding
    private var workbook: WritableWorkbook? = null
    private val permissionCode = 101
    private var isPermissionGranted: Boolean = false


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)
        checkPermission()
    }

    private fun checkPermission() {
        if (ContextCompat.checkSelfPermission(
                this, Manifest.permission.WRITE_EXTERNAL_STORAGE
            ) !=
            PackageManager.PERMISSION_GRANTED && ContextCompat.checkSelfPermission(
                this, Manifest.permission.READ_EXTERNAL_STORAGE
            ) !=
            PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                this,
                arrayOf(
                    Manifest.permission.WRITE_EXTERNAL_STORAGE,
                    Manifest.permission.READ_EXTERNAL_STORAGE
                ), 101
            )
            return
        }
        isPermissionGranted = true
        setUpClickListener()
    }

    override fun onRequestPermissionsResult(
        requestCode: Int, permissions: Array<String?>,
        grantResults: IntArray,
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == permissionCode) {
            if (grantResults.isNotEmpty() && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                isPermissionGranted = true
                Toast.makeText(this, "Permission are granted", Toast.LENGTH_SHORT).show()
                setUpClickListener()
            }
        }
    }

    private fun setUpClickListener() {
        binding.createFile.setOnClickListener {
            if (isPermissionGranted) {
                createExcelSheet()
            }
        }


    }

    private fun createExcelSheet() {
        //File futureStudioIconFile = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), fileName);
        val csvFile = "${binding.fileName.text}.xls"
        val futureStudioIconFile = File(
            Environment
                .getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
                .toString() + "/" + csvFile
        )
        val wbSettings = WorkbookSettings()
        wbSettings.locale = Locale("en", "EN")
        try {
            workbook = Workbook.createWorkbook(futureStudioIconFile, wbSettings)
            createFirstSheet()
            //closing cursor
            workbook?.let {
                it.write()
                it.close()
                Toast.makeText(this, "Successfully created excel file", Toast.LENGTH_LONG).show()
            }
        } catch (e: java.lang.Exception) {
            e.printStackTrace()
        }
    }

    private fun createFirstSheet() {
        try {
            val listData = ArrayList<User>()
            listData.add(User("name1", "name1@gmail.com"))
            listData.add(User("name1", "name1@gmail.com"))
            listData.add(User("name1", "name1@gmail.com"))
            listData.add(User("name1", "name1@gmail.com"))
            //excel sheet name
            val workSheet = workbook?.createSheet("UserSheet", 0)
            workSheet?.apply {
                addCell(Label(0, 0, "firstName"))
                addCell(Label(1, 0, "email"))
            }
            for (i in listData.indices) {
                workSheet?.apply {
                    addCell((Label(0, i + 1, listData[i].name)))
                    addCell((Label(1, i + 1, listData[i].email)))
                }
            }

        } catch (e: Exception) {
            e.printStackTrace()
        }


    }


}