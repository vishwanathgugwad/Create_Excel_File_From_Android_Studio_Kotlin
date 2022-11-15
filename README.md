# Create_Excel_File_From_Android_Studio_Kotlin
Writing a excel file form list in android studio using kotlin 2022

## Easy Steps 

 ## --------------Step 1:  Download jxl latest version and extract your zip file.---------

Firstly you can download the latest version of jxl library and extract it. Copy your JXL jar file and Past it into the Project base libe folder.

[![Download Library](http://www.java2s.com/Code/Jar/j/Downloadjxl26jar.htm)]

 

 ## --------------Step 2: Add your jar file in your module base Gradle file.------------

  Add library and sync your project after successfully syncing your project. open your java file. and implement a writable workbook for creating Xls file.
  
 
  
  ## ---------Step 3:Add this Permission Into your Manifest File.---------------------------------

 ## Code
```Kotlin
// code

    <uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" />
    <uses-permission android:name="android.permission.READ_EXTERNAL_STORAGE" />
    
   ```

---
    
 
 
## ------------Step 4 Now Checkout the MainActivity.kt file.-----------------   

*---------You can able to Create Your #Excel File Into Your Device Extralnal or Internal Storage 



## ------------Step 5 Dont Forget To Use Permission . Without Permission it will not work  ------------------

 
 ## Add Permission
```kotlin
// code
 
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


```

---

## Authors

* **Vishwanath Gugwad** -- [Application Developer](https://github.com/vishwanathgugwad)


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* For any Query Email Me - vishwanath.gugwad@gmail.com
* Thanks



