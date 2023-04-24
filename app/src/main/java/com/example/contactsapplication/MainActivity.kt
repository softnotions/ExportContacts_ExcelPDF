package com.example.contactsapplication

import android.Manifest
import android.annotation.SuppressLint
import android.content.ActivityNotFoundException
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.provider.ContactsContract
import android.util.Log
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.foundation.Image
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.res.colorResource
import androidx.compose.ui.res.painterResource
import androidx.compose.ui.res.stringResource
import androidx.compose.ui.text.font.Font
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.tooling.preview.Preview
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.core.content.FileProvider
import com.example.contactsapplication.ui.theme.ContactsApplicationTheme
import com.itextpdf.text.Document
import com.itextpdf.text.Paragraph
import com.itextpdf.text.Phrase
import com.itextpdf.text.pdf.PdfPTable
import com.itextpdf.text.pdf.PdfWriter
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException
import java.util.*


class MainActivity : ComponentActivity() {

    private val contacts = ArrayList<Pair<String, String>>()
    private val requestCodeReadContacts = 100


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        // Checks if the app has permission to read contacts. If not, requests the permission.
        // If the permission has already been granted, retrieves the contact list.
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.READ_CONTACTS)
            != PackageManager.PERMISSION_GRANTED
        ) {
            // Permission has already been Not-granted
            ActivityCompat.requestPermissions(
                this,
                arrayOf(Manifest.permission.READ_CONTACTS),
                requestCodeReadContacts
            )
        } else {
            // Permission has already been granted
            getContactList()
        }


        if (ContextCompat.checkSelfPermission(this, Manifest.permission.MANAGE_EXTERNAL_STORAGE)
            != PackageManager.PERMISSION_GRANTED
        ) {
            // Permission has already been Not-granted
            ActivityCompat.requestPermissions(
                this,
                arrayOf(Manifest.permission.MANAGE_EXTERNAL_STORAGE),
                requestCodeReadContacts
            )
        }

        // Sets the content of the app's UI to a ContactsScreen composable function wrapped in a ContactsApplicationTheme
        setContent {
            ContactsApplicationTheme {
                // Creates a surface container using the 'background' color from the theme,
                // and passes in the list of contacts to the ContactsScreen composable function.
                ContactsScreen(contacts)
            }
        }
    }

    /* This function retrieves a list of contacts and adds them to a list of pairs
    where each pair contains the contact name and phone number.*/
    @SuppressLint("Range")
    private fun getContactList() {
        val cursor = contentResolver.query(
            ContactsContract.CommonDataKinds.Phone.CONTENT_URI,
            null,
            null,
            null,
            ContactsContract.CommonDataKinds.Phone.DISPLAY_NAME + getString(R.string.asc)
        )
        while (cursor?.moveToNext() == true) {
            val name =
                cursor.getString(cursor.getColumnIndex(ContactsContract.CommonDataKinds.Phone.DISPLAY_NAME))
            val number =
                cursor.getString(cursor.getColumnIndex(ContactsContract.CommonDataKinds.Phone.NUMBER))
            contacts.add(Pair(name, number))
        }
        cursor?.close()
    }
}


@Composable
fun ContactsScreen(contacts: ArrayList<Pair<String, String>>) {

    val proximaNovaFontFamily = FontFamily(
        Font(R.font.proxima_nova, FontWeight.Normal),
    )

    Column(
        modifier = Modifier.fillMaxSize(),

        ) {
        // Creates a surface container with a white background color, and adds a Greeting composable
        // function as its content, which is passed the list of contacts as a parameter.

        Surface(
            modifier = Modifier
                .fillMaxWidth()
                .height(80.dp),
            color = colorResource(id = R.color.Header)
        ) {


            Text(
                modifier = Modifier
                    .fillMaxSize()
                    .wrapContentSize(align = Alignment.Center),
                text = stringResource(R.string.home),
                color = Color.White,
                textAlign = TextAlign.Center,
                fontSize = 24.sp,
                fontFamily = proximaNovaFontFamily
            )
        }


        Surface(
            modifier = Modifier
                .fillMaxWidth()
                .padding(top = 60.dp)
                .height(100.dp),
            color = Color.Transparent
        ) {
            Box(
                modifier = Modifier.fillMaxSize(),
                contentAlignment = Alignment.Center
            ) {
                Image(
                    modifier = Modifier.fillMaxWidth().fillMaxHeight().padding(all = 10.dp),
                    painter = painterResource(id = R.drawable.logo),
                    contentDescription = stringResource(R.string.image)
                )
            }
        }


        Surface(
            modifier = Modifier
                .fillMaxWidth()
                .padding(top = 40.dp, start = 40.dp, end = 40.dp)
                .height(80.dp),
            color = colorResource(id = R.color.Button),
            shape = RoundedCornerShape(8.dp),
        ) {
            ExcelButtonCall(contacts,stringResource(id = R.string.export_contacts),proximaNovaFontFamily)
        }
        // Creates another surface container with a white background color, and adds a Greeting2
        // composable function as its content, which is also passed the list of contacts as a parameter.
        Surface(
            modifier = Modifier
                .fillMaxWidth()
                .padding(top = 40.dp, start = 40.dp, end = 40.dp)
                .height(80.dp),
            color = colorResource(id = R.color.Button),
            shape = RoundedCornerShape(8.dp),
        ) {
            PdfButtonCall(contacts, stringResource(id = R.string.exportPdf_contacts),proximaNovaFontFamily)
        }

        Column(
            modifier = Modifier.fillMaxSize(),
            verticalArrangement = Arrangement.Bottom,
            horizontalAlignment = Alignment.CenterHorizontally
        ) {
            Surface(
                modifier = Modifier
                    .fillMaxWidth()
                    .height(28.dp),
                color = colorResource(id = R.color.Footer)
            ) {
                Text(
                    modifier = Modifier
                        .fillMaxSize()
                        .wrapContentSize(align = Alignment.Center),
                    text = stringResource(R.string.version),
                    color = Color.White,
                    textAlign = TextAlign.Center,
                    fontSize = 12.sp,
                    fontFamily = proximaNovaFontFamily
                )
            }
        }
    }
}


@Composable
fun ExcelButtonCall(
    name: ArrayList<Pair<String, String>>,
    stringResource: String,
    proximaNovaFontFamily: FontFamily
) {
    val context = LocalContext.current
    // Creates a button with a blue background color and white text, and sets its onClick listener to
    // call the createExcelSpreadsheet function and pass in the current context and the list of names
    // as parameters.

    Box(
        modifier = Modifier.clickable {
            // Handle click event here
            createExcelSpreadsheet(context, name)
        }
    ) {
        Row(verticalAlignment = Alignment.CenterVertically) {
            Image(
                modifier = Modifier
                    .padding(top = 10.dp, start = 10.dp, bottom = 10.dp),
                painter = painterResource(id = R.drawable.excel),
                contentDescription = stringResource(R.string.image)
            )

            Text(
                modifier = Modifier
                    .fillMaxSize()
                    .wrapContentSize(align = Alignment.Center)
                    .padding(end = 30.dp),
                text = stringResource,
                color = Color.White,
                textAlign = TextAlign.Center,
                fontSize = 24.sp,
                fontFamily = proximaNovaFontFamily
            )
        }
    }
}

@Composable
fun PdfButtonCall(
    name: ArrayList<Pair<String, String>>,
    string: String,
    proximaNovaFontFamily: FontFamily
) {
    val context = LocalContext.current
    // Creates a button with a blue background color and white text, and sets its onClick listener to
    // call the createPdfFromArrayList function and pass in the list of names and the current context
    // as parameters.

    Box(
        modifier = Modifier.clickable {
            // Handle click event here
            createPdfFromArrayList(name, context)
        }
    ) {
        Row(verticalAlignment = Alignment.CenterVertically) {
            Image(
                modifier = Modifier
                    .padding(top = 10.dp, start = 10.dp, bottom = 10.dp),
                painter = painterResource(id = R.drawable.pdf),
                contentDescription = stringResource(R.string.image)
            )

            Text(
                modifier = Modifier
                    .fillMaxSize()
                    .wrapContentSize(align = Alignment.Center)
                    .padding(end = 30.dp),
                text = string,
                color = Color.White,
                textAlign = TextAlign.Center,
                fontSize = 24.sp,
                fontFamily = proximaNovaFontFamily
            )
        }
    }
}


private fun createExcelSpreadsheet(context: Context, name: ArrayList<Pair<String, String>>) {
    // Create a new Excel workbook
    val workbook = XSSFWorkbook()
    // Create a new sheet in the workbook
    val sheet = workbook.createSheet(context.getString(R.string.mySheet))

    // Add some data to the sheet
    val headers = arrayOf(context.getString(R.string.name), context.getString(R.string.number))

    // Add the headers to the first row
    val headerRow = sheet.createRow(0)
    headerRow.heightInPoints = 40f
    for (i in headers.indices) {
        val cell = headerRow.createCell(i)
        cell.setCellValue(headers[i])
    }

    // Add the data to subsequent rows
    for (i in name.indices) {
        val dataRow = sheet.createRow(i + 1)
        dataRow.heightInPoints = 40f
        sheet.setColumnWidth(0, 15 * 500)
        sheet.setColumnWidth(1, 15 * 500)
        val nameCell = dataRow.createCell(0)
        val numberCell = dataRow.createCell(1)
        nameCell.setCellValue(name[i].first)
        numberCell.setCellValue(name[i].second)
    }

    // Write the workbook to a file
    val file = File(context.getExternalFilesDir(null), context.getString(R.string.MyWorkbook))
    val outputStream = FileOutputStream(file)
    workbook.write(outputStream)
    outputStream.close()

    // Open the Excel with a Excel viewer app
    val intent = Intent(Intent.ACTION_VIEW)
    val photoURI = FileProvider.getUriForFile(
        Objects.requireNonNull(context),
        BuildConfig.APPLICATION_ID + context.getString(R.string.provider),
        file
    )
    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    intent.setDataAndType(photoURI, context.getString(R.string.excel))
    try {
        context.startActivity(intent)
    } catch (e: ActivityNotFoundException) {
        e.printStackTrace()
    }

}


@Throws(IOException::class)
fun createPdfFromArrayList(
    data: ArrayList<Pair<String, String>>,
    context: Context
) {
    try {
        // Create a new document
        val document = Document()
        val file = File(context.getExternalFilesDir(null), context.getString(R.string.my_pdf_file))
        Log.e("File",file.absoluteFile.toString())
        val outputStream = FileOutputStream(file)
        // Create a new PDF writer
        PdfWriter.getInstance(document, outputStream)
        // Open the document
        document.open()
        // Create a table with two columns
        val table = PdfPTable(2)
        // Add headers to the table
        table.addCell(Phrase(context.getString(R.string.name)))
        table.addCell(Phrase(context.getString(R.string.number)))
        // Add data to the table
        for (item in data) {
            table.addCell(Paragraph(item.first))
            table.addCell(Paragraph(item.second))
        }

        // Add the table to the document
        document.add(table)

        // Close the document
        document.close()

        // Open the PDF with a PDF viewer app
        val intent = Intent(Intent.ACTION_VIEW)
        val photoURI = FileProvider.getUriForFile(
            Objects.requireNonNull(context),
            BuildConfig.APPLICATION_ID + context.getString(R.string.provider),
            file
        )
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        intent.setDataAndType(photoURI, context.getString(R.string.pdf))
        try {
            context.startActivity(intent)
        } catch (e: ActivityNotFoundException) {
            e.printStackTrace()
        }
    } catch (e: Exception) {
        e.printStackTrace()
    }
}

@Preview
@Composable
fun PreviewFunction() {
    val contacts = ArrayList<Pair<String, String>>()
    ContactsScreen(contacts)
}
