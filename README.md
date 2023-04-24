# Android Application: ExcelPDF-Contacts
This Android application fetches user contacts from their mobile device using content provider and
exports it into a .xlsx and PDF file. The app also lists the installed apps that support .xlsx and
PDF reader, allowing the user to open the exported file with contact data.

**Created Date**: 11-04-2023

## Technologies Used
- IDE: Android Studio Electric Eel | 2022.1.1 Patch 2
- Minimum SDK Version: 26, Maximum SDK Version: 33
- Design: Jetpack Compose
- Programming Language: Kotlin

## Third-Party Libraries
- org.apache.poi:poi-ooxml:5.2.0 (for creating Excel files)
- com.itextpdf:itextpdf:5.5.13.2 (for creating PDF files)

## Features
1. Fetches user contacts from the mobile device using content provider
2. Exports the contact data into .xlsx and PDF files
3. Lists the installed apps that support .xlsx and PDF reader
4. Opens the exported file with contact data using the user's preferred app

## To get started with the project

1. Clone the repository to their local machine
2. Open the project in Android Studio
3. Build and run the application on an Android device or emulator

## Usage 

To use the app, simply install it on your Android device and grant it the necessary permissions to
access your contacts. Then, open the app and select "Export Contacts" to create the .xlsx and PDF
files. Finally, select "Open File" to view the exported file using the preferred app.

## Where users can get help with your project

If users have any issues or questions regarding the application, they can reach out to the project's maintainers or contributors. They can do this by submitting an issue on the project's GitHub repository or by contacting the maintainers via email.

## Who maintains and contributes to the project

This project is maintained and contributed to by a group of developers who are passionate about creating useful and intuitive Android applications. The project's main contributors are listed in the CONTRIBUTORS.md file, and anyone is welcome to contribute to the project by submitting pull requests or reporting issues.

## Media Folder

**PDF File**:

The PDF file can be found at the following location:

`/storage/emulated/0/Android/data/com.example.contactsapplication/files/my_contact_file.pdf`

**Excel File**:

The Excel file can be found at the following location:

`/storage/emulated/0/Android/data/com.example.contactsapplication/files/my_contact_file.xlsx`

## Prototype URL
A prototype for this app has not been created.

## Contact Information

If you have any questions, feedback, or suggestions for improving the app, please feel free to reach out to us:

- Email: dev@softnotions.com
- Twitter: https://twitter.com/softnotions


## Authors
- [@softnotions](https://github.com/softnotions)


