# 📄 WEP Combiner

A Google Apps Script tool that combines multiple Google Docs into a single merged PDF file. Designed for efficiently processing large batches of documents with in-memory processing for optimal performance.

## ✨ Features

- **Fast Processing:** Handles 170+ documents in under 3 minutes
- **In-Memory Processing:** No temporary files created, all operations done in memory
- **Simple Setup:** Easy configuration through Google Sheets interface
- **Status Tracking:** Real-time status updates for each document
- **Error Handling:** Clear error messages and visual indicators
- **Custom Menu:** User-friendly interface with help documentation
- **Automatic PDF Merging:** Uses pdf-lib for reliable PDF manipulation

## 🚀 Performance

- **Tested:** 170 documents processed in ~3 minutes
- **Memory Efficient:** All PDFs handled in memory without temporary storage
- **Scalable:** Designed to handle large document batches

## 📋 Prerequisites

- Google Account with access to Google Sheets and Google Drive
- Google Docs to be merged (must have view/edit access)
- Destination folder in Google Drive for the output PDF

## 🛠️ Installation

1. **Create a new Google Sheet** or open an existing one

2. **Set up the required sheets:**
   - **Config** sheet with the following rows:
     - Output Folder ID
     - PDF Name
   - **DocIDs** sheet with columns:
     - Column A: Student Name (or description)
     - Column B: Document ID
     - Column C: Status (auto-populated)
     - Column D: Error Message (auto-populated)

3. **Open Apps Script Editor:**
   - Click `Extensions` → `Apps Script`

4. **Add the project files:**
   - Copy `Code.gs` content
   - Copy `Utils.gs` content  
   - Add `Help.html` file
   - Add the pdf-lib library file

5. **Configure appsscript.json** (if needed):
   ```json
   {
     "timeZone": "America/New_York",
     "dependencies": {},
     "exceptionLogging": "STACKDRIVER",
     "runtimeVersion": "V8"
   }
   ```

6. **Save and refresh** your Google Sheet
   - The "📄 WEP Combiner" menu should appear

## 📖 Usage

### Setup

1. **Get your Output Folder ID:**
   - Open the Google Drive folder where you want the final PDF saved
   - Copy the folder ID from the URL: `drive.google.com/drive/folders/[FOLDER_ID]`
   - Paste it into the Config sheet

2. **Set your PDF Name:**
   - Enter the desired name for your merged PDF (without .pdf extension)
   - Example: `2024_WEP_Combined`

3. **Add your documents:**
   - For each document, add a row to the DocIDs sheet
   - Column A: Student name or description
   - Column B: Document ID from the URL: `docs.google.com/document/d/[DOC_ID]/edit`

### Running the Process

1. Click `📄 WEP Combiner` menu → `▶️ Create & Merge PDFs`

2. **Monitor progress:**
   - Config sheet shows overall status
   - DocIDs sheet shows individual document status
   - Green checkmarks (✓) indicate success
   - Red "ERROR" indicates issues (with error messages in column D)

3. **Download your PDF:**
   - When complete, the final PDF URL appears in the Config sheet
   - Click the link to open/download your merged PDF

### Additional Features

- **Reset Status:** Clears all status indicators before a new run
- **Help:** Opens detailed setup instructions

## 🏗️ Technical Architecture

### File Structure

```
├── Code.gs           # Main processing logic
├── Utils.gs          # Helper functions and data management
├── Help.html         # Help dialog interface
├── appsscript.json   # Project configuration
└── pdf-lib.gs        # PDF manipulation library (minified)
```

### Key Functions

- **`createAndMergePDFs()`** - Main function that orchestrates the entire process
- **`loadConfig()`** - Loads settings from Config sheet
- **`loadDocumentData()`** - Retrieves document list from DocIDs sheet
- **`validateConfig()`** - Ensures all required settings are present
- **`updateDocumentStatus()`** - Updates processing status for each document
- **`resetProcessStatus()`** - Clears all status fields

### Processing Flow

1. Load and validate configuration
2. Load document list from DocIDs sheet
3. Convert each Google Doc to PDF blob (in memory)
4. Merge all PDF blobs using pdf-lib
5. Save final merged PDF to output folder
6. Update status and provide download link

## 🔧 Configuration Options

### Config Sheet Settings

| Setting | Required | Description |
|---------|----------|-------------|
| Output Folder ID | Yes | Google Drive folder ID for output PDF |
| PDF Name | Yes | Name for the merged PDF (no extension) |
| Process Status | Auto | Current process status |
| Total Documents | Auto | Number of documents being processed |
| Start Time | Auto | Process start timestamp |
| Completion Time | Auto | Process completion timestamp |
| Final PDF URL | Auto | Link to download merged PDF |
| Final PDF ID | Auto | Google Drive file ID of merged PDF |

## 🐛 Troubleshooting

### Common Issues

**"No documents found"**
- Ensure DocIDs sheet has document IDs in column B
- Check that rows aren't empty

**"Invalid folder ID"**
- Verify the Output Folder ID is correct
- Ensure you have write access to the folder

**Individual document errors**
- Check document ID is correct
- Verify you have access to the document
- Ensure the document exists and isn't trashed

**"Working..." message doesn't go away**
- This is normal for alert dialogs - click OK to dismiss
- The HTML help dialog doesn't have this issue

### Performance Tips

- Process documents in batches if you have 500+ documents
- Ensure good internet connection for Drive API calls
- Check Executions log for detailed progress: `Extensions` → `Apps Script` → `Executions`

## 📊 Version History

See [CHANGELOG.md](CHANGELOG.md) for detailed version history and changes.

## 🤝 Contributing

This is a personal project, but suggestions and improvements are welcome:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📝 License

This project is provided as-is for educational and personal use.

## 🙏 Acknowledgments

- Built with [Google Apps Script](https://developers.google.com/apps-script)
- PDF manipulation powered by [pdf-lib](https://pdf-lib.js.org/)
- Developed with assistance from Claude (Anthropic)

## 📞 Support

For issues or questions:
- Check the built-in Help dialog (`📄 WEP Combiner` → `ℹ️ Help`)
- Review the [CHANGELOG.md](CHANGELOG.md)
- Check Google Apps Script execution logs

---

**Current Version:** 1.0.0  
**Last Updated:** December 17, 2025  
**Status:** ✅ Stable and Production Ready
