# Changelog

All notable changes to the WEP Combiner project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-12-17

### 🎉 Initial Stable Release

**Performance:** Successfully processes 170+ documents in under 3 minutes

### Added
- Direct Google Doc to PDF conversion (in-memory processing)
- PDF merging using pdf-lib library
- Config sheet for settings management
  - Output Folder ID
  - PDF Name
  - Process status tracking
- DocIDs sheet for document management
  - Student Name column
  - Document ID column
  - Status column (auto-populated)
  - Error Message column (auto-populated)
- Custom menu with the following options:
  - Create & Merge PDFs
  - Reset Status
  - Help
- HTML help dialog with detailed setup instructions
- Automatic status updates during processing
- Color-coded success/error indicators in DocIDs sheet
- Final PDF URL and ID tracking in Config sheet

### Technical Details
- Uses pdf-lib (minified, static version in .gs file)
- All processing done in memory (no temporary files)
- Includes setTimeout workaround for pdf-lib compatibility
- SpreadsheetApp.flush() for immediate status updates

### Removed
- All batching logic (no longer needed with in-memory processing)
- Temporary folder operations
- Time-based triggers and scheduling code
- Batch number tracking

---

## [Unreleased]

### Future Considerations
- None at this time - system is stable and performant

---

## Version History

- **1.0.0** - Stable release with in-memory processing
- **0.x** - Development versions with batching (deprecated)

---

## Notes

### Known Issues
- None

### Dependencies
- pdf-lib (minified, embedded in project)
- Google Apps Script (V8 runtime)

### Browser Compatibility
- Works in all modern browsers via Google Sheets interface

### Tested With
- 170 documents: ~3 minutes processing time
- Various document lengths and formats
- Multiple Google Drive folders
