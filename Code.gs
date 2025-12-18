/**
 * Main function - Converts all Google Docs to PDFs and merges them into a single file
 * Uses in-memory processing (no temporary files created)
 * 
 * Process:
 * 1. Load configuration and document list
 * 2. Convert each Google Doc to PDF blob (in memory)
 * 3. Merge all PDF blobs using pdf-lib
 * 4. Save final merged PDF to output folder
 */
async function createAndMergePDFs() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Load and validate configuration
    const config = loadConfig();
    
    if (!validateConfig(config)) {
      ui.alert('❌ Configuration Error', 
        'Please check the Config sheet and ensure all required fields are filled in.', 
        ui.ButtonSet.OK);
      return;
    }
    
    // Load document IDs from DocIDs sheet
    const docData = loadDocumentData();
    
    if (docData.length === 0) {
      ui.alert('❌ No Documents', 
        'No document IDs found in the DocIDs sheet. Please add documents to process.', 
        ui.ButtonSet.OK);
      return;
    }
    
    // Initialize process tracking
    Logger.log(`Starting process: ${docData.length} documents`);
    updateConfigValue('Process Status', 'RUNNING');
    updateConfigValue('Total Documents', docData.length);
    updateConfigValue('Start Time', new Date().toString());
    
    // Step 1: Convert each Google Doc to PDF (in memory)
    const pdfBlobs = [];
    
    for (let i = 0; i < docData.length; i++) {
      const row = docData[i];
      const studentName = row.studentName;
      const docId = row.docId;
      const rowNumber = i + 2; // +2 because row 1 is headers
      
      try {
        Logger.log(`Converting ${studentName} to PDF (${i + 1} of ${docData.length})`);
        
        // Direct conversion: Google Doc → PDF blob (no temp file needed)
        const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
        
        // Store PDF bytes in memory for merging
        pdfBlobs.push(new Uint8Array(pdfBlob.getBytes()));
        
        // Update status in DocIDs sheet
        updateDocumentStatus(rowNumber, '✓', '');
        
        Logger.log(`  ✓ Successfully converted ${studentName}`);
        
      } catch (error) {
        const errorMsg = error.toString().substring(0, 200);
        Logger.log(`  ✗ Error converting ${studentName}: ${errorMsg}`);
        updateDocumentStatus(rowNumber, 'ERROR', errorMsg);
      }
    }
    
    // Ensure we have at least one PDF to merge
    if (pdfBlobs.length === 0) {
      throw new Error('No PDFs were created successfully');
    }
    
    Logger.log(`Successfully created ${pdfBlobs.length} PDFs. Starting merge...`);
    
    // Step 2: Merge all PDFs using pdf-lib
    const mergedPdfDoc = await PDFLib.PDFDocument.create();
    
    for (let i = 0; i < pdfBlobs.length; i++) {
      try {
        Logger.log(`Merging PDF ${i + 1} of ${pdfBlobs.length}`);
        
        // Load PDF from bytes
        const pdfDoc = await PDFLib.PDFDocument.load(pdfBlobs[i]);
        const pageCount = pdfDoc.getPageCount();
        
        // Create array of all page indices [0, 1, 2, ...]
        const indices = Array.from({ length: pageCount }, (_, idx) => idx);
        
        // Copy all pages from this PDF to the merged document
        const copiedPages = await mergedPdfDoc.copyPages(pdfDoc, indices);
        copiedPages.forEach(page => mergedPdfDoc.addPage(page));
        
        Logger.log(`  ✓ Merged ${pageCount} pages`);
        
      } catch (mergeError) {
        Logger.log(`  ✗ Error merging PDF ${i}: ${mergeError.toString()}`);
      }
    }
    
    // Step 3: Save merged PDF to output folder
    Logger.log('Saving final merged PDF...');
    
    const mergedPdfBytes = await mergedPdfDoc.save();
    const mergedBlob = Utilities.newBlob(
      Array.from(mergedPdfBytes), 
      MimeType.PDF, 
      config.pdfName + '.pdf'
    );
    
    const outputFolder = DriveApp.getFolderById(config.outputFolderId);
    const finalPdfFile = outputFolder.createFile(mergedBlob);
    
    Logger.log(`✅ Final PDF created: ${finalPdfFile.getName()}`);
    
    // Update final status in Config sheet
    updateConfigValue('Process Status', 'COMPLETE ✓');
    updateConfigValue('Completion Time', new Date().toString());
    updateConfigValue('Final PDF URL', finalPdfFile.getUrl());
    updateConfigValue('Final PDF ID', finalPdfFile.getId());
    
    Logger.log(`Final PDF URL: ${finalPdfFile.getUrl()}`);
    
    // Show success message to user
    ui.alert('✅ Success!', 
      `Created merged PDF with ${docData.length} documents.\n\n` +
      `Check the Config sheet for the download link.`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error in createAndMergePDFs: ' + error.toString());
    ui.alert('❌ Error', 
      'An error occurred: ' + error.toString(), 
      ui.ButtonSet.OK);
    updateConfigValue('Process Status', 'ERROR');
    updateConfigValue('Error Message', error.toString());
  }
}

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 WEP Combiner')
    .addItem('▶️ Create & Merge PDFs', 'createAndMergePDFs')
    .addSeparator()
    .addItem('🧹 Reset Status', 'resetProcessStatus')
    .addSeparator()
    .addItem('ℹ️ Help', 'showHelp')
    .addToUi();
}