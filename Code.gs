/**
 * Main function - converts all docs to PDFs and merges them
 * No batching needed since we're not creating temp docs
 */
async function createAndMergePDFs() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Load configuration
    const config = loadConfig();
    
    // Validate configuration
    if (!validateConfig(config)) {
      ui.alert('❌ Configuration Error', 
        'Please check the Config sheet and ensure all required fields are filled in.', 
        ui.ButtonSet.OK);
      return;
    }
    
    // Load document IDs
    const docData = loadDocumentData();
    
    if (docData.length === 0) {
      ui.alert('❌ No Documents', 
        'No document IDs found in the DocIDs sheet. Please add documents to process.', 
        ui.ButtonSet.OK);
      return;
    }
    
    Logger.log(`Starting process: ${docData.length} documents`);
    updateConfigValue('Process Status', 'RUNNING');
    updateConfigValue('Total Documents', docData.length);
    updateConfigValue('Start Time', new Date().toString());
    
    // Step 1: Convert each Google Doc to PDF
    const pdfBlobs = [];
    const tempFolder = DriveApp.getFolderById(config.tempFolderId);
    
    for (let i = 0; i < docData.length; i++) {
      const row = docData[i];
      const studentName = row.studentName;
      const docId = row.docId;
      const rowNumber = i + 2;
      
      try {
        Logger.log(`Converting ${studentName} to PDF (${i + 1} of ${docData.length})`);
        
        // THIS IS YOUR WORKING METHOD - Direct conversion!
        const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
        //pdfBlob.setName(`${String(i).padStart(3, '0')}_${studentName}.pdf`);
        
        // Save individual PDF to temp folder
        //const pdfFile = tempFolder.createFile(pdfBlob);
        
        // Store for merging
        //pdfBlobs.push(new Uint8Array(pdfFile.getBlob().getBytes()));
        pdfBlobs.push(new Uint8Array(pdfBlob.getBytes()));
        
        // Update status
        updateDocumentStatus(rowNumber, '✓', '', 0);
        
        Logger.log(`  ✓ Successfully converted ${studentName}`);
        
      } catch (error) {
        const errorMsg = error.toString().substring(0, 200);
        Logger.log(`  ✗ Error converting ${studentName}: ${errorMsg}`);
        updateDocumentStatus(rowNumber, 'ERROR', errorMsg, 0);
      }
    }
    
    if (pdfBlobs.length === 0) {
      throw new Error('No PDFs were created successfully');
    }
    
    Logger.log(`Successfully created ${pdfBlobs.length} PDFs. Starting merge...`);
    
    // Step 2: Merge all PDFs using pdf-lib
    const mergedPdfDoc = await PDFLib.PDFDocument.create();
    
    for (let i = 0; i < pdfBlobs.length; i++) {
      try {
        Logger.log(`Merging PDF ${i + 1} of ${pdfBlobs.length}`);
        
        const pdfDoc = await PDFLib.PDFDocument.load(pdfBlobs[i]);
        const pageCount = pdfDoc.getPageCount();
        const indices = Array.from({ length: pageCount }, (_, idx) => idx);
        
        const copiedPages = await mergedPdfDoc.copyPages(pdfDoc, indices);
        copiedPages.forEach(page => mergedPdfDoc.addPage(page));
        
        Logger.log(`  ✓ Merged ${pageCount} pages`);
        
      } catch (mergeError) {
        Logger.log(`  ✗ Error merging PDF ${i}: ${mergeError.toString()}`);
      }
    }
    
    // Step 3: Save merged PDF
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
    
    // Step 4: Clean up individual PDFs from temp folder
    const files = tempFolder.getFilesByType(MimeType.PDF);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }
    Logger.log('Cleaned up temporary PDF files');
    
    // Update final status
    updateConfigValue('Process Status', 'COMPLETE ✓');
    updateConfigValue('Completion Time', new Date().toString());
    updateConfigValue('Final PDF URL', finalPdfFile.getUrl());
    updateConfigValue('Final PDF ID', finalPdfFile.getId());
    
    Logger.log(`Final PDF URL: ${finalPdfFile.getUrl()}`);
    
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
 * Updates the menu to use the new simpler approach
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 WEP Combiner')
    .addItem('▶️ Create & Merge PDFs', 'createAndMergePDFs')
    .addSeparator()
    .addItem('🗑️ Clear All Triggers', 'clearAllTriggers')
    .addItem('🧹 Reset Status', 'resetProcessStatus')
    .addSeparator()
    .addItem('ℹ️ Help', 'showHelp')
    .addToUi();
}