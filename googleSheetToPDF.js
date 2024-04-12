/**
 * @author u/juddaaaa <https://www.reddit.com/user/juddaaaaa/>
 * @description Creates and emails a PDF of a Google Sheet including only the sheet who's name is passed in
 * @license MIT
 * @version 1
 */

/**
 *
 * @param { String } sheetName
 * @returns void
 */
function googleSheetToPDF(sheetName) {
  // Get active spreadsheet
  const spreadsheet = SpreadsheetApp.getActive()

  // Get sheet to export to PDF
  const sheet = spreadsheet.getSheetByName(sheetName)

  // If sheet doesn't exist, exit the function
  if (!sheet) return

  // Hide all other visible sheets so they don't get included in PDF
  const hiddenSheets = spreadsheet.getSheets().reduce((sheets, currentSheet) => {
    if (currentSheet.getName() !== sheet.getName() && !currentSheet.isSheetHidden()) {
      currentSheet.hideSheet()
      sheets.push(currentSheet)
    }

    return sheets
  }, [])

  // Get the sheet we're exporting as a Blob and choose a name for the PDF file
  const pdfBlob = spreadsheet.getAs(MimeType.PDF)
  const pdfName = sheet.getName()

  // Create the PDF file in the user's Google Drive
  const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName)

  // Get URL of PDF file to include in email
  const pdfUrl = pdfFile.getUrl()

  // Email address of recipient
  const recipient = "user@example.com"

  // Body of email to send to recipient
  const template = HtmlService.createTemplateFromFile("googleSheetToPDF")
  template.fileName = pdfName
  template.pdfUrl = pdfUrl

  const htmlBody = HtmlService.createHtmlOutput(template.evaluate())

  // Send email to recipient with PDF file attached
  GmailApp.sendEmail(recipient, `${pdfName} successfully created`, "", {
    htmlBody,
    attachments: [pdfFile],
  })

  // Unhide sheets that were hidden earlier
  hiddenSheets.forEach(sheet => {
    sheet.showSheet()
  })
}
