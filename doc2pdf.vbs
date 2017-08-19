' Written by Rafael Winterhalter at 
' http://mydailyjava.blogspot.ca/2013/05/converting-microsoft-doc-or-docx-files.html

' See http://msdn2.microsoft.com/en-us/library/bb238158.aspx
Const wdExportFormatPDF = 17  ' PDF format
Const wdExportOptimizeForPrint = 0  ' high quality 

Const WdDoNotSaveChanges = 0
 
Dim arguments
Set arguments = WScript.Arguments
 
' Make sure that there are one or two arguments
Function CheckUserArguments()
  If arguments.Unnamed.Count < 1 Or arguments.Unnamed.Count > 2 Then
    WScript.Echo "Use:" + vbCrlf + _
        "  <script> input.docx" + vbCrlf + _
        "  <script> input.docx output.pdf"
    WScript.Quit 1
  End If
End Function
 
 
' Transforms a doc to a pdf
Function DocToPdf( docInputFile, pdfOutputFile )
  
  ''WScript.Echo "Hi"
 
  Dim fileSystemObject
  Dim wordApplication
  Dim wordDocument
  Dim baseFolder
 
  Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
  Set wordApplication = CreateObject("Word.Application")
 
  docInputFile = fileSystemObject.GetAbsolutePathName(docInputFile)
  baseFolder = fileSystemObject.GetParentFolderName(docInputFile)
 
  If Len(pdfOutputFile) = 0 Then
    pdfOutputFile = fileSystemObject.GetBaseName(docInputFile) + ".pdf"
  End If
 
  If Len(fileSystemObject.GetParentFolderName(pdfOutputFile)) = 0 Then
    pdfOutputFile = baseFolder + "\" + pdfOutputFile
  End If
  
  ''WScript.Echo "Almost Ready"
 
  ' Disable any potential macros of the word document.
  wordApplication.WordBasic.DisableAutoMacros
  
  ''WScript.Echo "Ready"
 
  Set wordDocument = wordApplication.Documents.Open(docInputFile)
  
  ''WScript.Echo "Opened"
 
  ' See http://msdn2.microsoft.com/en-us/library/bb221597.aspx 
  wordDocument.ExportAsFixedFormat _ 
    pdfOutputFile, wdExportFormatPDF, , wdExportOptimizeForPrint
  
  ''WScript.Echo "PDFed"
 
  wordDocument.Close WdDoNotSaveChanges
  wordApplication.Quit WdDoNotSaveChanges
 
  Set wordApplication = Nothing
  Set fileSystemObject = Nothing
  
  ''WScript.Echo "Done! PDF created at: " + pdfOutputFile
 
End Function
 
' Execute script
Call CheckUserArguments()
If arguments.Unnamed.Count = 2 Then
 Call DocToPdf( arguments.Unnamed.Item(0), arguments.Unnamed.Item(1) )
Else
 Call DocToPdf( arguments.Unnamed.Item(0), "" )
End If

Set arguments = Nothing

