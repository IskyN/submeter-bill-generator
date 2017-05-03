' Written by Rafael Winterhalter at 
' http://mydailyjava.blogspot.ca/2013/05/converting-microsoft-doc-or-docx-files.html

' See http://msdn2.microsoft.com/en-us/library/bb238158.aspx
Const wdFormatPDF = 17  ' PDF format. 

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
 
  ' Disable any potential macros of the word document.
  wordApplication.WordBasic.DisableAutoMacros
 
  Set wordDocument = wordApplication.Documents.Open(docInputFile)
 
  ' See http://msdn2.microsoft.com/en-us/library/bb221597.aspx 
  wordDocument.SaveAs pdfOutputFile, wdFormatPDF
 
  wordDocument.Close WdDoNotSaveChanges
  wordApplication.Quit WdDoNotSaveChanges
   
  Set wordApplication = Nothing
  Set fileSystemObject = Nothing
 
End Function
 
' Execute script
Call CheckUserArguments()
If arguments.Unnamed.Count = 2 Then
 Call DocToPdf( arguments.Unnamed.Item(0), arguments.Unnamed.Item(1) )
Else
 Call DocToPdf( arguments.Unnamed.Item(0), "" )
End If
 
Set arguments = Nothing