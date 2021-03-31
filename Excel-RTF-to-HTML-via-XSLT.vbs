' Script Version 0.9.1, (c) 30.03.2021, Matthias.Heim@hep-verlag.ch
Const xsltConversionFile As String = "excel-complex-style-transformation.xsl"
' There are three conversions publicly available:
' excel-full-transformation.xsl : declares block-appearance of cells as styles but all text-level styles as inline declarations
' excel-complex-style-transformation.xsl : retains all cell styles, including text-level styles in style-declaration, but declares exceptions to preset style-rule in style-attribute of individual cells, overriding the cell-class rules, character-level styles are presented as inline declarations
' excel-simple-style-transformation.xsl : retains cell styles, including text-level styles in style-declaration, but corrects class style-declaration based on exemptions from first cell using style. Returns very clean code for individual cells, but can result in incorrect styling if several cells use same cell style preset.

' Overwrite content of selected cells with their html-code
Sub turnSelectionRichTextIntoHtml()
    Application.ScreenUpdating = False
    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection
    ' Parse individual cells
    For Each cel In selectedRange.Cells
        'Debug.Print cel.Address, cel.Value(11)
        If Not (cel.Value = "") Then cel.Value = fnConvertXML2HTML(cel.Value(11))
    Next cel
    Application.ScreenUpdating = True
End Sub

' Function to convert an Excel-XML-String into html
Function fnConvertXML2HTML(originalXML As String) As String
    ' insert xml:space="preserve"-Attribute before string is parsed
    ' originalXML = Replace(originalXML, "Workbook ", "Workbook xml:space=""preserve"" ", 1, 1)
    Dim source As Object
    Dim stylesheet As Object
    Set source = CreateObject("MSXML2.DOMDocument")
    Set stylesheet = CreateObject("MSXML2.DOMDocument")
    ' Load data.
    source.async = False
    source.preserveWhiteSpace = True
    source.LoadXML (originalXML)
    Dim myErr
    If (source.parseError.ErrorCode <> 0) Then
       Set myErr = source.parseError
       MsgBox ("There has been an error parsing the xml-document in Excel " & myErr.reason)
    Else
       ' Load style sheet.
       stylesheet.async = True
       stylesheet.Load Application.ThisWorkbook.Path & "\" & xsltConversionFile
       If (stylesheet.parseError.ErrorCode <> 0) Then
          Set myErr = stylesheet.parseError
          MsgBox ("There has been an error loading the XSL-Stylesheet " & myErr.reason)
       Else
          ' XSL-Transformation
            originalXML = source.transformNode(stylesheet)
            fnConvertXML2HTML = originalXML
       End If
    End If
End Function

    ' copy html of selected cells to clipboard
Sub copySelectionAsHTML()
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText fnConvertXML2HTML(Application.Selection.Value(11))
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

Sub saveSelectionAsHTML()
    Dim html As String
    Dim innerHtml As String
    innerHtml = fnConvertXML2HTML(Application.Selection.Value(11))
    
    Dim styleEnd As Integer
    styleEnd = InStr(innerHtml, "</style>") + 8
    
    html = "<!DOCTYPE html>" & vbNewLine & "<html>" & vbNewLine & "<head>" & vbNewLine & "<meta charset=""UTF-8"">" & vbNewLine & "<meta name=""generator"" content=""github.com/HeimMatthias/Excel-HTML-Tools-Public"" />" & vbNewLine & "<title>" & ActiveWorkbook.Name & "</title>" & vbNewLine
    html = html & Left(innerHtml, styleEnd)
    html = html & "</head>" & vbNewLine & "<body>"
    html = html & Mid(innerHtml, styleEnd)
    html = html & "</body>" & vbNewLine & "</html>"
    
    Dim fileSaveName As Variant
    fileSaveName = Application.GetSaveAsFilename( _
    fileFilter:="HTML Files (*.html), *.html")
    Debug.Print (fileSaveName)
    If fileSaveName = False Then
        Debug.Print ("User cancelled")
    Else
        Dim returnCode As Integer
        returnCode = writeOut(html, CStr(fileSaveName))
    End If
End Sub

' from https://gist.github.com/JoBrad/1023484
' Function saves cText in file, and returns 1 if successful, 0 if not
Public Function writeOut(cText As String, file As String) As Integer
    On Error GoTo errHandler
    Dim fsT As Object
    Dim tFilePath As String

    tFilePath = file

    'Create Stream object
    Set fsT = CreateObject("ADODB.Stream")

    'Specify stream type - we want To save text/string data.
    fsT.Type = 2

    'Specify charset For the source text data.
    fsT.Charset = "utf-8"

    'Open the stream And write binary data To the object
    fsT.Open
    fsT.writetext cText

    'Save binary data To disk
    fsT.SaveToFile tFilePath, 2

    GoTo finish

errHandler:
    MsgBox (Err.Description)
    writeOut = 0
    Exit Function

finish:
    writeOut = 1
End Function

' Overwrite number formats of selected cells with their text value
' It is highly recommended to run this prior to html-Export - BUT SAVE BEFOREHAND, NOT AFTERWARDS
Sub turnSelectionNumbersIntoText()
    Application.ScreenUpdating = False
    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection
    ' Parse individual cells
    For Each cel In selectedRange.Cells
        'Debug.Print cel.Address, cel.Value(11)
        
        Dim vData As String
        vData = cel.Text
        ' If Not (cel.Value = "") Then cel.Value = fnConvertXML2HTML(cel.Value(11))
        Select Case True
        Case IsEmpty(cel)
            cel.Value = " "
        Case WorksheetFunction.IsText(cel)
        Case Else
            cel.NumberFormat = "@"
            cel.Value = vData
        End Select
    Next cel
    Application.ScreenUpdating = True
End Sub
