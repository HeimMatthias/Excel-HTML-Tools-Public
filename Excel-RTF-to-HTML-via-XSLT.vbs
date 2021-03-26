' Script Version 0.8, (c) 26.03.2021, Matthias.Heim@hep-verlag.ch
' Overwrite content of selected cells with their html-code
Sub TurnSelectionRichTextIntoHtml()
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
    originalXML = Replace(originalXML, "Workbook ", "Workbook xml:space=""preserve"" ", 1, 1)
    Debug.Print (originalXML)
    Dim source As Object
    Dim stylesheet As Object
    Set source = CreateObject("MSXML2.DOMDocument")
    Set stylesheet = CreateObject("MSXML2.DOMDocument")
    ' Load data.
    source.async = False
    ' source.Load App.Path & "\books.xml"
    source.LoadXML (originalXML)
    Dim myErr
    If (source.parseError.ErrorCode <> 0) Then
       Set myErr = source.parseError
       MsgBox ("There has been an error parsing the xml-document in Excel " & myErr.reason)
    Else
       'For later reference: changing the property or inserting the relevant attribute *after* LoadXML cannot work since it has already been discarded, xml:space="preserve" needs to be included in XML-String
       'source.preserveWhiteSpace = True
       'Dim xmlRoot As Variant
       'Set xmlRoot = source.SelectSingleNode("Workbook")
       'Dim temp As Variant
       'temp = xmlRoot.setAttribute("xml:space", "preserve")
       
       ' Load style sheet.
       stylesheet.async = True
       stylesheet.Load Application.ThisWorkbook.Path & "\excel-full-transformation.xsl"
       If (stylesheet.parseError.ErrorCode <> 0) Then
          Set myErr = stylesheet.parseError
          MsgBox ("There has been an error loading the XSL-Stylesheet " & myErr.reason)
       Else
          ' XSL-Transformation
            originalXML = source.transformNode(stylesheet)
            Debug.Print (originalXML)
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
  
' Log Excel-XML of selected cells to console
Sub consoleLogXML()
    Application.ScreenUpdating = False
    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection
    Debug.Print selectedRange.Value(11)

    'For Each cel In selectedRange.Cells
    '    Debug.Print cel.Address, cel.Value(11)
    'Next cel
    Application.ScreenUpdating = True
End Sub
