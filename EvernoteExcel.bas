Option Explicit

Const FILE_NAME = "D:\Ivan\APBP\PTPP-REG\Evernote\EvernoteExcel.enex"
Const HEADING1 = "<?xml version=""1.0"" encoding=""UTF-8""?>"
Const HEADING2 = "<!DOCTYPE en-note SYSTEM ""http://xml.evernote.com/pub/enml2.dtd"">"

Sub OutputXML()

Dim iRow As Long
Dim sTags As String
Dim i As Integer
Dim nTag As Integer
Dim aTags

With ActiveSheet
    Open FILE_NAME For Output As #1
    Print #1, HEADING1
    Print #1, "<en-export export-date=""20120202T073208Z"" application=""Evernote/Windows"" version=""4.x"">"
    For iRow = 2 To .Cells(.Rows.Count, "A").End(xlUp).Row
        Print #1, "<note>"
        Print #1, vbTab & "<title>" & .Cells(iRow, "B").Value & "</title>"
        Print #1, vbTab & "<content><![CDATA[" & HEADING1 & _
            HEADING2 & _
            "<en-note style=""word-wrap: break-word; -webkit-nbsp-mode: space; -webkit-line-break: after-white-space;"">"
        If (Trim(.Cells(iRow, "D").Value) <> "") Then
            Print #1, vbTab & vbTab & CBr(.Cells(iRow, "D").Value) 'Note
        End If
        Print #1, vbTab & "</en-note>]]></content>"
        Print #1, vbTab & "<created>20120202T042955Z</created>"
        Print #1, vbTab & "<updated>20120202T070232Z</updated>"
        ' Tags
        sTags = Trim(.Cells(iRow, "C").Value)
        If (sTags <> "") Then
            aTags = Split(sTags, ";")
            nTag = UBound(aTags)
            For i = 0 To nTag
                Print #1, vbTab & "<tag>" & Trim(aTags(i)) & "</tag>"
            Next
        End If
        ' Author
        Print #1, vbTab & "<note-attributes><author>" & .Cells(iRow, "A").Value & "</author></note-attributes>"
        Print #1, "</note>"
    Next iRow
    Print #1, "</en-export>"
    Close #1
End With

End Sub

Sub InputXML()
    Dim oDoc As MSXML2.DOMDocument
    Dim oContent As MSXML2.DOMDocument
    Dim oExport As MSXML2.IXMLDOMNode
    Dim oNote As MSXML2.IXMLDOMNode
    Dim oChild As MSXML2.IXMLDOMNode
    Dim oInside As MSXML2.DOMDocument
    Dim oWks As Worksheet
    Dim i As Integer
    Dim sCol As String
    Dim sTag As String
    Dim sText As String

    Set oWks = ActiveSheet
    Set oDoc = New MSXML2.DOMDocument
    oDoc.async = False
    oDoc.validateOnParse = False
    oDoc.Load (FILE_NAME)
    Set oExport = oDoc.DocumentElement
    Set oNote = oExport.FirstChild

    i = 1
    For Each oNote In oExport.ChildNodes
        sTag = ""
        i = i + 1
        For Each oChild In oNote.ChildNodes
            sCol = ""
            Select Case oChild.BaseName
                Case "title"
                    sCol = "B"
                Case "content"
                    sCol = "D"
                Case "note-attributes"
                    sCol = "A"
                Case "tag"
                    If (sTag <> "") Then
                        sTag = sTag & "; "
                    End If
                    sTag = sTag & oChild.Text
            End Select
            If (sCol <> "") Then
                sText = oChild.Text
                If (oChild.BaseName = "content") Then
                    sText = Replace(sText, HEADING1, "")
                    sText = Replace(sText, HEADING2, "")
                    Set oInside = New MSXML2.DOMDocument
                    oInside.LoadXML (sText)
                    sText = oInside.Text
                End If
                oWks.Cells(i, sCol).Value = sText
            End If
        Next oChild
        oWks.Cells(i, "C").Value = sTag
    Next oNote
End Sub

'parse hard breaks into to HTML breaks
Function CBr(val) As String
    CBr = Replace(val, Chr(13), "<br />")
End Function


