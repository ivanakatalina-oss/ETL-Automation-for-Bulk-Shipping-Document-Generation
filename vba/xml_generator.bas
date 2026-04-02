' =============================================================================
' Module: xml_generator.bas
' Purpose: Generate ERP-compatible XML file for bulk shipping document (DDT)
'          import from a structured Excel data sheet.
'
' How it works:
'   1. User inputs a date range (start / end)
'   2. Data is sorted chronologically
'   3. Rows are grouped by client + date into document blocks
'   4. Each block is written as an XML <Document> node
'   5. Within each document, product rows are built from multiple columns
'   6. Output: a single XML file ready for ERP bulk import
'
' Key design decisions:
'   - CleanXML() sanitizes all string fields before writing to XML
'     to prevent malformed output on special characters (&, <, >, ", ')
'   - Document grouping uses a composite key: ClientCode + ExitDate
'     so multiple clients on the same day generate separate documents
'   - Transport responsibility is mapped from internal notation
'     to ERP-required values (CONSIGNOR/CONSIGNEE → Mittente/Destinatario)
'   - Extra traceability fields are concatenated dynamically
'     from a predefined column list, reducing hardcoded row logic
'
' BACKLOG (OMS v1.0):
'   - Replace date InputBox with structured date picker
'   - Replace column letter references with named ranges
'   - Add validation log for rows skipped due to missing data
' =============================================================================

' -----------------------------------------------------------------------------
' CleanXML
' Sanitizes a value for safe XML embedding.
' Replaces the 5 reserved XML characters with their entity equivalents.
' Must be applied to ALL string fields before Print #1 output.
' -----------------------------------------------------------------------------
Function CleanXML(txt As Variant) As String
    Dim s As String
    s = Trim(CStr(txt))
    s = Replace(s, "&",  "&amp;")
    s = Replace(s, "<",  "&lt;")
    s = Replace(s, ">",  "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'",  "&apos;")
    CleanXML = s
End Function


' -----------------------------------------------------------------------------
' GenerateXML_DDT
' Main procedure. Reads data from the staging sheet, applies date filter,
' groups rows into documents, and writes the ERP XML output file.
' -----------------------------------------------------------------------------
Sub GenerateXML_DDT()

    ' --- DECLARATIONS ---
    Dim xmlFile    As String
    Dim i          As Long
    Dim lastRow    As Long
    Dim currentKey As String
    Dim prevKey    As String
    Dim sh         As Worksheet
    Dim docDate    As String
    Dim transport  As String
    Dim startDate  As Date
    Dim endDate    As Date
    Dim rowDate    As Variant
    Dim inputStart As String
    Dim inputEnd   As String

    ' --- TARGET SHEET ---
    ' Sheet "_XMLDATA" is the Power Query output staging area.
    ' All columns referenced below correspond to this sheet's structure.
    Set sh = Sheets("_XMLDATA")

    ' --- USER INPUT: DATE RANGE ---
    inputStart = InputBox("Enter START date (dd/mm/yyyy):", "DDT Filter", Date)
    inputEnd   = InputBox("Enter END date (dd/mm/yyyy):",   "DDT Filter", Date)

    If Not IsDate(inputStart) Or Not IsDate(inputEnd) Then
        MsgBox "Invalid dates. Operation cancelled.", vbCritical
        Exit Sub
    End If

    startDate = CDate(inputStart)
    endDate   = CDate(inputEnd)

    ' --- SORT DATA BY EXIT DATE (Column L) ---
    ' Ensures documents are generated in chronological order.
    ' Required for correct document grouping logic below.
    lastRow = sh.Cells(sh.Rows.Count, "P").End(xlUp).Row
    If lastRow > 1 Then
        With sh.Sort
            .SortFields.Clear
            .SortFields.Add Key:=sh.Range("L2:L" & lastRow), Order:=xlAscending
            .SetRange sh.Range("A1:BZ" & lastRow)
            .Header = xlYes
            .Apply
        End With
    End If

    ' --- OUTPUT FILE PATH ---
    ' Saved to user Desktop for immediate ERP import access.
    xmlFile = Environ("USERPROFILE") & "\Desktop\ERP_DDT_Import.xml"

    ' --- OPEN FILE AND WRITE XML HEADER ---
    Open xmlFile For Output As #1
    Print #1, "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
    Print #1, "<ERPDocuments AppVersion=""2"">"
    Print #1, "  <Documents>"

    ' --- MAIN LOOP: PROCESS EACH DATA ROW ---
    For i = 2 To lastRow

        ' --- DATE FILTER ---
        rowDate = sh.Cells(i, "L").Value
        If Not IsDate(rowDate) Then GoTo NextRow
        If CDate(rowDate) < startDate Or CDate(rowDate) > endDate Then GoTo NextRow

        ' --- VALIDITY CHECK ---
        ' Skip rows with no client name (column P)
        If Trim(sh.Cells(i, "P").Value) = "" Then GoTo NextRow

        ' --- DOCUMENT GROUPING KEY ---
        ' One XML document per unique ClientCode + ExitDate combination.
        ' Multiple rows with the same key = multiple line rows within one document.
        currentKey = Trim(sh.Cells(i, "O").Value) & "|" & _
                     Format(sh.Cells(i, "L").Value, "yyyymmdd")

        ' --- CLOSE PREVIOUS DOCUMENT BLOCK (if key changed) ---
        If currentKey <> prevKey And prevKey <> "" Then
            Print #1, "      </Rows>"
            Print #1, "    </Document>"
        End If

        ' --- OPEN NEW DOCUMENT BLOCK ---
        If currentKey <> prevKey Then
            docDate = Format(sh.Cells(i, "L").Value, "yyyy-mm-dd")

            ' Map internal transport notation to ERP-required values
            transport = Trim(sh.Cells(i, "BE").Value)
            If LCase(transport) = "consignor"  Then transport = "Sender"
            If LCase(transport) = "consignee"  Then transport = "Recipient"

            Print #1, "    <Document>"
            Print #1, "      <DocumentType>D</DocumentType>"
            Print #1, "      <CustomerCode>"      & CleanXML(sh.Cells(i, "O").Value)  & "</CustomerCode>"
            Print #1, "      <CustomerName>"      & CleanXML(sh.Cells(i, "P").Value)  & "</CustomerName>"
            Print #1, "      <Date>"              & docDate                            & "</Date>"
            Print #1, "      <DeliveryDate>"      & docDate                            & "</DeliveryDate>"
            Print #1, "      <TransportDate>"     & docDate                            & "</TransportDate>"
            Print #1, "      <TransportInCharge>" & CleanXML(transport)                & "</TransportInCharge>"
            Print #1, "      <TransportReason>"   & CleanXML(sh.Cells(i, "BC").Value) & "</TransportReason>"
            Print #1, "      <GoodsAppearance>"   & CleanXML(sh.Cells(i, "BD").Value) & "</GoodsAppearance>"
            Print #1, "      <Rows>"
        End If

        ' --- ROW BLOCK 1: PROCESSING LINE (conditional) ---
        ' Only written if the processing flag column (AX) is populated.
        ' Indicates goods routed to secondary processing unit.
        If sh.Cells(i, "AX").Value <> "" Then
            Print #1, "        <Row>" & _
                       "<Description>" & CleanXML(sh.Cells(i, "AX").Value) & "</Description>" & _
                       "<Qty>" & Replace(sh.Cells(i, "AY").Value, ",", ".") & "</Qty>" & _
                       "</Row>"
        End If

        ' --- ROW BLOCK 2: OFFAL LINE (conditional) ---
        ' Only written if offal description column (AZ) is populated.
        If sh.Cells(i, "AZ").Value <> "" Then
            Print #1, "        <Row>" & _
                       "<Description>" & CleanXML(sh.Cells(i, "AZ").Value) & "</Description>" & _
                       "<Qty>" & Replace(sh.Cells(i, "BA").Value, ",", ".") & "</Qty>" & _
                       "</Row>"
        End If

        ' --- VISUAL SEPARATOR ROW ---
        Print #1, "        <Row><Description> </Description></Row>"

        ' --- ROW BLOCK 3: AGGREGATED DATA LINE ---
        If sh.Cells(i, "BB").Value <> "" Then
            Print #1, "        <Row>" & _
                       "<Description>" & CleanXML(sh.Cells(i, "BB").Value) & "</Description>" & _
                       "</Row>"
        End If

        ' --- ROW BLOCK 4: DYNAMIC TRACEABILITY STRING ---
        ' Concatenates non-empty values from defined traceability columns.
        ' Format: "FIELDNAME: value | FIELDNAME: value | ..."
        ' Column list is predefined; header row used as field label source.
        Dim extra   As String
        Dim v       As String
        Dim cols    As Variant
        Dim colName As String
        Dim j       As Integer

        cols  = Array("AV", "AW", "AM", "AD", "AC", "AN", "AR")
        extra = ""

        For j = LBound(cols) To UBound(cols)
            v = Trim(sh.Cells(i, cols(j)).Value)
            If v <> "" Then
                colName = UCase(Replace(Replace( _
                    sh.Cells(1, cols(j)).Value, "TRACE.", ""), "XML", ""))
                extra = extra & Trim(colName) & ": " & v & " | "
            End If
        Next j

        If extra <> "" Then
            extra = Left(extra, Len(extra) - 3) ' Remove trailing " | "
            Print #1, "        <Row>" & _
                       "<Description>" & CleanXML(extra) & "</Description>" & _
                       "</Row>"
        End If

        ' --- ROW BLOCK 5: NOTES (conditional) ---
        ' Written only if the notes column (BF) is populated.
        ' Preceded by a visual separator row.
        If Trim(sh.Cells(i, "BF").Value) <> "" Then
            Print #1, "        <Row><Description> </Description></Row>"
            Print #1, "        <Row>" & _
                       "<Description>NOTE: " & CleanXML(sh.Cells(i, "BF").Value) & "</Description>" & _
                       "</Row>"
        End If

        prevKey = currentKey

NextRow:
    Next i

    ' --- CLOSE LAST DOCUMENT BLOCK ---
    If prevKey <> "" Then
        Print #1, "      </Rows>"
        Print #1, "    </Document>"
    End If

    ' --- CLOSE XML AND FILE ---
    Print #1, "  </Documents>"
    Print #1, "</ERPDocuments>"
    Close #1

    MsgBox "XML file generated successfully." & vbCrLf & _
           "Path: " & xmlFile, vbInformation

End Sub
