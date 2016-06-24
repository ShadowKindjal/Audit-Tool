Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports System.Environment
Imports System.Xml
Imports System.Windows.Forms
Imports System.IO.Compression
Imports System.IO

Public Class AuditAnalyzer

    Private Sub Report_Click(sender As Object, e As RibbonControlEventArgs) Handles Report.Click
        'Try
        SortData()
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        '    Globals.ThisAddIn.Application.ScreenUpdating = True
        'End Try
    End Sub

    Sub SortData()
        'Globals.ThisAddIn.Application.ScreenUpdating = False

        Dim Start, Row, Column, Offset, Index, Result As Integer
        Dim Source, Data As Range
        Dim Specimen(1), BranchCode(1), Inventory(1), Patient(1), Original_to_Destination(1), Department(1), WorksheetNumber(1), WorkSheetName(1), CustomerCare(1) As String
        Dim WS, Sheet As Excel.Worksheet
        Dim Report, DataTable As Excel.ListObject
        Dim bool, rbool As Boolean

        Globals.ThisAddIn.Application.DisplayAlerts = False
        For Each Sheet In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets
            If Sheet.Name <> "Sheet1" Then Sheet.Delete()
        Next Sheet
        Globals.ThisAddIn.Application.DisplayAlerts = True

        WS = CType(Globals.ThisAddIn.Application.Worksheets.Add(, Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count)), Excel.Worksheet)
        WS.Name = "Report"
        With WS
            .Range("A1").Value = "Specimen"
            .Range("B1").Value = "Specimen Barcode"
            .Range("C1").Value = "Branch Code"
            .Range("D1").Value = "Inventory"
            .Range("E1").Value = "Patient"
            .Range("F1").Value = "Orig/Dest"
            .Range("G1").Value = "Department"
            .Range("H1").Value = "Worksheet Number"
            .Range("I1").Value = "Worksheet Number Barcode"
            .Range("J1").Value = "WorkSheet Name"
            .Range("K1").Value = "Customer Care"
            .ListObjects.Add(XlListObjectSourceType.xlSrcRange, .Range("A1:K2"), , XlYesNoGuess.xlYes).Name = "Report"
            .ListObjects("Report").TableStyle = "TableStyleLight1"
            Report = .ListObjects("Report")
            With Report
                .ListColumns("Specimen").DataBodyRange.NumberFormat = "000-000-0000-0"
                With .ListColumns("Specimen Barcode").DataBodyRange
                    .Font.Name = "Free 3 of 9"
                    .Font.Size = 26
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                End With
                .ListColumns("Branch Code").DataBodyRange.NumberFormat = "000"
                .ListColumns("Department").DataBodyRange.NumberFormat = "000"
                .ListColumns("Worksheet Number").DataBodyRange.NumberFormat = "0000"
                With .ListColumns("Worksheet Number Barcode").DataBodyRange
                    .Font.Name = "Free 3 of 9"
                    .Font.Size = 26
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                End With
            End With
        End With

        Sheet = Globals.ThisAddIn.Application.Worksheets("Sheet1")
        Dim LastRow, LastColumn As Long
        Dim DataArray As Array
        Start = 3
        Index = 0

        With Sheet
            LastRow = Sheet.Cells(Sheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            LastColumn = Sheet.Cells(1, Sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
            Source = .Range("A1", .Cells(LastRow + 3, LastColumn + 3))
            DataArray = Source.Value
        End With
        For Row = 1 To UBound(DataArray, 1)
            'Specimen Check
            If Int32.TryParse(Left(DataArray(Row, 2), 1), Result) Then
                Specimen(Index) = DataArray(Row, 2)
                bool = True
            ElseIf Int32.TryParse(Left(DataArray(Row, 1), 1), Result) Then
                Specimen(Index) = DataArray(Row, 1)
                bool = True
            Else
                bool = False
            End If

            If bool Then
                BranchCode(Index) = Mid(Specimen(Index), 5, 3)

                If DataArray(Row + 1, 1) <> Nothing Then
                    Inventory(Index) = DataArray(Row + 1, 1)
                ElseIf DataArray(Row + 2, 1) <> Nothing Then
                    Inventory(Index) = DataArray(Row + 2, 1)
                End If

                'For Column = 3 To UBound(DataArray, 2)
                '    If DataArray(Row, Column) <> Nothing Then
                '        If Int32.TryParse(Right(DataArray(Row, Column), 1), Result) Then
                '            Patient(Index) = DataArray(Row - 1, Column)
                '            Start = Column + 1
                '            Exit For
                '        Else
                '            Patient(Index) = DataArray(Row, Column)
                '            Start = Column + 1
                '            Exit For
                '        End If
                '    End If
                'Next
                Offset = 0
                rbool = False
                For Column = 3 To UBound(DataArray, 2)
                    For Offset = 0 To 2
                        'MsgBox(Column) 211-005-1833-0
                        'If Specimen(Index) = "211-005-1833-0" Then MsgBox(Column & ", " & Row & vbCrLf & Offset & vbCrLf & DataArray(Row + Offset, Column) & vbCrLf & DataArray(Row - Offset, Column))
                        If DataArray(Row + Offset, Column) <> Nothing Then
                            If Int32.TryParse(Right(DataArray(Row + Offset, Column), 1), Result) And DataArray(Row + Offset - 1, Column) <> Nothing And
                                (Int32.TryParse(Right(DataArray(Row + Offset - 1, Column), 1), Result) Or Int32.TryParse(Mid(DataArray(Row + Offset - 1, Column), 2, 1), Result)) = False Then

                                Patient(Index) = DataArray(Row + Offset - 1, Column)
                                'Patient(Index) = Len(DataArray(Row + Offset, Column))
                                Start = Column + 1
                                rbool = True
                                Exit For
                            ElseIf DataArray(Row + Offset + 1, Column) = Nothing And (Int32.TryParse(Right(DataArray(Row + Offset, Column), 1), Result) Or Int32.TryParse(Mid(DataArray(Row + Offset, Column), 2, 1), Result)) = False Then
                                Patient(Index) = DataArray(Row + Offset, Column)
                                Start = Column + 1
                                rbool = True
                                Exit For
                            ElseIf Int32.TryParse(DataArray(Row + Offset, Column + 1), Result) = True And DataArray(Row + Offset - 1, Column + 1) = Nothing _
                                And Int32.TryParse(DataArray(Row + Offset, Column), Result) = False Then
                                Patient(Index) = DataArray(Row - Offset - 1, Column) & DataArray(Row - Offset, Column)
                                Start = Column + 1
                                rbool = True
                                Exit For
                            End If
                        End If
                        'If Specimen(Index) = "211-005-1833-0" Then MsgBox(Column & ", " & Row & vbCrLf & Offset)
                        If DataArray(Row - Offset, Column) <> Nothing Then
                            If Int32.TryParse(Right(DataArray(Row - Offset, Column), 1), Result) And DataArray(Row - Offset - 1, Column) <> Nothing And
                                (Int32.TryParse(Right(DataArray(Row - Offset - 1, Column), 1), Result) Or Int32.TryParse(Mid(DataArray(Row - Offset - 1, Column), 2, 1), Result)) = False Then

                                Patient(Index) = DataArray(Row - Offset - 1, Column)
                                'Patient(Index) = Len(DataArray(Row + Offset, Column))
                                Start = Column + 1
                                rbool = True
                                Offset = Offset * -1
                                Exit For
                            ElseIf DataArray(Row - Offset + 1, Column) = Nothing And (Int32.TryParse(Right(DataArray(Row - Offset, Column), 1), Result) Or Int32.TryParse(Mid(DataArray(Row - Offset, Column), 2, 1), Result)) = False Then
                                Patient(Index) = DataArray(Row - Offset, Column)
                                Start = Column + 1
                                rbool = True
                                Offset = Offset * -1
                                Exit For
                            ElseIf Int32.TryParse(DataArray(Row - Offset, Column + 1), Result) = True And DataArray(Row - Offset - 1, Column + 1) = Nothing _
                                And Int32.TryParse(DataArray(Row - Offset, Column), Result) = False Then
                                Patient(Index) = DataArray(Row - Offset - 1, Column) & DataArray(Row - Offset, Column)
                                Start = Column + 1
                                rbool = True
                                Offset = Offset * -1
                                Exit For
                            End If
                        End If
                        'If Specimen(Index) = "211-005-1833-0" Then MsgBox(Column & ", " & Row & vbCrLf & Offset)
                    Next
                    If rbool = True Then Exit For
                Next

                If Patient(Index).Contains("Ctrl Number") Or Patient(Index) = "" Then
                    Patient(Index) = ""
                    Offset = 0
                End If
                If Inventory(Index).Contains("ESP Reference Audit") Then
                    Inventory(Index) = Offset
                End If

                rbool = False
                For Column = Start To UBound(DataArray, 2)
                    For Offset = 0 To 2
                        If CStr(DataArray(Row + Offset, Column)) = "/" Then
                            Original_to_Destination(Index) = DataArray(Row + Offset - 1, Column - 1) & " to " & DataArray(Row + Offset, Column - 1)
                            Department(Index) = DataArray(Row + Offset, Column + 1)
                            WorksheetNumber(Index) = DataArray(Row, Column + 2)
                            Start = Column + 3
                            Exit For
                        ElseIf CStr(DataArray(Row - Offset, Column)) = "/" Then
                            Original_to_Destination(Index) = DataArray(Row - Offset - 1, Column - 1) & " to " & DataArray(Row - Offset, Column - 1)
                            Department(Index) = DataArray(Row - Offset, Column + 1)
                            WorksheetNumber(Index) = DataArray(Row, Column + 2)
                            Start = Column + 3
                            Offset = Offset * -1
                            Exit For
                        End If
                    Next
                    If rbool = True Then Exit For
                Next
                    For Column = Start To UBound(DataArray, 2)
                        If DataArray(Row, Column) <> Nothing Then
                            WorkSheetName(Index) = DataArray(Row, Column)
                            Start = Column + 1
                            Exit For
                        ElseIf DataArray(Row - 1, Column) <> Nothing Then
                            WorkSheetName(Index) = DataArray(Row - 1, Column)
                            Start = Column + 1
                            Exit For
                        End If
                    Next
                    For Column = Start To UBound(DataArray, 2)
                        If DataArray(Row, Column) <> Nothing Then
                            CustomerCare(Index) = DataArray(Row, Column)
                            Exit For
                        ElseIf DataArray(Row - 1, Column) <> Nothing Then
                            CustomerCare(Index) = DataArray(Row - 1, Column)
                            Exit For
                        End If
                    Next

                    Index += 1
                    ReDim Preserve Specimen(Index + 1)
                    ReDim Preserve BranchCode(Index + 1)
                    ReDim Preserve Inventory(Index + 1)
                    ReDim Preserve Patient(Index + 1)
                    ReDim Preserve Original_to_Destination(Index + 1)
                    ReDim Preserve Department(Index + 1)
                    ReDim Preserve WorksheetNumber(Index + 1)
                    ReDim Preserve WorkSheetName(Index + 1)
                    ReDim Preserve CustomerCare(Index + 1)

                End If
        Next

        Dim TableArray(UBound(Specimen, 1), 11) As String

        For Index = 0 To UBound(TableArray, 1)
            TableArray(Index, 0) = Specimen(Index)
            'If Specimen(Index) <> "" Then TableArray(Index, 1) = " * " & Specimen(Index).Replace(" - ", "") & " * "
            TableArray(Index, 2) = BranchCode(Index)
            TableArray(Index, 3) = Inventory(Index)
            TableArray(Index, 4) = Patient(Index)
            TableArray(Index, 5) = Original_to_Destination(Index)
            TableArray(Index, 6) = Department(Index)
            If WorksheetNumber(Index) <> "" Then
                Do Until WorksheetNumber(Index).Length = 4
                    WorksheetNumber(Index) = "0" & WorksheetNumber(Index)
                Loop
                TableArray(Index, 7) = WorksheetNumber(Index)
                TableArray(Index, 8) = "*" & WorksheetNumber(Index) & "*"
            End If
            TableArray(Index, 7) = WorksheetNumber(Index)
            TableArray(Index, 8) = "*" & WorksheetNumber(Index) & "*"
            TableArray(Index, 9) = WorkSheetName(Index)
            TableArray(Index, 10) = CustomerCare(Index)
        Next

        Data = WS.Range("A2", WS.Cells(UBound(TableArray, 1), UBound(TableArray, 2)))
        Data.Value = TableArray


        DataTable = Table("BranchReport", "Branch, Branch Code, Instances, Percentage")
        DataTable.HeaderRowRange(2).EntireColumn.NumberFormat = "000"
        Call Quantities("Branch Code", DataTable, "BranchReport", True, False)
        DataTable.ListColumns("Branch Code").DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight
        DataTable.ListColumns("Percentage").DataBodyRange.NumberFormat = "0.00%"
        DataTable = Table("WorksheetReport", "Worksheet Name, Worksheet Number, Instances, Percentage")
        DataTable.HeaderRowRange(2).EntireColumn.NumberFormat = "0000"
        Call Quantities("Worksheet Number", DataTable, "WorksheetReport", False, True)
        DataTable.ListColumns("Percentage").DataBodyRange.NumberFormat = "0.00%"
        DataTable = Table("DepartmentReport", "Department, Instances, Percentage")
        DataTable.HeaderRowRange(1).EntireColumn.NumberFormat = "000"
        Call Quantities("Department", DataTable, "DepartmentReport", False, False)
        DataTable.ListColumns("Percentage").DataBodyRange.NumberFormat = "0.00%"

        Report.ShowTotals = True
        Report.ListColumns("Branch Code").TotalsCalculation = XlTotalsCalculation.xlTotalsCalculationCount
        Report.ListColumns(2).DataBodyRange.Formula = "=""*"" & [@Specimen] & ""*"""
        With Report.Sort
            .SortFields.Clear()
            .SortFields.Add(Report.ListColumns("Branch Code").DataBodyRange, XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, XlSortDataOption.xlSortNormal)
            .Header = XlYesNoGuess.xlYes
            .MatchCase = False
            .Orientation = XlSortOrientation.xlSortColumns
            .SortMethod = XlSortMethod.xlPinYin
            .Apply()
        End With

        With WS
            .Range("A1").EntireRow.Insert()
            .Range("A1").Value = "Total Specimen"
            .Range("B1").Value = "=COUNTA(Report[Branch Code])"
        End With
        For Each Sheet In Globals.ThisAddIn.Application.Worksheets
            For Each LO In Sheet.ListObjects
                LO.Range.EntireColumn.AutoFit()
            Next
        Next

        WS.Activate()

        'Dim EventCalc As Excel.DocEvents_CalculateEventHandler = New Excel.DocEvents_CalculateEventHandler(AddressOf Filter)
        'AddHandler WS.Calculate, EventCalc

        Globals.ThisAddIn.Application.ScreenUpdating = True
    End Sub

    Function Table(ByRef Name As String, ByRef Headers As String) As Excel.ListObject
        Dim WS As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets.Add(, Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count)), Excel.Worksheet)
        Dim Data As Range = WS.Range("A1")
        WS.Name = Name
        Dim HeaderNames As String() = Split(Headers, ", ")
        For x = 0 To HeaderNames.Length - 1
            Data.Offset(, x).Value = HeaderNames(x)
        Next
        With WS
            .ListObjects.Add(XlListObjectSourceType.xlSrcRange, Data.Resize(1, HeaderNames.Length), , XlYesNoGuess.xlYes).Name = Name
            .ListObjects(Name).TableStyle = "TableStyleLight1"
            Table = .ListObjects(Name)
        End With
    End Function

    Sub Quantities(ByVal Column As String, ByVal Table As Excel.ListObject, ByVal Name As String, ByVal BranchCode As Boolean, ByVal Worksheet As Boolean)
        Dim WS As Excel.Worksheet = Globals.ThisAddIn.Application.Worksheets("Report")
        Dim RS As Excel.Worksheet = Globals.ThisAddIn.Application.Worksheets(Name)
        Dim Report As Excel.ListObject = WS.ListObjects("Report")
        Dim Index, tIndex, Count As Integer
        Dim Data As Range
        Dim DataArray, WkshtArray As Array
        Dim Duplicate As Boolean
        Dim Colshift As Integer = 0
        If Not (BranchCode Or Worksheet) Then Colshift = 1
        'Data = Report.ListColumns(Column).DataBodyRange.AdvancedFilter()
        DataArray = Report.ListColumns(Column).DataBodyRange.Value
        If Worksheet Then WkshtArray = Report.ListColumns("WorkSheet Name").DataBodyRange.Value
        Count = 0
        For Index = 1 To UBound(DataArray, 1)
            Duplicate = False
            For tIndex = Index + 1 To UBound(DataArray, 1)
                If DataArray(Index, 1) = DataArray(tIndex, 1) Then
                    Duplicate = True
                End If
            Next
            If Not Duplicate Then
                Count += 1
            End If
        Next
        Dim TableArray(Count, 3) As String
        Count = 0
        For Index = 1 To UBound(DataArray, 1)
            Duplicate = False
            For tIndex = Index + 1 To UBound(DataArray, 1)
                If DataArray(Index, 1) = DataArray(tIndex, 1) Then
                    Duplicate = True
                End If
            Next
            If Not Duplicate Then
                TableArray(Count, 1 - Colshift) = DataArray(Index, 1)
                If BranchCode Then TableArray(Count, 0) = Branch(DataArray(Index, 1))
                If Worksheet Then TableArray(Count, 0) = WkshtArray(Index, 1)
                TableArray(Count, 2 - Colshift) = 0
                Count += 1
                End If
        Next
        For Index = 0 To UBound(TableArray, 1) - 1
            For tIndex = 1 To UBound(DataArray, 1)
                If TableArray(Index, 1 - Colshift) = DataArray(tIndex, 1) Then TableArray(Index, 2 - Colshift) += 1
            Next
        Next


        Data = RS.Range("A2", RS.Cells(UBound(TableArray, 1), UBound(TableArray, 2)))
        Data.Value = TableArray

        With Table.ListColumns("Instances").DataBodyRange
            .NumberFormat = "General"
            .Value = .Value
        End With
        Table.ListColumns("Percentage").DataBodyRange.Formula = "=[@Instances]/SUM([Instances])"
    End Sub

    Function Exists(Code As String, Data As Range) As Range
        Exists = Nothing
        Dim NativeWorksheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.Worksheets("Report")
        Dim cell As Range
        For Each cell In Data.Cells
            If cell.Text = Code Then
                Exists = cell
                Exit For
            Else
                Exists = NativeWorksheet.Range("A1")
            End If
        Next cell
    End Function

    Function Branch(Code As String) As String
        Branch = String.Empty
        Dim XML As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer\BranchCodes.XML"
        Dim XMLDoc As New XmlDocument
        XMLDoc.Load(XML)
        For Each Node As XmlNode In XMLDoc.DocumentElement.ChildNodes
            If Node.InnerText = Code Then Branch = Node.Name.Replace("_", " ").Replace("--", "/").Replace("..", "'").Replace(".-", ")").Replace("-.", "(")
        Next
    End Function

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim Manager As Form = New Branch_Code_Manager
        Manager.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Barcode()
    End Sub

    Private Sub Barcode()
        Dim zipPath As String = Environment.GetFolderPath(SpecialFolder.MyDocuments) & "font.zip"
        Dim extractPath As String = Environment.GetFolderPath(SpecialFolder.Fonts)
        Dim fso
        Dim bool As Boolean
        Dim res As Long
        Const WM_FONTCHANGE As Integer = &H1D
        Const HWND_BROADCAST As Integer = &HFFFF&
        fso = CreateObject("Scripting.FileSystemObject") 'Creating File System 
        Try
            If My.Computer.FileSystem.FileExists(zipPath) Then My.Computer.FileSystem.DeleteFile(zipPath)
            My.Computer.Network.DownloadFile("http://www.squaregear.net/fonts/free3of9.zip", zipPath)
            Using archive As ZipArchive = ZipFile.OpenRead(zipPath)
                For Each entry As ZipArchiveEntry In archive.Entries
                    If entry.FullName.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) Then
                        If Not fso.FileExists(Path.Combine(extractPath, entry.FullName)) Then
                            bool = True
                            entry.ExtractToFile(Path.Combine(extractPath, entry.FullName))
                        Else
                            bool = False
                            Try
                                fso.DeleteFile(Path.Combine(extractPath, entry.FullName))
                                entry.ExtractToFile(Path.Combine(extractPath, entry.FullName))
                            Catch ex As Exception

                            End Try
                        End If
                        res = AddFontResource(Path.Combine(extractPath, entry.FullName))
                        SendMessage(New System.IntPtr(HWND_BROADCAST), WM_FONTCHANGE, 0, Nothing)
                    End If
                Next
            End Using
            If My.Computer.FileSystem.FileExists(zipPath) Then My.Computer.FileSystem.DeleteFile(zipPath)
            MsgBox("Successfully installed barcode font.")
        Catch ex As Exception
            If bool Then MsgBox("There was an error installing the barcode font. Barcodes may not be displayed correctly in Excel. To address this issue restart Excel with admin rights.")
        End Try
    End Sub

    Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
    Public Function SendMessage(ByVal hWnd As Integer, ByVal Msg As UInteger, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        Return Nothing
    End Function
End Class


