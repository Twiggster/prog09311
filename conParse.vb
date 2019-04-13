Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Windows.Forms

Module conParse
    Public lstData As New List(Of clsDataExtraction)
    Sub Main()
        'Prompting the user for input
        Console.WriteLine("Please enter a valid file path for reading and parsing.")
        Dim strInputFile As String = Console.ReadLine

        'If logic for entering a proper file 
        If Not File.Exists(strInputFile) Then
            Console.WriteLine("This is not a valid file path. Auf Wiedersehen!")
            Console.ReadLine()
            Return
        End If

        ProcessFile(strInputFile, lstData)
        StartExcel()

    End Sub

    Sub ProcessFile(strInputFile As String, lstData As List(Of clsDataExtraction))
        Dim intCount As Integer

        Using objFileStream As StreamReader = File.OpenText(strInputFile)
            While Not objFileStream.EndOfStream

                'Reads a line from the file
                Dim data = New clsDataExtraction(objFileStream.ReadLine.Split(","))

                'Will hold all of the lines of the file once its done
                intCount += 1

                lstData.Add(data)
            End While
        End Using
        Console.WriteLine("Processing Completed..")
        Console.ReadLine()
    End Sub

    Sub StartExcel()
        Dim CheckExcel As Object
        Dim anExcelDoc As Excel.Application
        Dim intFileCount As Integer
        Dim intLoop As Integer
        Dim intStatsBeg As Integer


        Try
            CheckExcel = GetObject(, "Excel.Application")
        Catch ex As Exception

        End Try

        If CheckExcel Is Nothing Then
            anExcelDoc = New Excel.Application()
            anExcelDoc.Visible = True
        Else
            anExcelDoc = CheckExcel
            anExcelDoc.Visible = True
        End If

        'Gets the amount of lines read in
        intFileCount = conParse.lstData.Count()

        'Adds the Excel workbook
        anExcelDoc.Workbooks.Add()
        anExcelDoc.Sheets.Add()


        'For intFileCount = 1 To 5
        '    anExcelDoc.Cells(intFileCount, 1) = 100 * intFileCount
        'Next


        'For intLoop = 1 To intFileCount
        '    anExcelDoc.Cells(intLoop, 1) = clsDataExtraction.GetStoreNumber()
        'Next


        'Adding the store number to ColumA(FileCount)

        For intLoop = 1 To intFileCount

            For Each store In lstData

                anExcelDoc.Cells(intLoop, 1) = store.GetStoreNumber

            Next
        Next

        'Where the stats calcs will be populated 
        intStatsBeg = intFileCount + 2

        'This begins two lines beneath the file length
        anExcelDoc.Cells(intStatsBeg, 1) = "Ave:"
        anExcelDoc.Cells(intStatsBeg, 2) = "=average(a1..intFileCount)"
        intStatsBeg += 1
        'After each time we add something we skip down a cell
        anExcelDoc.Cells(intStatsBeg, 1) = "Min:"
        intStatsBeg += 1
        'Move down one cell
        anExcelDoc.Cells(intStatsBeg, 1) = "Total:"
        intStatsBeg += 1
        'Move to the last one which will land 5 bellow the file length
        anExcelDoc.Cells(intStatsBeg, 1) = "Max"

        MessageBox.Show("Data has been appended to the workbook")
    End Sub
End Module
