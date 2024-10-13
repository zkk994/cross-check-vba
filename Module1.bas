Attribute VB_Name = "Module1"
Sub CompareExcelWithCSV()
    Dim ws As Worksheet
    Dim csvFilePath As String
    Dim csvFullPath As String
    Dim excelData As Variant
    Dim csvData As Variant
    Dim folderPath As String
    Dim i As Long, k As Long
    Dim csvFileName As String
    Dim csvWorkbook As Workbook
    Dim outputSheet As Worksheet
    Dim rowCount As Long
    Dim excelValue As Variant
    Dim csvValue As Variant
    Dim maxRows As Long, maxCols As Long
    Dim excelRowCount As Long, csvRowCount As Long
    Dim excelColCount As Long, csvColCount As Long

    ' CSV dosyalar�n�n bulundu�u klas�r yolunu ayarlay�n
    folderPath = "C:\Users\90536\Downloads\" ' Dosya yolunu g�ncelleyin

    ' "Farklar" adl� sayfa olu�turun veya varsa se�in
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets("Farklar")
    On Error GoTo 0
    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Sheets.Add
        outputSheet.Name = "Farklar"
    End If
    
    ' Farklar sayfas�n� temizle
    outputSheet.Cells.Clear
    rowCount = 1
    outputSheet.Cells(rowCount, 1).Value = "M��teri Ad�"
    outputSheet.Cells(rowCount, 2).Value = "�r�n"
    outputSheet.Cells(rowCount, 3).Value = "S�tun"
    outputSheet.Cells(rowCount, 4).Value = "Excel'deki De�er"
    outputSheet.Cells(rowCount, 5).Value = "CSV'deki De�er"
    rowCount = rowCount + 1
    
    ' Her sayfa i�in d�ng�
    For Each ws In ThisWorkbook.Sheets
        ' "Farklar" sayfas�n� atla
        If ws.Name <> "Farklar" Then
            ' CSV dosya ad�n� sayfa ad�na g�re olu�tur
            csvFileName = ws.Name & ".csv"
            csvFullPath = folderPath & csvFileName
            
            ' CSV dosyas�n� kontrol et ve a�
            If Dir(csvFullPath) <> "" Then
                ' CSV dosyas�n� Workbooks.Open ile a�
                Set csvWorkbook = Workbooks.Open(csvFullPath)
                
                ' Excel'deki ve CSV'deki verileri al
                excelData = ws.UsedRange.Value
                csvData = csvWorkbook.Sheets(1).UsedRange.Value
                
                ' Sat�r ve s�tun say�s�n� bul
                excelRowCount = UBound(excelData, 1)
                csvRowCount = UBound(csvData, 1)
                excelColCount = UBound(excelData, 2)
                csvColCount = UBound(csvData, 2)
                
                ' Sat�r say�s� kontrol�
                If excelRowCount <> csvRowCount Then
                    outputSheet.Cells(rowCount, 1).Value = ws.Name
                    outputSheet.Cells(rowCount, 2).Value = "Sat�r fark�"
                    outputSheet.Cells(rowCount, 3).Value = "Excel Sat�r Say�s�"
                    outputSheet.Cells(rowCount, 4).Value = excelRowCount
                    outputSheet.Cells(rowCount, 5).Value = csvRowCount
                    rowCount = rowCount + 1
                End If
                
                ' S�tun say�s� kontrol�
                If excelColCount <> csvColCount Then
                    outputSheet.Cells(rowCount, 1).Value = ws.Name
                    outputSheet.Cells(rowCount, 2).Value = "S�tun fark�"
                    outputSheet.Cells(rowCount, 3).Value = "Excel S�tun Say�s�"
                    outputSheet.Cells(rowCount, 4).Value = excelColCount
                    outputSheet.Cells(rowCount, 5).Value = csvColCount
                    rowCount = rowCount + 1
                End If
                
                ' En k���k sat�r ve s�tun say�s�n� bul
                maxRows = Application.Min(excelRowCount, csvRowCount)
                maxCols = Application.Min(excelColCount, csvColCount)
                
                ' Sat�r ve s�tun d�ng�s� ile farklar� kontrol et
                For i = 2 To maxRows ' 2. sat�rdan ba�la (ba�l�k sat�r� atlan�r)
                    For k = 2 To maxCols ' 2. s�tundan ba�la (ba�l�k s�tunu atlan�r)
                        excelValue = excelData(i, k)
                        csvValue = csvData(i, k)
                        
                        ' Kar��la�t�rma yap ve farklar� kaydet
                        If excelValue <> csvValue Then
                            outputSheet.Cells(rowCount, 1).Value = ws.Name
                            outputSheet.Cells(rowCount, 2).Value = excelData(i, 1) ' �r�n ismi (ilk s�tun)
                            outputSheet.Cells(rowCount, 3).Value = ws.Cells(1, k).Value ' S�tun ad� (ba�l�k sat�r�ndan al)
                            outputSheet.Cells(rowCount, 4).Value = excelValue
                            outputSheet.Cells(rowCount, 5).Value = csvValue
                            rowCount = rowCount + 1
                        End If
                    Next k
                Next i
                
                ' CSV dosyas�n� kapat
                csvWorkbook.Close False
            Else
                MsgBox "CSV dosyas� bulunamad�: " & csvFileName
            End If
        End If
    Next ws
    
    MsgBox "Kar��la�t�rma tamamland�, farklar 'Farklar' sayfas�na yaz�ld�."
End Sub

