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

    ' CSV dosyalarýnýn bulunduðu klasör yolunu ayarlayýn
    folderPath = "C:\Users\90536\Downloads\" ' Dosya yolunu güncelleyin

    ' "Farklar" adlý sayfa oluþturun veya varsa seçin
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets("Farklar")
    On Error GoTo 0
    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Sheets.Add
        outputSheet.Name = "Farklar"
    End If
    
    ' Farklar sayfasýný temizle
    outputSheet.Cells.Clear
    rowCount = 1
    outputSheet.Cells(rowCount, 1).Value = "Müþteri Adý"
    outputSheet.Cells(rowCount, 2).Value = "Ürün"
    outputSheet.Cells(rowCount, 3).Value = "Sütun"
    outputSheet.Cells(rowCount, 4).Value = "Excel'deki Deðer"
    outputSheet.Cells(rowCount, 5).Value = "CSV'deki Deðer"
    rowCount = rowCount + 1
    
    ' Her sayfa için döngü
    For Each ws In ThisWorkbook.Sheets
        ' "Farklar" sayfasýný atla
        If ws.Name <> "Farklar" Then
            ' CSV dosya adýný sayfa adýna göre oluþtur
            csvFileName = ws.Name & ".csv"
            csvFullPath = folderPath & csvFileName
            
            ' CSV dosyasýný kontrol et ve aç
            If Dir(csvFullPath) <> "" Then
                ' CSV dosyasýný Workbooks.Open ile aç
                Set csvWorkbook = Workbooks.Open(csvFullPath)
                
                ' Excel'deki ve CSV'deki verileri al
                excelData = ws.UsedRange.Value
                csvData = csvWorkbook.Sheets(1).UsedRange.Value
                
                ' Satýr ve sütun sayýsýný bul
                excelRowCount = UBound(excelData, 1)
                csvRowCount = UBound(csvData, 1)
                excelColCount = UBound(excelData, 2)
                csvColCount = UBound(csvData, 2)
                
                ' Satýr sayýsý kontrolü
                If excelRowCount <> csvRowCount Then
                    outputSheet.Cells(rowCount, 1).Value = ws.Name
                    outputSheet.Cells(rowCount, 2).Value = "Satýr farký"
                    outputSheet.Cells(rowCount, 3).Value = "Excel Satýr Sayýsý"
                    outputSheet.Cells(rowCount, 4).Value = excelRowCount
                    outputSheet.Cells(rowCount, 5).Value = csvRowCount
                    rowCount = rowCount + 1
                End If
                
                ' Sütun sayýsý kontrolü
                If excelColCount <> csvColCount Then
                    outputSheet.Cells(rowCount, 1).Value = ws.Name
                    outputSheet.Cells(rowCount, 2).Value = "Sütun farký"
                    outputSheet.Cells(rowCount, 3).Value = "Excel Sütun Sayýsý"
                    outputSheet.Cells(rowCount, 4).Value = excelColCount
                    outputSheet.Cells(rowCount, 5).Value = csvColCount
                    rowCount = rowCount + 1
                End If
                
                ' En küçük satýr ve sütun sayýsýný bul
                maxRows = Application.Min(excelRowCount, csvRowCount)
                maxCols = Application.Min(excelColCount, csvColCount)
                
                ' Satýr ve sütun döngüsü ile farklarý kontrol et
                For i = 2 To maxRows ' 2. satýrdan baþla (baþlýk satýrý atlanýr)
                    For k = 2 To maxCols ' 2. sütundan baþla (baþlýk sütunu atlanýr)
                        excelValue = excelData(i, k)
                        csvValue = csvData(i, k)
                        
                        ' Karþýlaþtýrma yap ve farklarý kaydet
                        If excelValue <> csvValue Then
                            outputSheet.Cells(rowCount, 1).Value = ws.Name
                            outputSheet.Cells(rowCount, 2).Value = excelData(i, 1) ' Ürün ismi (ilk sütun)
                            outputSheet.Cells(rowCount, 3).Value = ws.Cells(1, k).Value ' Sütun adý (baþlýk satýrýndan al)
                            outputSheet.Cells(rowCount, 4).Value = excelValue
                            outputSheet.Cells(rowCount, 5).Value = csvValue
                            rowCount = rowCount + 1
                        End If
                    Next k
                Next i
                
                ' CSV dosyasýný kapat
                csvWorkbook.Close False
            Else
                MsgBox "CSV dosyasý bulunamadý: " & csvFileName
            End If
        End If
    Next ws
    
    MsgBox "Karþýlaþtýrma tamamlandý, farklar 'Farklar' sayfasýna yazýldý."
End Sub

