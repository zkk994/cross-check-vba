Sub ImportCSVFilesToSheets()

    ' CSV dosyalarının bulunduğu klasörü seçmek için bir dosya diyalogu açılır
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSV dosyalarının bulunduğu klasörü seçin"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "Klasör seçilmedi!", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Klasördeki tüm CSV dosyalarını bul
    Dim csvFile As String
    csvFile = Dir(folderPath & "*.csv")
    
    ' CSV dosyalarını sırayla işlemek
    Do While csvFile <> ""
        ' CSV dosyasının tam yolu
        Dim csvFilePath As String
        csvFilePath = folderPath & csvFile
        
        ' CSV dosyasının adını uzantısız al
        Dim sheetName As String
        sheetName = Replace(csvFile, ".csv", "")
        
        ' Yeni bir sheet ekleyin ve ismini CSV dosyasının ismi yapın
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
        
        ' CSV dosyasını açıp, veriyi bu sheet'e aktar
        With ws.QueryTables.Add(Connection:="TEXT;" & csvFilePath, Destination:=ws.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFileColumnDataTypes = Array(1)
            .Refresh BackgroundQuery:=False
        End With
        
        ' Sonraki CSV dosyasına geç
        csvFile = Dir
    Loop
    
    MsgBox "Tüm CSV dosyaları başarıyla aktarıldı!", vbInformation

End Sub
