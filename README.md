Sub VerileriAyirTopla()

    Dim ws As Worksheet
    Dim i As Long
    Dim baslangicSatir As Long
    Dim toplam As Double
    Dim sonSatir As Long

    ' "gy.baz malzeme" adlı sayfayı seçin
    Set ws = ThisWorkbook.Sheets("gy.baz malzeme")
    
    ' Veriler yalnızca 2811 satıra kadar olduğu için, son satır 2811 olarak sabitlenmiştir.
    sonSatir = 2811

    ' 1. Adım: A sütunundaki aynı verileri gruplayıp G sütunundaki değerleri toplamak
    baslangicSatir = 2 ' Veriler 2. satırdan başlıyor diye varsayıyorum
    
    For i = 2 To sonSatir
        ' A sütunundaki veri, bir önceki satırın verisinden farklı mı?
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' Eğer farklı veri tespit edildiyse, önceki grup için toplam hesaplanır
            If i - 1 >= baslangicSatir Then
                toplam = WorksheetFunction.Sum(ws.Range("G" & baslangicSatir & ":G" & i - 1))
                ws.Cells(baslangicSatir, "H").Value = toplam
                
                ' I sütununa formülü yazdır (aynı grup için aynı formül)
                Dim j As Long
                For j = baslangicSatir To i - 1
                    ws.Cells(j, "I").Formula = "=bazmalzeme[@[Kullanılan Miktar]]/H$" & baslangicSatir
                Next j
            End If
            ' Yeni grup başlangıcı
            baslangicSatir = i
        End If
    Next i

    ' Son grup için toplam işlemi ve I sütununa formül yazdırma
    toplam = WorksheetFunction.Sum(ws.Range("G" & baslangicSatir & ":G" & sonSatir))
    ws.Cells(baslangicSatir, "H").Value = toplam

    ' Son grup için formülü yazdır (aynı grup için aynı formül)
    For j = baslangicSatir To sonSatir
        ws.Cells(j, "I").Formula = "=bazmalzeme[@[Kullanılan Miktar]]/H$" & baslangicSatir
    Next j

    ' 2. Adım: E, F, G sütunları tamamen boş olan satırları sil
    For i = sonSatir To 2 Step -1
        If IsEmpty(ws.Cells(i, "E")) And IsEmpty(ws.Cells(i, "F")) And IsEmpty(ws.Cells(i, "G")) Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' 3. Adım: G sütunundaki bilgileri J sütununa kopyalama
    ' G sütunundaki veri G2'den itibaren başlar, J sütununa kopyalanacak
    For i = 2 To sonSatir
        If Not IsEmpty(ws.Cells(i, "G")) Then
            ' G hücresini J sütununa kopyala
            ws.Cells(i, "J").Value = ws.Cells(i, "G").Value
            ' Biçimleri de kopyala
            ws.Cells(i, "J").Interior.Color = ws.Cells(i, "G").Interior.Color
            ws.Cells(i, "J").Font.Color = ws.Cells(i, "G").Font.Color
            ws.Cells(i, "J").Font.Bold = ws.Cells(i, "G").Font.Bold
            ws.Cells(i, "J").Borders.LineStyle = ws.Cells(i, "G").Borders.LineStyle
        End If
    Next i
    
    ' 4. Adım: Farklı veriler arasında kalın çizgi çekme
    ' Çizgileri en son ekliyoruz, böylece diğer işlemler kalın çizginin üstüne gelmez
    For i = 2 To sonSatir
        ' Eğer A sütunundaki mevcut veri, bir önceki satırdan farklıysa
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' Farklı veri tespit edildiğinde, bir önceki satıra kalın çizgi ekle
            With ws.Rows(i - 1).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThick
            End With
        End If
    Next i

End Sub

