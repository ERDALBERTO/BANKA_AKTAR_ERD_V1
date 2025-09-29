Option Explicit

' === SABİT AYARLAR ===
Public Const SHEET_OZET As String = "OZET"
Public Const SHEET_ARINDIR As String = "ARINDIR"
Public Const SHEET_GB As String = "GB"
Public Const SHEET_BNKLIST As String = "BNKLIST"
Public Const SHEET_CHLIST As String = "CHLIST"
Public Const SHEET_BNK_ARINDIR As String = "BNK_ARINDIR"
Public Const MAX_ARINDIR_TOKENS As Long = 200

' === STOP TOKEN BELLEĞİ ===
Dim gStopTokens() As String
Dim gStopDict As Object
Dim gStopTokensCount As Long

' === BNK_ARINDIR TOKEN BELLEĞİ ===
Dim gBnkTokens() As String
Dim gBnkDict As Object
Dim gBnkTokensCount As Long

' === GB'DEN OZET'E VERİ ÇEKME ===
Public Sub GB_EKSTRE_GETIR_HEPSI()
    Dim srcWS As Worksheet, dstWS As Worksheet
    Dim srcLast As Long, existingLast As Long
    Dim i As Long
    Dim txt As String
    Dim regex As Object, matches As Object
    Dim rate As String, dolar As Integer

    Set srcWS = ThisWorkbook.Worksheets(SHEET_GB)
    Set dstWS = ThisWorkbook.Worksheets(SHEET_OZET)

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .IgnoreCase = True
        .Global = False
        .pattern = "(?:Döviz Kuru|döviz kuru|kur|kurundan|amerikan doları)\s*[:\-]?\s*([0-9]+[.,][0-9]+)"
    End With

    existingLast = dstWS.Cells(dstWS.Rows.count, "E").End(xlUp).Row
    If existingLast < 5 Then existingLast = 5
    dstWS.Range("B5:T" & existingLast).ClearContents

    srcLast = srcWS.Cells(srcWS.Rows.count, "A").End(xlUp).Row
    If srcLast < 5 Then
        MsgBox "Kaynak sayfada yeterli veri bulunamadı.", vbExclamation
        Exit Sub
    End If

    For i = 5 To srcLast
        dstWS.Cells(i, "E").Value = srcWS.Cells(i, "A").Value
        dstWS.Cells(i, "F").Value = srcWS.Cells(i, "B").Value
        dstWS.Cells(i, "G").Value = srcWS.Cells(i, "D").Value

        If srcWS.Cells(i, "D").Value < 0 Then
            dstWS.Cells(i, "H").Value = "B"
        Else
            dstWS.Cells(i, "H").Value = "A"
        End If

        If dstWS.Cells(i, "G").Value < 0 Then
            dstWS.Cells(i, "I").Value = -dstWS.Cells(i, "G").Value
        Else
            dstWS.Cells(i, "I").Value = dstWS.Cells(i, "G").Value
        End If

        txt = srcWS.Cells(i, "B").Value

        If InStr(1, txt, "USD", vbTextCompare) _
           Or InStr(1, txt, "EUR", vbTextCompare) _
           Or InStr(1, txt, "TL", vbTextCompare) _
           Or InStr(1, txt, "dolar", vbTextCompare) _
           Or InStr(1, txt, "amerikan doları", vbTextCompare) Then

            dstWS.Cells(i, "J").Value = "DÖVİZLİ"

            If InStr(1, txt, "USD", vbTextCompare) _
               Or InStr(1, txt, "dolar", vbTextCompare) _
               Or InStr(1, txt, "amerikan doları", vbTextCompare) Then
                dolar = 1
                dstWS.Cells(i, "K").Value = dolar
            ElseIf InStr(1, txt, "EUR", vbTextCompare) _
                   Or InStr(1, txt, "euro", vbTextCompare) Then
                dstWS.Cells(i, "K").Value = 2
            ElseIf InStr(1, txt, "TL", vbTextCompare) Then
                dstWS.Cells(i, "K").Value = 0
            End If

            If regex.Test(txt) Then
                Set matches = regex.Execute(txt)
                rate = matches(0).SubMatches(0)
                rate = Replace(rate, ",", ".")
                dstWS.Cells(i, "L").Value = rate
            Else
                dstWS.Cells(i, "L").Value = ""
            End If
        Else
            dstWS.Cells(i, "J").Value = ""
            dstWS.Cells(i, "K").Value = ""
            dstWS.Cells(i, "L").Value = ""
        End If
    Next i

    Set regex = Nothing
    Set srcWS = Nothing
    Set dstWS = Nothing
End Sub

' === BNK_ARINDIR STOP TOKEN CACHE GÜNCELLEME ===
Public Sub BA_RebuildBnkStopList()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_BNK_ARINDIR)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' BNK_ARINDIR sayfası yoksa boş cache oluştur
        gBnkTokensCount = 0
        Set gBnkDict = CreateObject("Scripting.Dictionary")
        Exit Sub
    End If
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Dim tmp() As String, count As Long, i As Long, raw As String, norm As String
    
    For i = 2 To lastRow
        raw = Trim$(ws.Cells(i, 1).Value)
        If Len(raw) > 0 Then
            norm = BA_NormalizeText(raw)
            If Not BA_InArr(norm, tmp, count) Then
                count = count + 1
                ReDim Preserve tmp(1 To count)
                tmp(count) = norm
            End If
            If count >= MAX_ARINDIR_TOKENS Then Exit For
        End If
    Next i
    
    gBnkTokensCount = count
    If count > 0 Then
        ReDim gBnkTokens(1 To count)
    End If
    Set gBnkDict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To count
        gBnkTokens(i) = tmp(i)
        If Not gBnkDict.Exists(tmp(i)) Then
            gBnkDict.Add tmp(i), True
        End If
    Next i
End Sub

' === BANKA MI DOLDURMA (GELİŞTİRİLMİŞ) ===
Public Sub BANKA_MI_DOLDUR()
    Dim wsOzet As Worksheet, wsBnkList As Worksheet
    Set wsOzet = ThisWorkbook.Worksheets(SHEET_OZET)
    Set wsBnkList = ThisWorkbook.Worksheets(SHEET_BNKLIST)
    
    Dim bankalar As Object: Set bankalar = CreateObject("Scripting.Dictionary")
    Dim lastBnk As Long, r As Long, i As Long
    Dim nRow As Long, aciklama As String, bankaMi As String
    Dim dövizKelimeler As Variant, key As Variant

    ' Banka kriterlerini genişlettik
    dövizKelimeler = Array("DOLAR", "USD", "EURO", "EUR", "AMERIKAN DOLARI", "TL", "TÜRK LİRASI", _
                          "KART", "K.KARTI", "KREDİ KARTI", "BONUS", "WORLD", "MAXIMUM", "AXESS", _
                          "GARANTI", "İŞ BANKASI", "AKBANK", "YAPIKRED", "ZIRAAT", "HALKBANK", _
                          "VEB", "TEB", "DENIZBANK", "ING", "QNB", "ÖDEME")

    lastBnk = wsBnkList.Cells(wsBnkList.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastBnk
        Dim bankaAdi As String
        bankaAdi = UCase(Trim(wsBnkList.Cells(i, "A").Value))
        If Len(bankaAdi) > 0 Then
            bankalar(bankaAdi) = True
        End If
    Next i

    nRow = wsOzet.Cells(wsOzet.Rows.count, "F").End(xlUp).Row
    For r = 5 To nRow
        aciklama = UCase(Trim(wsOzet.Cells(r, "F").Value))
        bankaMi = "H"
        
        ' Önce banka listesinden kontrol et
        For Each key In bankalar.Keys
            If InStr(aciklama, key) > 0 Then
                bankaMi = "E": Exit For
            End If
        Next key
        
        ' Döviz ve kart kelimelerini kontrol et
        If bankaMi = "H" Then
            For i = LBound(dövizKelimeler) To UBound(dövizKelimeler)
                If InStr(aciklama, dövizKelimeler(i)) > 0 Then
                    bankaMi = "E": Exit For
                End If
            Next i
        End If
        
        wsOzet.Cells(r, "N").Value = bankaMi
    Next r
End Sub

' === ARINDIR STOP TOKEN CACHE GÜNCELLEME ===
Public Sub BA_RebuildStopList()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_ARINDIR)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Dim tmp() As String, count As Long, i As Long, raw As String, norm As String
    
    For i = 2 To lastRow
        raw = Trim$(ws.Cells(i, 1).Value)
        If Len(raw) > 0 Then
            norm = BA_NormalizeText(raw)
            If Not BA_InArr(norm, tmp, count) Then
                count = count + 1
                ReDim Preserve tmp(1 To count)
                tmp(count) = norm
            End If
            If count >= MAX_ARINDIR_TOKENS Then Exit For
        End If
    Next i
    
    gStopTokensCount = count
    If count > 0 Then
        ReDim gStopTokens(1 To count)
    End If
    Set gStopDict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To count
        gStopTokens(i) = tmp(i)
        If Not gStopDict.Exists(tmp(i)) Then
            gStopDict.Add tmp(i), True
        End If
    Next i
End Sub

' === Yardımcı: Dizi İçinde Var mı? ===
Public Function BA_InArr(val As String, arr() As String, cnt As Long) As Boolean
    Dim i As Long
    For i = 1 To cnt
        If arr(i) = val Then BA_InArr = True: Exit Function
    Next i
    BA_InArr = False
End Function

' === BANKA İÇİN ÖZEL ARINDIRMA ===
Public Function BA_CleanBankDescription(ByVal s As String) As String
    Dim u As String
    u = BA_NormalizeText(s)

    ' Sayıları ve özel karakterleri temizle (*, yıldız vs)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "[0-9\*\-\.]+"
    u = regex.Replace(u, " ")
    
    ' Özel karakterleri temizle
    regex.pattern = "[^A-Z ]"
    u = regex.Replace(u, " ")

    ' Sabit ibareleri sil
    Dim removeWords As Variant
    removeWords = Array("ACIKLAMA", "LEHDAR", "FAST", "ANLIK", "ODEME", "ÖDEME", "HESAP", "TAHSILAT", _
                       "NOLU", "MUSTERI", "MÜŞTERİ", "TAKSIT", "POLICE", "POLİÇE", "LEASING", _
                       "ALL", "RISK", "KARTI", "KART")
    Dim w As Long
    For w = LBound(removeWords) To UBound(removeWords)
        u = Replace(u, UCase(removeWords(w)), " ")
    Next w

    u = BA_CollapseSpaces(u)

    ' BNK_ARINDIR listesindeki kelimeleri çıkar
    If gBnkDict Is Nothing Then Call BA_RebuildBnkStopList
    
    Dim parts() As String, out() As String, word As Variant, wCnt As Long
    parts = Split(u, " ")
    ReDim out(0 To UBound(parts))
    
    For Each word In parts
        If Len(word) > 1 And Not IsNumeric(word) Then
            If Not gBnkDict.Exists(word) Then
                out(wCnt) = word
                wCnt = wCnt + 1
            End If
        End If
    Next word
    
    If wCnt = 0 Then
        BA_CleanBankDescription = ""
    Else
        ReDim Preserve out(0 To wCnt - 1)
        BA_CleanBankDescription = Join(out, " ")
    End If
End Function

' === NORMAL ARINDIRMA (ESKİ HALİ) ===
Public Function BA_CleanDescription(ByVal s As String) As String
    Dim u As String
    u = BA_NormalizeText(s)

    ' Sabit ibareleri sil
    Dim removeWords As Variant
    removeWords = Array("ACIKLAMA=", "LEHDAR=", "FAST ANLIK ODEME", "FAST ANLIK ÖDEME", "ODEME", "ÖDEME", "HESAP", "TAHSILAT")
    Dim w As Long
    For w = LBound(removeWords) To UBound(removeWords)
        u = Replace(u, UCase(removeWords(w)), " ")
    Next w

    ' Özel karakter ve rakamları boşluk yap
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "[^A-Z ]"
    u = regex.Replace(u, " ")

    ' Birleşik kelimeleri ayır
    u = Replace(u, "ACIKLAMA", " ACIKLAMA ")
    u = Replace(u, "LEHDAR", " LEHDAR ")
    u = Replace(u, "ODEME", " ODEME ")
    u = Replace(u, "HESAP", " HESAP ")
    u = Replace(u, "TAHSILAT", " TAHSILAT ")

    u = BA_CollapseSpaces(u)

    ' Sabit kelimeleri tekrar sil
    For w = LBound(removeWords) To UBound(removeWords)
        u = Replace(u, UCase(removeWords(w)), " ")
        u = Replace(u, Replace(UCase(removeWords(w)), "=", ""), " ")
    Next w

    u = BA_CollapseSpaces(u)

    ' Normal stop token listesini kullan
    If gStopDict Is Nothing Then Call BA_RebuildStopList
    
    Dim parts() As String, out() As String, word As Variant, wCnt As Long
    parts = Split(u, " ")
    ReDim out(0 To UBound(parts))
    
    For Each word In parts
        If Len(word) > 1 And Not IsNumeric(word) Then
            If Not gStopDict.Exists(word) Then
                out(wCnt) = word
                wCnt = wCnt + 1
            End If
        End If
    Next word
    
    If wCnt = 0 Then
        BA_CleanDescription = ""
    Else
        ReDim Preserve out(0 To wCnt - 1)
        BA_CleanDescription = Join(out, " ")
    End If
End Function

' === Yardımcı: Fazla boşlukları tek yap ===
Public Function BA_CollapseSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    BA_CollapseSpaces = Trim(s)
End Function

' === Yardımcı: Büyük harf ve boşluk normalizasyonu ===
Public Function BA_NormalizeText(ByVal s As String) As String
    Dim u As String: u = UCase(s)
    u = Replace(u, "İ", "I")
    u = Replace(u, "Ş", "S")
    u = Replace(u, "Ç", "C")
    u = Replace(u, "Ğ", "G")
    u = Replace(u, "Ü", "U")
    u = Replace(u, "Ö", "O")
    Do While InStr(u, "  ") > 0: u = Replace(u, "  ", " "): Loop
    BA_NormalizeText = Trim(u)
End Function

' === OZET SAYFASI TOPLU ARINDIRMA (GELİŞTİRİLMİŞ) ===
Public Sub BA_TopluArindirOzet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_OZET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "F").End(xlUp).Row
    Dim r As Long
    
    For r = 5 To lastRow
        If Len(ws.Cells(r, "F").Value) > 0 Then
            ' N sütununda banka kontrolü yap
            If ws.Cells(r, "N").Value = "E" Then
                ' Banka ise P sütununa özel arındırma yap, D sütunu boş kalsın
                ws.Cells(r, "P").Value = BA_CleanBankDescription(ws.Cells(r, "F").Value)
                ws.Cells(r, "D").ClearContents
            Else
                ' Banka değilse D sütununa normal arındırma yap
                ws.Cells(r, "D").Value = BA_CleanDescription(ws.Cells(r, "F").Value)
                ws.Cells(r, "P").ClearContents
            End If
        Else
            ws.Cells(r, "D").ClearContents
            ws.Cells(r, "P").ClearContents
        End If
    Next r
End Sub

' === GELİŞTİRİLMİŞ CHLIST YAKIN EŞLEŞME ===
Public Sub CHLIST_YAKIN_ESLESME()
    Dim wsOzet As Worksheet, wsCHList As Worksheet
    Set wsOzet = ThisWorkbook.Worksheets(SHEET_OZET)
    Set wsCHList = ThisWorkbook.Worksheets(SHEET_CHLIST)
    Dim lastRowOzet As Long, lastRowCH As Long
    Dim r As Long, i As Long
    Dim desc As String, chDesc As String, chCode As String
    Dim arrDesc() As String, arrCH() As String
    Dim bestScore As Double, bestChoice As String, bestCode As String
    Dim tempScore As Double, totalWordsDesc As Double, totalWordsCH As Double

    lastRowOzet = wsOzet.Cells(wsOzet.Rows.count, "N").End(xlUp).Row
    lastRowCH = wsCHList.Cells(wsCHList.Rows.count, "B").End(xlUp).Row

    For r = 5 To lastRowOzet
        wsOzet.Cells(r, "O").Value = ""
        
        If wsOzet.Cells(r, "N").Value = "H" Then
            ' Normal müşteri - D sütunundaki arındırılmış metni kullan
            desc = BA_NormalizeText(wsOzet.Cells(r, "D").Value)
        ElseIf wsOzet.Cells(r, "N").Value = "E" Then
            ' Banka - P sütunundaki arındırılmış metni kullan
            desc = BA_NormalizeText(wsOzet.Cells(r, "P").Value)
        Else
            GoTo NextRow
        End If
        
        If lastRowCH < 2 Or desc = "" Then GoTo NextRow

        arrDesc = Split(desc, " ")
        totalWordsDesc = UBound(arrDesc) - LBound(arrDesc) + 1
        bestScore = 0: bestChoice = "": bestCode = ""
        
        For i = 2 To lastRowCH
            chDesc = BA_NormalizeText(wsCHList.Cells(i, "B").Value)
            chCode = Trim(wsCHList.Cells(i, "A").Value) ' A sütunundaki kod
            arrCH = Split(chDesc, " ")
            totalWordsCH = UBound(arrCH) - LBound(arrCH) + 1
            
            ' Eşleşen kelime sayısı
            tempScore = CommonWordCount(arrCH, arrDesc)
            
            ' Geliştirilmiş eşleştirme mantığı:
            ' 1. En az 1 kelime eşleşmeli
            ' 2. CHLIST'teki kelimelerin en az %60'ı eşleşmeli VEYA
            ' 3. Açıklamadaki kelimelerin en az %40'ı eşleşmeli
            If tempScore >= 1 And _
               (tempScore / totalWordsCH >= 0.6 Or tempScore / totalWordsDesc >= 0.4) Then
                If tempScore > bestScore Then
                    bestScore = tempScore
                    bestChoice = wsCHList.Cells(i, "B").Value
                    bestCode = chCode
                End If
            End If
        Next i
        
        If bestScore > 0 Then
            ' A sütunu kodu ve B sütunu açıklamayı birleştir
            If Len(bestCode) > 0 Then
                wsOzet.Cells(r, "O").Value = bestCode & " - " & bestChoice
            Else
                wsOzet.Cells(r, "O").Value = bestChoice
            End If
        End If
        
NextRow:
    Next r
End Sub

' === Yardımcı: Ortak kelime sayısı ===
Function CommonWordCount(arrA As Variant, arrB As Variant) As Long
    Dim i As Long, cnt As Long
    For i = LBound(arrA) To UBound(arrA)
        If InArray(arrA(i), arrB) Then cnt = cnt + 1
    Next i
    CommonWordCount = cnt
End Function

' === Yardımcı: Dizi içinde var mı? ===
Function InArray(ByVal val As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then InArray = True: Exit Function
    Next i
    InArray = False
End Function

' === ANA MAKRO (TEK TUŞ) ===
Public Sub BANKA_AKTAR_FULL_SIRA()
    Call GB_EKSTRE_GETIR_HEPSI
    Call BANKA_MI_DOLDUR
    Call BA_RebuildStopList
    Call BA_RebuildBnkStopList
    Call BA_TopluArindirOzet
    Call CHLIST_YAKIN_ESLESME
    MsgBox "Tüm süreç tamamlandı!", vbInformation
End Sub

