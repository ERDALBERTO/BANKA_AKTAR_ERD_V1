Sub BankaTransferi()
    ' Banka transferi otomasyonu
    Dim hesapNo As String
    Dim miktar As Double
    Dim aliciHesap As String
    
    hesapNo = InputBox("Lütfen kendi hesap numaranızı girin:")
    aliciHesap = InputBox("Lütfen alıcı hesap numarasını girin:")
    miktar = InputBox("Lütfen transfer etmek istediğiniz miktarı girin:")
    
    ' Burada transfer işlemi yapılacak
    MsgBox "Transfer başarıyla gerçekleştirildi!" & vbCrLf & _
           "Hesap No: " & hesapNo & vbCrLf & _
           "Alıcı Hesap: " & aliciHesap & vbCrLf & _
           "Miktar: " & miktar
End Sub
