'# VBA Kodlar

'## Müşteri Teklif Oluşturma Formu Sayfa Kaynak Kodları

'---
Option Explicit

'=============================================================================
' --- SAYFA DÜZEYİNDE DEĞİŞKENLER (EN ÜSTE )
'=============================================================================
Dim EskiSatirSayisi As Long
Dim EskiSutunSayisi As Long

'=============================================================================
' 1. ADIM: MEVCUT YAPININ TAKİBİ
'=============================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    'Sayfa sınırlarını hafızaya alıyoruz ki değişiklik olursa fark edelim
    If EskiSatirSayisi = 0 Then
        EskiSatirSayisi = Me.UsedRange.Rows.Count
        EskiSutunSayisi = Me.UsedRange.Columns.Count
    End If
End Sub

'=============================================================================
' 2. ADIM: DEĞİŞİKLİK OLAYI
'=============================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    
    '--- A) SATIR/SÜTUN EKLEME-SİLME KONTROLÜ ---
    On Error Resume Next
    Dim YeniSatirSayisi As Long
    Dim YeniSutunSayisi As Long
    Dim CevapYapi As Integer
    
    'Güncel sınırları al
    YeniSatirSayisi = Me.UsedRange.Rows.Count
    YeniSutunSayisi = Me.UsedRange.Columns.Count
    
    'İlk açılış kontrolü
    If EskiSatirSayisi = 0 Then EskiSatirSayisi = YeniSatirSayisi
    If EskiSutunSayisi = 0 Then EskiSutunSayisi = YeniSutunSayisi
    
    'Eğer satır veya sütun sayısı değişmişse (Yani ekleme/silme yapılmışsa)
    If YeniSatirSayisi <> EskiSatirSayisi Or YeniSutunSayisi <> EskiSutunSayisi Then
        
        CevapYapi = MsgBox("DİKKAT: Sayfada Satır veya Sütun Ekleme/Silme işlemi algılandı!" & vbNewLine & vbNewLine & _
                       "Bu işlem, formun yapısını ve makro kodlarındaki hücre referanslarını (Örn: C10, L34) bozabilir." & vbNewLine & vbNewLine & _
                       "Yine de işleme devam etmek istiyor musunuz?" & vbNewLine & _
                       "(EVET derseniz düzen bozulabilir, kodları güncellemeniz gerekebilir.)", _
                       vbYesNo + vbCritical, "Yapısal Değişiklik Uyarısı")
                       
        If CevapYapi = vbNo Then
            'HAYIR: İşlemi Geri Al (Undo)
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
            
            MsgBox "İşlem iptal edildi. Yapı korundu.", vbInformation
            
            'Değişkenleri eski haline döndür ve ÇIK
            YeniSatirSayisi = EskiSatirSayisi
            YeniSutunSayisi = EskiSutunSayisi
            Exit Sub
        Else
            'EVET: Değişikliği kabul et, hafızayı güncelle ve devam et
            EskiSatirSayisi = YeniSatirSayisi
            EskiSutunSayisi = YeniSutunSayisi
        End If
    End If
   
    
    ' GÖREV: Kullanıcı hücrelerde değişiklik yaptığında devreye giren bekçi.
    ' 1. Teklif No değişirse -> Revizyonu temizle.
    ' 2. Revizyon elle değiştirilirse -> Uyarı ver.
    '=============================================================================
    
    Dim Cevap As Integer
    
    '--- DURUM 1: Teklif No (L13) Değişirse ---
    If Not Intersect(Target, Range("L13")) Is Nothing Then
        'Olayları geçici durdur (Kendi kendini tetiklemesin diye)
        Application.EnableEvents = False
        Range("L16") = "" 'Revizyonu sil, çünkü bu artık yeni bir teklif
        Application.EnableEvents = True 'Olayları tekrar aç
    End If
    
    '--- DURUM 2: Revizyon (L16) Elle Değiştirilirse ---
    If Not Intersect(Target, Range("L16")) Is Nothing Then
        'Eğer hücre boşaltıldıysa uyarı verme (zaten temizlemek istiyor olabilir)
        If Range("L16").Value = "" Then Exit Sub
        
        'Kullanıcıya sor
        Cevap = MsgBox("DİKKAT: Revizyon numarasını (L16) elle değiştiriyorsunuz!" & vbNewLine & vbNewLine & _
                        "Sistem normalde bunu otomatik (R1, R2...) atar." & vbNewLine & _
                        "Yine de elle yazdığınız bu değeri kullanmak istiyor musunuz?", _
                        vbYesNo + vbQuestion, "Manuel Müdahale Uyarısı")
                        
        If Cevap = vbNo Then
            'Hayır derse değişikliği geri al (Hücreyi temizle)
            Application.EnableEvents = False
            Range("L16").ClearContents
            Application.EnableEvents = True
        End If
        'Evet derse hiçbir şey yapma, kullanıcının yazdığı kalsın.
    End If
    
End Sub


'---

'## Genel Müşteri Teklif Özeti Sayfa Kaynak Kodları

'---

Option Explicit

' --- MEVCUT SAYFA DÜZEYİNDE DEĞİŞKENLER ---
Dim EskiSutunSayisi As Integer
Dim SecilenHucreEskiIsmi As Variant
Dim YapisalIslemAktifMi As Boolean
'  Hücrenin silinmeden önceki değerini tutmak için
Dim EskiVeriDegeri As Variant
'  Satır silme kontrolü için satır sayısı hafızası
Dim EskiSatirSayisi As Long

Private Sub Worksheet_Activate()
    'Sayfa ilk açıldığında sütun ve satır sayısını hafızaya al
    If EskiSutunSayisi = 0 Then
        EskiSutunSayisi = Me.Cells(1, Me.Columns.Count).End(xlToLeft).Column
    End If
    'Sayfa açıldığında dolu satır sayısını al (E sütununa göre - Teklif No)
    EskiSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
    
    YapisalIslemAktifMi = False
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    
    '--- 1. SÜTUN SAYISINI KONTROL ET (MEVCUT YAPI) ---
    Dim YeniSutunSayisi As Integer
    YeniSutunSayisi = Me.Cells(1, Me.Columns.Count).End(xlToLeft).Column
    
    If EskiSutunSayisi = 0 Then EskiSutunSayisi = YeniSutunSayisi
    
    '--- DURUM A: YAPISAL DEĞİŞİKLİK ---
    If YeniSutunSayisi <> EskiSutunSayisi Then
        YapisalIslemAktifMi = True
        
        If YeniSutunSayisi > EskiSutunSayisi Then
            '--- SÜTUN EKLENDİ ---
            Dim CevapEkle As Integer
            CevapEkle = MsgBox("Araya sütun eklemeye çalışıyorsun." & vbNewLine & vbNewLine & _
                               "Sütun eklensin mi?", vbYesNo + vbQuestion, "Yapısal Değişiklik Uyarısı")
            
            If CevapEkle = vbNo Then
                Application.EnableEvents = False
                Application.Undo
                Application.EnableEvents = True
                MsgBox "İşlem iptal edildi, sütun eklenmedi.", vbInformation
                YeniSutunSayisi = EskiSutunSayisi
            Else
                MsgBox "Sütun başarıyla eklendi." & vbNewLine & vbNewLine & _
                       "DİKKAT: Eğer bu sütuna kodlar aracılığıyla dinamik veri çekecekseniz, " & _
                       "VBA kaynak kodunuza ilgili sütun başlığını tanımlayan eklemeyi yapmalısınız.", _
                       vbExclamation, "Geliştirici Notu"
                SecilenHucreEskiIsmi = ""
            End If
            
        ElseIf YeniSutunSayisi < EskiSutunSayisi Then
            '--- SÜTUN SİLİNDİ ---
        End If
        
        EskiSutunSayisi = YeniSutunSayisi
        YapisalIslemAktifMi = False
        Exit Sub
    End If
    
    '--- DURUM B: NORMAL HÜCRE SEÇİMİ ---
    YapisalIslemAktifMi = False
    
    ' Başlık kontrolü için hafıza
    If Target.Row = 1 And Target.Count = 1 Then
        SecilenHucreEskiIsmi = Target.Value
    Else
        SecilenHucreEskiIsmi = ""
    End If
    
    ' ---  Eski veriyi hafızaya al ---
    ' Sadece E, F, K sütunlarında ve veri satırlarındaysak
    If Target.Row > 1 And Not Intersect(Target, Range("E:E,F:F,K:K")) Is Nothing Then
        If Target.Cells.Count = 1 Then
            EskiVeriDegeri = Target.Value
        Else
            EskiVeriDegeri = "CokluSecim" 'Çoklu seçim kontrolü
        End If
    Else
        EskiVeriDegeri = ""
    End If
    
    ' --- SATIR SAYISINI HER SEÇİMDE GÜNCELLE ---
    ' Böylece silme işleminden hemen önceki durumu biliriz
    EskiSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    ' Eğer Form üzerinden işlem yapılıyorsa kodları durdur
    If FormIslemiAktif = True Then Exit Sub
    
    ' Eğer yapısal işlem varsa durdur
    If YapisalIslemAktifMi = True Then Exit Sub
    
    '--- SÜTUN İSMİ DEĞİŞTİRME KONTROLÜ (MEVCUT YAPI) ---
    If Target.Row = 1 And Target.Count = 1 Then
        If CStr(Target.Value) <> CStr(SecilenHucreEskiIsmi) Then
            Dim CevapDegistir As Integer
            Dim EskiGoster As String
            If CStr(SecilenHucreEskiIsmi) = "" Then EskiGoster = "(Boş)" Else EskiGoster = SecilenHucreEskiIsmi
            
            CevapDegistir = MsgBox("DİKKAT: Bir sütun ismini değiştiriyorsunuz!" & vbNewLine & _
                                   "Eski: " & EskiGoster & " -> Yeni: " & Target.Value & vbNewLine & _
                                   "Onaylıyor musunuz?", vbYesNo + vbCritical)
            
            If CevapDegistir = vbNo Then
                Application.EnableEvents = False
                Application.Undo
                Application.EnableEvents = True
                MsgBox "Değişiklik iptal edildi.", vbInformation
            Else
                SecilenHucreEskiIsmi = Target.Value
            End If
        End If
    End If
    
    ' ==============================================================================
    ' --- GELİŞMİŞ SİLME/TEMİZLEME KONTROLÜ (SATIR SİLME DAHİL) ---
    ' ==============================================================================
    
    ' 1. Sadece E, F veya K sütunlarındaki değişiklikleri izle
    If Target.Row > 1 And Not Intersect(Target, Range("E:E,F:F,K:K")) Is Nothing Then
        
        Dim YeniDeger As String
        Dim SilmeVarMi As Boolean
        Dim SatirSilindiMi As Boolean
        Dim YeniSatirSayisi As Long
        
        SilmeVarMi = False
        SatirSilindiMi = False
        
        ' --- A. SATIR SİLME KONTROLÜ (Ctrl - veya Sağ Tık Sil) ---
        YeniSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
        
        If YeniSatirSayisi < EskiSatirSayisi Then
            ' Satır sayısı azalmış, demek ki komple satır silinmiş
            SatirSilindiMi = True
            SilmeVarMi = True
        Else
            ' --- B. HÜCRE İÇERİĞİ SİLME KONTROLÜ (Delete Tuşu) ---
            ' Eğer satır silinmediyse ve tek hücre işlem görüyorsa içeriğe bak
            If Target.Cells.Count = 1 Then
                YeniDeger = CStr(Target.Value)
                If (EskiVeriDegeri <> "" And EskiVeriDegeri <> "CokluSecim") And YeniDeger = "" Then
                    SilmeVarMi = True
                End If
            End If
        End If
        
        ' --- UYARI MEKANİZMASI ---
        If SilmeVarMi = True Then
            
            Application.EnableEvents = False ' Olayları geçici durdur
            
            Dim CevapSil As Integer
            Dim MesajDetay As String
            
            If SatirSilindiMi Then
                MesajDetay = "BİR VEYA DAHA FAZLA SATIRI KOMPLE SİLDİNİZ!"
            Else
                MesajDetay = "Bir hücrenin içeriğini temizlediniz."
            End If
            
            CevapSil = MsgBox("DİKKAT: " & MesajDetay & vbNewLine & vbNewLine & _
                              "Eğer SİL (Evet) derseniz:" & vbNewLine & _
                              "Birim Bazlı Müşteri Teklif Özet sayfasındaki analiz sonuçları hatalı olabilir. O sayfadan da ilgili detayları MANUEL silmelisiniz." & vbNewLine & vbNewLine & _
                              "Eğer İPTAL (Hayır) derseniz:" & vbNewLine & _
                              "İşlem geri alınacak ve veri silinmeyecektir." & vbNewLine & vbNewLine & _
                              "İşlemi onaylıyor musunuz?", vbYesNo + vbExclamation, "Veri Tutarlılığı Uyarısı")
            
            If CevapSil = vbNo Then
                ' HAYIR DEDİ -> Ctrl+Z görevi görür, veriyi (veya satırı) geri getirir
                Application.Undo
                ' Geri alınca satır sayısı eski haline döner, hafızayı güncelle
                EskiSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
            Else
                ' EVET DEDİ -> Silmeye izin ver
                ' Satır sayısını güncelle ki sonraki işlemde hata vermesin
                EskiSatirSayisi = YeniSatirSayisi
            End If
            
            Application.EnableEvents = True ' Olayları tekrar aç
        End If
        
    End If
       
    ' ==============================================================================
    ' --- SADECE SÜTUN SİLME KONTROL BLOĞU ---
    ' ==============================================================================
    Dim GuncelSutunSayisi As Integer
    GuncelSutunSayisi = Me.Cells(1, Me.Columns.Count).End(xlToLeft).Column

    ' Eğer mevcut sütun sayısı, hafızadaki sayıdan küçükse sütun silinmiş demektir
    If GuncelSutunSayisi < EskiSutunSayisi Then
        Application.EnableEvents = False ' Olayları durdur (Sonsuz döngü engeli)
        
        Dim SutunCevap As Integer
        SutunCevap = MsgBox("DİKKAT: Sayfadan bir veya daha fazla SÜTUN siliyorsunuz!" & vbNewLine & vbNewLine & _
                            "Bu işlem sütundaki tüm verileri ve formülleri kalıcı olarak kaldıracaktır. Eğer Bu Sütune Önceden Dinamik Veri Ekleniyorsa Silerseniz Sütune Ait Dinamik Veriler Gelmez" & vbNewLine & _
                            "Sütun silme işlemini onaylıyor musunuz?", vbYesNo + vbCritical, "Sütun Silme Onayı")
        
        If SutunCevap = vbNo Then
            ' HAYIR -> İşlemi Geri Al (Undo)
            Application.Undo
            MsgBox "Sütun silme işlemi iptal edildi ve geri alındı.", vbInformation
        Else
            ' EVET -> Silmeye izin ver ve yeni sütun sayısını hafızaya kaydet
            EskiSutunSayisi = GuncelSutunSayisi
        End If
        
        Application.EnableEvents = True ' Olayları tekrar aç
    End If
    ' ==============================================================================

End Sub

'---

'## Birim Bazında MüşeriTeklif Özeti Sayfa Kaynak Kodları

'---

Option Explicit

' --- MEVCUT SAYFA DÜZEYİNDE DEĞİŞKENLER ---
Dim EskiSutunSayisi As Integer
Dim SecilenHucreEskiIsmi As Variant
Dim YapisalIslemAktifMi As Boolean
'  Hücrenin silinmeden önceki değerini tutmak için
Dim EskiVeriDegeri As Variant
'  Satır silme kontrolü için satır sayısı hafızası
Dim EskiSatirSayisi As Long

Private Sub Worksheet_Activate()
    'Sayfa ilk açıldığında sütun ve satır sayısını hafızaya al
    If EskiSutunSayisi = 0 Then
        EskiSutunSayisi = Me.Cells(1, Me.Columns.Count).End(xlToLeft).Column
    End If
    'Sayfa açıldığında dolu satır sayısını al (E sütununa göre - Teklif No)
    EskiSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
    
    YapisalIslemAktifMi = False
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    
    '--- 1. SÜTUN SAYISINI KONTROL ET ---
    Dim YeniSutunSayisi As Integer
    YeniSutunSayisi = Me.Cells(1, Me.Columns.Count).End(xlToLeft).Column
    
    If EskiSutunSayisi = 0 Then EskiSutunSayisi = YeniSutunSayisi
    
    '--- DURUM A: YAPISAL DEĞİŞİKLİK ---
    If YeniSutunSayisi <> EskiSutunSayisi Then
        YapisalIslemAktifMi = True
        
        If YeniSutunSayisi > EskiSutunSayisi Then
            '--- SÜTUN EKLENDİ ---
            Dim CevapEkle As Integer
            CevapEkle = MsgBox("Araya sütun eklemeye çalışıyorsun." & vbNewLine & vbNewLine & _
                               "Sütun eklensin mi?", vbYesNo + vbQuestion, "Yapısal Değişiklik Uyarısı")
            
            If CevapEkle = vbNo Then
                Application.EnableEvents = False
                Application.Undo
                Application.EnableEvents = True
                MsgBox "İşlem iptal edildi, sütun eklenmedi.", vbInformation
                YeniSutunSayisi = EskiSutunSayisi
            Else
                MsgBox "Sütun başarıyla eklendi." & vbNewLine & vbNewLine & _
                       "DİKKAT: Eğer bu sütuna kodlar aracılığıyla dinamik veri çekecekseniz, " & _
                       "VBA kaynak kodunuza ilgili sütun başlığını tanımlayan eklemeyi yapmalısınız.", _
                       vbExclamation, "Geliştirici Notu"
                SecilenHucreEskiIsmi = ""
            End If
            
        ElseIf YeniSutunSayisi < EskiSutunSayisi Then
            '--- SÜTUN SİLİNDİ ---
        End If
        
        EskiSutunSayisi = YeniSutunSayisi
        YapisalIslemAktifMi = False
        Exit Sub
    End If
    
    '--- DURUM B: NORMAL HÜCRE SEÇİMİ ---
    YapisalIslemAktifMi = False
    
    ' Başlık kontrolü için hafıza
    If Target.Row = 1 And Target.Count = 1 Then
        SecilenHucreEskiIsmi = Target.Value
    Else
        SecilenHucreEskiIsmi = ""
    End If
    
    ' --- Eski veriyi hafızaya al ---
    ' DİKKAT: Burada L sütununu (Müşteri Adı) baz aldık. E, F, L
    If Target.Row > 1 And Not Intersect(Target, Range("E:E,F:F,L:L")) Is Nothing Then
        If Target.Cells.Count = 1 Then
            EskiVeriDegeri = Target.Value
        Else
            EskiVeriDegeri = "CokluSecim" 'Çoklu seçim kontrolü
        End If
    Else
        EskiVeriDegeri = ""
    End If
    
    ' --- SATIR SAYISINI HER SEÇİMDE GÜNCELLE ---
    EskiSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    ' Eğer Form üzerinden işlem yapılıyorsa kodları durdur
    ' (Modüldeki FormIslemiAktif değişkeni burada da çalışır)
    If FormIslemiAktif = True Then Exit Sub
    
    ' Eğer yapısal işlem varsa durdur
    If YapisalIslemAktifMi = True Then Exit Sub
    
    '--- SÜTUN İSMİ DEĞİŞTİRME KONTROLÜ ---
    If Target.Row = 1 And Target.Count = 1 Then
        If CStr(Target.Value) <> CStr(SecilenHucreEskiIsmi) Then
            Dim CevapDegistir As Integer
            Dim EskiGoster As String
            If CStr(SecilenHucreEskiIsmi) = "" Then EskiGoster = "(Boş)" Else EskiGoster = SecilenHucreEskiIsmi
            
            CevapDegistir = MsgBox("DİKKAT: Bir sütun ismini değiştiriyorsunuz!" & vbNewLine & _
                                   "Eski: " & EskiGoster & " -> Yeni: " & Target.Value & vbNewLine & _
                                   "Onaylıyor musunuz?", vbYesNo + vbCritical)
            
            If CevapDegistir = vbNo Then
                Application.EnableEvents = False
                Application.Undo
                Application.EnableEvents = True
                MsgBox "Değişiklik iptal edildi.", vbInformation
            Else
                SecilenHucreEskiIsmi = Target.Value
            End If
        End If
    End If
    
    ' ==============================================================================
    ' --- DETAY SAYFASI İÇİN SİLME KONTROLÜ ---
    ' ==============================================================================
    
    ' 1. Sadece E, F veya L sütunlarındaki değişiklikleri izle (L = Müşteri Adı varsayıldı)
    If Target.Row > 1 And Not Intersect(Target, Range("E:E,F:F,L:L")) Is Nothing Then
        
        Dim YeniDeger As String
        Dim SilmeVarMi As Boolean
        Dim SatirSilindiMi As Boolean
        Dim YeniSatirSayisi As Long
        
        SilmeVarMi = False
        SatirSilindiMi = False
        
        ' --- A. SATIR SİLME KONTROLÜ ---
        YeniSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
        
        If YeniSatirSayisi < EskiSatirSayisi Then
            SatirSilindiMi = True
            SilmeVarMi = True
        Else
            ' --- B. HÜCRE İÇERİĞİ SİLME KONTROLÜ ---
            If Target.Cells.Count = 1 Then
                YeniDeger = CStr(Target.Value)
                If (EskiVeriDegeri <> "" And EskiVeriDegeri <> "CokluSecim") And YeniDeger = "" Then
                    SilmeVarMi = True
                End If
            End If
        End If
        
        ' --- UYARI MEKANİZMASI (MESAJ GÜNCELLENDİ) ---
        If SilmeVarMi = True Then
            
            Application.EnableEvents = False
            
            Dim CevapSil As Integer
            Dim MesajDetay As String
            
            If SatirSilindiMi Then
                MesajDetay = "BİRİM BAZLI LİSTEDEN BİR SATIRI KOMPLE SİLDİNİZ!"
            Else
                MesajDetay = "Bir birim detayını siliyor veya temizliyorsunuz."
            End If
            
            CevapSil = MsgBox("DİKKAT: " & MesajDetay & vbNewLine & vbNewLine & _
                              "Eğer SİL (Evet) derseniz:" & vbNewLine & _
                              "Bu işlem kalıcı olacaktır. Genel Özet sayfasındaki toplam tutarlarla uyumsuzluk oluşabilir. Onun İçin Birim Bazındaki Müşteri Teklifinin Genel Müşteri Teklif Özet Sayfasındanki Ana Özet Teklifinide Manuel Olarak Siliniz. Aksi halde verilerde uyumsuzluklar oluşur. " & vbNewLine & vbNewLine & _
                              "Eğer İPTAL (Hayır) derseniz:" & vbNewLine & _
                              "İşlem geri alınacak ve veri silinmeyecektir." & vbNewLine & vbNewLine & _
                              "Silme işlemini onaylıyor musunuz?", vbYesNo + vbExclamation, "Birim Silme Onayı")
            
            If CevapSil = vbNo Then
                ' HAYIR -> Geri Al
                Application.Undo
                EskiSatirSayisi = Me.Cells(Me.Rows.Count, "E").End(xlUp).Row
            Else
                ' EVET -> İzin Ver
                EskiSatirSayisi = YeniSatirSayisi
            End If
            
            Application.EnableEvents = True
        End If
        
    End If

    ' ==============================================================================
    ' --- SADECE SÜTUN SİLME KONTROL BLOĞU ---
    ' ==============================================================================
    Dim GuncelSutunSayisi As Integer
    GuncelSutunSayisi = Me.Cells(1, Me.Columns.Count).End(xlToLeft).Column

    ' Eğer mevcut sütun sayısı, hafızadaki sayıdan küçükse sütun silinmiş demektir
    If GuncelSutunSayisi < EskiSutunSayisi Then
        Application.EnableEvents = False ' Olayları durdur (Sonsuz döngü engeli)
        
        Dim SutunCevap As Integer
        SutunCevap = MsgBox("DİKKAT: Sayfadan bir veya daha fazla SÜTUN siliyorsunuz!" & vbNewLine & vbNewLine & _
                            "Bu işlem sütundaki tüm verileri ve formülleri kalıcı olarak kaldıracaktır. Eğer Bu Sütune Önceden Dinamik Veri Ekleniyorsa Silerseniz Sütune Ait Dinamik Veriler Gelmez" & vbNewLine & _
                            "Sütun silme işlemini onaylıyor musunuz?", vbYesNo + vbCritical, "Sütun Silme Onayı")
        
        If SutunCevap = vbNo Then
            ' HAYIR -> İşlemi Geri Al (Undo)
            Application.Undo
            MsgBox "Sütun silme işlemi iptal edildi ve geri alındı.", vbInformation
        Else
            ' EVET -> Silmeye izin ver ve yeni sütun sayısını hafızaya kaydet
            EskiSutunSayisi = GuncelSutunSayisi
        End If
        
        Application.EnableEvents = True ' Olayları tekrar aç
    End If
    ' ==============================================================================

End Sub

'---

'## Module1 - Ana Veri Çekme Kaynak Kodları

'---

Option Explicit

Sub TeklifSistemi_AnaKayit()
    '=============================================================================
    ' PROJE: GELİŞMİŞ TEKLİF YÖNETİM SİSTEMİ (V20 - GÜVENLİ MOD + TABLO DESTEĞİ)
    '=============================================================================
    
    Dim wsForm As Worksheet, wsOzet As Worksheet, wsDetay As Worksheet
    Dim sonSatirOzet As Long, sonSatirDetay As Long, i As Long
    Dim KlasorYolu As String, DosyaAdi As String, TamYol As String
    
    'Form Verileri
    Dim T_No As String, T_Musteri As String, T_Rev As String, T_Tarih As Date
    
    '*** ToplamT Değişkeni ***
    Dim ToplamT As Double, KayitliTutar As Double
    Dim YeniRevizyon As String
    Dim DonemAy As String, DonemYil As Integer
    Dim GMid As Long, BMid As Long
    Dim KayitYapilsinMi As Boolean
    
    'Sayfa Tanımları
    Set wsForm = ThisWorkbook.Sheets("Müşteri Teklif Oluşturma Formu")
    Set wsOzet = ThisWorkbook.Sheets("Genel Müşteri Teklif Özeti")
    Set wsDetay = ThisWorkbook.Sheets("Birim Bazında MüşeriTeklifÖzeti")
    
    Application.ScreenUpdating = False
    
    '--- 1. TEMEL KONTROLLER ---
    If wsForm.Range("C10").Value = "" Or wsForm.Range("L13").Value = "" Then
        MsgBox "Lütfen 'Müşteri Adı' ve 'Teklif No' alanlarını doldurunuz!", vbCritical
        Exit Sub
    End If

    '--- SON KONTROL ONAYI ---
    Dim Onay As Integer
    Onay = MsgBox("Form üzerindeki SON KONTROLLERİ YAPTINIZ MI?" & vbNewLine & _
                  "(Teklif PDF'e dönüştürülüp kaydedilecek)", vbYesNo + vbQuestion, "Kayıt Onayı")
    
    If Onay = vbNo Then
        Application.ScreenUpdating = True
        MsgBox "İşlem iptal edildi.", vbInformation
        Exit Sub
    End If
    
    'Verileri Al
    T_No = wsForm.Range("L13").Value
    T_Musteri = wsForm.Range("C10").Value
    
    '*** N37 HÜCRESİNDEN DEĞERİ AL ***
    ToplamT = wsForm.Range("N37").Value
    
    T_Rev = wsForm.Range("L16").Value
    
    If IsDate(wsForm.Range("L14").Value) Then T_Tarih = wsForm.Range("L14").Value Else T_Tarih = Date
    DonemAy = Format(T_Tarih, "mmmm")
    DonemYil = Year(T_Tarih)
    
    '-----------------------------------------------------------------------------
    ' DİNAMİK SÜTUN TESPİTİ
    '-----------------------------------------------------------------------------
    Dim c_Oz_TeklifNo As Integer, c_Oz_Musteri As Integer
    Dim c_Oz_Rev As Integer, c_Oz_Tutar As Integer
    
    'Özet sayfasındaki kritik başlıkları buluyoruz
    c_Oz_TeklifNo = SutunNoBul(wsOzet, "Teklif No")
    c_Oz_Musteri = SutunNoBul(wsOzet, "Müşteri Adı")
    c_Oz_Rev = SutunNoBul(wsOzet, "Revizyon")
    c_Oz_Tutar = SutunNoBul(wsOzet, "Toplam")
    
    If c_Oz_TeklifNo = 0 Or c_Oz_Musteri = 0 Then
        MsgBox "HATA: Özet sayfasında 'Teklif No' veya 'Müşteri Adı' başlıkları bulunamadı.", vbCritical
        Exit Sub
    End If

    '--- 2. VERİTABANI ARAMA ---
    Dim BulunanSatir As Long
    Dim SonKayitliRev As String
    
    BulunanSatir = 0
    sonSatirOzet = wsOzet.Cells(wsOzet.Rows.Count, c_Oz_TeklifNo).End(xlUp).Row
    
    For i = sonSatirOzet To 2 Step -1
        If wsOzet.Cells(i, c_Oz_TeklifNo).Value = T_No And wsOzet.Cells(i, c_Oz_Musteri).Value = T_Musteri Then
            BulunanSatir = i
            Exit For
        End If
    Next i


'    =============================================================================
    '--- 3. KARAR VE REVİZYON MEKANİZMASI İLK KAYIT YAZISI) ---
    '=============================================================================
    KayitYapilsinMi = False
    
    If BulunanSatir = 0 Then
        '--- DURUM: İLK KAYIT ---
        
        YeniRevizyon = "İlk Kayıt"
        wsForm.Range("L16").Value = "" 'Form üzerinde "İlk Kayıt" yazmasına gerek yok, temiz kalsın
        KayitYapilsinMi = True
    Else
        '--- DURUM: KAYIT VAR (REVİZYON KONTROLÜ) ---
        SonKayitliRev = wsOzet.Cells(BulunanSatir, c_Oz_Rev).Value
        
        If c_Oz_Tutar > 0 Then
            KayitliTutar = wsOzet.Cells(BulunanSatir, c_Oz_Tutar).Value
        Else
            KayitliTutar = 0
        End If
        
        Dim DegisiklikVar As Boolean
        Dim RevizyonArtir As Boolean
        Dim DevamOnayi As Integer
        
        DegisiklikVar = (Round(ToplamT, 2) <> Round(KayitliTutar, 2))
        RevizyonArtir = False
        
        '--- SENARYO 1: Formdaki Revizyon DB ile Aynı veya Boş ---
        'Not: DB'de "İlk Kayıt" yazıyorsa ve Form boşsa, bunlar aynı kabul edilmeli mi?
        'Kullanıcı formda L16'yı boş bırakmışsa ve DB'de kayıt varsa revizyon süreci başlar.
        
        If T_Rev = SonKayitliRev Or T_Rev = "" Or (SonKayitliRev = "İlk Kayıt" And T_Rev = "") Then
            
            If DegisiklikVar = False Then
                DevamOnayi = MsgBox("BU TEKLİF ZATEN LİSTEDE VAR!" & vbNewLine & vbNewLine & _
                                    "Teklif No: " & T_No & vbNewLine & _
                                    "Revizyon: " & SonKayitliRev & vbNewLine & _
                                    "Tutar: " & Format(ToplamT, "#,##0.00") & vbNewLine & vbNewLine & _
                                    "Yine de YENİ BİR REVİZYON (R+1) oluşturmak istiyor musunuz?", _
                                    vbYesNo + vbExclamation, "Mükerrer Kayıt Uyarısı")
                
                If DevamOnayi = vbNo Then
                    MsgBox "İşlem iptal edildi.", vbInformation
                    Exit Sub
                Else
                    RevizyonArtir = True
                End If
            Else
                RevizyonArtir = True
            End If
            
            '--- REVİZYON OLUŞTURMA ---
            If RevizyonArtir Then
                Dim RevizyonSecimi As Integer
                RevizyonSecimi = MsgBox("Revizyon Numarası Nasıl Belirlensin?" & vbNewLine & vbNewLine & _
                                        "EVET: Otomatik Ata (R1, R2...)" & vbNewLine & _
                                        "HAYIR: Elle Girmek İstiyorum" & vbNewLine & _
                                        "İPTAL: Vazgeç", vbYesNoCancel + vbQuestion, "Revizyon Yöntemi")
                
                If RevizyonSecimi = vbYes Then
                    'OTOMATİK REVİZYON
                    If SonKayitliRev = "" Or SonKayitliRev = "İlk Kayıt" Then
                        YeniRevizyon = "R1" 'Eğer önceki "İlk Kayıt" ise R1'den başla
                    Else
                        On Error Resume Next
                        YeniRevizyon = "R" & (CInt(Replace(SonKayitliRev, "R", "")) + 1)
                        If Err.Number <> 0 Then YeniRevizyon = "R1"
                        On Error GoTo 0
                    End If
                    
                ElseIf RevizyonSecimi = vbNo Then
                    'ELLE GİRİŞ
                    Dim GirilenRev As String
                    GirilenRev = InputBox("Lütfen Revizyon Adını Yazınız:" & vbNewLine & "(Örn: Final, v2, R5)", "Manuel Giriş")
                    If Trim(GirilenRev) = "" Then Exit Sub
                    If GirilenRev = SonKayitliRev Then
                        MsgBox "Girdiğiniz revizyon son kayıtla aynı.", vbCritical
                        Exit Sub
                    End If
                    YeniRevizyon = GirilenRev
                Else
                    Exit Sub 'İptal
                End If
                KayitYapilsinMi = True
            End If
            
        Else
            '--- SENARYO 2: Kullanıcı Formda Elle Farklı Revizyon Yazmış ---
            If T_Rev = SonKayitliRev Then
                 MsgBox "Bu revizyon (" & T_Rev & ") zaten sistemde kayıtlı.", vbCritical
                 Exit Sub
            End If
            YeniRevizyon = T_Rev
            KayitYapilsinMi = True
        End If
        
        'Revizyon numarasını (R1, R2...) forma yaz
        Application.EnableEvents = False
        wsForm.Range("L16").Value = YeniRevizyon
        Application.EnableEvents = True
    End If
    
    If KayitYapilsinMi = False Then Exit Sub

'    =============================================================================
'    --- 3.5 DETAYLI MÜKERRER ÜRÜN KONTROLÜ (GÜVENLİ) ---
'    =============================================================================
    Dim KontrolUrun As String
    Dim KontrolTutar As Double
    Dim KayitSayisi As Double

    Dim c_Det_TeklifNo As Integer, c_Det_Musteri As Integer
    Dim c_Det_Rev As Integer, c_Det_Urun As Integer, c_Det_Tutar As Integer

    c_Det_TeklifNo = SutunNoBul(wsDetay, "Teklif No")
    c_Det_Musteri = SutunNoBul(wsDetay, "Müşteri Adı")
    c_Det_Rev = SutunNoBul(wsDetay, "Revizyon")
    c_Det_Urun = SutunNoBul(wsDetay, "Ürün Adı")

    'Tutar Başlığı Kontrolü
    c_Det_Tutar = SutunNoBul(wsDetay, "Genel Satır Toplam")
    If c_Det_Tutar = 0 Then c_Det_Tutar = SutunNoBul(wsDetay, "Toplam")

    'Sadece sütunlar bulunursa kontrol et
    If c_Det_TeklifNo > 0 And c_Det_Tutar > 0 Then
        For i = 22 To 31
            If wsForm.Cells(i, "B").Value <> "" Then
                KontrolUrun = wsForm.Cells(i, "B").Value
                KontrolTutar = wsForm.Cells(i, "L").Value

                KayitSayisi = Application.WorksheetFunction.CountIfs( _
                                wsDetay.Columns(c_Det_TeklifNo), T_No, _
                                wsDetay.Columns(c_Det_Musteri), T_Musteri, _
                                wsDetay.Columns(c_Det_Rev), YeniRevizyon, _
                                wsDetay.Columns(c_Det_Urun), KontrolUrun, _
                                wsDetay.Columns(c_Det_Tutar), KontrolTutar)

                If KayitSayisi > 0 Then
                    Application.ScreenUpdating = True
                    MsgBox "KRİTİK HATA: Bu revizyon için bu ürünler zaten kaydedilmiş!" & vbNewLine & vbNewLine & _
                           "Ürün: " & KontrolUrun & vbNewLine & _
                           "Veritabanında çift kayıt oluşması engellendi.", vbCritical
                    Exit Sub
                End If
            End If
        Next i
    End If



'    '--- 4. PDF KAYDETME  istediğiniz konuma kayıt için burasını Aşağıdaki 4.pdf yazan yere kadar olan yorumları kaldırın Yorumlaru.---
    ' Yorum satırından çıkarmakak için ' simgesini silin. Eğer Aşağıdakileri yorum satırında çıkarırsanız bu kodun altındaki masaüstüne otomatik kayıt kısmındaki kodlara Yorum
    ' satırı haline getirin.  Yorum Satırını ' simgesi ile ekelyebilisiniz.
    ' İlgili konum yolunu değiştirmek için klasörün konum yolunu KlasorYolu Kısmındaki D:\TEKLİFLER\MÜŞTERİ TEKLİFLERİ\ ile değiştiriniz.

'    KlasorYolu = "C:\Users\Osman\Desktop\Teklifler\"
'    If Dir(KlasorYolu, vbDirectory) = "" Then
'        MsgBox "HATA: Ağ yolu bulunamadı!" & vbNewLine & KlasorYolu, vbCritical
'        Exit Sub
'    End If
'
'    Dim TemizMusteri As String
'    TemizMusteri = Replace(Replace(T_Musteri, ".", ""), " ", "_")
'
'    DosyaAdi = T_No
'    If YeniRevizyon <> "" Then DosyaAdi = DosyaAdi & "_" & YeniRevizyon
'    DosyaAdi = DosyaAdi & "_" & Format(T_Tarih, "dd.mm.yyyy") & "_" & TemizMusteri & ".pdf"
'
'    TamYol = KlasorYolu & DosyaAdi
'
'    On Error Resume Next
'    wsForm.ExportAsFixedFormat Type:=xlTypePDF, Filename:=TamYol, Quality:=xlQualityStandard
'    On Error GoTo 0

'   '=============================================================================
    '--- 4. PDF KAYDETME (MASAÜSTÜNE OTOMATİK KAYIT) Eğer istediğiniz konuma kayıt etmek isterseniz de yukarıdaki yorum satırında olan kodları yorum satırından çıkarıtn ---
    ' Yorum satırından çıkarmakak için ' simgesini silin. Aşağıdaki 5. Genel bölüme kadar olan kodlara yorum satır ekleyin. Yorum Satırını ' simgesi ile ekelyebilisiniz.
    '=============================================================================
        
    '--- OTOMATİK MASAÜSTÜ YOLU BULMA ---
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")

    'Her kullanıcının masaüstü yolunu dinamik olarak bulur (OneDrive dahil)
    KlasorYolu = wshShell.SpecialFolders("Desktop") & "\"
    Set wshShell = Nothing 'Hafızayı temizle
    '----------------------------------------------

    Dim TemizMusteri As String
    TemizMusteri = Replace(Replace(T_Musteri, ".", ""), " ", "_")

    DosyaAdi = T_No
    If YeniRevizyon <> "" Then DosyaAdi = DosyaAdi & "_" & YeniRevizyon
    DosyaAdi = DosyaAdi & "_" & Format(T_Tarih, "dd.mm.yyyy") & "_" & TemizMusteri & ".pdf"

    TamYol = KlasorYolu & DosyaAdi

    On Error Resume Next
    wsForm.ExportAsFixedFormat Type:=xlTypePDF, Filename:=TamYol, Quality:=xlQualityStandard

    'Hata kontrolü (Örn: Dosya açıksa üzerine yazamaz)
    If Err.Number <> 0 Then
        MsgBox "PDF kaydedilemedi! Dosya şu an açık olabilir veya yol hatalı." & vbNewLine & _
               "Hata: " & Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    '*********************************************************
    ' UYARI SİSTEMİNİ GEÇİCİ OLARAK SUSTUR
    FormIslemiAktif = True
    '*********************************************************
    
    '=============================================================================
    '--- 5. GENEL ÖZETE KAYIT (TABLO DESTEKLİ VE GÜVENLİ) ---
    '=============================================================================
    
    Dim c_GMid As Integer
    c_GMid = SutunNoBul(wsOzet, "GMid")
    
    '--- DEĞİŞİKLİK: TABLO KONTROLÜ ---
    If wsOzet.ListObjects.Count > 0 Then
        'Tablo varsa içine satır ekle
        sonSatirOzet = wsOzet.ListObjects(1).ListRows.Add.Range.Row
    Else
        'Tablo yoksa klasik yöntem
        sonSatirOzet = wsOzet.Cells(wsOzet.Rows.Count, c_GMid).End(xlUp).Row + 1
    End If
    '----------------------------------
    
    If sonSatirOzet = 2 Then GMid = 1 Else GMid = Application.WorksheetFunction.Max(wsOzet.Columns(c_GMid)) + 1
    
    Call VeriYaz(wsOzet, sonSatirOzet, "GMid", GMid)
    Call VeriYaz(wsOzet, sonSatirOzet, "Dönem", DonemAy)
    Call VeriYaz(wsOzet, sonSatirOzet, "Yıl", DonemYil)
    Call VeriYaz(wsOzet, sonSatirOzet, "Teklif Tarihi", T_Tarih)
    Call VeriYaz(wsOzet, sonSatirOzet, "Teklif No", T_No)
    Call VeriYaz(wsOzet, sonSatirOzet, "Revizyon", YeniRevizyon)
    
    Dim c_LinkOzet As Integer
    c_LinkOzet = SutunNoBul(wsOzet, "PDF Linki")
    If c_LinkOzet > 0 Then wsOzet.Hyperlinks.Add Anchor:=wsOzet.Cells(sonSatirOzet, c_LinkOzet), Address:=TamYol, TextToDisplay:="PDF"
    
    Call VeriYaz(wsOzet, sonSatirOzet, "T.Bitiş Tarihi", wsForm.Range("L15").Value)
    Call VeriYaz(wsOzet, sonSatirOzet, "Teklif Veren", wsForm.Range("C15").Value)
    Call VeriYaz(wsOzet, sonSatirOzet, "Müşteri Adı", T_Musteri)
    Call VeriYaz(wsOzet, sonSatirOzet, "Durumu", "Beklemede")
    Call VeriYaz(wsOzet, sonSatirOzet, "Önem Derecesi", "Kritik")
    Call VeriYaz(wsOzet, sonSatirOzet, "Para Birimi", wsForm.Range("N22").Value)
    'Call VeriYaz(wsOzet, sonSatirOzet, "Önem Derecesi", wsForm.Range("L2").Value)
    
    '--- KRİTİK VERİ YAZMA ALANI (GÜVENLİK KONTROLLÜ) ---
    
    '1. İSKONTO KONTROLÜ (İskonta mı, İskonto mu?)
    Dim BaslikG_Isk As String
    If SutunNoBul(wsOzet, "G.İskonta %") > 0 Then
        BaslikG_Isk = "G.İskonta %"
    Else
        BaslikG_Isk = "G.İskonto %" 'Alternatif
    End If
    Call VeriYaz(wsOzet, sonSatirOzet, BaslikG_Isk, wsForm.Range("L34").Value)

    Call VeriYaz(wsOzet, sonSatirOzet, "KDV %", wsForm.Range("L36").Value)
    Call VeriYaz(wsOzet, sonSatirOzet, "Genel Toplam", wsForm.Range("N33").Value)
    
    '2. TOPLAM İSKONTO KONTROLÜ
    Dim BaslikTopIsk As String
    If SutunNoBul(wsOzet, "Toplam İskonta") > 0 Then
        BaslikTopIsk = "Toplam İskonta"
    Else
        BaslikTopIsk = "Toplam İskonto" 'Alternatif
    End If
    Call VeriYaz(wsOzet, sonSatirOzet, BaslikTopIsk, wsForm.Range("N34").Value)

    Call VeriYaz(wsOzet, sonSatirOzet, "Ara Toplam", wsForm.Range("N35").Value)
    
    '3. KDV TOPLAM ve GENEL TOPLAM (Manuel Kontrol)
    Dim s_KDV_Top As Integer
    s_KDV_Top = SutunNoBul(wsOzet, "KDV Toplam")
    If s_KDV_Top > 0 Then
        wsOzet.Cells(sonSatirOzet, s_KDV_Top).Value = wsForm.Range("N36").Value
    Else
        MsgBox "UYARI: 'KDV Toplam' sütun başlığı bulunamadı! Veri yazılamadı.", vbExclamation
    End If
    
    Dim s_Genel_Top As Integer
    s_Genel_Top = SutunNoBul(wsOzet, "Toplam")
    If s_Genel_Top > 0 Then
        wsOzet.Cells(sonSatirOzet, s_Genel_Top).Value = ToplamT
    Else
        MsgBox "UYARI: 'Toplam' sütun başlığı bulunamadı! Veri yazılamadı.", vbExclamation
    End If
    
    Call VeriYaz(wsOzet, sonSatirOzet, "PDF Link", TamYol)
    
   
   '=============================================================================
    '--- 6. DETAYLARA KAYIT (TABLO DESTEKLİ) ---
    '=============================================================================
    
    Dim c_BMid As Integer
    c_BMid = SutunNoBul(wsDetay, "BMid")
    If c_BMid = 0 Then c_BMid = 1
    
    For i = 22 To 31
        If wsForm.Cells(i, "B").Value <> "" Then
            
            '--- DEĞİŞİKLİK: TABLO KONTROLÜ ---
            If wsDetay.ListObjects.Count > 0 Then
                sonSatirDetay = wsDetay.ListObjects(1).ListRows.Add.Range.Row
            Else
                sonSatirDetay = wsDetay.Cells(wsDetay.Rows.Count, c_BMid).End(xlUp).Row + 1
            End If
            '----------------------------------
            
            If sonSatirDetay = 2 Then BMid = 1 Else BMid = Application.WorksheetFunction.Max(wsDetay.Columns(c_BMid)) + 1
            
            Call VeriYaz(wsDetay, sonSatirDetay, "BMid", BMid)
            Call VeriYaz(wsDetay, sonSatirDetay, "Dönem", DonemAy)
            Call VeriYaz(wsDetay, sonSatirDetay, "Yıl", DonemYil)
            Call VeriYaz(wsDetay, sonSatirDetay, "Teklif Tarihi", T_Tarih)
            Call VeriYaz(wsDetay, sonSatirDetay, "Teklif No", T_No)
            Call VeriYaz(wsDetay, sonSatirDetay, "Revizyon", YeniRevizyon)
            
            Dim c_LinkDetay As Integer
            c_LinkDetay = SutunNoBul(wsDetay, "PDF linki")
            If c_LinkDetay > 0 Then wsDetay.Hyperlinks.Add Anchor:=wsDetay.Cells(sonSatirDetay, c_LinkDetay), Address:=TamYol, TextToDisplay:="PDF"
            
            Call VeriYaz(wsDetay, sonSatirDetay, "T.Bitiş Tarihi", wsForm.Range("L15").Value)
            Call VeriYaz(wsDetay, sonSatirDetay, "Teklif Veren", wsForm.Range("C15").Value)
            Call VeriYaz(wsDetay, sonSatirDetay, "Müşteri Adı", T_Musteri)
            Call VeriYaz(wsDetay, sonSatirDetay, "Durumu", "Beklemede")
            Call VeriYaz(wsDetay, sonSatirDetay, "Önem Derecesi", "Kritik")
            Call VeriYaz(wsDetay, sonSatirDetay, "Para Birimi", wsForm.Range("N22").Value)
            'Call VeriYaz(wsDetay, sonSatirDetay, "Önem Derecesi", wsForm.Range("M2").Value)
            
            'İskonta/İskonto Kontrolü (Başlık Bulma)
            Dim BaslikDet_Isk As String
            If SutunNoBul(wsDetay, "İskonta %") > 0 Then
                BaslikDet_Isk = "İskonta %"
            Else
                BaslikDet_Isk = "İskonto %"
            End If
            
            '--- Eğer L34 (Genel İskonto) %0 ise Detay sayfasına da 0 yaz ---
            'Normalde P sütunundan alır ama %0 ise direkt 0 yazarız.
            If wsForm.Range("L34").Value <= 0 Then
                 Call VeriYaz(wsDetay, sonSatirDetay, BaslikDet_Isk, 0)
            Else
                 Call VeriYaz(wsDetay, sonSatirDetay, BaslikDet_Isk, wsForm.Range("L34").Value)
            End If
            
            Call VeriYaz(wsDetay, sonSatirDetay, "Ürün Adı", wsForm.Cells(i, "B").Value)
            Call VeriYaz(wsDetay, sonSatirDetay, "Miktar", wsForm.Cells(i, "H").Value)
            Call VeriYaz(wsDetay, sonSatirDetay, "Birim Fiyat", wsForm.Cells(i, "K").Value)
            
            'Genel Satır İskonto Kontrolü (Başlık Bulma)
            Dim BaslikGen_Isk_Top As String
            If SutunNoBul(wsDetay, "Genel Satır İskonta Toplamı") > 0 Then
                BaslikGen_Isk_Top = "Genel Satır İskonta Toplamı"
            Else
                BaslikGen_Isk_Top = "Genel Satır İskonto Toplamı"
            End If
            
            '--- Eğer L34 (Genel İskonto) %0 ise TUTAR da 0 olmalı ---
            If wsForm.Range("L34").Value <= 0 Then
                Call VeriYaz(wsDetay, sonSatirDetay, BaslikGen_Isk_Top, 0)
            Else
                Call VeriYaz(wsDetay, sonSatirDetay, BaslikGen_Isk_Top, wsForm.Cells(i, "P").Value)
            End If
            
            Call VeriYaz(wsDetay, sonSatirDetay, "Genel Satır Toplam", wsForm.Cells(i, "L").Value)
            Call VeriYaz(wsDetay, sonSatirDetay, "PDF Link", TamYol)
            
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "İşlem Başarılı!" & vbNewLine & vbNewLine & _
           "Müşteri: " & T_Musteri & vbNewLine & _
           "Revizyon: " & IIf(YeniRevizyon = "", "İlk Kayıt", YeniRevizyon), vbInformation
           
    '*********************************************************
    ' UYARI SİSTEMİNİ TEKRAR DEVREYE AL
    FormIslemiAktif = False
    '*********************************************************

End Sub

Sub FormuTemizle()
    'ORİJİNAL FORM TEMİZLEME KODU
    Dim wsForm As Worksheet
    Dim Cevap As Integer
    Set wsForm = ThisWorkbook.Sheets("Müşteri Teklif Oluşturma Formu")
    
    Cevap = MsgBox("Form üzerindeki tüm veriler temizlenecek." & vbNewLine & _
                   "Emin misiniz?", vbYesNo + vbQuestion, "Formu Temizle")
    If Cevap = vbNo Then Exit Sub
    
    With wsForm
        .Range("C10") = ""
        .Range("C12") = ""
        .Range("C15") = ""
        .Range("C16") = ""
        .Range("L13") = ""
        .Range("L14").Value = Date
        .Range("L15") = ""
        .Range("L16") = ""
        .Range("B22:B31") = ""
        .Range("H22:H31") = ""
        .Range("K22:K31") = ""
        .Range("L34") = 0
        .Range("C10").Select
    End With
    
    MsgBox "Form temizlendi, yeni teklif için hazır.", vbInformation
End Sub

'-----------------------------------------------------------------------------
' YARDIMCI FONKSİYONLAR
'-----------------------------------------------------------------------------

Function SutunNoBul(SayfaAdi As Worksheet, BaslikAdi As String) As Integer
    Dim Bulunan As Range
    Set Bulunan = SayfaAdi.Rows(1).Find(What:=BaslikAdi, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not Bulunan Is Nothing Then
        SutunNoBul = Bulunan.Column
    Else
        SutunNoBul = 0
    End If
End Function

Sub VeriYaz(Sayfa As Worksheet, Satir As Long, Baslik As String, Deger As Variant)
    Dim SutunNo As Integer
    Dim HedefHucre As Range
    
    SutunNo = SutunNoBul(Sayfa, Baslik)
    
    If SutunNo > 0 Then
        Set HedefHucre = Sayfa.Cells(Satir, SutunNo)
        
        ' 1. Önce Veriyi Yaz
        HedefHucre.Value = Deger
        
        ' 2. Veri Doğrulama (Açılır Liste) Koruması - GÜÇLENDİRİLMİŞ MOD
        ' Eğer üst satırda (Satir - 1) veri doğrulama varsa kopyala.
        ' ARTIK: Başlık satırı (Satir 1) dahil kontrol ediyoruz.
        ' Çünkü tablo boşaldığında referans alabileceğimiz tek yer başlıktır.
        
        On Error Resume Next
        If Satir > 1 Then 'Satır 1 (Başlık) hariç her yer için çalışır
            
            ' Üst satırın (Bu başlık da olabilir, normal satır da) doğrulama kuralı var mı?
            If Not Sayfa.Cells(Satir - 1, SutunNo).Validation Is Nothing Then
                
                ' Sadece Validation (Doğrulama) özelliğini kopyala
                Sayfa.Cells(Satir - 1, SutunNo).Copy
                HedefHucre.PasteSpecial Paste:=xlPasteValidation
                Application.CutCopyMode = False
                
            End If
        End If
        On Error GoTo 0
    End If
End Sub

'---

'Module2 - Ana Form Aktif Etme Kaynak Kodları

'---

Public FormIslemiAktif As Boolean      
