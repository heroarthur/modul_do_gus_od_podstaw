Imports System.IO
Imports System.Net
Imports System.Xml



Module NarodowyBankPolski


    Private Const NBPformatDaty As String = "yyyy-MM-dd"
    Private Const Kod_waluty_dlugosc As Integer = 3
    Private Const Limit_zapytania_w_dniach As Integer = 92


    Enum Komunikaty_NBP
        poprawne_wyszukanie = 0
        niepoprawny_kod_waluty = 1
        puste_wyszukanie = 2
        kod_waluty_empty_lub_null = 3
        poprawny_format = 4
        throwned_exception = 5
        przekroczono_limit_92_dni = 6
        brak_aktualizacji_w_tym_okresie = 7
    End Enum


    Public Class KursWaluty
        Public nazwaWaluty,
        kodWaluty,
        przeliczonyKursŚredniWaluty,
        przeliczonyKursKupnaWaluty,
        przeliczonyKursSprzedażyWaluty As String
        Public dataPublikacji As Date

        Public komunikat_diagnostyczny As Komunikaty_NBP

        Public Sub New()
            nazwaWaluty = "" : kodWaluty = "" : przeliczonyKursŚredniWaluty = ""
            przeliczonyKursKupnaWaluty = "" : przeliczonyKursSprzedażyWaluty = "" : dataPublikacji = Nothing
        End Sub

        Public Sub New(nazwaWaluty_ As String, kodWaluty_ As String, przeliczonyKursŚredniWaluty_ As String,
            przeliczonyKursKupnaWaluty_ As String, przeliczonyKursSprzedażyWaluty_ As String, dataPublikacji_ As Date)
            nazwaWaluty = nazwaWaluty_
            kodWaluty = kodWaluty_
            przeliczonyKursŚredniWaluty = przeliczonyKursŚredniWaluty_
            przeliczonyKursKupnaWaluty = przeliczonyKursKupnaWaluty_
            przeliczonyKursSprzedażyWaluty = przeliczonyKursSprzedażyWaluty_
            dataPublikacji = dataPublikacji_
        End Sub

    End Class


    Public Class PrzedzialKursuWaluty
        Public kursyWaluty As ArrayList
        Public komunikat_diagnostyczny As Komunikaty_NBP
        Public okresCzasu As OkresCzasu

        Public Sub New(okresCzasu_ As OkresCzasu)
            kursyWaluty = New ArrayList : okresCzasu = okresCzasu_
        End Sub

        Public Sub Uzupelnij_liste_kursow(nazwaWaluty As String, kodWaluty As String,
                 datyAktualizacjiKursu As XmlNodeList, kursySrednie As XmlNodeList,
                 kursyKupna As XmlNodeList, kursySprzedazy As XmlNodeList)
            Dim iloscWpisow As Integer = datyAktualizacjiKursu.Count
            For i = 0 To iloscWpisow - 1
                kursyWaluty.Add(New KursWaluty(nazwaWaluty,
                                                kodWaluty,
                                                kursyKupna(i).InnerText,
                                                kursySprzedazy(i).InnerText,
                                                kursySprzedazy(i).InnerText,
                                                datyAktualizacjiKursu(i).InnerText))
            Next i
        End Sub
    End Class


    Public Class OkresCzasu
        Public dataPoczatkowa As Date
        Public dataKoncowa As Date
        Private Const earlierOrEqual = 0

        Public Sub New(dataPoczatkowa_ As Date, dataKoncowa_ As Date)
            Debug.Assert(Not IsNothing(dataPoczatkowa_) And Not IsNothing(dataKoncowa_))
            Dim result As Integer = DateTime.Compare(dataPoczatkowa_, dataKoncowa_)
            Debug.Assert(result <= earlierOrEqual)
            dataPoczatkowa = dataPoczatkowa_ : dataKoncowa = dataKoncowa_
        End Sub
    End Class


    Public Class NBPapi

        Private Const NBP_api_rates_url As String = "http://api.nbp.pl/api/exchangerates/rates"
        Private Const TabelaKursówŚrednichWalutObcych_A As String = "a"
        Private Const TabelaKursówKupnaISprzedażyWalutObcych_C As String = "c"
        Private Const xmlFormat = "?format=xml"


        Private Function Adres_srednich_kursow_walut_obcych(kodWaluty As String) As String
            Return (NBP_api_rates_url + "/" + TabelaKursówŚrednichWalutObcych_A + "/" + kodWaluty + "/" + xmlFormat)
        End Function


        Private Function Adres_kupna_sprzedazy_walut_obcych(kodWaluty As String) As String
            Return (NBP_api_rates_url + "/" + TabelaKursówKupnaISprzedażyWalutObcych_C + "/" + kodWaluty + "/" + xmlFormat)
        End Function


        Private Function Adres_tabeli_kupna_sprzedazy_waluty_przedzial(kodWaluty As String, okresCzasu As OkresCzasu) As String
            Return (NBP_api_rates_url + "/" + TabelaKursówKupnaISprzedażyWalutObcych_C + "/" + kodWaluty +
                "/" + Format(okresCzasu.dataPoczatkowa, NBPformatDaty) +
                "/" + Format(okresCzasu.dataKoncowa, NBPformatDaty) + "/" + xmlFormat)
        End Function


        Private Function Adres_tabeli_srednich_kursow_waluty_przedzial(kodWaluty As String, okresCzasu As OkresCzasu) As String
            Return (NBP_api_rates_url + "/" + TabelaKursówŚrednichWalutObcych_A + "/" + kodWaluty +
                "/" + Format(okresCzasu.dataPoczatkowa, NBPformatDaty) +
                "/" + Format(okresCzasu.dataKoncowa, NBPformatDaty) + "/" + xmlFormat)
        End Function


        Public Function Pobierz_tabele_xml(urlTabeli As String) As XmlDocument
            Dim WinHttpReq As Object
            WinHttpReq = CreateObject("Microsoft.XMLHTTP")
            WinHttpReq.Open("GET", urlTabeli, False)
            WinHttpReq.send()
            Dim text As String = WinHttpReq.responseText
            Dim doc As New XmlDocument()
            doc.LoadXml(text)
            Return doc
        End Function


        Public Function Pobierz_tabele_string(urlTabeli As String) As String
            Dim WinHttpReq As Object
            WinHttpReq = CreateObject("Microsoft.XMLHTTP")
            WinHttpReq.Open("GET", urlTabeli, False)
            WinHttpReq.send()
            Dim text As String = WinHttpReq.responseText
            Return text
        End Function


        Private Function Poprawnosc_formatu_kodu_waluty(kodWaluty As String) As Komunikaty_NBP
            If String.IsNullOrEmpty(kodWaluty) Or IsNothing(kodWaluty) Then
                Return Komunikaty_NBP.kod_waluty_empty_lub_null
            ElseIf Char.IsNumber(kodWaluty) Or kodWaluty.Length <> Kod_waluty_dlugosc Then
                Return Komunikaty_NBP.niepoprawny_kod_waluty
            End If
            Return Komunikaty_NBP.poprawny_format
        End Function


        Private Function Poprawnosc_zapytania_o_przedzial(kodWaluty As String, przedzial As OkresCzasu) As Komunikaty_NBP
            Dim roznicaWDniach As Integer = DateDiff("d", przedzial.dataPoczatkowa, przedzial.dataKoncowa)
            If roznicaWDniach > Limit_zapytania_w_dniach Then
                Return Komunikaty_NBP.przekroczono_limit_92_dni
            Else
                Return Poprawnosc_formatu_kodu_waluty(kodWaluty)
            End If
        End Function


        Private Function Poprawnosc_wyszukania_xml(ByRef xmlDownloandedData As String) As Komunikaty_NBP
            If String.IsNullOrEmpty(xmlDownloandedData) Or IsNothing(xmlDownloandedData) Then
                Return Komunikaty_NBP.puste_wyszukanie
            End If
            Return Komunikaty_NBP.poprawne_wyszukanie
        End Function


        Private Sub Wypelnij_kurs_waluty(ByRef tabela_A_text As String,
                                         ByRef tabela_C_text As String,
                                         ByRef kurs As KursWaluty)

            Dim data_publikacji_tabeli_A As Date
            Dim data_publikacji_tabeli_C As Date
            Dim tabela_A As New XmlDocument() : tabela_A.LoadXml(tabela_A_text)
            Dim tabela_C As New XmlDocument() : tabela_C.LoadXml(tabela_C_text)

            data_publikacji_tabeli_A = tabela_A.GetElementsByTagName("EffectiveDate")(0).InnerXml
            data_publikacji_tabeli_C = tabela_C.GetElementsByTagName("EffectiveDate")(0).InnerXml
            Debug.Assert(data_publikacji_tabeli_A = data_publikacji_tabeli_C) 'NBP should set this dates equal

            kurs.nazwaWaluty = tabela_A.GetElementsByTagName("Currency")(0).InnerXml
            kurs.kodWaluty = tabela_A.GetElementsByTagName("Code")(0).InnerXml
            kurs.dataPublikacji = tabela_A.GetElementsByTagName("EffectiveDate")(0).InnerXml
            kurs.przeliczonyKursŚredniWaluty = tabela_A.GetElementsByTagName("Mid")(0).InnerXml
            kurs.przeliczonyKursKupnaWaluty = tabela_C.GetElementsByTagName("Bid")(0).InnerXml
            kurs.przeliczonyKursSprzedażyWaluty = tabela_C.GetElementsByTagName("Ask")(0).InnerXml
        End Sub


        Public Function Daj_aktualny_kurs_waluty(kodWaluty As String) As KursWaluty
            Dim kurs As New KursWaluty()
            Try
                kurs.komunikat_diagnostyczny = Poprawnosc_formatu_kodu_waluty(kodWaluty)
                If kurs.komunikat_diagnostyczny = Komunikaty_NBP.poprawny_format Then
                    Dim srednieKursyWalutURL As String = Adres_srednich_kursow_walut_obcych(kodWaluty)
                    Dim kursyKupnaIsprzedazyURL As String = Adres_kupna_sprzedazy_walut_obcych(kodWaluty)
                    Dim tabela_A_text As String = Pobierz_tabele_string(srednieKursyWalutURL)
                    Dim tabela_C_text As String = Pobierz_tabele_string(kursyKupnaIsprzedazyURL)
                    If Poprawnosc_wyszukania_xml(tabela_A_text) = Komunikaty_NBP.poprawne_wyszukanie And
                            Poprawnosc_wyszukania_xml(tabela_C_text) = Komunikaty_NBP.poprawne_wyszukanie Then
                        Wypelnij_kurs_waluty(tabela_A_text, tabela_C_text, kurs)
                        kurs.komunikat_diagnostyczny = Komunikaty_NBP.poprawne_wyszukanie
                    Else
                        kurs.komunikat_diagnostyczny = Komunikaty_NBP.puste_wyszukanie
                    End If
                End If

                Return kurs
            Catch ex As Exception
                kurs.komunikat_diagnostyczny = Komunikaty_NBP.brak_aktualizacji_w_tym_okresie
                Return kurs
            End Try
        End Function


        Public Sub Wypelnij_przedzial_kursu_waluty(tabela_A_text As String,
                                                    tabela_C_text As String,
                                                    kursyWaluty As PrzedzialKursuWaluty)
            Dim tabela_A As New XmlDocument() : tabela_A.LoadXml(tabela_A_text)
            Dim tabela_C As New XmlDocument() : tabela_C.LoadXml(tabela_C_text)

            Dim nazwaWaluty As String = tabela_A.GetElementsByTagName("Currency")(0).InnerText
            Dim kodWaluty As String = tabela_A.GetElementsByTagName("Code")(0).InnerText
            Dim datyAktualizacjiKursu = tabela_A.GetElementsByTagName("EffectiveDate")
            Dim kursySrednie = tabela_A.GetElementsByTagName("Mid")
            Dim kursyKupna = tabela_C.GetElementsByTagName("Bid")
            Dim kursySprzedazy = tabela_C.GetElementsByTagName("Ask")

            kursyWaluty.Uzupelnij_liste_kursow(nazwaWaluty, kodWaluty, datyAktualizacjiKursu,
                                               kursySrednie, kursyKupna, kursySprzedazy)
        End Sub


        Public Function Daj_kurs_w_okresie_czasu(kodWaluty As String, okresCzasu As OkresCzasu) As PrzedzialKursuWaluty
            Dim kursyWaluty As New PrzedzialKursuWaluty(okresCzasu)
            Try
                kursyWaluty.komunikat_diagnostyczny = Poprawnosc_zapytania_o_przedzial(kodWaluty, okresCzasu)
                If kursyWaluty.komunikat_diagnostyczny = Komunikaty_NBP.poprawny_format Then
                    Dim tabelaSrednichKursowUrl As String = Adres_tabeli_srednich_kursow_waluty_przedzial(kodWaluty, okresCzasu)
                    Dim tabelaKupnaSprzedazyUrl As String = Adres_tabeli_kupna_sprzedazy_waluty_przedzial(kodWaluty, okresCzasu)
                    Dim tabela_A_text As String = Pobierz_tabele_string(tabelaSrednichKursowUrl)
                    Dim tabela_C_text As String = Pobierz_tabele_string(tabelaKupnaSprzedazyUrl)
                    If Poprawnosc_wyszukania_xml(tabela_A_text) = Komunikaty_NBP.poprawne_wyszukanie And
                            Poprawnosc_wyszukania_xml(tabela_C_text) = Komunikaty_NBP.poprawne_wyszukanie Then
                        Wypelnij_przedzial_kursu_waluty(tabela_A_text, tabela_C_text, kursyWaluty)
                        kursyWaluty.komunikat_diagnostyczny = Komunikaty_NBP.poprawne_wyszukanie
                    Else
                        kursyWaluty.komunikat_diagnostyczny = Komunikaty_NBP.puste_wyszukanie
                    End If
                End If
                Return kursyWaluty
            Catch ex As Exception
                kursyWaluty.komunikat_diagnostyczny = Komunikaty_NBP.brak_aktualizacji_w_tym_okresie
                Return kursyWaluty
            Finally

            End Try

        End Function



    End Class

End Module