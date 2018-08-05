Imports System.IO
Imports System.Net
Imports System.Xml



Module NarodowyBankPolski


    Private Const NBPformatDaty As String = "dd/MM/yyyy"
    Const Kod_waluty_dlugosc As Integer = 3


    Enum Komunikaty_NBP
        poprawne_wyszukanie = 0
        niepoprawny_kod_waluty = 1
        puste_wyszukanie = 2
        kod_waluty_empty_lub_null = 3
        poprawny_format = 4
        throwned_exception = 5
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

    End Class


    Public Class OkresCzasu
        Public dataPoczatkowa As Date
        Public dataKoncowa As Date

        Public Sub New(dataPoczatkowa_ As Date, dataKoncowa_ As Date)
            dataPoczatkowa = dataPoczatkowa_ : dataKoncowa = dataKoncowa_
        End Sub
    End Class


    Public Class NBPapi

        Private Const NBP_api_url As String = "http://api.nbp.pl/api/exchangerates/rates"
        Private Const TabelaKursówŚrednichWalutObcych_A As String = "a"
        Private Const TabelaKursówKupnaISprzedażyWalutObcych_C As String = "c"
        Private Const xmlFormat = "?format=xml"


        Private Function Adres_srednich_kursow_walut_obcych(kodWaluty As String) As String
            Return (NBP_api_url + "/" + TabelaKursówŚrednichWalutObcych_A + "/" + kodWaluty + "/" + xmlFormat)
        End Function


        Private Function Adres_kupna_sprzedazy_walut_obcych(kodWaluty As String) As String
            Return (NBP_api_url + "/" + TabelaKursówKupnaISprzedażyWalutObcych_C + "/" + kodWaluty + "/" + xmlFormat)
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

        Private Function Poprawnosc_formatu_kodu_waltuty(kodWaluty As String) As Komunikaty_NBP
            If String.IsNullOrEmpty(kodWaluty) Or IsNothing(kodWaluty) Then
                Return Komunikaty_NBP.kod_waluty_empty_lub_null
            ElseIf Char.IsNumber(kodWaluty) Or kodWaluty.Length <> Kod_waluty_dlugosc Then
                Return Komunikaty_NBP.niepoprawny_kod_waluty
            End If
            Return Komunikaty_NBP.poprawny_format
        End Function

        Private Function Poprawnosc_wyszukania(ByRef xmlDownloandedData As String) As Komunikaty_NBP
            If String.IsNullOrEmpty(xmlDownloandedData) Or IsNothing(xmlDownloandedData) Then
                Return Komunikaty_NBP.puste_wyszukanie
            End If
            Return Komunikaty_NBP.poprawne_wyszukanie
        End Function


        Private Sub Wypelnij_kurs_waluty(ByRef srednieKursyWalutURL As String,
                                         ByRef kursyKupnaIsprzedazyURL As String,
                                         ByRef kurs As KursWaluty)
            Dim tabela_A As XmlDocument = Pobierz_tabele_xml(srednieKursyWalutURL)
            Dim tabela_C As XmlDocument = Pobierz_tabele_xml(kursyKupnaIsprzedazyURL)
            Dim data_publikacji_tabeli_A As Date
            Dim data_publikacji_tabeli_C As Date

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
                kurs.komunikat_diagnostyczny = Poprawnosc_formatu_kodu_waltuty(kodWaluty)
                If kurs.komunikat_diagnostyczny = Komunikaty_NBP.poprawny_format Then
                    Dim srednieKursyWalutURL As String = Adres_srednich_kursow_walut_obcych(kodWaluty)
                    Dim kursyKupnaIsprzedazyURL As String = Adres_kupna_sprzedazy_walut_obcych(kodWaluty)
                    If Poprawnosc_wyszukania(srednieKursyWalutURL) = Komunikaty_NBP.poprawne_wyszukanie And
                            Poprawnosc_wyszukania(kursyKupnaIsprzedazyURL) = Komunikaty_NBP.poprawne_wyszukanie Then
                        Wypelnij_kurs_waluty(srednieKursyWalutURL, kursyKupnaIsprzedazyURL, kurs)
                        kurs.komunikat_diagnostyczny = Komunikaty_NBP.poprawne_wyszukanie
                    Else
                        kurs.komunikat_diagnostyczny = Komunikaty_NBP.puste_wyszukanie
                    End If
                End If

                Return kurs
            Catch ex As Exception
                kurs.komunikat_diagnostyczny = Komunikaty_NBP.throwned_exception
                Return kurs
            End Try
        End Function

        Public Function Daj()



    End Class

End Module