Imports System.IO
Imports System.Net
Imports System.Xml


Module NarodowyBankPolski

    Private Const NBPformatDaty As String = "dd/MM/yyyy"

    Public Class KursWaluty
        Public nazwaWaluty,
        kodWaluty,
        przeliczonyKursŚredniWaluty,
        przeliczonyKursKupnaWaluty,
        przeliczonyKursSprzedażyWaluty As String
        Public dataPublikacji As Date

        Public Sub New()
            nazwaWaluty = "" : kodWaluty = "" : przeliczonyKursŚredniWaluty = ""
            przeliczonyKursKupnaWaluty = "" : przeliczonyKursSprzedażyWaluty = "" : dataPublikacji = Nothing
        End Sub

    End Class


    Public Class NBPapi

        Private Const NBP_api_url As String = "http://api.nbp.pl/api/exchangerates/rates"
        Private Const TabelaKursówŚrednichWalutObcych_A As String = "a"
        Private Const TabelaKursówKupnaISprzedażyWalutObcych_C As String = "c"
        Private Const xmlFormat = "?format=xml"


        Private Function adres_srednich_kursow_walut_obcych(kodWaluty As String) As String
            Return (NBP_api_url + "/" + TabelaKursówŚrednichWalutObcych_A + "/" + kodWaluty + "/" + xmlFormat)
        End Function


        Private Function adres_kupna_sprzedazy_walut_obcych(kodWaluty As String) As String
            Return (NBP_api_url + "/" + TabelaKursówKupnaISprzedażyWalutObcych_C + "/" + kodWaluty + "/" + xmlFormat)
        End Function


        Public Function pobierz_tabele_xml(urlTabeli As String) As XmlDocument
            Dim WinHttpReq As Object
            WinHttpReq = CreateObject("Microsoft.XMLHTTP")
            WinHttpReq.Open("GET", urlTabeli, False)
            WinHttpReq.send()
            Dim text As String = WinHttpReq.responseText
            Dim doc As New XmlDocument()
            doc.LoadXml(text)
            Return doc
            'MsgBox("url tabeli NBP niepoprawny (prawdopodobnie zly kod waluty): " + vbCrLf + urlTabeli)
            'Err.Raise("url tabeli NBP niepoprawny")
        End Function


        Public Function Daj_aktualny_kurs_waluty(kodWaluty As String) As KursWaluty
            Dim kurs As New KursWaluty()
            Try
                If String.IsNullOrEmpty(kodWaluty) Or IsNothing(kodWaluty) Then
                    Throw New ArgumentNullException(NameOf(kodWaluty))
                End If
                Dim srednieKursyWalutURL As String = adres_srednich_kursow_walut_obcych(kodWaluty)
                Dim kursyKupnaIsprzedazyURL As String = adres_kupna_sprzedazy_walut_obcych(kodWaluty)
                Dim tabela_A As XmlDocument = pobierz_tabele_xml(srednieKursyWalutURL)
                Dim tabela_C As XmlDocument = pobierz_tabele_xml(kursyKupnaIsprzedazyURL)

                kurs.nazwaWaluty = tabela_A.GetElementsByTagName("Currency")(0).InnerXml
                kurs.kodWaluty = tabela_A.GetElementsByTagName("Code")(0).InnerXml
                kurs.dataPublikacji = tabela_A.GetElementsByTagName("EffectiveDate")(0).InnerXml
                kurs.przeliczonyKursŚredniWaluty = tabela_A.GetElementsByTagName("Mid")(0).InnerXml
                kurs.przeliczonyKursKupnaWaluty = tabela_C.GetElementsByTagName("Bid")(0).InnerXml
                kurs.przeliczonyKursSprzedażyWaluty = tabela_C.GetElementsByTagName("Ask")(0).InnerXml

                Return kurs
            Catch ex As Exception
                'MsgBox("zly url dla kodu waluty: " & vbCrLf & kodWaluty)
                Return kurs
            End Try


        End Function


    End Class

End Module