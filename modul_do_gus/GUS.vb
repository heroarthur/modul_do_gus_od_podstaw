﻿Imports System.ServiceModel
Imports System.Xml

' git in visual studio
' https://services.github.com/on-demand/windows/visual-studio

' remember to add Service reference "UslugaBIRpubl" from https://wyszukiwarkaregontest.stat.gov.pl/wsBIR/wsdl/UslugaBIRzewnPubl.xsd



Module Module1


    Const NIP_ilosc_cyfr = 10


    Enum Komunikaty_GusApi
        poprawnie_wyszukano_dzialalnosc = 0
        niepoprawny_format_nip = 1
        brak_danej_dzialalnosci = 2
        nip_empty_lub_null = 3
        poprawny_format = 4
        nieudane_logowanie_pusty_SID = 5
        throwned_exception = 6
    End Enum


    Public Class Podstawowe_dane_dzialalnosci
        Public regon,
        nazwa,
        wojewodztwo,
        powiat,
        gmina,
        miejscowosc,
        kodpocztowy,
        ulica,
        typ,
        silosId As String

        Public komunikat_diagnostyczny As Komunikaty_GusApi

        Public Sub New()
            regon = "" : nazwa = "" : wojewodztwo = "" : powiat = "" : gmina = ""
            miejscowosc = "" : kodpocztowy = "" : ulica = "" : typ = "" : silosId = ""
        End Sub
    End Class


    Public Class Pelny_raport_dzialalnosci
        Public regon,
        nip,
        nazwa,
        nazwaSkrocona,
        numerWRejestrzeEwidencji As String

        Public Sub New()
            regon = "" : nip = "" : nazwa = "" : nazwaSkrocona = "" : numerWRejestrzeEwidencji = ""
        End Sub
    End Class






    Public Class GusApi
        Private myBinding As WSHttpBinding
        Private ea As EndpointAddress
        Private cc As UslugaBIRpubl.UslugaBIRzewnPublClient
        Private requestMessage As Channels.HttpRequestMessageProperty

        Private Const gusUrl As String = "https://wyszukiwarkaregontest.stat.gov.pl/wsBIR/UslugaBIRzewnPubl.svc"
        Private Const gus_key As String = "abcde12345abcde12345"
        Private strSID As String


        Public Sub New()
            myBinding = New WSHttpBinding
            myBinding.Security.Mode = SecurityMode.Transport
            myBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None
            myBinding.MessageEncoding = WSMessageEncoding.Mtom

            ea = New EndpointAddress(gusUrl)

            cc = New UslugaBIRpubl.UslugaBIRzewnPublClient(myBinding, ea)

            requestMessage = New Channels.HttpRequestMessageProperty
            cc.Open()
        End Sub

        Protected Overrides Sub Finalize()
            cc.Wyloguj(strSID)
            cc.Close()
            MyBase.Finalize()
        End Sub

        Private Sub Uaktualnij_sesje_gus()
            strSID = cc.Zaloguj(gus_key)
        End Sub


        Private Function Pobierz_podstawowe_dane(Nip As String) As String
            Uaktualnij_sesje_gus()
            Using (New OperationContextScope(cc.InnerChannel))
                requestMessage.Headers("sid") = strSID
                OperationContext.Current.OutgoingMessageProperties(Channels.HttpRequestMessageProperty.Name) = requestMessage

                Dim objParametryGR1 As New UslugaBIRpubl.ParametryWyszukiwania
                objParametryGR1.Nip = Nip
                Return cc.DaneSzukaj(objParametryGR1)
            End Using
        End Function


        Private Function Poprawnosc_formatu_nip(Nip As String) As Komunikaty_GusApi
            If String.IsNullOrEmpty(Nip) Or IsNothing(Nip) Then
                Return Komunikaty_GusApi.nip_empty_lub_null
            ElseIf Not Char.IsNumber(Nip) Or Nip.Length <> NIP_ilosc_cyfr Then
                Return Komunikaty_GusApi.niepoprawny_format_nip
            End If
            Return Komunikaty_GusApi.poprawny_format
        End Function


        Private Function Poprawnosc_wyszukania_dzialalnosci(ByRef xmlBasicData As String) As Komunikaty_GusApi
            If String.IsNullOrEmpty(strSID) Or IsNothing(strSID) Then
                Return Komunikaty_GusApi.nieudane_logowanie_pusty_SID
            ElseIf String.IsNullOrEmpty(xmlBasicData) Or IsNothing(xmlBasicData) Then
                Return Komunikaty_GusApi.brak_danej_dzialalnosci
            End If
            Return Komunikaty_GusApi.poprawnie_wyszukano_dzialalnosc
        End Function


        Private Sub Wypelnij_podstawowe_dane(xmlBasicData As String, ByRef dane As Podstawowe_dane_dzialalnosci)
            Dim doc As New XmlDocument()
            doc.LoadXml(xmlBasicData)

            dane.regon = doc.GetElementsByTagName("Regon")(0).InnerXml
            dane.nazwa = doc.GetElementsByTagName("Nazwa")(0).InnerXml
            dane.wojewodztwo = doc.GetElementsByTagName("Wojewodztwo")(0).InnerXml
            dane.powiat = doc.GetElementsByTagName("Powiat")(0).InnerXml
            dane.gmina = doc.GetElementsByTagName("Gmina")(0).InnerXml
            dane.miejscowosc = doc.GetElementsByTagName("Miejscowosc")(0).InnerXml
            dane.kodpocztowy = doc.GetElementsByTagName("KodPocztowy")(0).InnerXml
            dane.ulica = doc.GetElementsByTagName("Ulica")(0).InnerXml
            dane.typ = doc.GetElementsByTagName("Typ")(0).InnerXml
            dane.silosId = doc.GetElementsByTagName("SilosID")(0).InnerXml
        End Sub


        Public Function Daj_podstawowe_dane_dzialalnosci(Nip As String) As Podstawowe_dane_dzialalnosci
            Dim dane As Podstawowe_dane_dzialalnosci = New Podstawowe_dane_dzialalnosci
            Try
                dane.komunikat_diagnostyczny = Poprawnosc_formatu_nip(Nip)
                If dane.komunikat_diagnostyczny = Komunikaty_GusApi.poprawny_format Then
                    Dim xmlBasicData As String = Pobierz_podstawowe_dane(Nip)
                    dane.komunikat_diagnostyczny = Poprawnosc_wyszukania_dzialalnosci(xmlBasicData)
                    If dane.komunikat_diagnostyczny = Komunikaty_GusApi.poprawnie_wyszukano_dzialalnosc Then
                        Wypelnij_podstawowe_dane(xmlBasicData, dane)
                    End If
                End If
                Return dane
            Catch ex As Exception 'shouldnt happen
                dane.komunikat_diagnostyczny = Komunikaty_GusApi.throwned_exception
                Return dane
            End Try
        End Function

    End Class





    Sub Main()

        Dim nip1 As String = "9370003171"
        Dim nip2 As String = "9542685559"
        Dim nip3 As String = "9372509153"


        'zapytania o dzialalnosci do Gus
        Dim gusApi As New GusApi

        Dim dane1 As Podstawowe_dane_dzialalnosci = gusApi.Daj_podstawowe_dane_dzialalnosci("6920000013")
        Dim dane2 As Podstawowe_dane_dzialalnosci = gusApi.Daj_podstawowe_dane_dzialalnosci("")
        Dim dane3 As Podstawowe_dane_dzialalnosci = gusApi.Daj_podstawowe_dane_dzialalnosci(Nothing)
        Dim unini As String
        Dim dane4 As Podstawowe_dane_dzialalnosci = gusApi.Daj_podstawowe_dane_dzialalnosci(unini)
        Dim dane5 As Podstawowe_dane_dzialalnosci = gusApi.Daj_podstawowe_dane_dzialalnosci("tekst")





        'NBP zapytania o walute
        Dim nbp As New NBPapi

        Dim kurs As KursWaluty = nbp.Daj_aktualny_kurs_waluty("chf")
        Dim kurs1 = nbp.Daj_aktualny_kurs_waluty("someWrongCode")

        Dim przedzial1 As New OkresCzasu(DateValue("Jun 19, 2010"), DateValue("Jun 28, 2010"))
        Dim przedzial2 As New OkresCzasu(DateValue("Jun 19, 2010"), DateValue("Jun 28, 2011"))
        'Dim przedzial3 As New OkresCzasu(DateValue("Jun 19, 2010"), DateValue("Jun 28, 2009")) asercja zla kolejnosc dat
        Dim przedzial4 As New OkresCzasu(DateValue("Jun 19, 2010"), DateValue("Jun 19, 2010"))
        Dim przedzial5 As New OkresCzasu(DateValue("Jun 19, 2010"), DateValue("Jun 20, 2010"))

        Dim kursy1 = nbp.Daj_kurs_w_okresie_czasu("chf", przedzial1)
        Dim kursy2 = nbp.Daj_kurs_w_okresie_czasu("chf", przedzial2)
        Dim kursy4 = nbp.Daj_kurs_w_okresie_czasu("chf", przedzial4)
        Dim kursy5 = nbp.Daj_kurs_w_okresie_czasu("chf", przedzial5)
        Dim kursy6 = nbp.Daj_kurs_w_okresie_czasu(Nothing, przedzial1)
        Dim kursy7 = nbp.Daj_kurs_w_okresie_czasu(Nothing, Nothing)

    End Sub  'ustaw punkt przerwania tutaj by sprawdzic ustawione dane i raporty

End Module