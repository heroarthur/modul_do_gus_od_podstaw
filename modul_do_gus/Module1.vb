Imports System.ServiceModel
Imports System.Xml

' git in visual studio
' https://services.github.com/on-demand/windows/visual-studio

' remember to add Service reference "UslugaBIRpubl" from https://wyszukiwarkaregontest.stat.gov.pl/wsBIR/wsdl/UslugaBIRzewnPubl.xsd



Module Module1


    Public Class Podstawowe_dane
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

        Public Sub New()
            regon = "" : nazwa = "" : wojewodztwo = "" : powiat = "" : gmina = ""
            miejscowosc = "" : kodpocztowy = "" : ulica = "" : typ = "" : silosId = ""
        End Sub
    End Class



    Public Class Pelny_raport
        Public regon,
        nip,
        nazwa,
        nazwaskrocona,
        numerWRejestrzeEwidencji As String

        Public Sub New()
            regon = "" : nip = "" : nazwa = "" : nazwaskrocona = "" : numerWRejestrzeEwidencji = ""
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

        Private Sub Update_gus_session()
            Static session_valid As String = "1"
            Static session_state As String = "StatusSesji"
            Dim session_value = cc.GetValue(session_state)
            If session_value <> session_valid Then
                strSID = cc.Zaloguj(gus_key)
            End If
        End Sub


        Private Function Downloand_basic_data(Nip As String) As String
            Update_gus_session()
            Using (New OperationContextScope(cc.InnerChannel))
                requestMessage.Headers("sid") = strSID
                OperationContext.Current.OutgoingMessageProperties(Channels.HttpRequestMessageProperty.Name) = requestMessage

                Dim objParametryGR1 As New UslugaBIRpubl.ParametryWyszukiwania
                objParametryGR1.Nip = Nip
                Return cc.DaneSzukaj(objParametryGR1)
            End Using
        End Function


        Private Function Downloand_full_raport(Regon As String) As String
            Update_gus_session()
            Using (New OperationContextScope(cc.InnerChannel))
                requestMessage.Headers("sid") = strSID
                OperationContext.Current.OutgoingMessageProperties(Channels.HttpRequestMessageProperty.Name) = requestMessage

                Return cc.DanePobierzPelnyRaport(Regon, "PublDaneRaportPrawna")
            End Using
        End Function


        Public Function Get_basic_data(Nip As String) As Podstawowe_dane
            Dim dane As Podstawowe_dane = New Podstawowe_dane
            Try
                If String.IsNullOrEmpty(Nip) Or IsNothing(Nip) Then
                    Throw New ArgumentNullException(NameOf(Nip))
                End If
                Dim xmlBasicData As String = Downloand_basic_data(Nip)
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
                Return dane
            Catch ex As Exception
                Return dane
            End Try
        End Function


        Public Function Get_full_raport(Regon As String) As Pelny_raport
            Dim raport As Pelny_raport = New Pelny_raport
            Try
                If String.IsNullOrEmpty(Regon) Or IsNothing(Regon) Then
                    Throw New ArgumentNullException(NameOf(Regon))
                End If
                Dim xmlBasicData As String = Downloand_full_raport(Regon)
                Dim doc As New XmlDocument()
                doc.LoadXml(xmlBasicData)

                raport.regon = doc.GetElementsByTagName("praw_regon14")(0).InnerXml
                raport.nip = doc.GetElementsByTagName("praw_nip")(0).InnerXml
                raport.nazwa = doc.GetElementsByTagName("praw_nazwa")(0).InnerXml
                raport.nazwaskrocona = doc.GetElementsByTagName("praw_nazwaSkrocona")(0).InnerXml
                raport.numerWRejestrzeEwidencji = doc.GetElementsByTagName("praw_numerWrejestrzeEwidencji")(0).InnerXml
                Return raport
            Catch ex As Exception
                Return raport
            End Try
        End Function


    End Class





    Sub Main()

        Dim gusApi As New GusApi


        Dim dane1 As Podstawowe_dane = gusApi.Get_basic_data("6920000013")
        Dim dane2 As Podstawowe_dane = gusApi.Get_basic_data("")
        Dim dane3 As Podstawowe_dane = gusApi.Get_basic_data(Nothing)
        Dim unini As String
        Dim dane4 As Podstawowe_dane = gusApi.Get_basic_data(unini)
        Dim dane5 As Podstawowe_dane = gusApi.Get_basic_data("spitfire")


        Dim raport1 As Pelny_raport = gusApi.Get_full_raport("39002176400000")
        Dim raport2 As Pelny_raport = gusApi.Get_full_raport("")
        Dim raport3 As Pelny_raport = gusApi.Get_full_raport("spitfire")
        Dim raport4 As Pelny_raport = gusApi.Get_full_raport(Nothing)
        Dim raport5 As Pelny_raport = gusApi.Get_full_raport("32222222222222")

    End Sub

End Module