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

        Public Function Daj_aktualny_kurs_waluty(kodWaluty As String) As KursWaluty

        End Function


    End Class

End Module
