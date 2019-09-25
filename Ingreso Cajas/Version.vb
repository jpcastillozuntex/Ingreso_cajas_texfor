
Imports System.Deployment.Application
Public Class Version

    Private Sub Version_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        obtiene_version()
    End Sub

    Private Sub obtiene_version()
        versiones.Text = Application.ProductVersion
        empresa.Text = "Texsun S.A."
        copyw.Text = "2019"
    End Sub

End Class