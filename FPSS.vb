Imports Microsoft.Office.Tools.Ribbon
Imports System.Deployment.Application


Public Class FPSS

    Private Sub FPSS_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        Call SAP.SAP()


    End Sub

    Private Sub btn_update_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_update.Click



    End Sub
End Class
