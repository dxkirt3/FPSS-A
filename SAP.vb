
Module SAP

    Public Sub SAP()

        Dim application, SapGuiAuto, Connection, System, Session

        System = "GBP GCM Production"


        If application Is Nothing Then
            SapGuiAuto = GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
        End If
        If Connection Is Nothing Then
            On Error Resume Next
            Connection = application.OpenConnection(System, True)
            If Err.Number <> 0 Then

                MsgBox("Unable to Connect to SAP system: " & "system")
                Exit Sub
            End If
            'Set Connection = Application.Children(0)
        End If
        If Session Is Nothing Then
            Session = Connection.Children(0)
        End If
        'If WScript IsNot Nothing Then
        'WScript.ConnectObject(Session, "on")
        'WScript.ConnectObject(application, "on")
        'End If



    End Sub

End Module
