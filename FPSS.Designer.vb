Partial Class FPSS
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btn_update = Me.Factory.CreateRibbonButton
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Label1)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.Label = "Button1"
        Me.Button1.Name = "Button1"
        '
        'Label1
        '
        Me.Label1.Label = "Version 3.8"
        Me.Label1.Name = "Label1"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btn_update)
        Me.Group3.Label = "Updates"
        Me.Group3.Name = "Group3"
        '
        'btn_update
        '
        Me.btn_update.Label = "Check For Updates"
        Me.btn_update.Name = "btn_update"
        '
        'Tab2
        '
        Me.Tab2.Groups.Add(Me.Group2)
        Me.Tab2.Label = "FPSS AddIn"
        Me.Tab2.Name = "Tab2"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Label2)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
        '
        'Label2
        '
        Me.Label2.Label = "Version 3.8"
        Me.Label2.Name = "Label2"
        '
        'FPSS
        '
        Me.Name = "FPSS"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_update As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property FPSS() As FPSS
        Get
            Return Me.GetRibbon(Of FPSS)()
        End Get
    End Property
End Class
