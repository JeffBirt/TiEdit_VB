Public Class frmAbout
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtAbout As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
        Me.txtAbout = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtAbout
        '
        Me.txtAbout.Location = New System.Drawing.Point(16, 16)
        Me.txtAbout.Multiline = True
        Me.txtAbout.Name = "txtAbout"
        Me.txtAbout.ReadOnly = True
        Me.txtAbout.Size = New System.Drawing.Size(256, 88)
        Me.txtAbout.TabIndex = 0
        Me.txtAbout.Text = ""
        '
        'frmAbout
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(288, 125)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtAbout})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.Text = "TI-Edit 1.0"
        Me.ResumeLayout(False)

    End Sub

#End Region

    '' **-SUB frmAbout_Load
    '  handles frmAbout.Load, displays text in text bok
    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim sMessage As String = "TI Basic Editor 1.0" + vbCrLf + "This program is an editing utility for " + _
                        "TI basic programs (TI-89, TI-92, V200).  It is released as Freeware and is a work in progresss." + _
                        vbCrLf + "Contact: Jeff Birt--birt_j@earthlink.net"
        txtAbout.Text = sMessage
        txtAbout.SelectionStart = 0
    End Sub


End Class
