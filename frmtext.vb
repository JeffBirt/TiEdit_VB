Imports MDI

Public Class frmText
    Inherits System.Windows.Forms.Form

    Private sPath As String                 ' path excluding file name
    Private sFileName As String             ' entire file name
    Private sFolderName As String               ' folder name only
    Private sVariableName As String             ' variable name only
    Private sExtension As String            ' extension type only
    Private boolChanged As Boolean = False  ' has text box been changed

    Private sCalcname As String
    Private sProgComment As String
    Private byteType As Byte


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
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmText))
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtFile
        '
        Me.txtFile.Location = New System.Drawing.Point(8, 8)
        Me.txtFile.Multiline = True
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(400, 400)
        Me.txtFile.TabIndex = 0
        Me.txtFile.Text = ""
        '
        'frmText
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(416, 414)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtFile})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmText"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Properties"

    '' Property path, gets or sets path of this child
    '  used by parent to load and save files
    Public Property path() As String
        Get
            Return sPath
        End Get
        Set(ByVal sValue As String)
            sPath = sValue
        End Set
    End Property

    '' Property file, gets file name
    '  file name split from path in detokenize class
    Public Property fileName() As String
        Get
            Return sFileName
        End Get
        Set(ByVal sValue As String)
            sFileName = sValue
        End Set
    End Property

    '' Property folder, gets or sets folder name
    '  name split in detokenize class
    Public Property folder() As String
        Get
            Return sFolderName
        End Get
        Set(ByVal sValue As String)
            sFolderName = sValue
        End Set
    End Property

    '' Property variable, gets or sets folder name
    '  name split in detokenize class
    Public Property variable() As String
        Get
            Return sVariableName
        End Get
        Set(ByVal sValue As String)
            sVariableName = sValue
        End Set
    End Property

    '' Property file, gets extension
    '  file name split from path when path is set
    Public Property extension() As String
        Get
            Return sExtension
        End Get
        Set(ByVal sValue As String)
            sExtension = sValue
        End Set
    End Property

    '' Property calcName, gets or sets
    '  calculator name
    Public Property calcName() As String
        Get
            Return sCalcname
        End Get
        Set(ByVal sValue As String)
            sCalcname = sValue
        End Set
    End Property

    '' Property ProgComment, gets or sets
    '  calculator name
    Public Property progComment() As String
        Get
            Return sProgComment
        End Get
        Set(ByVal sValue As String)
            sProgComment = sValue
        End Set
    End Property

    '' Property changed, get or sets
    '  file changed property
    Public Property changed() As Boolean
        Get
            Return boolChanged
        End Get
        Set(ByVal boolValue As Boolean)
            boolChanged = boolValue
        End Set
    End Property

#End Region

#Region "Methods"


    Private Sub txtFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFile.TextChanged
        boolChanged = True
    End Sub


    '' Sub frmText_Load handles MyBase.Load
    '  Change Form title to match what is sent from parent and saved in sVariable name
    Private Sub frmText_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Text = sFolderName + "." + sVariableName

    End Sub

    Private Sub frmText_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If boolChanged = True Then
            CType(Me.MdiParent, Form1).subPrompt2Save()
        End If

    End Sub


#End Region



End Class
