Option Strict On
Imports System.Text
Imports System.IO
Imports System.Drawing.Printing

Public Class Form1

    Inherits System.Windows.Forms.Form

    Private sSourceFile As String
    Private sDestFile As String
    Private tpTextBox As New TextPrint("")
    Private tpFont As New Font("Ti92Pluspc", 12)

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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuECut As System.Windows.Forms.MenuItem
    Friend WithEvents mnuECopy As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWindow As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWCascade As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWHorizontal As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWVertical As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWArrange As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFNew As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFLoad As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFSave As System.Windows.Forms.MenuItem
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents dlgSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFClose As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFSaveAs As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFPrint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFPSetup As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEUndo As System.Windows.Forms.MenuItem
    Friend WithEvents mnuERedo As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEDelete As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuESelectAll As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEFind As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEClear As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTGroup As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTUngroup As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFPpreview As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFRecent As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuFNew = New System.Windows.Forms.MenuItem()
        Me.mnuFLoad = New System.Windows.Forms.MenuItem()
        Me.mnuFClose = New System.Windows.Forms.MenuItem()
        Me.mnuFSave = New System.Windows.Forms.MenuItem()
        Me.mnuFSaveAs = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuFPrint = New System.Windows.Forms.MenuItem()
        Me.mnuFPpreview = New System.Windows.Forms.MenuItem()
        Me.mnuFPSetup = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.mnuFRecent = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.mnuFExit = New System.Windows.Forms.MenuItem()
        Me.mnuEdit = New System.Windows.Forms.MenuItem()
        Me.mnuEUndo = New System.Windows.Forms.MenuItem()
        Me.mnuERedo = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.mnuECut = New System.Windows.Forms.MenuItem()
        Me.mnuECopy = New System.Windows.Forms.MenuItem()
        Me.mnuEPaste = New System.Windows.Forms.MenuItem()
        Me.mnuEDelete = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.mnuESelectAll = New System.Windows.Forms.MenuItem()
        Me.mnuEFind = New System.Windows.Forms.MenuItem()
        Me.mnuEClear = New System.Windows.Forms.MenuItem()
        Me.mnuTools = New System.Windows.Forms.MenuItem()
        Me.mnuTGroup = New System.Windows.Forms.MenuItem()
        Me.mnuTUngroup = New System.Windows.Forms.MenuItem()
        Me.mnuWindow = New System.Windows.Forms.MenuItem()
        Me.mnuWCascade = New System.Windows.Forms.MenuItem()
        Me.mnuWHorizontal = New System.Windows.Forms.MenuItem()
        Me.mnuWVertical = New System.Windows.Forms.MenuItem()
        Me.mnuWArrange = New System.Windows.Forms.MenuItem()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuHAbout = New System.Windows.Forms.MenuItem()
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgSave = New System.Windows.Forms.SaveFileDialog()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.mnuTools, Me.mnuWindow, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFNew, Me.mnuFLoad, Me.mnuFClose, Me.mnuFSave, Me.mnuFSaveAs, Me.MenuItem2, Me.mnuFPrint, Me.mnuFPpreview, Me.mnuFPSetup, Me.MenuItem6, Me.mnuFRecent, Me.MenuItem3, Me.mnuFExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuFNew
        '
        Me.mnuFNew.Index = 0
        Me.mnuFNew.Text = "&New"
        '
        'mnuFLoad
        '
        Me.mnuFLoad.Index = 1
        Me.mnuFLoad.Text = "&Open"
        '
        'mnuFClose
        '
        Me.mnuFClose.Index = 2
        Me.mnuFClose.Text = "&Close"
        '
        'mnuFSave
        '
        Me.mnuFSave.Index = 3
        Me.mnuFSave.Text = "&Save"
        '
        'mnuFSaveAs
        '
        Me.mnuFSaveAs.Index = 4
        Me.mnuFSaveAs.Text = "Save &As"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 5
        Me.MenuItem2.Text = "-"
        '
        'mnuFPrint
        '
        Me.mnuFPrint.Index = 6
        Me.mnuFPrint.Text = "&Print"
        '
        'mnuFPpreview
        '
        Me.mnuFPpreview.Index = 7
        Me.mnuFPpreview.Text = "Pr&int Preview"
        '
        'mnuFPSetup
        '
        Me.mnuFPSetup.Index = 8
        Me.mnuFPSetup.Text = "Page Set&up"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 9
        Me.MenuItem6.Text = "-"
        '
        'mnuFRecent
        '
        Me.mnuFRecent.Index = 10
        Me.mnuFRecent.Text = "Recent Files"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 11
        Me.MenuItem3.Text = "-"
        '
        'mnuFExit
        '
        Me.mnuFExit.Index = 12
        Me.mnuFExit.Text = "E&xit"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 1
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEUndo, Me.mnuERedo, Me.MenuItem7, Me.mnuECut, Me.mnuECopy, Me.mnuEPaste, Me.mnuEDelete, Me.MenuItem9, Me.mnuESelectAll, Me.mnuEFind, Me.mnuEClear})
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEUndo
        '
        Me.mnuEUndo.Index = 0
        Me.mnuEUndo.Text = "&Undo"
        '
        'mnuERedo
        '
        Me.mnuERedo.Index = 1
        Me.mnuERedo.Text = "&Redo"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.Text = "-"
        '
        'mnuECut
        '
        Me.mnuECut.Index = 3
        Me.mnuECut.Text = "Cu&t"
        '
        'mnuECopy
        '
        Me.mnuECopy.Index = 4
        Me.mnuECopy.Text = "&Copy"
        '
        'mnuEPaste
        '
        Me.mnuEPaste.Index = 5
        Me.mnuEPaste.Text = "&Paste"
        '
        'mnuEDelete
        '
        Me.mnuEDelete.Index = 6
        Me.mnuEDelete.Text = "&Delete"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 7
        Me.MenuItem9.Text = "-"
        '
        'mnuESelectAll
        '
        Me.mnuESelectAll.Index = 8
        Me.mnuESelectAll.Text = "Select &All"
        '
        'mnuEFind
        '
        Me.mnuEFind.Index = 9
        Me.mnuEFind.Text = "&Find"
        '
        'mnuEClear
        '
        Me.mnuEClear.Index = 10
        Me.mnuEClear.Text = "&Clear"
        '
        'mnuTools
        '
        Me.mnuTools.Index = 2
        Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuTGroup, Me.mnuTUngroup})
        Me.mnuTools.Text = "&Tools"
        '
        'mnuTGroup
        '
        Me.mnuTGroup.Index = 0
        Me.mnuTGroup.Text = "&Group"
        '
        'mnuTUngroup
        '
        Me.mnuTUngroup.Index = 1
        Me.mnuTUngroup.Text = "&Ungroup"
        '
        'mnuWindow
        '
        Me.mnuWindow.Index = 3
        Me.mnuWindow.MdiList = True
        Me.mnuWindow.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuWCascade, Me.mnuWHorizontal, Me.mnuWVertical, Me.mnuWArrange})
        Me.mnuWindow.Text = "&Window"
        '
        'mnuWCascade
        '
        Me.mnuWCascade.Index = 0
        Me.mnuWCascade.Text = "&Cascade"
        '
        'mnuWHorizontal
        '
        Me.mnuWHorizontal.Index = 1
        Me.mnuWHorizontal.Text = "Tile &Horizontal"
        '
        'mnuWVertical
        '
        Me.mnuWVertical.Index = 2
        Me.mnuWVertical.Text = "Tile &Vertical"
        '
        'mnuWArrange
        '
        Me.mnuWArrange.Index = 3
        Me.mnuWArrange.Text = "&Arrange Icons"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 4
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHAbout
        '
        Me.mnuHAbout.Index = 0
        Me.mnuHAbout.Text = "&About"
        '
        'dlgSave
        '
        Me.dlgSave.FileName = "doc1"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(928, 537)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.Text = "TI-Edit 1.2"

    End Sub

#End Region

#Region "File Loading and Saving"

    '' SUB mnuFNew_Click handles mnfFNew.Click
    '  creates new empty child and gives it a unique frm number
    Private Sub mnuFNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFNew.Click

        Dim frm As frmText                              ' create frm as type frmText
        Static intCount As Integer                      ' creating as static make variable persistant
        frm = New frmText()                             ' creates a new text form
        intCount += 1                                   ' increment the caption counter.
        frm.path = "c:\"
        frm.folder = "main"                             ' set the caption to be unique.
        frm.variable = intCount.ToString()
        frm.extension = "v2p"
        frm.calcName = "v200****"
        frm.progComment = "Created by TI-Edit 1.2"
        frm.MdiParent = Me                              ' set new forms parent

        frm.changed = True ' try creating reference to text box
        frm.Show()                                      ' show the form 

    End Sub

    '' SUB mnuFLoad_Click handles mnuFLoad.Click
    '  
    Private Sub mnuFOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFLoad.Click

        Dim byteCalcProgram() As Byte                   ' used to temp. store loaded file
        '                                               ' file is then transferred to detokenize class
        Dim sSourceFile As String = funOpenPath()       ' get path with open dialog
        mnuFNew_Click(Nothing, Nothing)                 ' open new child with new child subroutine
        Dim activeChild As Form = Me.ActiveMdiChild     ' create a reference to new child

        If (Not activeChild Is Nothing) Then            ' if we were sucessful in creating child
            Try                                         ' try creating reference to text box
                Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)
                If (Not theBox Is Nothing) Then         ' if the text box was found
                    Try                                 ' try to open file for binary read
                        Dim binReader As New BinaryReader(File.Open(sSourceFile, FileMode.Open))
                        Dim intFileLength As Integer = CInt(FileLen(sSourceFile))       ' get file length
                        ReDim byteCalcProgram(intFileLength - 1)        ' resize array to hold file
                        Try                                             ' and try to read it in
                            byteCalcProgram = binReader.ReadBytes(intFileLength + 1)    ' read file in
                            binReader.Close()                           ' close file
                            Dim myDetokenize As New detokenize()        ' create instance of detokenize class
                            myDetokenize.calcProgFile = byteCalcProgram ' set program file in detokenize class
                            myDetokenize.path = sSourceFile             ' set path property of detokenize class
                            DirectCast(activeChild, frmText).path = myDetokenize.path ' save path to child
                            DirectCast(activeChild, frmText).folder = myDetokenize.folderName    ' update folder name displayed
                            DirectCast(activeChild, frmText).variable = myDetokenize.variableName ' update variable name displayed
                            DirectCast(activeChild, frmText).extension = myDetokenize.extension ' save extension to child
                            DirectCast(activeChild, frmText).calcName = myDetokenize.calcName ' save calc name
                            DirectCast(activeChild, frmText).progComment = myDetokenize.progComment
                            theBox.Text = myDetokenize.ProgramText      ' update text box with detokenized file
                        Catch EndOfFile As EndOfStreamException         ' handles end of stream exception
                            MessageBox.Show("End Of File", "Look!", MessageBoxButtons.OK)
                        End Try                                         '
                    Catch FileNotFound As Exception                     ' handles file not found exception
                        MessageBox.Show("File Not Found", "Error!", MessageBoxButtons.OK)
                    End Try
                End If                                  '    
            Catch                                       '
                MessageBox.Show("You need to select a TextBox.")    ' if text box not found
            End Try                                     '
        End If                                          '

    End Sub

    '' Function funOpenPath
    '  Returns chosen open path as String
    Private Function funOpenPath() As String

        Dim sSourceFile As String                               ' holds source path
        With dlgOpen                                            ' use standard open dialog
            .Filter = "All (*.*)|*.*|Text (*.txt)|*.txt"        ' with these extension filters
            .AddExtension = True                                ' display extension 
            .CheckFileExists = True                             ' check that file exists
            .CheckPathExists = True                             ' check that path exists
            .InitialDirectory = IO.Path.GetDirectoryName("c:\") ' set initial directory to C drice
            .Multiselect = False                                ' only allow one file to be selected

            If .ShowDialog() = DialogResult.OK Then             ' if user pointed to a file
                If IO.File.Exists(.FileName) Then               ' and it exists
                    sSourceFile = .FileName                     ' save its path
                End If                                          '
            End If                                              '
        End With                                                '

        Return sSourceFile                                      ' return path string

    End Function

    '' SUB mnuFSave_Click handles mnuFSave.Click
    '  
    Private Sub mnuFSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles mnuFSave.Click, mnuFSaveAs.Click

        Dim byteCalcProgram() As Byte                   ' holds program file packaged to save
        Dim activeChild As Form = Me.ActiveMdiChild     ' create a reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox) ' create reference to childs text box
        Dim myDetokenize As New detokenize()            ' create instance of detokenize class

        sDestFile = DirectCast(activeChild, frmText).path               ' build save path string, with path
        sDestFile += DirectCast(activeChild, frmText).folder + "."      ' and folder
        sDestFile += DirectCast(activeChild, frmText).variable + "."    ' and variable name
        sDestFile += DirectCast(activeChild, frmText).extension         ' and extension

        myDetokenize.calcName = DirectCast(activeChild, frmText).calcName
        myDetokenize.progComment = DirectCast(activeChild, frmText).progComment
        myDetokenize.path = sDestFile                   ' send full path to detokenize class

        If sender Is mnuFSaveAs Then                    ' if sender was Save As then use save dialog
            Dim sTempFile As String = myDetokenize.folderName + "." + myDetokenize.variableName
            sDestFile = funSavePath(myDetokenize.path, sTempFile, myDetokenize.extension)   ' start at current path and file
            myDetokenize.path = sDestFile               ' save new path\folder.variable.ext from 'Save As'
            DirectCast(activeChild, frmText).path = myDetokenize.path ' save path to child
            DirectCast(activeChild, frmText).folder = myDetokenize.folderName    ' update folder name displayed
            DirectCast(activeChild, frmText).variable = myDetokenize.variableName ' update variable name displayed
            DirectCast(activeChild, frmText).extension = myDetokenize.extension ' save extension to child
        End If

        myDetokenize.ProgramText = theBox.Text          ' get program text to be packaged to save
        byteCalcProgram = myDetokenize.calcProgFile()   ' gets converted file from detokenize class

        Try                                             ' try writing file
            Dim binWriter As New BinaryWriter(File.Open(sDestFile, FileMode.OpenOrCreate))
            binWriter.Write(byteCalcProgram)            ' write program file to disk
            binWriter.Close()                           ' close file when done
            activeChild.Text = DirectCast(activeChild, frmText).folder + "." + DirectCast(activeChild, frmText).variable ' set title of child
            DirectCast(activeChild, frmText).changed = False                ' reset changed flag
        Catch eIO As Exception                          ' handles write exceptions
            Dim drEx As DialogResult                    '
            drEx = MessageBox.Show("Error, Could not write to file!", _
            "File Write Error", MessageBoxButtons.OK)   '
        End Try

    End Sub

    '' Function funSavePath
    'Returns chosen open path as String
    Private Function funSavePath(ByVal sPath As String, ByVal sFile As String, ByVal sExt As String) As String

        Dim sDestFile As String                             ' holds destination path
        With dlgSave                                        ' use standard save dialog
            .FileName = sFile                               ' initial file name
            .CheckFileExists = False                        ' don't check to see if it exists
            .CheckPathExists = True                         ' check that path exists
            .DefaultExt = sExt                              ' default save extension
            .Filter = sExt + " (*." + sExt + ")|*." + sExt + "|All (*.*)|*.*|Text (*.txt)|*.txt" ' extension filters
            .AddExtension = True                            ' add extension to path
            .InitialDirectory = IO.Path.GetDirectoryName(sPath) ' initial directory
            .OverwritePrompt = True                         ' prompt before overwriting
            .CreatePrompt = False                           ' don't prompt to create file

            If .ShowDialog() = DialogResult.OK Then         ' if a path was pointed to above
                sDestFile = .FileName                       ' save it into sDestFile
            End If                                          '
        End With                                            '

        Return sDestFile                                    ' return path string

    End Function

#End Region

#Region "Form exiting, closing, prompt to save before closing"

    '' When a child is closed from file menu the childs closing method is called
    '  the childs closing event handler checks to see if file has unsaved changes and
    '  calls subPrompt2Save to give user the chance to save changes.  When a child
    '  is closed from its title bar X it closing event handler operates as above.
    '  When the parent is closed from the title bar X its closing event handler calls
    '  each childs close method.  When the parent is closed from the File menu
    '  Exit command its event handler calls the form closing handler.

    '' SUB mnuFClose_Click handles mnuFClose.Click
    '  closes active child form
    Private Sub mnuFClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFClose.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' create reference to active child
        If (Not activeChild Is Nothing) Then            ' if a child does exist
            activeChild.Close()                         ' close it, close event handler in child catches
        End If                                          ' if file has been changed

    End Sub

    '' SUB mnuFExit_Click handles mnuFExit.Click
    '  shuts program down
    Private Sub mnuFExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFExit.Click

        Me.Close()

    End Sub

    '' Private Sub Form1_Closing handles Form1.Closing
    '  calls exit code in mnuFExit_Click
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        While (Not activeChild Is Nothing)              ' if active child exists
            activeChild.Close()                         ' close active child
            activeChild = Me.ActiveMdiChild             ' find next active child
        End While                                       ' loop until all children closed

    End Sub

    '' Public Sub subPrompt2Save, prompts user to save changed file
    '  before closing it, called from child
    Public Sub subPrompt2Save()

        Dim activechild As Form = Me.ActiveMdiChild     ' create reference to active child
        If MessageBox.Show("Save changes to file?", "Save Changes?", MessageBoxButtons.YesNo) _
        = DialogResult.Yes Then                         ' if user said yes
            mnuFSave_Click(Nothing, Nothing)            ' call save routine
        End If                                          ' done

    End Sub

#End Region

#Region "Menu command enabling/ disabling"

    '' SUB mnuFile_Popup handles mnuFile.Popup
    '  disables Save menu item if no child exists
    Private Sub mnuFile_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFile.Popup

        Dim boolEnabled As Boolean = (Not Me.ActiveMdiChild Is Nothing) '
        mnuFSave.Enabled = boolEnabled      ' disable if no child exists
        mnuFSaveAs.Enabled = boolEnabled    ' disable if no child exists
        mnuFClose.Enabled = boolEnabled     ' disable if no child exists
        mnuFPrint.Enabled = boolEnabled     ' disable if no child exists
        mnuFPpreview.Enabled = boolEnabled  ' disable if no child exists
        mnuFPSetup.Enabled = boolEnabled    ' disable if no child exists
        mnuFRecent.Enabled = False          ' diasable until feature is supported

    End Sub

    '' SUB mnuEdit_Popup handles mnuEdit.Popup
    '  disables Edit menu item if no child exists
    Private Sub mnuEdit_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEdit.Popup
        ' Create boolean variable whose value indicates if a child exists
        Dim boolEnabled As Boolean = (Not Me.ActiveMdiChild Is Nothing) '
        mnuEUndo.Enabled = boolEnabled          ' enable if child exists
        mnuERedo.Enabled = boolEnabled          ' enable if child exists
        mnuECut.Enabled = boolEnabled           ' enable if child exists
        mnuECopy.Enabled = boolEnabled          ' enable if child exists
        mnuEPaste.Enabled = boolEnabled         ' enable if child exists
        mnuEDelete.Enabled = boolEnabled        ' enable if child exists
        mnuESelectAll.Enabled = boolEnabled     ' enable if child exists
        mnuEFind.Enabled = False                ' disable until feature supported
        mnuEClear.Enabled = boolEnabled         ' enable if child exists

    End Sub

    '' SUB mnuTools_Popup handles mnuTools.Popup
    '  disables Tool menu items until supported
    Private Sub mnuTools_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTools.Popup

        mnuTGroup.Enabled = False
        mnuTUngroup.Enabled = False

    End Sub

    '' mnuWindow_Popup handles mnuWindow.Popup
    '  disables mnuWCascade, mnuWHorizontal, mnuWVertical, mnuWArrange if no child exists
    Private Sub mnuWindow_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuWindow.Popup

        ' Create boolean variable whose value indicates if a child exists
        Dim boolEnabled As Boolean = (Not Me.ActiveMdiChild Is Nothing) '
        mnuWCascade.Enabled = boolEnabled       ' enable if child exists
        mnuWHorizontal.Enabled = boolEnabled    ' enable if child exists
        mnuWVertical.Enabled = boolEnabled      ' enable if child exists
        mnuWArrange.Enabled = boolEnabled       ' enable if child exists

    End Sub

#End Region

#Region "Printing"

    '' Private Sub mnuFPrint_Click handles mnuFPrint.Click
    '  displays print dialog and prints page
    Private Sub btnPrintDialog_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles mnuFPrint.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        Dim tpDialog As New PrintDialog()               ' create print dialog
        tpTextBox.Text = theBox.Text                    ' set text box to print
        tpTextBox.Font = tpFont                         ' set font property ("Ti-92pluspc",12)
        tpDialog.Document = tpTextBox                   ' set documnet to print
        'tpDialog.AllowPrintToFile = True               ' not used
        'tpDialog.AllowSelection = True                 ' not used
        'tpDialog.AllowSomePages = True                 ' not used
        If tpDialog.ShowDialog = DialogResult.OK Then   ' if result of dialog was ok thei
            tpTextBox.Print()                           ' Issue print command
        End If                                          ' done

    End Sub

    '' Private Sub mnuFPSetup_Click handles mnuFPSetup.click
    '  displays page set up dialog
    Private Sub mnuFPSetup_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles mnuFPSetup.Click

        Dim dlg As New PageSetupDialog()
        dlg.Document = tpTextBox
        'dlg.AllowMargins = False
        'dlg.AllowOrientation = False
        'dlg.AllowPaper = False
        'dlg.AllowPrinter = False
        dlg.ShowDialog()

    End Sub

    '' Private Sub mnuFPreview_Click handles mnuFPreview.Click
    '  displays print preview dialog of current document
    Private Sub mnuFPreview_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles mnuFPpreview.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        tpTextBox.Text = theBox.Text                    ' set text property of TextPrint
        tpTextBox.Font = tpFont                         ' set font property ("Ti-92pluspc",12)
        Dim dlg As New PrintPreviewDialog()             ' create preview dialog
        dlg.Document = tpTextBox                        ' set document to preview
        dlg.WindowState = FormWindowState.Maximized     ' set to open maximized
        dlg.ShowDialog()                                ' display preview dialog

    End Sub


#End Region

#Region "Edit Menu"

    '' Private Sub mnuEUndo_Click handles mnuEUndo.Click
    '  undoes last editing change
    Private Sub mnuEUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEUndo.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.Undo()

    End Sub

    '' Private Sub mnuERedo_Click handles mnuERedo.Click
    '  redoes last undo
    Private Sub mnuERedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuERedo.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.Undo()

    End Sub

    '' Private Sub mnuECut_Click handles mnuECut.Click
    '  cuts selected text
    Private Sub mnuECut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuECut.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.Cut()

    End Sub

    '' Private Sub mnuECopy_Click handles mnuECopy.Click
    '  copies selected text
    Private Sub mnuECopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuECopy.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.Copy()

    End Sub

    '' Private Sub mnuEPaste_Click handles mnuEpaste.Click
    '  paste text from clipboard to current cursor position
    Private Sub mnuEPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEPaste.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.Paste()

    End Sub

    '' Private Sub mnuEdelete_Click handles mnuEDelete.Click
    '  deletes selected text
    Private Sub mnuEDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEDelete.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.Cut()

    End Sub

    '' Private Sub mnuESelectAll_Click handles mnuESelectAll.Click
    '  selects entire docoment
    Private Sub mnuESelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuESelectAll.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.SelectAll()

    End Sub

    '' Private Sub mnuEClear_Click handles mnuEClear.Click
    '  clears entire document
    Private Sub mnuEClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEClear.Click

        Dim activeChild As Form = Me.ActiveMdiChild     ' get reference to active child
        Dim theBox As TextBox = CType(activeChild.ActiveControl, TextBox)   ' reference to childs text box
        theBox.Clear()

    End Sub

#End Region

#Region "Window Menu"

    '' SUB mnuWLayout handles mnuWCascade.Click, mnuWHorizontal.Click, mnuWVertical.Click,
    '   mnuWLayout.Click.  Arrange children if they exist layout according to sender.
    Private Sub mnuWLayout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles mnuWCascade.Click, mnuWArrange.Click, mnuWHorizontal.Click, mnuWVertical.Click

        If sender Is mnuWCascade Then               ' if cascade
            Me.LayoutMdi(MdiLayout.Cascade)         ' 
        ElseIf sender Is mnuWHorizontal Then        ' if hoizontal
            Me.LayoutMdi(MdiLayout.TileHorizontal)  '
        ElseIf sender Is mnuWVertical Then          ' if vertical
            Me.LayoutMdi(MdiLayout.TileVertical)    '
        ElseIf sender Is mnuWArrange Then           ' if arrange
            Me.LayoutMdi(MdiLayout.ArrangeIcons)    '
        End If                                      '

    End Sub

#End Region


End Class