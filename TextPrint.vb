Public Class TextPrint

    Inherits Printing.PrintDocument         ' Inherits functionality of PrintDocument

    Private fntPrintFont As Font            ' holds default font
    Private strText As String               ' holds text string to print


#Region "Initialization Code"

    '' Public Sub New, This code runs when a new instace of TextPrint is created
    '  it copies the text string sent with call to strText
    Public Sub New(ByVal Text As String)

        MyBase.New()                        ' Sets the file stream
        strText = Text                      ' copy string sent to strText

    End Sub


#End Region

#Region "Properties"

    '' Public Property Text, gets or sets text to print
    '  not sure if this is needed
    Public Property Text() As String        '
        Get                                 ' 
            Return strText                  ' returns strText
        End Get                             '
        Set(ByVal Value As String)          ' 
            strText = Value                 ' sets strText with string sent
        End Set                             '
    End Property                            '

    '' Public Property Font
    '  gets or sets font print with
    Public Property Font() As Font          '
        Get                                 '
            Return fntPrintFont             ' returns font currently used
        End Get                             '
        Set(ByVal Value As Font)            '
            fntPrintFont = Value            ' sets font with value sent
        End Set                             '
    End Property                            '

#End Region

    '' Protected Overides Sub OnBeginPrint
    '  Checks to see if font is set, if not sets it to ("Times New Roman", 12)
    Protected Overrides Sub OnBeginPrint(ByVal ev As Printing.PrintEventArgs)

        MyBase.OnBeginPrint(ev)             ' Run base code
        If fntPrintFont Is Nothing Then     ' If font not set
            fntPrintFont = New Font("Times New Roman", 12) ' Sets the default font
        End If                              '

    End Sub


    '' Protected Overrides Sub OnPrintPage
    '  Provides the print logic for our document
    Protected Overrides Sub OnPrintPage(ByVal ev As Printing.PrintPageEventArgs)

        MyBase.OnPrintPage(ev)              ' Run base code
        Static intCurrentChar As Integer    ' Persistant variable, charecter position
        Dim intPrintAreaHeight As Integer   ' holds print area height
        Dim intPrintAreaWidth As Integer    ' holds print area width
        Dim intMarginLeft As Integer        ' holds left margin
        Dim intMarginTop As Integer         ' holds right margin

        ' Set printing area boundaries and margin coordinates
        With MyBase.DefaultPageSettings     ' calculate print area
            intPrintAreaHeight = .PaperSize.Height - .Margins.Top - .Margins.Bottom
            intPrintAreaWidth = .PaperSize.Width - .Margins.Left - .Margins.Right
            intMarginLeft = .Margins.Left   ' X
            intMarginTop = .Margins.Top     ' Y
        End With

        ' If Landscape set, swap printing height/width
        If MyBase.DefaultPageSettings.Landscape Then    ' if landscape is set
            Dim intTemp As Integer                      ' create temp variable
            intTemp = intPrintAreaHeight                ' copy height to temp
            intPrintAreaHeight = intPrintAreaWidth      ' copy width to height
            intPrintAreaWidth = intTemp                 ' copy temp to width
        End If

        Dim intLineCount As Int32 = CInt(intPrintAreaHeight / Font.Height) ' Calculate total number of lines
        Dim rectPrintingArea As New RectangleF(intMarginLeft, _
            intMarginTop, intPrintAreaWidth, intPrintAreaHeight)    ' Initialise rectangle printing area
        Dim objSF As New StringFormat(StringFormatFlags.LineLimit)  ' Initialise StringFormat class, for text layout
        Dim intLinesFilled As Int32
        Dim intCharsFitted As Int32

        ' Figure out how many lines will fit into rectangle
        ev.Graphics.MeasureString(Mid(strText, UpgradeZeros(intCurrentChar)), Font, _
            New SizeF(intPrintAreaWidth, intPrintAreaHeight), objSF, intCharsFitted, intLinesFilled)

        ' Print the text to the page
        ev.Graphics.DrawString(Mid(strText, UpgradeZeros(intCurrentChar)), Font, Brushes.Black, rectPrintingArea, objSF)
        intCurrentChar += intCharsFitted        ' Increase current char count
        If intCurrentChar < strText.Length Then ' Check whether we need to print more
            ev.HasMorePages = True              ' yes, set more pages property
        Else                                    ' or
            ev.HasMorePages = False             ' no, clear more pages property
            intCurrentChar = 0                  ' reset current charecter pointer
        End If                                  '

    End Sub


    Public Function UpgradeZeros(ByVal Input As Integer) As Integer
        ' Upgrades all zeros to ones
        ' - used as opposed to defunct IIF or messy If statements
        If Input = 0 Then
            Return 1
        Else
            Return Input
        End If
    End Function


End Class
