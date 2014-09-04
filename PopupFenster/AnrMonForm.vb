Imports System.ComponentModel
Imports System.Drawing.Drawing2D

'<System.ComponentModel.DefaultPropertyAttribute("Content"), System.ComponentModel.DesignTimeVisible(False)> _
Friend Class AnrMonForm
    Inherits System.Windows.Forms.Form

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Copyright �1996-2011 VBnet/Randy Birch, All Rights Reserved.
    ' Some pages may also contain other copyrights by the author.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Distribution: You can freely use this code in your own
    '               applications, but you may not reproduce 
    '               or publish this code on any web site,
    '               online service, or distribute as source 
    '               on any media without express permission.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sub New(ByVal vParent As F_AnrMon, ByRef vCommon As CommonFenster)
        P_Parent = vParent
        P_Common = vCommon
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        Me.SetStyle(ControlStyles.ResizeRedraw, True)
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'PopUpAnrMonForm
        '
        Me.ClientSize = New System.Drawing.Size(392, 66)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "PopUpAnrMonForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.ResumeLayout(True)
    End Sub

    Private bMouseOnClose As Boolean = False
    Private bMouseOnLink As Boolean = False
    Private bMouseOnOptions As Boolean = False
    Private iHeightOfTitle As Integer
    Private iHeightOfAnrName As Integer
    Private iHeightOfTelNr As Integer
    Private iTitleOrigin As Integer

    Public Event LinkClick(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event CloseClick(ByVal sender As Object, ByVal e As System.EventArgs)

#Region "Properties"
    Protected Overrides ReadOnly Property ShowWithoutActivation() As Boolean
        Get
            Return True
        End Get
    End Property

    Private pnParent As F_AnrMon
    Shadows Property P_Parent() As F_AnrMon
        Get
            Return pnParent
        End Get
        Set(ByVal value As F_AnrMon)
            pnParent = value
        End Set
    End Property

    Private pnCmn As CommonFenster
    Shadows Property P_Common() As CommonFenster
        Get
            Return pnCmn
        End Get
        Set(ByVal value As CommonFenster)
            pnCmn = value
        End Set
    End Property

    'Protected Overrides ReadOnly Property CreateParams As CreateParams

    '    Get
    '        Dim baseParams As System.Windows.Forms.CreateParams = MyBase.CreateParams
    '        ' WS_EX_NOACTIVATE = 0x08000000,
    '        ' WS_EX_TOOLWINDOW = 0x00000080,
    '        ' baseParams.ExStyle |= ( int )( 
    '        '  Win32.ExtendedWindowStyles.WS_EX_NOACTIVATE | 
    '        '  Win32.ExtendedWindowStyles.WS_EX_TOOLWINDOW );
    '        'baseParams.ExStyle = baseParams.ExStyle Or CInt((WindowStyles.WS_EX_NOACTIVATE Or WindowStyles.WS_EX_TOOLWINDOW))

    '        Return baseParams
    '    End Get
    'End Property
#End Region

#Region "Functions & Private properties"
    Private ReadOnly Property RectTelNr() As RectangleF
        Get
            If P_Parent.Image IsNot Nothing Then
                Return New RectangleF(P_Parent.ImagePosition.X + P_Parent.ImageSize.Width + P_Common.TextPadding.Left, CSng(P_Common.TextPadding.Top + iHeightOfTitle + 1.5 * P_Common.HeaderHeight), Me.Width - P_Parent.ImageSize.Width - P_Parent.ImagePosition.X - P_Common.TextPadding.Left - P_Common.TextPadding.Right, iHeightOfTelNr)
            Else
                Return New RectangleF(P_Common.TextPadding.Left, CSng(P_Common.TextPadding.Top + iHeightOfTitle + 1.5 * P_Common.HeaderHeight), Me.Width - P_Common.TextPadding.Left - P_Common.TextPadding.Right, iHeightOfTelNr)
            End If
        End Get
    End Property

    Private ReadOnly Property RectAnrName() As RectangleF

        Get
            If P_Parent.Image IsNot Nothing Then
                Return New RectangleF(P_Parent.ImagePosition.X + P_Parent.ImageSize.Width + P_Common.TextPadding.Left, _
                                      CSng(P_Common.TextPadding.Top + iHeightOfTitle + 1.5 * P_Common.HeaderHeight + iHeightOfTelNr), _
                                      Me.Width - P_Parent.ImageSize.Width - P_Parent.ImagePosition.X - P_Common.TextPadding.Left - P_Common.TextPadding.Right, iHeightOfAnrName)
            Else
                Return New RectangleF(P_Common.TextPadding.Left, CSng(P_Common.TextPadding.Top + iHeightOfTitle + 1.5 * P_Common.HeaderHeight + iHeightOfTelNr), _
                                      Me.Width - P_Common.TextPadding.Left - P_Common.TextPadding.Right, iHeightOfAnrName)
            End If
        End Get
    End Property

    Private ReadOnly Property RectFirma() As RectangleF
        Get
            If P_Parent.Image IsNot Nothing Then
                Return New RectangleF(P_Parent.ImagePosition.X + P_Parent.ImageSize.Width + P_Common.TextPadding.Left, Me.Height - P_Common.TextPadding.Bottom - iHeightOfTitle, Me.Width - P_Parent.ImageSize.Width - P_Parent.ImagePosition.X - P_Common.TextPadding.Left - P_Common.TextPadding.Right, iHeightOfTitle)
            Else
                Return New RectangleF(P_Common.TextPadding.Left, Me.Height - iHeightOfTitle - P_Common.TextPadding.Bottom, Me.Width - P_Common.TextPadding.Left - P_Common.TextPadding.Right, iHeightOfTitle)
            End If
        End Get
    End Property

    Private ReadOnly Property RectClose() As Rectangle
        Get
            Return New Rectangle(Me.Width - 5 - 16, 12, 16, 16)
        End Get
    End Property

    Private ReadOnly Property RectOptions() As Rectangle
        Get
            Return New Rectangle(Me.Width - 5 - 16, 12 + 16 + 5, 16, 16)
        End Get
    End Property

    Private ReadOnly Property RectImage() As Rectangle
        Get
            If P_Parent.Image IsNot Nothing Then
                Return New Rectangle(P_Parent.ImagePosition, P_Parent.ImageSize)
            End If
        End Get
    End Property

#End Region

#Region "Events"

    Private Sub AnrMonForm_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Me.Finalize()
    End Sub

    Private Sub Me_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        If P_Common.CloseButton Then
            If RectClose.Contains(e.X, e.Y) Then
                bMouseOnClose = True
            Else
                bMouseOnClose = False
            End If
        End If
        If P_Common.OptionsButton Then
            If RectOptions.Contains(e.X, e.Y) Then
                bMouseOnOptions = True
            Else
                bMouseOnOptions = False
            End If
        End If
        If RectAnrName.Contains(e.X, e.Y) Then
            bMouseOnLink = True
        Else
            bMouseOnLink = False
        End If
        Invalidate()
    End Sub

    Private Sub Me_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        If RectClose.Contains(e.X, e.Y) Then
            RaiseEvent CloseClick(Me, EventArgs.Empty)
        End If
        If RectAnrName.Contains(e.X, e.Y) Then
            RaiseEvent LinkClick(Me, EventArgs.Empty)
        End If
        If RectOptions.Contains(e.X, e.Y) Then
            If P_Parent.OptionsMenu IsNot Nothing Then
                P_Parent.OptionsMenu.Show(Me, New Point(RectOptions.Right - P_Parent.OptionsMenu.Width, RectOptions.Bottom))
                P_Parent.bShouldRemainVisible = True
            End If
        End If
    End Sub

    Private Sub Me_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim iTelNameL�nge As Integer
        Dim iUhrzeitL�nge As Integer
        Dim iAnrNameL�nge As Integer
        Dim sUhrzeit As String
        Dim sTelName As String
        Dim L�nge As Integer

        Dim rcBody As New Rectangle(0, 0, Me.Width, Me.Height)

        Dim rcHeader As New Rectangle(0, 0, Me.Width, P_Common.HeaderHeight)
        Dim rcForm As New Rectangle(0, 0, Me.Width - 1, Me.Height - 1)
        Dim brBody As New LinearGradientBrush(rcBody, P_Common.BodyColor, P_Common.GetLighterColor(P_Common.BodyColor), LinearGradientMode.Vertical)
        Dim drawFormatCenter As New StringFormat()
        Dim drawFormatRight As New StringFormat()

        Dim brHeader As New LinearGradientBrush(rcHeader, P_Common.HeaderColor, P_Common.GetDarkerColor(P_Common.HeaderColor), LinearGradientMode.Vertical)
        Dim RectZeit As RectangleF
        Dim RectTelName As RectangleF

        drawFormatCenter.Alignment = StringAlignment.Center
        drawFormatRight.Alignment = StringAlignment.Far

        With e.Graphics
            .Clip = New Region(rcBody)
            .FillRectangle(brBody, rcBody)
            .FillRectangle(brHeader, rcHeader)
            .DrawRectangle(New Pen(P_Common.BorderColor), rcForm)
            If P_Common.CloseButton Then
                If bMouseOnClose Then
                    .FillRectangle(New SolidBrush(P_Common.ButtonHoverColor), RectClose)
                    .DrawRectangle(New Pen(P_Common.ButtonBorderColor), RectClose)
                End If
                .DrawLine(New Pen(P_Common.ContentColor, 2), RectClose.Left + 4, RectClose.Top + 4, RectClose.Right - 4, RectClose.Bottom - 4)
                .DrawLine(New Pen(P_Common.ContentColor, 2), RectClose.Left + 4, RectClose.Bottom - 4, RectClose.Right - 4, RectClose.Top + 4)
            End If
            If P_Common.OptionsButton Then
                If bMouseOnOptions Then
                    .FillRectangle(New SolidBrush(P_Common.ButtonHoverColor), RectOptions)
                    .DrawRectangle(New Pen(P_Common.ButtonBorderColor), RectOptions)
                End If
                .FillPolygon(New SolidBrush(ForeColor), New Point() {New Point(RectOptions.Left + 4, RectOptions.Top + 6), New Point(RectOptions.Left + 12, RectOptions.Top + 6), New Point(RectOptions.Left + 8, RectOptions.Top + 4 + 6)})
            End If
            iHeightOfTitle = CInt(.MeasureString("A", P_Common.TitleFont).Height)
            iHeightOfAnrName = CInt(.MeasureString("A", P_Common.ContentFont).Height)
            iHeightOfTelNr = CInt(.MeasureString("A", P_Common.TelNrFont).Height)
            iTitleOrigin = P_Common.TextPadding.Left
            If P_Parent.Image IsNot Nothing Then
                Dim showim As Image = New Bitmap(P_Parent.ImageSize.Width, P_Parent.ImageSize.Height)
                Dim g1 As Graphics = Graphics.FromImage(showim)
                g1.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g1.DrawImage(P_Parent.Image, 0, 0, P_Parent.ImageSize.Width, P_Parent.ImageSize.Height)
                g1.Dispose()
                .DrawImage(showim, P_Parent.ImagePosition)
                .DrawRectangle(New Pen(P_Common.ButtonBorderColor), RectImage)
            End If
            L�nge = P_Parent.Size.Width - P_Common.TextPadding.Right - 21 - iTitleOrigin + P_Common.TextPadding.Left
            sUhrzeit = CDate(P_Parent.Uhrzeit).ToString("dddd, dd. MMMM yyyy HH:mm:ss")
            sTelName = P_Parent.TelName
            iTelNameL�nge = CInt(.MeasureString(sTelName, P_Common.TitleFont).Width)
            iUhrzeitL�nge = CInt(.MeasureString(sUhrzeit, P_Common.TitleFont).Width)
            If iTelNameL�nge + iUhrzeitL�nge > L�nge Then
                sUhrzeit = CDate(P_Parent.Uhrzeit).ToString("dddd, dd. MMM. yy HH:mm:ss")
                iUhrzeitL�nge = CInt(.MeasureString(sUhrzeit, P_Common.TitleFont).Width)
                If iTelNameL�nge + iUhrzeitL�nge > L�nge Then
                    sUhrzeit = CDate(P_Parent.Uhrzeit).ToString("ddd, dd.MM.yy HH:mm:ss")
                    iUhrzeitL�nge = CInt(.MeasureString(sUhrzeit, P_Common.TitleFont).Width)
                    If iTelNameL�nge + iUhrzeitL�nge > L�nge Then
                        sUhrzeit = CDate(P_Parent.Uhrzeit).ToString("dd.MM.yy HH:mm:ss")
                        iUhrzeitL�nge = CInt(.MeasureString(sUhrzeit, P_Common.TitleFont).Width)
                    End If
                End If
            End If
            RectZeit = New RectangleF(iTitleOrigin + P_Common.TextPadding.Left, P_Common.TextPadding.Top + P_Common.HeaderHeight, .MeasureString(sUhrzeit, P_Common.TitleFont).Width, iHeightOfTitle)
            RectTelName = New RectangleF(RectZeit.Right, RectZeit.Top, RectClose.Left - RectZeit.Right, iHeightOfTitle)

            .DrawString(sUhrzeit, P_Common.TitleFont, New SolidBrush(P_Common.TitleColor), RectZeit)
            If iTelNameL�nge > RectTelName.Width Then
                RectTelName.Y = P_Common.HeaderHeight
                RectTelName.Size = New Size(CInt(RectTelName.Width), CInt(RectTelName.Height * 2 - 3))
            End If
            .DrawString(sTelName, P_Common.TitleFont, New SolidBrush(P_Common.TitleColor), RectTelName, drawFormatRight)
            .DrawString(P_Parent.TelNr, P_Common.TelNrFont, New SolidBrush(P_Common.TitleColor), RectTelNr, drawFormatCenter)
            .DrawString(P_Parent.Firma, P_Common.TitleFont, New SolidBrush(P_Common.TitleColor), RectFirma, drawFormatCenter)

            Dim tempfont As New Font(P_Common.DefFontName, 16, P_Common.DefFontStyle, P_Common.DefGraphicsUnit, P_Common.DefgdiCharSet)
            Dim sAnrName As String
            sAnrName = P_Parent.AnrName
            iAnrNameL�nge = CInt(.MeasureString(sAnrName, tempfont, 0, StringFormat.GenericTypographic).Width)

            If iAnrNameL�nge > RectAnrName.Width Then
                Dim iFontSize As Integer
                iFontSize = CInt(((RectAnrName.Width - P_Common.TextPadding.Right - P_Common.TextPadding.Left) * (tempfont.Size / 72 * .DpiX - 1.5 * P_Common.TextPadding.Top)) / (iAnrNameL�nge - 2 * P_Common.TextPadding.Left))
                iFontSize = CInt(IIf(iFontSize < 8, 8, iFontSize))
                tempfont = New Font(P_Common.DefFontName, 16, P_Common.DefFontStyle, P_Common.DefGraphicsUnit, P_Common.DefgdiCharSet)
            End If

            If bMouseOnLink Then
                Me.Cursor = Cursors.Hand
                .DrawString(P_Parent.AnrName, tempfont, New SolidBrush(P_Common.LinkHoverColor), RectAnrName, drawFormatCenter)
            Else
                Me.Cursor = Cursors.Default
                .DrawString(P_Parent.AnrName, tempfont, New SolidBrush(P_Common.ContentColor), RectAnrName, drawFormatCenter)
            End If
        End With
    End Sub

#End Region

    Protected Overrides Sub Finalize()
        Me.Hide()
        MyBase.Finalize()
    End Sub
End Class

Public Class F_AnrMon
    Inherits Component

    Private cmnPrps As New CommonFenster
    Public Event LinkClick(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event Close(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event Closed(ByVal sender As Object, ByVal e As System.EventArgs)

    Public Event ToolStripMenuItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs)

    Private WithEvents fPopup As New AnrMonForm(Me, cmnPrps)
    'Private WithEvents tmAnimation As New Timer
    Private WithEvents tmWait As New Timer

    Private bAppearing As Boolean = True
    Public bShouldRemainVisible As Boolean = False
    Private i As Integer = 0

    Private bMouseIsOn As Boolean = False
    Private iMaxPosition As Integer
    Private dMaxOpacity As Double
    Private dummybool As Boolean

    Private CompContainer As New System.ComponentModel.Container()
    Private WithEvents AnrMonContextMenuStrip As New ContextMenuStrip(CompContainer)
    Private ToolStripMenuItemKontakt�ffnen As New ToolStripMenuItem()
    Private ToolStripMenuItemR�ckruf As New ToolStripMenuItem()
    Private ToolStripMenuItemKopieren As New ToolStripMenuItem()

    Enum eStartPosition
        BottomRight
        BottomLeft
        TopLeft
        TopRight
    End Enum

    Enum eMoveDirection
        Y
        X
    End Enum

#Region "Properties"

    Private WithEvents ctContextMenu As ContextMenuStrip = Nothing
    Public Property OptionsMenu() As ContextMenuStrip
        Get
            Return ctContextMenu
        End Get
        Set(ByVal value As ContextMenuStrip)
            ctContextMenu = value
        End Set
    End Property

    Private iShowDelay As Integer = 3000
    Public Property ShowDelay() As Integer
        Get
            Return iShowDelay
        End Get
        Set(ByVal value As Integer)
            iShowDelay = value
        End Set
    End Property

    Private szSize As Size = New Size(400, 100)
    Public Property Size() As Size
        Get
            Return szSize
        End Get
        Set(ByVal value As Size)
            szSize = value
        End Set
    End Property

    Private bAutoAusblenden As Boolean = True
    Public Property AutoAusblenden() As Boolean
        Get
            Return bAutoAusblenden
        End Get
        Set(ByVal value As Boolean)
            bAutoAusblenden = value
        End Set
    End Property

    Private szPositionsKorrektur As Size = New Size(0, 0)
    Public Property PositionsKorrektur() As Size
        Get
            Return szPositionsKorrektur
        End Get
        Set(ByVal value As Size)
            szPositionsKorrektur = value
        End Set
    End Property

    Private bEffektTransparenz As Boolean = True
    Public Property EffektTransparenz() As Boolean
        Get
            Return bEffektTransparenz
        End Get
        Set(ByVal value As Boolean)
            bEffektTransparenz = value
        End Set
    End Property

    Private bEffektMove As Boolean = True
    Public Property EffektMove() As Boolean
        Get
            Return bEffektMove
        End Get
        Set(ByVal value As Boolean)
            bEffektMove = value
        End Set
    End Property

    Private iEffektMoveGeschwindigkeit As Integer = 5
    Public Property EffektMoveGeschwindigkeit() As Integer
        Get
            Return iEffektMoveGeschwindigkeit
        End Get
        Set(ByVal value As Integer)
            iEffektMoveGeschwindigkeit = value
        End Set
    End Property

    Private pStartpunkt As eStartPosition
    Public Property Startpunkt() As eStartPosition
        Get
            Return pStartpunkt
        End Get
        Set(ByVal value As eStartPosition)
            pStartpunkt = value
        End Set
    End Property

    Private _MoveDirection As eMoveDirection
    Public Property MoveDirecktion() As eMoveDirection
        Get
            Return _MoveDirection
        End Get
        Set(ByVal value As eMoveDirection)
            _MoveDirection = value
        End Set
    End Property

    Private ptImagePosition As Point = New Point(12, 32) 'New Point(12, 21)
    Public Property ImagePosition() As Point
        Get
            Return ptImagePosition
        End Get
        Set(ByVal value As Point)
            ptImagePosition = value

        End Set
    End Property

    Private szImageSize As Size = New Size(48, 48) 'New Size(0, 0)
    Public Property ImageSize() As Size
        Get
            If szImageSize.Width = 0 Then
                If Image IsNot Nothing Then
                    Return Image.Size
                Else
                    Return New Size(32, 32)
                End If
            Else
                Return szImageSize
            End If
        End Get
        Set(ByVal value As Size)
            szImageSize = value
        End Set
    End Property

    Private imImage As Image = Nothing
    Public Property Image() As Image
        Get
            Return imImage
        End Get
        Set(ByVal value As Image)
            imImage = value
        End Set
    End Property

    Private sAnrName As String
    Public Property AnrName() As String
        Get
            Return sAnrName
        End Get
        Set(ByVal value As String)
            sAnrName = value
        End Set
    End Property

    Private sUhrzeit As String
    Public Property Uhrzeit() As String
        Get
            Return sUhrzeit
        End Get
        Set(ByVal value As String)
            sUhrzeit = value
        End Set
    End Property

    Private sTelNr As String
    Public Property TelNr() As String
        Get
            Return sTelNr
        End Get
        Set(ByVal value As String)
            sTelNr = value
        End Set
    End Property

    Private sTelName As String
    Public Property TelName() As String
        Get
            Return sTelName
        End Get
        Set(ByVal value As String)
            sTelName = value
        End Set
    End Property

    Private sFirma As String
    Public Property Firma() As String
        Get
            Return sFirma
        End Get
        Set(ByVal value As String)
            sFirma = value
        End Set
    End Property

#End Region

    Public Sub New()
        With fPopup
            .FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            .StartPosition = System.Windows.Forms.FormStartPosition.Manual
            .ShowInTaskbar = True
        End With
        InitializeComponentContextMenuStrip()
    End Sub

    Public Sub Popup()
        Dim X As Integer
        Dim Y As Integer
        Dim retVal As Boolean

        tmWait.Interval = 200
        With fPopup
            .TopMost = True
            .Size = Size
            .Opacity = CDbl(IIf(bEffektTransparenz, 0, 1))

            Select Case Startpunkt
                Case eStartPosition.BottomLeft
                    X = Screen.PrimaryScreen.WorkingArea.Left + 10 - PositionsKorrektur.Width
                    Y = Screen.PrimaryScreen.WorkingArea.Bottom - 10 - .Height - PositionsKorrektur.Height
                Case eStartPosition.TopLeft
                    X = Screen.PrimaryScreen.WorkingArea.Left + 10 - PositionsKorrektur.Width
                    Y = Screen.PrimaryScreen.WorkingArea.Top + 10 - PositionsKorrektur.Height
                Case eStartPosition.BottomRight
                    X = Screen.PrimaryScreen.WorkingArea.Right - 10 - .Size.Width - PositionsKorrektur.Width
                    Y = Screen.PrimaryScreen.WorkingArea.Bottom - 10 - .Height - PositionsKorrektur.Height
                Case eStartPosition.TopRight
                    X = Screen.PrimaryScreen.WorkingArea.Right - 10 - .Size.Width - PositionsKorrektur.Width
                    Y = Screen.PrimaryScreen.WorkingArea.Top + 10 - PositionsKorrektur.Height
            End Select

            If bEffektMove Then
                Select Case MoveDirecktion
                    Case eMoveDirection.X
                        Select Case Startpunkt
                            Case eStartPosition.BottomLeft, eStartPosition.TopLeft
                                X = Screen.PrimaryScreen.WorkingArea.Left - fPopup.Size.Width + 2
                            Case eStartPosition.BottomRight, eStartPosition.TopRight
                                X = Screen.PrimaryScreen.WorkingArea.Right + 2
                        End Select
                    Case eMoveDirection.Y
                        Select Case Startpunkt
                            Case eStartPosition.TopLeft, eStartPosition.TopRight
                                Y = Screen.PrimaryScreen.WorkingArea.Top - fPopup.Height + 1
                            Case eStartPosition.BottomRight, eStartPosition.BottomLeft
                                Y = Screen.PrimaryScreen.WorkingArea.Bottom - 1
                        End Select
                End Select

            End If

            .Location = New Point(X, Y)
            .Text = AnrName & CStr(IIf(TelNr = "", "", " (" & TelNr & ")"))

            retVal = OutlookSecurity.SetWindowPos(.Handle, hWndInsertAfterFlags.HWND_TOPMOST, 0, 0, 0, 0, _
                                                  CType(SetWindowPosFlags.DoNotActivate + _
                                                  SetWindowPosFlags.IgnoreMove + _
                                                  SetWindowPosFlags.IgnoreResize + _
                                                  SetWindowPosFlags.DoNotChangeOwnerZOrder, SetWindowPosFlags))

            .Show()
        End With

        'tmAnimation.Interval = 1 'iEffektMoveGeschwindigkeit
        'tmAnimation.Start()
    End Sub

    ''' <summary>
    ''' Initialisierungsroutine des ehemaligen AnrMonForm. Es wird das ContextMenuStrip an Sich initialisiert 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeComponentContextMenuStrip()
        '
        'ContextMenuStrip
        '
        With Me.AnrMonContextMenuStrip
            .Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemKontakt�ffnen, Me.ToolStripMenuItemR�ckruf, Me.ToolStripMenuItemKopieren})
            .Name = "AnrMonContextMenuStrip"
            .RenderMode = System.Windows.Forms.ToolStripRenderMode.System
            .Size = New System.Drawing.Size(222, 70)
        End With
        '
        'ToolStripMenuItemKontakt�ffnen
        '
        With Me.ToolStripMenuItemKontakt�ffnen
            '.Image = ToolStripMenuItemKontakt�ffnenImage
            .ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
            .Name = "ToolStripMenuItemKontakt�ffnen"
            .Size = New System.Drawing.Size(221, 22)
            '.Text = ToolStripMenuItemKontakt�ffnenText '"Kontakt �ffnen"
        End With
        '
        'ToolStripMenuItemR�ckruf
        '
        With Me.ToolStripMenuItemR�ckruf
            '.Image = ToolStripMenuItemKontakt�ffnenImage
            .ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
            .Name = "ToolStripMenuItemR�ckruf"
            .Size = New System.Drawing.Size(221, 22)
            '.Text = ToolStripMenuItemR�ckrufText '"R�ckruf"
        End With
        '
        'ToolStripMenuItemKopieren
        '
        With Me.ToolStripMenuItemKopieren
            '.Image = ToolStripMenuItemKopierenImage
            .ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
            .Name = "ToolStripMenuItemKopieren"
            .Size = New System.Drawing.Size(221, 22)
            '.Text = ToolStripMenuItemKopierenText '"In Zwischenablage kopieren"
        End With
        Me.OptionsMenu = Me.AnrMonContextMenuStrip
    End Sub

    Public Sub Hide()
        bMouseIsOn = False
        tmWait.Stop()
        'tmAnimation.Start()
    End Sub

#Region "Eigene Events"
    Private Sub fPopup_CloseClick() Handles fPopup.CloseClick
        RaiseEvent Close(Me, EventArgs.Empty)
        Me.Finalize()
    End Sub

    Private Sub fPopup_LinkClick() Handles fPopup.LinkClick
        RaiseEvent LinkClick(Me, EventArgs.Empty)
    End Sub
    Private Sub fPopup_ToolStripMenuItemR�ckrufClickClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ctContextMenu.ItemClicked
        RaiseEvent ToolStripMenuItemClicked(Me, e)
    End Sub
#End Region

    Private Function GetOpacityBasedOnPosition() As Double

        Dim iCentPurcent As Integer
        Dim iCurrentlyShown As Integer
        Dim dPourcentOpacity As Double

        Select Case MoveDirecktion
            Case eMoveDirection.X
                iCentPurcent = fPopup.Width
                Select Case Startpunkt
                    Case eStartPosition.BottomLeft, eStartPosition.TopLeft
                        iCurrentlyShown = fPopup.Right
                    Case eStartPosition.BottomRight, eStartPosition.TopRight
                        iCurrentlyShown = Screen.PrimaryScreen.WorkingArea.Width - fPopup.Left
                End Select
                dPourcentOpacity = iCurrentlyShown * 100 / iCentPurcent
            Case eMoveDirection.Y
                iCentPurcent = fPopup.Height
                Select Case Startpunkt
                    Case eStartPosition.BottomLeft, eStartPosition.BottomRight
                        iCurrentlyShown = Screen.PrimaryScreen.WorkingArea.Height - fPopup.Top
                    Case eStartPosition.TopLeft, eStartPosition.TopRight
                        iCurrentlyShown = fPopup.Bottom
                End Select
                dPourcentOpacity = iCentPurcent / 100 * iCurrentlyShown
        End Select

        Return dPourcentOpacity / 100
    End Function

    Public Sub tmAnimation_Tick() '(ByVal sender As Object, ByVal e As System.EventArgs) ' Handles tmAnimation.Tick
        Dim StoppAnimation As Boolean = False
        With fPopup
            .Invalidate()
            If bEffektMove Then
                If bAppearing Then 'Einblenden
                    Select Case MoveDirecktion
                        Case eMoveDirection.X
                            Select Case Startpunkt
                                Case eStartPosition.BottomLeft, eStartPosition.TopLeft
                                    .Left += 2
                                    StoppAnimation = .Left = Screen.PrimaryScreen.WorkingArea.Left + 10 - PositionsKorrektur.Width
                                Case eStartPosition.BottomRight, eStartPosition.TopRight
                                    .Left -= 2
                                    StoppAnimation = .Left = Screen.PrimaryScreen.WorkingArea.Right - fPopup.Size.Width - 10 - PositionsKorrektur.Width
                            End Select
                        Case eMoveDirection.Y
                            Select Case Startpunkt
                                Case eStartPosition.BottomLeft, eStartPosition.BottomRight
                                    .Top -= 1
                                    StoppAnimation = .Top + .Height = Screen.PrimaryScreen.WorkingArea.Bottom - 10 - PositionsKorrektur.Height
                                Case eStartPosition.TopLeft, eStartPosition.TopRight
                                    .Top += 1
                                    StoppAnimation = .Top = Screen.PrimaryScreen.WorkingArea.Top + 10 - PositionsKorrektur.Height
                            End Select
                    End Select

                    If StoppAnimation Then
                        'tmAnimation.Stop()
                        bAppearing = False
                        iMaxPosition = .Top
                        dMaxOpacity = .Opacity
                        If bAutoAusblenden Then tmWait.Start()
                    End If

                    Try
                        .Opacity = CDbl(IIf(bEffektTransparenz, GetOpacityBasedOnPosition(), 1))
                    Catch : End Try


                Else 'Ausblenden
                    If Not tmWait.Enabled Then

                        If bMouseIsOn Then
                            .Top = iMaxPosition
                            .Opacity = dMaxOpacity
                            'tmAnimation.Stop()
                            tmWait.Start()
                        Else
                            Select Case MoveDirecktion
                                Case eMoveDirection.X
                                    Select Case Startpunkt
                                        Case eStartPosition.BottomLeft, eStartPosition.TopLeft
                                            .Left -= 2
                                            StoppAnimation = .Right < Screen.PrimaryScreen.WorkingArea.Left
                                        Case eStartPosition.BottomRight, eStartPosition.TopRight
                                            .Left += 2
                                            StoppAnimation = .Left > Screen.PrimaryScreen.WorkingArea.Right
                                    End Select
                                Case eMoveDirection.Y
                                    Select Case Startpunkt
                                        Case eStartPosition.BottomLeft, eStartPosition.BottomRight
                                            .Top += 1
                                            StoppAnimation = .Top > Screen.PrimaryScreen.WorkingArea.Bottom - PositionsKorrektur.Width
                                        Case eStartPosition.TopLeft, eStartPosition.TopRight
                                            .Top -= 1
                                            StoppAnimation = .Bottom < Screen.PrimaryScreen.WorkingArea.Top - PositionsKorrektur.Width
                                    End Select
                            End Select

                            If StoppAnimation Then
                                'tmAnimation.Stop()
                                .TopMost = False
                                .Close()
                                bAppearing = True
                                RaiseEvent Closed(Me, EventArgs.Empty)
                            End If

                            .Opacity = CDbl(IIf(bEffektTransparenz, GetOpacityBasedOnPosition(), 1))
                        End If
                    End If
                End If
            Else
                If bAppearing Then
                    .Opacity += CDbl(IIf(bEffektTransparenz, 0.05, 1))
                    If .Opacity = 1 Then
                        'tmAnimation.Stop()
                        bAppearing = False
                        If bAutoAusblenden Then tmWait.Start()
                    End If
                Else
                    If Not tmWait.Enabled Then
                        If bMouseIsOn Then
                            fPopup.Opacity = 1
                            'tmAnimation.Stop()
                            tmWait.Start()
                        Else
                            .Opacity -= CDbl(IIf(bEffektTransparenz, 0.05, 1))
                            If .Opacity = 0 Then
                                'tmAnimation.Stop()
                                .TopMost = False
                                .Close()
                                bAppearing = True
                                RaiseEvent Closed(Me, EventArgs.Empty)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub tmWait_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmWait.Tick
        i += tmWait.Interval
        If i > ShowDelay Then
            tmWait.Stop()
            'tmAnimation.Start()
        End If
        fPopup.Invalidate()

    End Sub

    Private Sub fPopup_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles fPopup.MouseEnter
        bMouseIsOn = True
    End Sub

    Private Sub fPopup_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles fPopup.MouseLeave
        If Not bShouldRemainVisible Then bMouseIsOn = False
    End Sub

    Private Sub ctContextMenu_Closed(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosedEventArgs) Handles ctContextMenu.Closed
        bShouldRemainVisible = False
        bMouseIsOn = False
        'tmAnimation.Start()
    End Sub

End Class