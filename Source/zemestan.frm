VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form MSBG 
   BorderStyle     =   0  'None
   Caption         =   "MSBG"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   Icon            =   "zemestan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   741
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   3600
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash2 
      Height          =   10470
      Left            =   4320
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   6660
      _cx             =   4206051
      _cy             =   4212772
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "FFFFFF"
      SWRemote        =   ""
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   11520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      _cx             =   4221397
      _cy             =   4214624
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "NoBorder"
      DeviceFont      =   -1  'True
      EmbedMovie      =   -1  'True
      BGColor         =   "000000"
      SWRemote        =   ""
   End
End
Attribute VB_Name = "MSBG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_TEST = &H4
Private Const DISP_CHANGE_SUCCESSFUL = 0

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private Type DevMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public W As Integer, H As Integer
Dim Tnid As NOTIFYICONDATA
Dim msg As Long

Public Function ShamsiDate() As String
    On Error Resume Next
    Dim YearEqual(2, 2) As Integer
    Dim AddOneDay As Boolean
    Dim AddFarDay As Boolean
    Dim AddToDays As Byte, FarDay As Byte
    Dim ThisDay As Byte
    Dim ThisMonth As Byte
    Dim ThisYear As Integer
    Dim YearDif1 As Integer, YearDif2 As Integer
    Dim TestRange1 As Integer, TestRange2 As Integer
    Dim FarsiRange1 As Integer, FarsiRange2 As Integer
    Dim p As Integer
    Dim CurM As String, CurD As String
    Dim sYear As Variant, sMonth As Variant, sDay As Variant
Rem -------
    YearEqual(1, 1) = 1997
    YearEqual(1, 2) = 1998
    YearEqual(2, 1) = 1376
    YearEqual(2, 2) = 1377
    ThisDay = Day(Date)
    ThisMonth = Month(Date)
    ThisYear = Year(Date)
    YearDif1 = ThisYear - 1997
    YearDif2 = ThisYear - 1998
    TestRange1 = 1996 - (100 * 4)
    TestRange2 = 1996 + (100 * 4)
    FarsiRange1 = 1375 - (100 * 4)
    FarsiRange2 = 1375 + (100 * 4)
    AddOneDay = False
Rem -------
    For p = TestRange1 To TestRange2 Step 4
        If ThisYear = p Then
            AddOneDay = True
            Exit For
        End If
    Next
    If AddOneDay Then AddToDays = 1 Else AddToDays = 0
Rem -------
    If ((ThisMonth = 3 And ThisDay < 20 + AddToDays) Or (ThisMonth < 3)) Then
        YearDif1 = YearDif1 - 1
    End If
Rem -------
    If (ThisYear Mod 2 <> 0 And ((ThisMonth = 3 And ThisDay > (20 - AddToDays)) Or (ThisMonth > 4))) Then
        CurrentYear = YearEqual(2, 1) + YearDif1
    Else
        CurrentYear = YearEqual(2, 1) + YearDif2
        For p = FarsiRange1 To FarsiRange2 Step 4
            If CurrentYear = p Then AddFarDay = True: Exit For
        Next p
        If AddFarDay Then FarDay = 1 Else FarDay = 0
        If ((ThisMonth = 3 And ThisDay > 20 - (AddToDays) + FarDay) Or ThisMonth > 3) Then
            CurrentYear = CurrentYear + 1
        End If
    End If
    If AddToDays = 1 Then FarDay = 0
Rem -------
    Select Case ThisMonth
        Case 1
            If ThisDay < (21 - FarDay) Then
                CurrentMonth = 10
                CurrentDay = (ThisDay + 10) + FarDay
            Else
                CurrentMonth = 11
                CurrentDay = (ThisDay - 20) + FarDay
            End If
        Case 2
            If ThisDay < (20 - FarDay) Then
                CurrentMonth = 11
                CurrentDay = (ThisDay + 11) + FarDay
            Else
                CurrentMonth = 12
                CurrentDay = (ThisDay - 19) + FarDay
            End If
        Case 3
            If ThisDay < (21 - AddToDays) Then
                CurrentMonth = 12
                CurrentDay = (ThisDay + 9) + AddToDays + FarDay
            Else
                CurrentMonth = 1
                CurrentDay = (ThisDay - 20) + AddToDays
            End If
        Case 4
           If ThisDay < (21 - AddToDays) Then
                CurrentMonth = 1
                CurrentDay = (ThisDay + 11) + AddToDay
            Else
                CurrentMonth = 2
                CurrentDay = (ThisDay - 20) + AddToDays
            End If
        Case 5
            If ThisDay < (22 - AddToDays) Then
                CurrentMonth = 2
                CurrentDay = (ThisDay + 10) + AddToDay
            Else
                CurrentMonth = 3
                CurrentDay = (ThisDay - 21) + AddToDays
            End If
        Case 6
            If ThisDay < (22 - AddToDays) Then
                CurrentMonth = 3
                CurrentDay = (ThisDay + 10) + AddToDay
            Else
                CurrentMonth = 4
                CurrentDay = (ThisDay - 21) + AddToDays
            End If
        Case 7
            If ThisDay < (23 - AddToDays) Then
                CurrentMonth = 4
                CurrentDay = (ThisDay + 9) + AddToDay
            Else
                CurrentMonth = 5
                CurrentDay = (ThisDay - 22) + AddToDays
            End If
        Case 8
            If ThisDay < (23 - AddToDays) Then
                CurrentMonth = 5
                CurrentDay = (ThisDay + 9) + AddToDays
            Else
                CurrentMonth = 6
                CurrentDay = (ThisDay - 22) + AddToDays
            End If
        Case 9
            If ThisDay < (23 - AddToDays) Then
                CurrentMonth = 6
                CurrentDay = (ThisDay + 9) + AddToDays
            Else
                CurrentMonth = 7
                CurrentDay = (ThisDay - 22) + AddToDays
            End If
        Case 10
            If ThisDay < (23 - AddToDays) Then
                CurrentMonth = 7
                CurrentDay = (ThisDay + 8) + AddToDays
            Else
                CurrentMonth = 8
                CurrentDay = (ThisDay - 22) + AddToDays
            End If
        Case 11
            If ThisDay < (22 - AddToDays) Then
                CurrentMonth = 8
                CurrentDay = (ThisDay + 9) + AddToDays
            Else
                CurrentMonth = 9
                CurrentDay = (ThisDay - 21) + AddToDays
            End If
        Case 12
            If ThisDay < (22 - AddToDays) Then
                CurrentMonth = 9
                CurrentDay = (ThisDay + 9) + AddToDays
            Else
                CurrentMonth = 10
                CurrentDay = (ThisDay - 21) + AddToDays
            End If
    End Select
Rem -------
    CurM = Trim(Str(CurrentMonth))
    CurD = Trim(Str(CurrentDay))
Rem -------
    If CurrentMonth < 10 Then CurM = "0" & Trim(Str(CurrentMonth))
    If CurrentDay < 10 Then CurD = "0" & Trim(Str(CurrentDay))
    ShamsiDate = Trim(Str(CurrentYear)) & "/" & CurM & "/" & CurD
End Function

Private Sub ChangeDisplay(WResolution As Integer, HResolution As Integer, Color As Integer)
    Dim DevM As DevMODE

    EnumDisplaySettings 0, 0, DevM
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = WResolution
    DevM.dmPelsHeight = HResolution
    DevM.dmBitsPerPel = Color
    If ChangeDisplaySettings(DevM, CDS_TEST) = DISP_CHANGE_SUCCESSFUL Then ChangeDisplaySettings DevM, 0
End Sub

Private Sub Form_Load()
    W = Screen.Width / Screen.TwipsPerPixelX
    H = Screen.Height / Screen.TwipsPerPixelY
    ChangeDisplay 1024, 768, 32
    ShockwaveFlash1.Movie = App.Path + "\gui.mnu"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X / Screen.TwipsPerPixelX = WM_LBUTTONDBLCLK Then
        Tnid.cbSize = Len(TNotifyIconData)
        Tnid.hwnd = Picture1.hwnd
        Tnid.uId = 1
        Shell_NotifyIcon NIM_DELETE, Tnid
        ShowWindow MSBG.hwnd, SW_SHOW
        ShockwaveFlash1.SetFocus
    End If
End Sub

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
    If args <> "SlideShow" Then Timer2.Enabled = False
    Select Case command
        Case "product"
            Text1.Visible = False
            Timer1.Enabled = False
            ShockwaveFlash2.Visible = True
            ShockwaveFlash2.Movie = App.Path + "\store.db"
            Select Case args
                Case "case"
                    ShockwaveFlash2.GotoFrame (1)
                Case "speaker"
                    ShockwaveFlash2.GotoFrame (63)
                Case "keyboard/mouse"
                    ShockwaveFlash2.GotoFrame (85)
                Case "set"
                    ShockwaveFlash2.GotoFrame (91)
                Case "other"
                    ShockwaveFlash2.GotoFrame (101)
            End Select
        Case "about"
            Text1.Visible = False
            Timer1.Enabled = False
        Case "quit"
            ChangeDisplay W, H, 32
            End
        Case "Navigator"
            Select Case args
                Case "Previous"
                    If ShockwaveFlash2.FrameNum > 1 Then ShockwaveFlash2.Back
                Case "Next"
                    If ShockwaveFlash2.FrameNum < 101 Then ShockwaveFlash2.Forward
                Case "StepBackward"
                    If ShockwaveFlash2.FrameNum > 100 Then
                        ShockwaveFlash2.GotoFrame (91)
                        GoTo Skip
                    End If
                    If ShockwaveFlash2.FrameNum > 90 Then
                        ShockwaveFlash2.GotoFrame (85)
                        GoTo Skip
                    End If
                    If ShockwaveFlash2.FrameNum > 84 Then
                        ShockwaveFlash2.GotoFrame (63)
                        GoTo Skip
                    End If
                    If ShockwaveFlash2.FrameNum > 62 Then
                        ShockwaveFlash2.GotoFrame (1)
                        GoTo Skip
                    End If
                Case "StepForward"
                    If ShockwaveFlash2.FrameNum < 63 Then
                        ShockwaveFlash2.GotoFrame (63)
                        GoTo Skip
                    End If
                    If ShockwaveFlash2.FrameNum < 85 Then
                        ShockwaveFlash2.GotoFrame (85)
                        GoTo Skip
                    End If
                    If ShockwaveFlash2.FrameNum < 91 Then
                        ShockwaveFlash2.GotoFrame (91)
                        GoTo Skip
                    End If
                    If ShockwaveFlash2.FrameNum < 101 Then
                        ShockwaveFlash2.GotoFrame (101)
                        GoTo Skip
                    End If
                Case "ZoomIn"
                    ShockwaveFlash2.Zoom 50
                Case "ZoomOut"
                    ShockwaveFlash2.Zoom 200
                Case "Refresh"
                    ShockwaveFlash2.Zoom 0
                Case "SlideShow"
                    If Timer2.Enabled = False Then
                        ShockwaveFlash2.GotoFrame 1
                        Timer2.Enabled = True
                    Else
                        Timer2.Enabled = False
                    End If
            End Select
        Case "HideBrowser"
            Timer2.Enabled = False
            ShockwaveFlash2.Visible = False
            Timer1.Enabled = True
            Text1.Visible = True
        Case "sYSTEMtRAY"
            Picture1.Picture = MSBG.Icon
            ShowWindow MSBG.hwnd, SW_HIDE
            Tnid.cbSize = Len(TNotifyIconData)
            Tnid.ucallbackMessage = WM_LBUTTONDOWN
            Tnid.hwnd = Picture1.hwnd
            Tnid.uId = 1
            Tnid.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
            Tnid.hIcon = Picture1.Picture
            Tnid.szTip = "Double Click for shown again..." & Chr$(0)
            Shell_NotifyIcon NIM_DELETE, Tnid
            Shell_NotifyIcon NIM_ADD, Tnid
    End Select
Skip:
End Sub

Private Sub Timer1_Timer()
    Text1.Text = ShamsiDate()
End Sub

Private Sub Timer2_Timer()
    ShockwaveFlash2.Forward
End Sub
