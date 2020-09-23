VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alert Window - By Muhammad Waqas Iqbal"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4800
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   1320
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -120
      ScaleHeight     =   825
      ScaleWidth      =   6465
      TabIndex        =   7
      Top             =   -120
      Width           =   6495
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   195
         Left            =   3480
         MouseIcon       =   "Form1.frx":08CA
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upgrade Your Product."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         MouseIcon       =   "Form1.frx":0BD4
         TabIndex        =   8
         Top             =   360
         Width           =   1860
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form1.frx":0EDE
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upgrade Your Product."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   6480
         MouseIcon       =   "Form1.frx":17A8
         TabIndex        =   10
         Top             =   360
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   3360
      Width           =   4215
      Begin VB.Label Label3 
         Caption         =   "Do not resize this picture box"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label10 
         Caption         =   "If you change its HEIGHT you have to change its top to your need in Timer1"
         Height          =   435
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   2865
      End
      Begin VB.Label Label9 
         Caption         =   "You can change its WIDTH with your need, actuall height 2655"
         Height          =   435
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   3060
      End
   End
   Begin VB.Timer TmrClose 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   1800
   End
   Begin VB.Timer TmrShow 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   1800
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   -120
      ScaleHeight     =   1755
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      Begin VB.OptionButton Option2 
         Caption         =   "Remind me later"
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Whats New"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Download Now"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label7 
         BackColor       =   &H00424242&
         Caption         =   "Label6"
         Height          =   375
         Left            =   2700
         TabIndex        =   12
         Top             =   1120
         Width           =   1300
      End
      Begin VB.Label Label6 
         BackColor       =   &H00424242&
         Caption         =   "Label6"
         Height          =   375
         Left            =   1260
         TabIndex        =   11
         Top             =   1120
         Width           =   1300
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Update version to your product is available."
         Height          =   795
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   3105
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image QuestionIcon 
      Height          =   480
      Left            =   5520
      Picture         =   "Form1.frx":1AB2
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Update2Icon 
      Height          =   480
      Left            =   5520
      Picture         =   "Form1.frx":237C
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Update1Icon 
      Height          =   480
      Left            =   5040
      Picture         =   "Form1.frx":2C46
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image ErrorIcon 
      Height          =   480
      Left            =   5040
      Picture         =   "Form1.frx":3510
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image InformationIcon 
      Height          =   480
      Left            =   5040
      Picture         =   "Form1.frx":3DDA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image ExclamationIcon 
      Height          =   480
      Left            =   5520
      Picture         =   "Form1.frx":46A4
      Top             =   2160
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    End Type
' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Dim resto As Long
Dim TaskBar As Long
Private ClsGradient As New CGradient
Private Enum TransType
    byColor
    byValue
End Enum
Private Enum IconType
    Error = 1
    Exclamation = 2
    Information = 3
    Update1 = 4
    Update2 = 5
    Question = 6
End Enum
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long




Public Function MakePopUpSize(Index As Integer)
Dim HeightInit
    Dim WindowRect As RECT
    Me.Height = Picture1.Height
    Me.Width = Picture1.Width
    HeightInit = Me.Height
    Picture3.Width = Me.Width + 100
    Picture4.Width = Me.Width + 100
    Me.Caption = "Alert Window - By Muhammad Waqas Iqbal"
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    TaskBar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.PAbj) * Screen.TwipsPerPixelX
    
    Me.Left = Screen.Width - Me.ScaleWidth - 220
    Me.Top = Screen.Height
    resto = Me.Top - ((Me.Height * (Index)) + TaskBar)
    Me.Top = resto + Me.ScaleHeight
    Me.Show
End Function


Private Sub Command1_Click()
Dim L&
If Command1.Caption = "Download" Then
MsgBox "Your will automatically be redirected to the website.", vbInformation
TmrClose.Enabled = True
If Form3.Check2.Value = vbChecked Then
FadeOut (True)
Else
End If
Else
TmrClose.Enabled = True
If Form3.Check2.Value = vbChecked Then
FadeOut (True)
Else
End If
End If
End Sub

Private Sub Command2_Click()
MsgBox "MSN style alert now in windows form with " _
& "title bar, it neither requires activeX control nor " _
& "any DLL file. You can even change sound of Alert " _
& "Window, just by changing 'PlaySoundResource' in " _
& "FormLoad event form 101 to 112. You can also add " _
& "your own sound, just add wave file in the resource " _
& "file under the heading 'WAVE'." _
& vbNewLine & vbNewLine & "Gradient effect is " _
& "now included in the version of Alert Window in prior " _
& "version you could use only one color of heading area " _
& "but in this version you can use gradient effect of " _
& "two colors of your desire. You can not only change " _
& "the colors of gradient of your of desire but also " _
& "change the angle of gardient. For example: 90 or -90. " _
& "All you need to do just change color and anlge in " _
& "FormLoad Event as simple as that." _
& vbNewLine & vbNewLine & "Another upgrade " _
& "version of Alert Window is now here, in this version " _
& "of Alert Window FadeIn and FadeOut effect of form is " _
& "included. You can set FadeIn and FadeOut duration as " _
& "you like by adding 'Step' code in ForNext loop in " _
& "FormLoad Event from 1 to 255 depends on you need. More...", vbInformation, App.Title & " - Whats New"
End Sub

Private Sub Form_Initialize()

Dim x As Long
x = InitCommonControls
End Sub

Private Sub Form_Load()
On Error GoTo err:
'Please Do not open Alert Window more than 3 at
'a time it can be harmful to your system, it could
'low windows resources
   ' Icon Type
   ' Error = 1
   ' Exclamation = 2
   ' Information = 3
   ' Update1 = 4
   ' Update2 = 5
   ' Question = 6
' You can change icon from 1 to 6
SetIcon (Form3.Label1.Caption)
Title "Upgrade Your Product."
SetOnTop (True)
'AlwaysOnTop Form1, True
If Form3.Check3.Value = vbChecked Then
TitleShadow (True)
Else
TitleShadow (False)
End If
TitleShadowColor (&H424242)

'For WindowsXP it is better to set buttonShadow = False
'But for Windows 9x you can set it to True
ButtonShadow (False)


If Form3.Check1.Value = vbChecked Then
Option1.Visible = False
Option2.Visible = False
Command1.Enabled = True
Else
Option1.Visible = True
Option2.Visible = True
End If
Label8.Caption = Label2.Caption
If Len(Label2.Caption) > 30 Then
Label2.Caption = Left$(Label2.Caption, 30) & "..."
Label5.Caption = Left$(Label5.Caption, 30) & "..."
End If
Label4.Caption = Form3.Text1.Text
Dim FormCount As Long
If Form3.Label2.Caption = 0 Then
MakePopUpSize 1
FormCount = FormCount + 1
TmrShow.Enabled = True
ElseIf Form3.Label2.Caption = 1 Then
MakePopUpSize 2
Timer3.Enabled = True
ElseIf Form3.Label2.Caption = 2 Then
MakePopUpSize 3
Timer4.Enabled = True
End If
With ClsGradient
        .Angle = 90
        .Color2 = Form3.Picture3.BackColor ' &H80000002  '&H80000015 'RGB(255, 0, 0)'RGB(215, 226, 243) '
        .Color1 = Form3.Picture2.BackColor '&H80000003 '&H80000016 'RGB(242, 245, 240)'RGB(112, 155, 208) '
        .Draw Picture4
    End With

If Form3.Check4.Value = vbChecked Then
Call PlaySoundResource(Form3.Text2.Text)
Else
End If


If Form3.Check2.Value = vbChecked Then
FadeIn (True)
Else
End If

perr:
Screen.MousePointer = 0
Exit Sub
err:
MsgBox err.Description, vbCritical
End
Resume perr:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Label2.Caption = Form3.Label2.Caption - 1
If Form3.Check2.Value = vbChecked Then
FadeOut (True)
Else
End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = False
Label1.MousePointer = 0
Label2.FontUnderline = False
Label2.MousePointer = 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = True
Label1.MousePointer = 99
End Sub

Private Sub Label2_Click()
Label2.FontUnderline = False
Label2.MousePointer = 0
'MsgBox Label4.Caption & vbNewLine & vbNewLine & "Created By: Muhammad Waqas Iqbal" & vbCrLf & "E-mail: mwaqasiq007@hotmail.com" & vbNewLine & "            pakistani_muslims@yahoo.com", vbInformation, "Alert Window - " & Label8.Caption
Form2.Show vbModal, Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label2.FontUnderline = True
Label2.MousePointer = 99
End Sub

Private Sub Option1_Click()
Command1.Caption = "Download"
Command1.Enabled = True
End Sub

Private Sub Option2_Click()
Command1.Enabled = True
Command1.Caption = "Close"
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = False
Label1.MousePointer = 0
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = False
Label1.MousePointer = 0
Label2.FontUnderline = False
Label2.MousePointer = 0
End Sub

Private Sub TmrShow_Timer()
Me.Top = Me.Top - 100
If Me.Top <= resto Then
TmrShow.Enabled = False
End If
Debug.Print Me.Top
End Sub

Private Sub TmrClose_Timer()
Me.Top = Me.Top + 100
If Me.Top > 10000 Then
TmrClose.Enabled = False
'MsgBox "Program will now close.", vbInformation
Unload Me 'End
End If
End Sub
Private Sub CreateTransparentWindowStyle(lHwnd)
'-----------------------------------
'this is used to create the window style needed
'to allow transparency to be set/altered with
'calls to SetLayeredWindowAttributes
'-----------------------------------
 On Error GoTo Err_Handler:
 
'VARIABLES:
  Dim Ret As Long
'CODE:
       'Set the window style to 'Layered'
       Ret = GetWindowLong(lHwnd, GWL_EXSTYLE)
       Ret = Ret Or WS_EX_LAYERED
       SetWindowLong lHwnd, GWL_EXSTYLE, Ret
'END CODE:
 
Exit Sub
perr:
Screen.MousePointer = vbDefault
Exit Sub
Err_Handler:
    err.Source = err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox err.Number & vbTab & err.Source & err.Description, vbCritical
    err.Clear
    Resume perr:
End Sub




Private Sub WindowTransparency(lHwnd&, TransparencyBy As TransType, _
                                      Optional Clr As Long, _
                                      Optional TransVal As Long)
On Error GoTo Err_Handler:
'---------------------------------
'sets window transparency
'proper window style must be set first
'with call to CreateTransparentWindowStyle
'that call only has to be made once for the
'life of the form.  After that, this sub
'may be called multiple times by itself
'---------------------------------
'CODE:
    'first create the window style cabable of transparancies
    Call CreateTransparentWindowStyle(lHwnd)
    
    If TransparencyBy = byColor Then
         'the color specified in Clr becomes totally transparent
         SetLayeredWindowAttributes lHwnd, Clr, 0, LWA_COLORKEY
         
    ElseIf TransparencyBy = byValue Then
         If TransVal < 0 Or TransVal > 255 Then
            'makes sure valid transparency number chosen
            '0=totally opaque    255= totally transparent
            err.Raise 2222, "Sub WindowTransparency", _
                    "must choose number between 0-255"
         End If
         SetLayeredWindowAttributes lHwnd, 0, TransVal, LWA_ALPHA
    End If
'END CODE:
Exit Sub
perr:
Exit Sub
Screen.MousePointer = vbDefault
Exit Sub
Err_Handler:
    err.Source = err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox err.Number & vbTab & err.Source & err.Description, vbCritical
Resume Next
End Sub
Private Sub SetOnTop(Value As Boolean)
 Dim i
If Value = True Then
i = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, vbNormal)
ElseIf Value = False Then
i = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, vbNormal)
End If

End Sub

Private Sub TitleShadow(Value As Boolean)
If Value = True Then
Label5.Caption = Label2.Caption
Label5.Left = Label2.Left + 25
Label5.Top = Label2.Top
Label5.Font = Label2.Font
Label5.FontSize = Label2.FontSize
Label5.FontName = Label2.FontName
Label5.Visible = True
ElseIf Value = False Then
Label5.Visible = False
End If
End Sub

Private Sub TitleShadowColor(Value As OLE_COLOR)
Label5.ForeColor = Value
End Sub
Private Sub ButtonShadow(Value As Boolean)
If Value = True Then
Label6.Visible = True
Label7.Visible = True
ElseIf Value = False Then
Label6.Visible = False
Label7.Visible = False
End If
End Sub

Private Sub SetIcon(Icon As IconType)
If Icon = 1 Then   ' Error = 1
Image1.Picture = ErrorIcon.Picture
ElseIf Icon = 2 Then   ' Exclamation = 2
Image1.Picture = ExclamationIcon.Picture
ElseIf Icon = 3 Then   ' Information = 3
Image1.Picture = InformationIcon.Picture
ElseIf Icon = 4 Then   ' Update1 = 4
Image1.Picture = Update1Icon.Picture
ElseIf Icon = 5 Then    ' Update2 = 5
Image1.Picture = Update2Icon.Picture
ElseIf Icon = 6 Then    ' Question = 6
Image1.Picture = QuestionIcon.Picture
ElseIf Icon <> 1 And Icon <> 2 And Icon <> 3 And Icon <> 4 And Icon <> 5 And Icon <> 6 Then
MsgBox "Wrong number of icon. Please enter valid icon number from 1 to 6", vbCritical, "Alert Window - Wrong Icon Number"
End If
End Sub

Private Sub FadeIn(Value As Boolean)
Dim L As Long
If Value = True Then
If WinVersion = "Windows XP" Then
WindowTransparency hWnd, byValue, , 0
Show
For L = 0 To 255 'Step 1 'You can also change fade in and fade out duration
     WindowTransparency hWnd, byValue, , L
     DoEvents
     Refresh
Next L
End If
End If
End Sub
Private Sub FadeOut(Value As Boolean)
 Dim L&
If Value = True Then
If WinVersion = "Windows XP" Then
  For L = 255 To 0 Step -1
     WindowTransparency hWnd, byValue, , L
     DoEvents
     Refresh
  Next L
  End If
  End If
End Sub
Private Sub Title(Text As String)
Dim msg
If Text = "" Then
    MsgBox "Please enter some sting in Title.", vbInformation
    Label2.Caption = "No Title"
ElseIf Text <> Label2.Caption Then
    msg = MsgBox("You have changed title manually; without entring title caption in Title Fuction. Please Enter Title in Title Function properly to avoid this message next time. Are you sure want to change the title.", vbQuestion + vbYesNo, "Alert Window - Change in Title found")
        If msg = vbYes Then
          Label2.Caption = Text
        Else
          Exit Sub
        End If
Else
    Label2.Caption = Text
End If
End Sub

Private Sub Timer3_Timer()
Me.Top = Me.Top - 100
If Me.Top < 3300 Then
Timer3.Enabled = False
End If
Debug.Print Me.Top
End Sub

Private Sub Timer4_Timer()
Me.Top = Me.Top - 100
If Me.Top < 650 Then
Timer4.Enabled = False
End If
Debug.Print Me.Top
End Sub
Private Function WinVersion() As String
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)


    With osinfo


        Select Case .dwPlatformId
            Case 1


            If .dwMinorVersion = 0 Then
                WinVersion = "Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                WinVersion = "Windows 98"
            End If
            Case 2


            If .dwMajorVersion = 3 Then
                WinVersion = "Windows NT 3.51"
            ElseIf .dwMajorVersion = 4 Then
                WinVersion = "Windows NT 4.0"
            ElseIf .dwMajorVersion >= 5 Then
                WinVersion = "Windows XP"
            End If
            Case Else
            WinVersion = "Failed"
        End Select
End With
End Function


