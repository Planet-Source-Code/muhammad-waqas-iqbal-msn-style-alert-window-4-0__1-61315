VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6600
   ClientLeft      =   300
   ClientTop       =   705
   ClientWidth     =   4695
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   4695
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   6000
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   1
      Orientation     =   2
   End
   Begin VB.Frame Frame3 
      Caption         =   "Gradient Colors"
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   4455
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   2880
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   20
         Top             =   240
         Width           =   1215
         Begin VB.CommandButton Command2 
            Caption         =   "Select"
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Select"
            Height          =   375
            Left            =   0
            TabIndex        =   21
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000002&
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   1395
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000003&
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   1395
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Color2"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Color1"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Form3.frx":000C
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Layout"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   4215
         TabIndex        =   2
         Top             =   240
         Width           =   4215
         Begin VB.Timer Timer1 
            Interval        =   200
            Left            =   3720
            Top             =   1200
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2760
            TabIndex        =   27
            Text            =   "106"
            Top             =   1620
            Width           =   615
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Play Sound"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show Title Show"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Enable Fade-In and Fade-Out Effect"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Hide Options Button "
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   3375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Update1 "
            Height          =   195
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Question "
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Update2 "
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Information "
            Height          =   195
            Left            =   2760
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Exclamation "
            Height          =   255
            Left            =   1200
            TabIndex        =   4
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Error"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Enter Resource Sound Number:"
            Height          =   195
            Left            =   360
            TabIndex        =   26
            Top             =   1680
            Width           =   2265
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Alert Window"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000

Private Sub Check4_Click()
If Check4.Value = vbChecked Then
Text2.Enabled = True
Else
Text2.Enabled = False
End If
End Sub

Private Sub Command1_Click()

Dim AlertWindow
If Label2.Caption = 3 Then
    MsgBox "Not more than three Alert Window can open.", vbInformation
    Exit Sub
ElseIf Label2.Caption < 0 Then
    MsgBox "An Unexpected error occured. Alert Window will now close.", vbCritical
    End
Else
    Set AlertWindow = New Form1
    AlertWindow.Show
End If
Label2.Caption = Label2.Caption + 1
    End Sub

Private Sub Command2_Click()
On Error GoTo err:
cd.DialogTitle = "Alert Window - Select Color"
cd.ShowColor
Picture2.BackColor = cd.Color
err:
If cd.CancelError Then
Exit Sub

End If
End Sub

Private Sub Command3_Click()
On Error GoTo err:
cd.DialogTitle = "Alert Window - Select Color"
cd.ShowColor
Picture3.BackColor = cd.Color
err:
If cd.CancelError Then
Exit Sub
End If
End Sub

Private Sub Command4_Click()
Form2.Show vbModal, Me
End Sub

Private Sub Form_Initialize()
Dim x As Long
x = InitCommonControls
End Sub

Private Sub Form_Load()
Dim style As Long
    style = GetWindowLong(Text2.hWnd, GWL_STYLE)
    SetWindowLong Text2.hWnd, GWL_STYLE, style Or ES_NUMBER
    
 Label2.Caption = 0
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Option1_Click()
Label1.Caption = 1
Text2.Text = "114"
End Sub

Private Sub Option2_Click()
Label1.Caption = 2
Text2.Text = "117"
End Sub

Private Sub Option3_Click()
Label1.Caption = 3
Text2.Text = "116"
End Sub

Private Sub Option4_Click()
Label1.Caption = 4
End Sub

Private Sub Option5_Click()
Label1.Caption = 5
End Sub

Private Sub Option6_Click()
Label1.Caption = 6
Text2.Text = "113"
End Sub

Private Sub Picture2_Click()
Picture2.BackColor = &H80000003
End Sub

Private Sub Picture3_Click()
Picture3.BackColor = &H80000002
End Sub
Private Function Check()
 
 'Get keys pressed
 
 If Text2.SelStart Then Timer1.Enabled = False
 
 With Text2
  
        '//Normal//
        If GetAsyncKeyState(vbKey0) Then
            .Text = .Text + "0"
        End If
        If GetAsyncKeyState(vbKey1) Then
            .Text = .Text + "1"
        End If
        If GetAsyncKeyState(vbKey2) Then
            .Text = .Text + "2"
        End If
        If GetAsyncKeyState(vbKey3) Then
            .Text = .Text + "3"
        End If
        If GetAsyncKeyState(vbKey4) Then
            .Text = .Text + "4"
        End If
        If GetAsyncKeyState(vbKey5) Then
            .Text = .Text + "5"
        End If
        If GetAsyncKeyState(vbKey6) Then
            .Text = .Text + "6"
        End If
        If GetAsyncKeyState(vbKey7) Then
            .Text = .Text + "7"
        End If
        If GetAsyncKeyState(vbKey8) Then
            .Text = .Text + "8"
        End If
        If GetAsyncKeyState(vbKey9) Then
            .Text = .Text + "9"
        End If
        
            '//Keypads//
             If GetAsyncKeyState(vbKeyNumpad0) Then
            .Text = .Text + "0"
        End If
        If GetAsyncKeyState(vbKeyNumpad1) Then
            .Text = .Text + "1"
        End If
        If GetAsyncKeyState(vbKeyNumpad2) Then
            .Text = .Text + "2"
        End If
        If GetAsyncKeyState(vbKeyNumpad3) Then
            .Text = .Text + "3"
        End If
        If GetAsyncKeyState(vbKeyNumpad4) Then
            .Text = .Text + "4"
        End If
        If GetAsyncKeyState(vbKeyNumpad5) Then
            .Text = .Text + "5"
        End If
        If GetAsyncKeyState(vbKeyNumpad6) Then
            .Text = .Text + "6"
        End If
        If GetAsyncKeyState(vbKeyNumpad7) Then
            .Text = .Text + "7"
        End If
        If GetAsyncKeyState(vbKeyNumpad8) Then
            .Text = .Text + "8"
        End If
        If GetAsyncKeyState(vbKeyNumpad9) Then
            .Text = .Text + "9"
        End If
        
        On Error Resume Next
        If GetAsyncKeyState(vbKeyBack) Then
           .Text = Left(.Text, Len(.Text) - 1)
        End If

    End With
    End Function

Private Sub Text2_Change()
'  If Text2.SelStart Then TmrShow.Enabled = True
        If Len(Text2.Text) = 4 Then
        Text2.Text = Left$(Text2.Text, 3)
        MsgBox "Only three Numbes are allowed.", vbInformation
        End If
        
End Sub

Private Sub Timer1_Timer()
    Call Check
End Sub

