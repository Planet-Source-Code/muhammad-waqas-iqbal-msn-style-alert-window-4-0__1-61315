VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Whats New"
   ClientHeight    =   5130
   ClientLeft      =   2925
   ClientTop       =   1245
   ClientWidth     =   6060
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":000C
      Top             =   2520
      Width           =   5775
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   -120
      ScaleHeight     =   1695
      ScaleWidth      =   6375
      TabIndex        =   2
      Top             =   -120
      Width           =   6375
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "........................."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   36
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Window"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   690
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   3510
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000002&
         BorderWidth     =   14
         Height          =   975
         Left            =   360
         Top             =   360
         Width           =   975
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000002&
         BorderWidth     =   9
         Height          =   855
         Left            =   4440
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000D&
         BorderWidth     =   9
         FillColor       =   &H00004080&
         Height          =   855
         Left            =   4800
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   12
         FillColor       =   &H00004080&
         Height          =   1095
         Left            =   -240
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2004 Muhammad Waqas Iqbal. All rights reserved."
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   4500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "If you find any bug please mail me"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: mwaqasiq007@hotmail.com                  pakistani_muslims@yahoo.com"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Muhammad Waqas Iqbal"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ClsGradient As New CGradient
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label7.Caption = "Version: " & App.Major & "." & App.Minor & " (Build " & App.Major & "." & App.Minor & ".0" & App.Revision & ")"
AlwaysOnTop Form2, True
With ClsGradient
        .Angle = 90
        .Color2 = &H80000002
        .Color1 = &H80000003
        .Draw Picture4
    End With
End Sub


