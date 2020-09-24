VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   3840
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4605
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2650.436
   ScaleMode       =   0  'User
   ScaleWidth      =   4324.333
   StartUpPosition =   2  'CenterScreen
   Begin MP3Library.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   53
   End
   Begin MP3Library.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   53
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      Top             =   0
      Width           =   4020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MP3 LIBRARY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: Visual Soft Development 2000-2006"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "www.visualsoft.com +377(44)206-926 visualsoftdevelopment@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1470
      Left            =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
