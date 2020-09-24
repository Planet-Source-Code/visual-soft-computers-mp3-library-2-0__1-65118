VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmshenimet 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   9255
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   2160
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   16113
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmshenimet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Header As ColumnHeader
Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "CDNr"
    Header.Width = ListView1.Width * 0.05
Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Emri i Skedarit"
    Header.Width = ListView1.Width * 0.4
Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Lokacioni"
    Header.Width = ListView1.Width * 0.3
    Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Madhësia"
    Header.Width = ListView1.Width * 0.3
    Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Data"
    Header.Width = ListView1.Width * 0.3
ItemIndex = 0
End Sub

