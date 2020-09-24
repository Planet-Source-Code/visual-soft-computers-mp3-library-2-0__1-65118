VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00EDA972&
   Caption         =   "MP3 Library 2.0"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1482
      ButtonWidth     =   2037
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      ImageList       =   "imlNormal"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add CD"
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Delete"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Edit"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Search Record"
            Object.Tag             =   ""
            ImageIndex      =   22
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   9600
         Picture         =   "MDIForm1.frx":2A8B2
         ScaleHeight     =   540
         ScaleWidth      =   4005
         TabIndex        =   1
         Top             =   120
         Width           =   3999
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31A06
            Key             =   "s_Key1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31FA0
            Key             =   "s_Key2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3253A
            Key             =   "s_Key3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":32AD4
            Key             =   "s_Key4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3306E
            Key             =   "s_Key5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":33608
            Key             =   "s_Key6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":33BA2
            Key             =   "s_Key7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3413C
            Key             =   "s_Key8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":346D6
            Key             =   "s_Key9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":34C70
            Key             =   "s_Key10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3520A
            Key             =   "s_Key11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":357A4
            Key             =   "s_Key12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":35D3E
            Key             =   "s_Key13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":362D8
            Key             =   "s_Key14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":36872
            Key             =   "s_Key15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":36E0C
            Key             =   "s_Key16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":373A6
            Key             =   "s_Key17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":37940
            Key             =   "s_Key18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":37EDA
            Key             =   "s_Key19"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlNormal 
      Left            =   4320
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   22
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":38474
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":390C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":39D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3A96A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3B5BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3C20E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3CE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3DAB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3E704
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3F356
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3FFA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":40BFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4184C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4249E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":430F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":43D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":44994
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":455E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":46238
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":46E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":47ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4872E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSkedar 
      Caption         =   "File"
      Tag             =   "Skedar"
      Begin VB.Menu mnuSkedarShtoCD 
         Caption         =   "Add CD"
         Tag             =   "Shto CD|#s_Key16"
      End
      Begin VB.Menu mnuSkedarRuajKombinimin 
         Caption         =   "Save"
         Tag             =   "Ruaj Kombinimin|#s_Key17"
      End
      Begin VB.Menu mnuSkedarFshijeKombinimin 
         Caption         =   "Delete"
         Tag             =   "Fshije Kombinimin|#s_Key18"
      End
      Begin VB.Menu mnuSkedarSep1 
         Caption         =   "-"
         Tag             =   "-"
      End
      Begin VB.Menu mnuSkedarDalja 
         Caption         =   "Exit"
         Tag             =   "Dalja|#s_Key6"
      End
   End
   Begin VB.Menu mnuLista 
      Caption         =   "List"
      Tag             =   "Lista"
      Begin VB.Menu mnuListaShikoListn 
         Caption         =   "Preview List"
         Tag             =   "Shiko Listën|#s_Key11"
      End
      Begin VB.Menu ppp 
         Caption         =   "Print List"
      End
   End
   Begin VB.Menu mnuKrkimi 
      Caption         =   "Search"
      Tag             =   "Kërkimi"
      Begin VB.Menu mnuKërkimiKrkoShnimin 
         Caption         =   "Search record"
         Tag             =   "Kërko Shënimin|#s_Key8"
      End
   End
   Begin VB.Menu mnuNdihm 
      Caption         =   "Help"
      Tag             =   "Ndihmë"
      Begin VB.Menu mnuNdihmëRrethProgramit 
         Caption         =   "About..."
         Tag             =   "Rreth Programit|#s_Key15"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents MenuEvents As CEvents
Attribute MenuEvents.VB_VarHelpID = -1
Dim DB As New ADODB.Connection
Dim Item1 As ListItem
Public ItemIndex As Integer
Public Sub UnloadAllForms()
Dim Form As Form
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End Sub


Private Sub scan_Click()
frmFindFile.Show
frmFindFile.cmdRefresh.SetFocus
End Sub

Private Sub Command1_Click()
objMenuEx.MenuDesigner Me.hwnd
End Sub

Private Sub mnuKërkimiKrkoShnimin_Click()
Dim a$
a = InputBox("Type name to search?", "Search record")
If a <> "" Then
Dim RS As New ADODB.Recordset
Dim x$
x = "select * from Shenimet where Emri like '%" & a & "%'"
Set RS = DB.Execute(x)
If RS.EOF <> True Then
frmFound.Show
Dim i%
i = 1
Do While RS.EOF <> True
frmFound.lstFind.ListItems.Add , , i
frmFound.lstFind.ListItems.Item(i).SubItems(1) = RS!CDNr
frmFound.lstFind.ListItems.Item(i).SubItems(2) = RS!Emri
frmFound.lstFind.ListItems.Item(i).SubItems(3) = Mid(RS!Skedari, 4, Len(RS!Skedari))
RS.MoveNext
i = i + 1
Loop
Else
MsgBox "Record not found", vbInformation, "Search!"
End If
Else
Exit Sub
End If
Set RS = Nothing

End Sub

Private Sub mnuListaShikoListn_Click()
        frmmain.cmdRead.SetFocus
        frmFindFile.Hide
End Sub

Private Sub mnuNdihmëRrethProgramit_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuSkedarDalja_Click()
End
End Sub

Private Sub mnuSkedarFshijeKombinimin_Click()
        frmFindFile.cmdDelete.SetFocus
End Sub

Private Sub mnuSkedarRuajKombinimin_Click()
        frmFindFile.cmdSave.SetFocus
End Sub

Private Sub mnuSkedarShtoCD_Click()
        frmFindFile.Show
        frmFindFile.cmdRefresh.SetFocus
End Sub

Private Sub ppp_Click()
dr1.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1:
        frmFindFile.Show
        frmFindFile.cmdRefresh.SetFocus
        Case 2: frmFindFile.Show
        frmFindFile.cmdSave.SetFocus
        Case 3: frmFindFile.Show
        frmFindFile.cmdDelete.SetFocus
        Case 4: frmmain.Show
        frmmain.cmdRead.SetFocus
        frmFindFile.Hide
        Case 5:
Dim a$
a = InputBox("Type name to search?", "Search record")
If a <> "" Then
Dim RS As New ADODB.Recordset
Dim x$
x = "select * from Shenimet where Emri like '%" & a & "%'"
Set RS = DB.Execute(x)
If RS.EOF <> True Then
frmFound.Show
Dim i%
i = 1
Do While RS.EOF <> True
frmFound.lstFind.ListItems.Add , , i
frmFound.lstFind.ListItems.Item(i).SubItems(1) = RS!CDNr
frmFound.lstFind.ListItems.Item(i).SubItems(2) = RS!Emri
frmFound.lstFind.ListItems.Item(i).SubItems(3) = Mid(RS!Skedari, 4, Len(RS!Skedari))
RS.MoveNext
i = i + 1
Loop
Else
MsgBox "Record not found", vbInformation, "Search!"
End If
Else
Exit Sub
End If
Set RS = Nothing
        Case 6: frmAbout.Show 1
        Case 7: End
        UnloadAllForms
    End Select
End Sub

Private Sub MDIForm_Load()
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\baza.mdb;Persist Security Info=False"
End Sub

