VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindFile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "frmFindFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   -1680
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   12120
      Width           =   990
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   12120
      Width           =   990
   End
   Begin VB.CommandButton cmdXit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   12120
      Width           =   990
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   12120
      Width           =   990
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   12120
      Width           =   990
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Re&fresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   12120
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   2175
      Left            =   -3960
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   16680
      TabIndex        =   12
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14208
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDA972&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -360
      ScaleHeight     =   825
      ScaleWidth      =   19170
      TabIndex        =   6
      Top             =   8640
      Width           =   19200
      Begin VB.TextBox TxtFilename 
         Height          =   285
         Left            =   8280
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox txtPath 
         Height          =   315
         Left            =   5640
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbCdId 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1425
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00653A21&
         Height          =   195
         Left            =   7800
         TabIndex        =   10
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00653A21&
         Height          =   195
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CD/DVD:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00653A21&
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   8415
      Left            =   105
      Top             =   150
      Width           =   15045
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00653A21&
      Height          =   255
      Left            =   13680
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00653A21&
      Height          =   255
      Left            =   12360
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00653A21&
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00653A21&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NR."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00653A21&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   14880
      Picture         =   "frmFindFile.frx":0442
      Stretch         =   -1  'True
      Top             =   165
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000016&
      X1              =   13425
      X2              =   13425
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000016&
      X1              =   12195
      X2              =   12195
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   7095
      X2              =   7095
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   915
      X2              =   915
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   120
      Picture         =   "frmFindFile.frx":124E
      Stretch         =   -1  'True
      Top             =   165
      Width           =   14745
   End
End
Attribute VB_Name = "frmFindFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As New ADODB.Connection
Dim Item1 As ListItem
Public ItemIndex As Integer



Private Sub cmbCDId_Click()
    Dim RS As ADODB.Recordset
    Set RS = DB.Execute("select * from Shenimet where CDNr='" & cmbCdId.Text & "'")
        lstFiles.ListItems.Clear
        Dim i%
        i = 1
        Do While RS.EOF <> True
            lstFiles.ListItems.Add , , i
            lstFiles.ListItems.Item(i).SubItems(1) = RS!Emri
            lstFiles.ListItems.Item(i).SubItems(2) = Mid(RS!Skedari, 4, Len(RS!Skedari))
            lstFiles.ListItems.Item(i).SubItems(3) = RS!Madhësia
            lstFiles.ListItems.Item(i).SubItems(4) = RS!Data
            RS.MoveNext
            i = i + 1
        Loop
If cmbCdId.Text = "" Then
Exit Sub
Else
cmdDelete.Enabled = True
End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
    If cmbCdId.Text <> "" Then
        If MsgBox("Are you sure you want to delete " & cmbCdId.Text, vbCritical + vbYesNo) = vbYes Then
            DB.Execute ("delete from Shenimet where CDNr='" & cmbCdId.Text & "'")
            lstFiles.ListItems.Clear
            ComboAdd
        Else
            Exit Sub
        End If
    Else
        MsgBox "Select CD to delete", vbInformation, "CD ID.?"
        cmbCdId.SetFocus
    End If
End Sub

Private Sub cmdEllipsis_Click()
    Dim RS As ADODB.Recordset
    Set RS = DB.Execute("select distinct(CDNr) from shenimet")
    Dim MaxNo As Long
    MaxNo = 0
    Do While Not RS.EOF
        If MaxNo < Val(Mid(RS("CDNr"), 4, Len(RS("CDNr")) - 3)) Then
            MaxNo = Val(Mid(RS("CDNr"), 4, Len(RS("CDNr")) - 3))
        Else
            MaxNo = MaxNo
        End If
        RS.MoveNext
    Loop
    Me.cmbCdId.Text = "CD" & MaxNo + 1
End Sub

Private Sub cmdDelete_GotFocus()
    If cmbCdId.Text <> "" Then
        If MsgBox("Are you sure you want to delete " & cmbCdId.Text, vbCritical + vbYesNo) = vbYes Then
            DB.Execute ("delete from shenimet where CDNr='" & cmbCdId.Text & "'")
            lstFiles.ListItems.Clear
            ComboAdd
        Else
            Exit Sub
        End If
    Else
        MsgBox "Select CD to delete", vbInformation, "CD ID.?"
        cmbCdId.SetFocus
    End If
    End Sub

Private Sub cmdFind_Click()
Dim a$
a = InputBox("Type name to search?", "Search Record")
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
MsgBox "Record not found!", vbInformation, "Search!"
End If
Else
Exit Sub
End If
Set RS = Nothing
End Sub


Private Sub cmdFind_GotFocus()
Dim a$
        a = InputBox("Type name to search?", "Kërko Shënimet")
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

Private Sub cmdRefresh_Click()
    lstFiles.ListItems.Clear
    SearchFromDir txtPath, TxtFilename, lstFiles
    ComboAdd
    MpSize
End Sub

Sub ViewReport()
On Error Resume Next
Dim Lines, x As Integer
Dim RS As ADODB.Recordset
Set RS = DB.Execute("select * from Shenimet where CDNr='" & cmbCdId.Text & "'")
Dim taxRegNo As String
Open App.Path & "\List.txt" For Output As #1
If RS.EOF <> True Then
    Print #1, "File Name                                      " & "Location                      "
    Print #1, "-----------------------------------------------------------------------------------------------"
    Print #1, "=========  "
    Print #1, RS!CDNr
    Print #1, "=========  "
    Do While RS.EOF <> True
        Lines = Lines + 1
        Print #1, EvenSpace(RS!Emri, 45) & "  " & EvenSpace(RS!Skedari, 50)
        RS.MoveNext
    Loop
    Close #1
    ShellExecute Me.hwnd, "Open", App.Path & "\List.txt", "", "", SW_MAXIMIZE
Else
    MsgBox "Select CD ID", vbInformation
End If
End Sub

Private Sub cmdRefresh_GotFocus()
    lstFiles.ListItems.Clear
    SearchFromDir txtPath, TxtFilename, lstFiles
    ComboAdd
    MpSize
End Sub

Private Sub cmdReport_Click()
    ViewReport
End Sub

Function EvenSpace(TableField As String, MaxLen As Integer) As Variant
Dim strlen As Integer
    strlen = Len(TableField)
    If strlen > MaxLen Then
        EvenSpace = Left(TableField, MaxLen)
    Else
        EvenSpace = TableField & String((MaxLen - strlen), Chr(32))
    End If
End Function

Private Sub cmdReport_GotFocus()
ViewReport
End Sub

Private Sub cmdSave_Click()
    Dim RS As ADODB.Recordset
    Set RS = DB.Execute("select * from Shenimet where CDNr='" & cmbCdId.Text & "'")
    If RS.EOF = True Then
        Dim i%, x$
        i = 1
        If cmbCdId.Text <> "" Then
            Do While i < lstFiles.ListItems.Count
                DB.Execute "insert into Shenimet(CDNr,Emri,Skedari,Madhësia,data)Values('" & cmbCdId & "','" & Replace(lstFiles.ListItems.Item(i).SubItems(1), "'", "~") & "','" & Replace(lstFiles.ListItems.Item(i).SubItems(2), "'", "~") & "','" & lstFiles.ListItems.Item(i).SubItems(3) & "','" & lstFiles.ListItems.Item(i).SubItems(4) & "')"
                i = i + 1
            Loop
            MsgBox "Save sucessfully", vbInformation, "Save..!"
        Else
            MsgBox "Please type CD id", vbInformation, "CD ID?.."
        End If
    Else
        MsgBox "CD already exist", vbCritical, "Error..!"
        Exit Sub
    End If
End Sub



Private Sub cmdSave_GotFocus()
    Dim RS As ADODB.Recordset
    Set RS = DB.Execute("select * from Shenimet where CDNr='" & cmbCdId.Text & "'")
    If RS.EOF = True Then
        Dim i%, x$
        i = 1
        If cmbCdId.Text <> "" Then
            Do While i < lstFiles.ListItems.Count
                DB.Execute "insert into Shenimet(CDNr,Emri,Skedari,Madhësia,data)Values('" & cmbCdId & "','" & Replace(lstFiles.ListItems.Item(i).SubItems(1), "'", "~") & "','" & Replace(lstFiles.ListItems.Item(i).SubItems(2), "'", "~") & "','" & lstFiles.ListItems.Item(i).SubItems(3) & "','" & lstFiles.ListItems.Item(i).SubItems(4) & "')"
                i = i + 1
            Loop
            MsgBox "save sucessfully", vbInformation, "save..!"
        Else
            MsgBox "Please type CD ID", vbInformation, "CD ID?.."
        End If
    Else
        MsgBox "CD already exist", vbCritical, "Error..!"
        Exit Sub
    End If
End Sub

Private Sub cmdXit_Click()
    MDIForm1.Visible = False
    frmMe.Show
End Sub



Private Sub Command1_Click()
lstFiles.ListItems.Clear
   SearchFromDir txtPath, TxtFilename, lstFiles
  ComboAdd
  MpSize
End Sub

Private Sub Command1_GotFocus()
   SearchFromDir txtPath, TxtFilename, lstFiles
  ComboAdd
  MpSize
End Sub

Private Sub Command2_Click()
lstFiles.ListItems.Clear
Dim r&, allDrives$, JustOneDrive$, pos%, DriveType&
Dim CDfound As Integer
allDrives$ = Space$(64)
r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
allDrives$ = Left$(allDrives$, r&)
Do
pos% = InStr(allDrives$, Chr$(0))
If pos% Then
        JustOneDrive$ = Left$(allDrives$, pos%)
        allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
        DriveType& = GetDriveType(JustOneDrive$)
        If DriveType& = DRIVE_CDROM Then
           CDfound% = True
           Exit Do
        End If
      End If
  Loop Until allDrives$ = "" Or DriveType& = DRIVE_CDROM
  If CDfound% Then
        DriveLetter = UCase$(JustOneDrive$)
  Else: MsgBox "Not found CD-Rom on this computer.", vbInformation, "CD Drive not found!"
  End If
  txtPath = DriveLetter
  TxtFilename = "*.mp3"
   SearchFromDir txtPath, TxtFilename, lstFiles
  ComboAdd
  MpSize
End Sub

Private Sub Command2_GotFocus()
lstFiles.ListItems.Clear
Dim r&, allDrives$, JustOneDrive$, pos%, DriveType&
Dim CDfound As Integer
allDrives$ = Space$(64)
r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
allDrives$ = Left$(allDrives$, r&)
Do
pos% = InStr(allDrives$, Chr$(0))
If pos% Then
        JustOneDrive$ = Left$(allDrives$, pos%)
        allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
        DriveType& = GetDriveType(JustOneDrive$)
        If DriveType& = DRIVE_CDROM Then
           CDfound% = True
           Exit Do
        End If
      End If
  Loop Until allDrives$ = "" Or DriveType& = DRIVE_CDROM
  If CDfound% Then
        DriveLetter = UCase$(JustOneDrive$)
  Else: MsgBox "Not found CD-Rom on this computer.", vbInformation, "CD Drive not found!"
  End If
  txtPath = DriveLetter
  TxtFilename = "*.mp3"
   SearchFromDir txtPath, TxtFilename, lstFiles
  ComboAdd
  MpSize
End Sub






Private Sub deletem_Click()
If lstFiles.ListItems.Count = 0 Then Exit Sub
If ItemIndex <> 0 Then
    Dim Ask As String
    Ask = MsgBox("Are you sure that you want to delete '" & lstFiles.ListItems.Item(ItemIndex).Text & "'?", vbYesNo + vbInformation, "Delete record")
    If Ask = vbYes Then
        Dim DB As Database
        Dim RS As Recordset
        Set DB = OpenDatabase(App.Path + "\baza.mdb")
        Set RS = DB.OpenRecordset("Shenimet")
        RS.Index = "CDNr"
        RS.MoveFirst
        RS.Seek "=", lstFiles.ListItems.Item(ItemIndex).TaG
        If RS.AbsolutePosition = adPosUnknown Then
            MsgBox ("The record can't be found in the table!"), vbOKOnly + vbCritical, "Error"
            ItemIndex = 0
            RS.Close
            DB.Close
            Exit Sub
        Else
            RS.Delete
            lstFiles.ListItems.Remove (ItemIndex)
        End If
        RS.Close
        DB.Close
    End If
ItemIndex = 0
End If
End Sub

Private Sub Command3_Click()
  CancelDirSearch = True
End Sub

Private Sub Command3_GotFocus()
  CancelDirSearch = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        End
    End If
End Sub
Sub MpSize()
On Error GoTo err:
Dim a%
Dim b As Double
    a = 1
    If lstFiles.ListItems.Count <> 0 Then
        Do While a <> lstFiles.ListItems.Count
            b = b + Val(lstFiles.ListItems.Item(a).SubItems(3))
            a = a + 1
        Loop
        lstFiles.ListItems.Add , , ""
        lstFiles.ListItems.Item(a + 1).SubItems(3) = Round(b, 2)
    End If
Exit Sub
err:
    MsgBox err.Description
End Sub
Sub ComboAdd()
Dim RS As New ADODB.Recordset
  Set RS = DB.Execute("select distinct(CDNr) from Shenimet")
    cmbCdId.Clear
    Do While RS.EOF <> True
        cmbCdId.AddItem RS!CDNr
        RS.MoveNext
    Loop
  Set RS = Nothing
End Sub

Private Sub Form_Load()
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\baza.mdb;Persist Security Info=False"
Dim r&, allDrives$, JustOneDrive$, pos%, DriveType&
Dim CDfound As Integer
allDrives$ = Space$(64)
r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
allDrives$ = Left$(allDrives$, r&)
Do
pos% = InStr(allDrives$, Chr$(0))
If pos% Then
        JustOneDrive$ = Left$(allDrives$, pos%)
        allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
        DriveType& = GetDriveType(JustOneDrive$)
        If DriveType& = DRIVE_CDROM Then
           CDfound% = True
           Exit Do
        End If
      End If
  Loop Until allDrives$ = "" Or DriveType& = DRIVE_CDROM
  If CDfound% Then
        DriveLetter = UCase$(JustOneDrive$)
  Else: MsgBox "Not found CD-Rom on this computer.", vbInformation, "CD Drive not found!"
  End If
  With lstFiles.ColumnHeaders
    .Add , , "Nr.", 800
    .Add , , "Emri i Skedarit", 6200
    .Add , , "Shtegu", 5080
    .Add , , "Madhësia", 1200, lvwColumnLeft
    .Add , , "Data", 1470
  End With
  txtPath = DriveLetter
  TxtFilename = "*.mp3"
  ComboAdd
  MpSize
End Sub
Private Sub Form_Unload(Cancel As Integer)
  CancelDirSearch = True
End Sub
Private Sub imgCloseC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    End
End Sub
Private Sub imgMin2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.WindowState = vbMinimized
End Sub
Private Sub lstFiles_DblClick()
On Error GoTo Error2:
If lstFiles.ListItems.Count = 0 Then
Exit Sub
Else
player.Show
player.Mplayer1.Filename = frmFindFile.lstFiles.SelectedItem.SubItems(2) & frmFindFile.lstFiles.SelectedItem.SubItems(1)
End If
Error2:
End Sub

Private Sub TxtFilename_GotFocus()
  With TxtFilename
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

