VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Formmain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   10320
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   10320
      Width           =   1695
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert new record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   10320
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear ListView"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   10320
      Width           =   1695
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   8655
      Left            =   135
      TabIndex        =   4
      Top             =   480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15266
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
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "CD Nr."
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
      TabIndex        =   14
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label9 
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
      TabIndex        =   13
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label8 
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
      TabIndex        =   12
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label7 
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
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Left            =   13560
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   135
      Picture         =   "Formmain.frx":2A8B2
      Stretch         =   -1  'True
      Top             =   165
      Width           =   14745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   930
      X2              =   930
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   7110
      X2              =   7110
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000016&
      X1              =   12210
      X2              =   12210
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000016&
      X1              =   13440
      X2              =   13440
      Y1              =   120
      Y2              =   480
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
      Left            =   255
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Emri i Skedarit"
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
      Left            =   1095
      TabIndex        =   8
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lokacioni"
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
      Left            =   7215
      TabIndex        =   7
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Madhësia"
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
      Left            =   12375
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
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
      Left            =   13575
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   9015
      Left            =   120
      Top             =   150
      Width           =   15045
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'you have to declare the variables you use
Public ItemIndex As Integer 'index of a row in listview
Private Sub Read()
If lstFiles.ListItems.Count > 0 Then cmdClear_Click
Dim DB As Database
Dim RS As Recordset
Set DB = OpenDatabase(App.Path + "\baza.mdb")
Set RS = DB.OpenRecordset("Shenimet")
    Dim a As Integer
    a = 1
    Do Until RS.EOF

        lstFiles.ListItems.Add , , RS!CDNr
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Emri
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Skedari
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Madhësia
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Data
        lstFiles.ListItems(a).TaG = RS!id
        a = a + 1
        RS.MoveNext
    Loop
RS.Close
DB.Close
End Sub

Private Sub cmdDelete_GotFocus()
If lstFiles.ListItems.Count = 0 Then Exit Sub
If ItemIndex <> 0 Then
    Dim Ask As String
    Ask = MsgBox("Are you sure you want to delete record", vbYesNo + vbInformation, "Confirm delete")
    If Ask = vbYes Then
        Dim DB As Database
        Dim RS As Recordset
        Set DB = OpenDatabase(App.Path + "\baza.mdb")
        ' Open recordset (table: People) in database DB
        Set RS = DB.OpenRecordset("Shenimet")
        
        'we will seek the record with ID equal to item's tag
        'first we set the table's index (this is the ID)
        'we need this because otherwise we can't use the seek function
        RS.Index = "ID"
         
        If RS.RecordCount = 0 Then
        Exit Sub
        Else
        RS.MoveFirst
        End If
            
        'here we check if the seek functions result is equal to the item's tag.
        'the item's tag contains the ID from the table.
        'we stored it when we were reading from the table
        RS.Seek "=", lstFiles.ListItems.Item(ItemIndex).TaG
    
        If RS.NoMatch Then
            'if there was no match (the ID couldn't be found in the table)
            MsgBox ("Record cannot found on the table!"), vbOKOnly + vbCritical, "Error"
            
            
            'we reset itemindex to 0, this means that nothing is selected in lstFiles
            ItemIndex = 0
            
            'close recordset and database
            RS.Close
            DB.Close
            
            Exit Sub
        Else
            'there was a match! we will now delete the record from the database
            RS.Delete
            
            'and we will delete the row from the listview
            lstFiles.ListItems.Remove (ItemIndex)
        End If
           
        'close recordset and database
        RS.Close
        DB.Close
    
    End If

'we reset itemindex to 0, this means that nothing is selected in lstFiles
ItemIndex = 0
    
End If



End Sub

Private Sub cmdInsert_Click()
' Now we want to insert a new record to the table People.
' We are going to show another form to do this.
frmInsert.Show

'we have to reset ItemIndex, otherwise in some cases the ItemIndex will be remembered
ItemIndex = 0

End Sub

Private Sub cmdDelete_Click()
If lstFiles.ListItems.Count = 0 Then Exit Sub
If ItemIndex <> 0 Then
    Dim Ask As String
    Ask = MsgBox("Are you sure you want to delete record", vbYesNo + vbInformation, "Confirm delete")
    If Ask = vbYes Then
        Dim DB As Database
        Dim RS As Recordset
        Set DB = OpenDatabase(App.Path + "\baza.mdb")
        ' Open recordset (table: People) in database DB
        Set RS = DB.OpenRecordset("Shenimet")
        
        'we will seek the record with ID equal to item's tag
        'first we set the table's index (this is the ID)
        'we need this because otherwise we can't use the seek function
        RS.Index = "ID"
         
        'we move the recordpointer to the first record, this way we can seek the whole table
        RS.MoveFirst
            
        'here we check if the seek functions result is equal to the item's tag.
        'the item's tag contains the ID from the table.
        'we stored it when we were reading from the table
        RS.Seek "=", lstFiles.ListItems.Item(ItemIndex).TaG
    
        If RS.NoMatch Then
            'if there was no match (the ID couldn't be found in the table)
            MsgBox ("Record cannot found on the table!"), vbOKOnly + vbCritical, "Error"
            
            
            'we reset itemindex to 0, this means that nothing is selected in lstFiles
            ItemIndex = 0
            
            'close recordset and database
            RS.Close
            DB.Close
            
            Exit Sub
        Else
            'there was a match! we will now delete the record from the database
            RS.Delete
            
            'and we will delete the row from the listview
            lstFiles.ListItems.Remove (ItemIndex)
        End If
           
        'close recordset and database
        RS.Close
        DB.Close
    
    End If

'we reset itemindex to 0, this means that nothing is selected in lstFiles
ItemIndex = 0
    
End If


            
End Sub

Private Sub cmdClear_Click()
lstFiles.ListItems.Clear
ItemIndex = 0
End Sub

Private Sub cmdRead_Click()
If lstFiles.ListItems.Count > 0 Then cmdClear_Click
Dim DB As Database
Dim RS As Recordset
Set DB = OpenDatabase(App.Path + "\baza.mdb")
Set RS = DB.OpenRecordset("Shenimet")
    Dim a As Integer
    a = 1
    Do Until RS.EOF

        lstFiles.ListItems.Add , , RS!id
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Emri
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Skedari
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Madhësia
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Data
        lstFiles.ListItems(a).TaG = RS!id
        a = a + 1
        RS.MoveNext
    Loop
RS.Close
DB.Close
End Sub

Private Sub cmdRead_GotFocus()
If lstFiles.ListItems.Count > 0 Then cmdClear_Click
Dim DB As Database
Dim RS As Recordset
Set DB = OpenDatabase(App.Path + "\baza.mdb")
Set RS = DB.OpenRecordset("Shenimet")
    Dim a As Integer
    a = 1
    Do Until RS.EOF

        lstFiles.ListItems.Add , , RS!CDNr
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Emri
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Skedari
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Madhësia
        lstFiles.ListItems(a).ListSubItems.Add , , RS!Data
        lstFiles.ListItems(a).TaG = RS!id
        a = a + 1
        RS.MoveNext
    Loop
RS.Close
DB.Close
End Sub

Private Sub Form_Load()
Read
Me.Top = 0
Me.Left = 0
  With lstFiles.ColumnHeaders
    .Add , , "Nr.", 800
    .Add , , "Emri i Skedarit", 6200
    .Add , , "Lokacioni", 5080
    .Add , , "Madhësia", 1200, lvwColumnLeft
    .Add , , "Data", 1470
  End With
ItemIndex = 0
MDIForm1.mnuSkedar.Visible = True
End Sub

Private Sub fshij_Click()
If lstFiles.ListItems.Count = 0 Then Exit Sub
If ItemIndex <> 0 Then
    Dim Ask As String
    Ask = MsgBox("Are you sure you want to delete record?", vbYesNo + vbInformation, "Confirm Delete")
    If Ask = vbYes Then
        Dim DB As Database
        Dim RS As Recordset
        Set DB = OpenDatabase(App.Path + "\baza.mdb")
        ' Open recordset (table: People) in database DB
        Set RS = DB.OpenRecordset("Shenimet")
        
        'we will seek the record with ID equal to item's tag
        'first we set the table's index (this is the ID)
        'we need this because otherwise we can't use the seek function
        RS.Index = "ID"
         
        'we move the recordpointer to the first record, this way we can seek the whole table
        RS.MoveFirst
            
        'here we check if the seek functions result is equal to the item's tag.
        'the item's tag contains the ID from the table.
        'we stored it when we were reading from the table
        RS.Seek "=", lstFiles.ListItems.Item(ItemIndex).TaG
    
        If RS.NoMatch Then
            'if there was no match (the ID couldn't be found in the table)
            MsgBox ("The record can't be found in the table!"), vbOKOnly + vbCritical, "Error"
            
            
            'we reset itemindex to 0, this means that nothing is selected in lstFiles
            ItemIndex = 0
            
            'close recordset and database
            RS.Close
            DB.Close
            
            Exit Sub
        Else
            'there was a match! we will now delete the record from the database
            RS.Delete
            
            'and we will delete the row from the listview
            lstFiles.ListItems.Remove (ItemIndex)
        End If
           
        'close recordset and database
        RS.Close
        DB.Close
    
    End If

'we reset itemindex to 0, this means that nothing is selected in lstFiles
ItemIndex = 0
    
End If
End Sub

Private Sub lstFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
ItemIndex = Item.Index
End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton _
    Then PopupMenu menyja.menu
End Sub
