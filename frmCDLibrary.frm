VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCDLibrary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CD Library"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   DrawMode        =   11  'Not Xor Pen
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNotes 
      Caption         =   "Notes"
      Height          =   315
      Left            =   6720
      TabIndex        =   14
      Top             =   240
      Width           =   1050
   End
   Begin VB.TextBox txtCDName 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtCDDate 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtCDID 
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2400
      Width           =   1050
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   1050
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Width           =   1150
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Top             =   2400
      Width           =   1050
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   2790
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&l"
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Shape sqr1 
      BorderColor     =   &H80000003&
      Height          =   255
      Left            =   6960
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblLabels 
      Caption         =   "CD ID"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "CD Name:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "CD Date:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmCDLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim mbDeleteFlag As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdCancel_Click()

On Error Resume Next

Data1.Recordset.CancelUpdate
Data1.UpdateControls

If Not IsEmpty(varBookmark) Then
        Data1.Recordset.Bookmark = varBookmark
    Else
        Data1.Recordset.MoveFirst
End If

StatusBar1.SimpleText = "Record " & Data1.Recordset.AbsolutePosition + 1 & " of " & Data1.Recordset.RecordCount
mbDataChanged = False
SetButtons True
mbEditFlag = False
mbAddNewFlag = False

End Sub

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub CheckRecordChanged()


If mbDeleteFlag Then Exit Sub

On Error Resume Next


Dim ctl As Control
For Each ctl In Me.Controls
If ctl.DataChanged Then mbDataChanged = ctl.DataChanged
Next ctl

If mbDataChanged = True Then

    Dim xMsg
    xMsg = MsgBox("Record has changed!" & vbCr & vbCr & "Save changes ?", vbYesNo, "Save changes")
        If xMsg = vbNo Then
            'Data1.Recordset.Edit
            Data1.Recordset.CancelUpdate
            'Data1.Recordset.Update
            Data1.UpdateControls
        End If
End If

mbDataChanged = False

Exit Sub
errhand:
Call MsgBox(Err.Number & " " & Err.Description)
Err.Clear

End Sub
Private Sub cmdDelete_Click()

On Error Resume Next

If Data1.Recordset.RecordCount = 0 Then MsgBox "No records to delete": Exit Sub

mbDeleteFlag = True

Dim xMsg
xMsg = MsgBox("Delete Record, are you sure ?", vbYesNo, "Delete Record")
If xMsg = vbNo Then mbDeleteFlag = False: Exit Sub

With Data1.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
End With


End Sub

Private Sub cmdEdir_Click()

End Sub

Private Sub cmdEdit_Click()

If Data1.Recordset.RecordCount = 0 Then MsgBox "No records to edit": Exit Sub

varBookmark = Data1.Recordset.Bookmark
StatusBar1.SimpleText = "Edit record"
mbEditFlag = True
SetButtons False
Data1.Recordset.Edit

txtCDName.SetFocus

End Sub

Private Sub cmdFind_Click()

varBookmark = Data1.Recordset.Bookmark
If Data1.Recordset.RecordCount = 0 Then MsgBox "No records in database": Exit Sub
frmFind.Show
FillLvFind frmFind.lvFind, Me.Data1.Recordset

End Sub

Private Sub cmdNew_Click()

  With Data1.Recordset
    If Not (.BOF And .EOF) And Not IsEmpty(varBookmark) Then
      varBookmark = .Bookmark
    End If
If Data1.Recordset.RecordCount > 0 Then .Edit: .Update
    .AddNew
    mbAddNewFlag = True
    StatusBar1.SimpleText = "Add record"
    SetButtons False
    
      End With

txtCDDate.Text = Date
txtCDName.SetFocus

End Sub


Private Sub cmdNotes_Click()

ShellExecute 0, "OPEN", "c:\notes.txt", "", "", 4

End Sub

Private Sub cmdUpdate_Click()

On Error Resume Next

Data1.UpdateRecord
  If mbAddNewFlag Then
    Data1.Recordset.MoveLast
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  mbDeleteFlag = False

StatusBar1.SimpleText = "Record " & Data1.Recordset.AbsolutePosition + 1 & " of " & Data1.Recordset.RecordCount & "  Update Successful . . "

End Sub


Private Sub Data1_Reposition()

StatusBar1.SimpleText = "Record " & Data1.Recordset.AbsolutePosition + 1 & " of " & Data1.Recordset.RecordCount

Exit Sub
errhand:
Call MsgBox(Err.Number & " " & Err.Description)
Err.Clear

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)


CheckRecordChanged

mbDataChanged = False

  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    
    Case vbKeyF2
        cmdFind_Click
        
    Case vbKeyEscape, vbKeyF12
      cmdClose_Click
    
    Case vbKeyEnd
      Data1.Recordset.MoveLast
    
    Case vbKeyHome
      Data1.Recordset.MoveFirst
      
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        Data1.Recordset.MoveFirst
      Else
        If Not Data1.Recordset.BOF Then Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        Data1.Recordset.MoveLast
      Else
If Not Data1.Recordset.EOF Then Data1.Recordset.MoveNext
If Data1.Recordset.EOF And Data1.Recordset.RecordCount > 0 Then Data1.Recordset.MoveLast
      End If
  End Select

End Sub

Private Sub Form_Load()

On Error Resume Next
'find and open the database file

Set Db = OpenDatabase(App.Path & "\" & "daodata.mdb")
Set Rs = Db.OpenRecordset("CD Library", dbOpenDynaset)
Rs.MoveLast: Rs.MoveFirst

'set the Data control on the form to use the table opened
Data1.DatabaseName = Db.Name
Data1.RecordSource = Rs.Name
Set Data1.Recordset = Rs

'just for esthetics. You can position anything you want with this subroutine
PositionControls

'set all the relevant controls on the form to use the recordset opened
txtCDID.DataField = Data1.Recordset.Fields(0).SourceField
txtCDName.DataField = Data1.Recordset.Fields(1).SourceField
txtCDDate.DataField = Data1.Recordset.Fields(2).SourceField

'this is a switch declared in the forms general section.
mbDataChanged = False

'if there are no records then display relevant message on status bar
If Data1.Recordset.RecordCount = 0 Then
StatusBar1.SimpleText = "Click 'New' to add a record"
End If


End Sub

Private Sub PositionControls()

'see form_load function
With sqr1
    .Left = Me.CurrentX + 20
    .Width = Me.Width - 120
    .Top = Me.CurrentY + 20
    .Height = Me.Height - 730
End With

End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdNew.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdFind.Visible = bVal
  Data1.Visible = bVal
  
  End Sub

Private Sub Form_Unload(Cancel As Integer)

Rs.Close
Set Rs = Nothing
Db.Close
Set Db = Nothing

End Sub

Private Sub txtCDDate_GotFocus()

If mbEditFlag Then
txtCDDate.SelStart = 0
txtCDDate.SelLength = Len(txtCDDate.Text)
End If

End Sub

Private Sub txtCDName_GotFocus()

If mbEditFlag Then
txtCDName.SelStart = 0
txtCDName.SelLength = Len(txtCDName.Text)
End If

End Sub

Private Sub txtCDName_LostFocus()

If mbAddNewFlag = True Or mbEditFlag = True Then txtCDName = UCase(txtCDName)

End Sub
