VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFind 
   Caption         =   "Find"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   315
      Left            =   8640
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pB1 
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   6840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return Found"
      Height          =   315
      Left            =   9840
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1250
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvFind 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10186
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":0000
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":015A
            Key             =   "Down"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Click Column Header to search particular column"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label lblColHdr 
      AutoSize        =   -1  'True
      Caption         =   "test"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()

Rs.Bookmark = varBookmark
Unload Me

End Sub



Private Sub cmdClose_Click()

End Sub

Private Sub cmdReturn_Click()

On Error Resume Next

Dim strFld As String, strRet As String
strFld = Rs(0).Name
strRet = lvFind.SelectedItem

    If lvFind.ColumnHeaders(1).Tag <> 4 Then
        Rs.FindFirst strFld & " = '" & strRet & "'"
    Else
        Rs.FindFirst strFld & " = " & strRet
    End If
    
Unload Me

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case vbKeyReturn
        cmdReturn_Click

    Case vbKeyEscape
        cmdCancel_Click
        
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then KeyAscii = 0  'No Beep!!

End Sub

Private Sub lvFind_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

Dim nIndex As Integer, TagData As Integer

nIndex = ColumnHeader.Index - 1
TagData = lvFind.ColumnHeaders(ColumnHeader.Index).Tag
  
Me.lblColHdr.Caption = lvFind.ColumnHeaders(ColumnHeader.Index).Text & " is like: "
Me.txtFind.Left = Me.lblColHdr.Left + Me.lblColHdr.Width + 200

With lvFind
    Select Case TagData
        Case 2, 3, 4, 5, 6, 17  'Numeric
        SortListView nIndex, lvFind, SortNumeric
      
        Case 202, 203, 11, 10  'Text
        SortListView nIndex, lvFind, SortText

        Case 7, 8 'Date
        SortListView nIndex, lvFind, SortDate
    
        Case Else  'if any other values are given, then end
        Exit Sub
    End Select
End With
SendKeys vbTab

End Sub


Private Sub lvFind_DblClick()

Unload Me
End Sub

Private Sub txtFind_Change()

FindInListview lvFind, (Trim$(txtFind.Text)), txtFind

End Sub

Private Sub txtFind_GotFocus()

txtFind.SelStart = 0
txtFind.SelLength = Len(txtFind.Text)

End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then lvFind.SetFocus

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)

 KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub
