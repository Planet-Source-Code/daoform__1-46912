Attribute VB_Name = "modDataUtilities"
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32.dll" _
  (ByVal hWnd As Long) As Long

Public Execute As Boolean

Public Enum lvxSortType
  SortText = 0
  SortNumeric = 1
  SortDate = 2
  SortHHMM = 3
  SortHHMMSS = 4
  SortFileDateTime = 5
  
  End Enum
  
Public Db As dao.Database, Rs As dao.Recordset
Public varBookmark As Variant


Public Function FillLvFind(xLv As ListView, xRs As dao.Recordset)

frmFind.lblStatus.Caption = "Getting Records . . ."
frmFind.Refresh
xLv.View = lvwReport
Dim x As Integer, y As Integer, Dt As Integer
Screen.MousePointer = vbHourglass

With xLv.ColumnHeaders
.Clear
    For x = 0 To xRs.Fields.Count - 1
        .Add , , xRs.Fields(x).Name
        Dt = xRs.Fields.Item(x).Type
        .Item(x + 1).Tag = Dt
    Next x
.Item(1).Icon = "Up"
End With

xLv.Sorted = False
xRs.Requery

    If xRs.RecordCount = 0 Then
        xLv.ListItems.Clear
        Exit Function
    End If


With xLv.ListItems
    .Clear
    xRs.MoveLast:    xRs.MoveFirst
    frmFind.pB1.Max = xRs.RecordCount
        For x = 1 To xRs.RecordCount
        .Add x, , xRs(0)
            For y = 1 To xRs.Fields.Count - 1
                If Not IsNull(xRs(y)) Then xLv.ListItems(x).SubItems(y) = xRs(y)
           Next y
        xRs.MoveNext
        frmFind.pB1 = x
        Next x

frmFind.pB1.Visible = False
frmFind.lblStatus.Caption = x - 1 & " Records Found"
End With

xLv.SortKey = 0
xLv.SortOrder = lvwAscending
xLv.Sorted = True
xLv.ListItems(1).Selected = True
frmFind.lblColHdr.Caption = xLv.ColumnHeaders(1).Text & " is like:"
frmFind.txtFind.Left = frmFind.lblColHdr.Left + frmFind.lblColHdr.Width + 100
Screen.MousePointer = vbNormal

End Function

Public Sub FindInListview(Lvx As ListView, lvxText, txtBox As TextBox)

On Error GoTo ErrorHandler

If Lvx.SortKey < 1 Then GoTo DoFindMain Else GoTo DoFindSub

DoFindMain:
Dim Lvfindtm As ListItem
Dim TempSelStart As Integer
Dim strTemp As String

Set Lvfindtm = Lvx.FindItem(lvxText, lvwText, , lvwPartial)
If Not Lvfindtm Is Nothing Then
Lvfindtm.EnsureVisible
Lvfindtm.Selected = True

If Execute Then
TempSelStart = txtBox.SelStart

If Not txtBox.Text = "" Then

txtBox.SelLength = Len(txtBox.Text) - TempSelStart
    End If
        End If
            End If
Exit Sub

DoFindSub:

Dim LastColumnClicked As Integer
    LastColumnClicked = Lvx.SortKey
'Search Subitems
Dim iSubItemIndex As Integer
Dim i As Integer

iSubItemIndex = Lvx.SortKey
For i = 1 To Lvx.ListItems.Count
If UCase(Lvx.ListItems(i).SubItems(iSubItemIndex)) Like lvxText & "*" Then  'you could also use the LIKE operator
Lvx.ListItems(i).Selected = True
Lvx.ListItems(i).EnsureVisible

Exit For
End If
Next

Exit Sub

ErrorHandler:

  Call MsgBox("Runtime error " & Err.Number & ": " & _
    vbCrLf & Err.Description, vbOKOnly + vbCritical)

End Sub

Public Sub SortListView(ByVal Index As Integer, _
  ByVal CurrentListView As ListView, _
  Optional vSortType As lvxSortType = SortText)

'On Error GoTo ErrorHandler

  Dim i As Integer
  Dim strFormat As String
  Dim strData() As String
  Dim lRet As Long
  Dim ColHdrName As String
  Dim ColHdrPos As Integer
  'On Error GoTo ErrorHandler
  
'ColHdrName = CurrentListView.ColumnHeaders(Index).Text
'ColHdrPos = CurrentListView.ColumnHeaders(CurrentListView.SortOrder)

  Select Case vSortType

    Case SortText
      With CurrentListView
      
        .Sorted = True

                If .SortKey = Index Then
                    .SortOrder = 1 - .SortOrder
                Else
                    .SortKey = Index
                    .SortOrder = 0
                End If
        

GoSub Clear_Column_Header_Icon:
        

'If .SortOrder = lvwAscending Then
'.ColumnHeaders(Index + 1).Icon = "Up"
'Else
'.ColumnHeaders(Index + 1).Icon = "Down"
'End If

        
        Exit Sub
      
      End With
    
    
    
    Case SortNumeric
      strFormat = String(30, "0") & "." & String(30, "0")
    
    
    Case SortDate
      strFormat = String(2, "0") & "." & String(2, "0") & "." & String(4, "0")
      
    
    Case SortHHMM
      strFormat = "hh:mm"
      
    
    Case SortHHMMSS
      strFormat = "hh:mm:ss"
      
    
    Case SortFileDateTime
      strFormat = String(2, "0") & "." & String(2, "0") & "." & String(4, "0")
      strFormat = strFormat & " " & "hh:mm:ss"
      
    
    Case Else
      Exit Sub
  End Select
    
  
  lRet = LockWindowUpdate(CurrentListView.Parent.hWnd)
  If lRet = 0& Then
    Call MsgBox("Can't lock window " & _
      CurrentListView.Parent.hWnd, _
      vbOKOnly + vbCritical)
    Exit Sub
  End If
    
  With CurrentListView
    With .ListItems
      If (Index > 0) Then
        For i = 1 To .Count
          With .Item(i).ListSubItems(Index)
            .Tag = .Text & vbNullChar & .Tag
            Select Case vSortType
              Case SortNumeric
                If IsNumeric(.Text) Then
                  .Text = Format$(CDbl(.Text), strFormat)
                End If
              Case SortDate
                If IsDate(.Text) Then
                  .Text = Format$(CDate(.Text), strFormat)
                End If
              Case SortHHMM
                .Text = Format$(.Text, strFormat)
              Case SortHHMMSS
                .Text = Format$(.Text, strFormat)
              Case SortFileDateTime
                .Text = Format$(.Text, strFormat)
            End Select
          End With
        Next i
      Else
        For i = 1 To .Count
          With .Item(i)
           
            .Tag = .Text & vbNullChar & .Tag
            Select Case vSortType
              Case SortNumeric
                If IsNumeric(.Text) Then
                  .Text = Format$(CDbl(.Text), strFormat)
                End If
              Case SortDate
                 If IsDate(.Text) Then
                  .Text = Format$(CDate(.Text), strFormat)
                End If
              Case SortHHMM
                .Text = Format$(.Text, strFormat)
              Case SortHHMMSS
                .Text = Format$(.Text, strFormat)
              Case SortFileDateTime
                .Text = Format$(.Text, strFormat)
            End Select
          End With
        Next i
      End If
    End With
        
    
      .Sorted = True


        
                If .SortKey = Index Then
                    .SortOrder = 1 - .SortOrder
                    
                Else
                    .SortKey = Index
                    .SortOrder = 0
                End If
        
GoSub Clear_Column_Header_Icon:
        
    With .ListItems
      If (Index > 0) Then
        For i = 1 To .Count
          With .Item(i).ListSubItems(Index)
            strData = Split(.Tag, vbNullChar)
            .Text = strData(0)
            .Tag = strData(1)
          End With
        Next i
      Else
        For i = 1 To .Count
          With .Item(i)
            strData = Split(.Tag, vbNullChar)
            .Text = strData(0)
            .Tag = strData(1)
          End With
        Next i
      End If
    End With
  End With
        

  lRet = LockWindowUpdate(0&)
  Exit Sub
    
    
    
Clear_Column_Header_Icon:
'This handles the up and down icons in the Column headers

Dim x As Integer
For x = 1 To CurrentListView.ColumnHeaders.Count
If CurrentListView.ColumnHeaders(x).Icon > 0 Then
CurrentListView.ColumnHeaders(x).Icon = 0
CurrentListView.ColumnHeaders(x).Text = CurrentListView.ColumnHeaders(x).Text
End If
Next x

If CurrentListView.SortOrder = lvwAscending Then
CurrentListView.ColumnHeaders(Index + 1).Icon = "Up"
Else
CurrentListView.ColumnHeaders(Index + 1).Icon = "Down"
End If

Return
    
Exit Sub
ErrorHandler:

  Call MsgBox("Runtime error " & Err.Number & ": " & _
    vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub


