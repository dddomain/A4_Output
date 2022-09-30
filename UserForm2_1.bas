Option Explicit

Private WithEvents cboAutoComplete As MSForms.ComboBox
Private cboStored As Object

Private Sub UserForm_Initialize()

    Call LoadData

    Dim i As Integer
    With ComboBox1
        For i = 1 To Agrs.Count
            .AddItem Agrs(i).CoName
        Next
        .MatchEntry = fmMatchEntryNone
    End With
    
      '保存用のComboBoxにリストをコピー
    Set cboAutoComplete = ComboBox1
    Set cboStored = CreateObject("Forms.ComboBox.1")
    cboStored.List = cboAutoComplete.List
  
End Sub
 
Private Sub cboAutoComplete_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  Dim accCbo As Office.IAccessible
  Dim accLst As Office.IAccessible
  Dim i As Integer
   
  Set accCbo = cboAutoComplete
  Select Case KeyCode
    '動作するキー指定 ※必要に応じて変更
    '変換(28),無変換(29)
    Case 28, 29, vbKeyReturn, vbKeyBack, vbKeySpace, vbKeyDelete, _
         vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9

      'フィルタリングしてアイテム追加
      cboAutoComplete.Clear
      For i = 0 To cboStored.ListCount - 1
        If cboStored.List(i) Like "*" & cboAutoComplete.Text & "*" Then
          cboAutoComplete.AddItem cboStored.List(i)
        End If
      Next
 
      '開いているドロップダウンを閉じる
      If accCbo.accName(&H2&) = "閉じる" Then
        Set accLst = accCbo.accChild(&H3&)
        accLst.accDoDefaultAction &H0&
        'DoEvents
      End If
 
      cboAutoComplete.DropDown
  End Select
  
End Sub

Private Sub CommandButton1_Click()
On Error GoTo catch:
    
    Dim gotName As String: gotName = ComboBox1.Text
    passName = gotName
    Unload Me
    UserForm2_2.Show
    
Exit Sub
catch:
    MsgBox "正しい企業名を入力してください。"
End Sub
