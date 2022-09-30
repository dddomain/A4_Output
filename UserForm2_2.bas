Option Explicit

Private WithEvents cboAutoComplete As MSForms.ComboBox
Private cboStored As Object

Private Sub UserForm_Initialize()
    
    'LoadData はForm10で呼び出し済み
    
    Dim i As Integer
    Dim AgrNames As Collection: Set AgrNames = New Collection
    For i = 1 To Agrs.Count
        If Agrs(i).CoName = passName Then
            ComboBox1.AddItem Agrs(i).AgrName
        End If
    Next
    'コンボボックスの初期値設定
    ComboBox1.ListIndex = 0
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

    '協定名がコンボボックスのテキストで、企業名がpassNameである、Agr.masterIdをシートに入力する
    Dim masterId As Integer
    Dim i As Integer
    Dim Agr As Variant
    For Each Agr In Agrs
        If Agr.AgrName = ComboBox1.Text And Agr.CoName = passName Then
            masterId = Agr.masterId
            Exit For
        End If
    Next
    Range("B3").value = masterId

    Unload Me

Exit Sub
catch:
    MsgBox "入力が不適切です。"
End Sub
