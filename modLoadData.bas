Option Explicit

Public MasterTable As ListObject
Public Agrs As Collection
Public passName As String 'ユーザーフォーム間の値渡し用のパブリック変数

Sub LoadData()
    Set MasterTable = Sheets("マスタテーブル").ListObjects(1)
    Dim Agr As Agreement
    Set Agrs = New Collection
    Dim i As Integer
    For i = 1 To MasterTable.ListRows.Count
        Set Agr = New Agreement
        Agr.initialize MasterTable.ListRows(i).Range
        Agrs.Add Agr
    Next
    Set Agr = Nothing
End Sub
