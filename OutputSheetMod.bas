Option Explicit

'ボタンに登録する関数
Sub CallUserForm1()
    UserForm1.Show
End Sub
Sub CallUserForm2()
    UserForm2.Show
End Sub
Sub CallUserForm3()
    UserForm3.Show
End Sub
Sub PrintSheet()
    Worksheets("【A4出力】").PrintOut preview:=True
End Sub
Sub Clear()
    Range("B3").Select
    Selection.ClearContents
    Range("B11:H11").Select
    Selection.ClearContents
    Range("B12:H12").Select
    Selection.ClearContents
    Range("B3").Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo catch:

    Call LoadData
    Dim i As Integer
    
    If Not (Intersect(Target, Range("B3")) Is Nothing) And Range("B3").value <> "" Then
        Dim AgrId As Integer: AgrId = Range("D3").value
        'マスタテーブルで協定IDを検索して、同じ協定ＩＤをもつ協定の企業ＩＤを配列に格納していく
        Dim CoNames As Collection: Set CoNames = New Collection
        For i = 1 To Agrs.Count
            If Agrs(i).AgrId = AgrId And Agrs(i).CoName <> Range("B6").value Then
                CoNames.Add Agrs(i).CoName
            End If
        Next
        'テキストの作成・入力
        Range("B11").value = LineUpText(CoNames)
    End If
    
    If Not (Intersect(Target, Range("B3")) Is Nothing) And Range("B3").value <> "" Then
        Dim CoId As Integer: CoId = Range("F3").value
        'マスタテーブルで協定IDを検索して、同じ協定ＩＤをもつ協定の企業ＩＤを配列に格納していく
        Dim AgrNames As Collection: Set AgrNames = New Collection
        For i = 1 To Agrs.Count
            If Agrs(i).CoId = CoId And Agrs(i).AgrName <> Range("B5").value Then
                AgrNames.Add Agrs(i).AgrName
            End If
        Next
        'テキストの作成・入力
        Range("B12").value = LineUpText(AgrNames)
    End If
    
Exit Sub
catch:
    MsgBox "入力が不適切です。"
End Sub

Function LineUpText(ByRef NamesArr As Collection)
    Dim i As Integer: i = 1
    Dim Text As String: Text = ""
    For i = 1 To NamesArr.Count
        Text = Text & NamesArr(i)
        If i <> NamesArr.Count Then
            Text = Text & ",　"
        End If
    Next
    LineUpText = Text
End Function
