Option Explicit

Private Sub UserForm_Initialize()
    With ComboBox1
        Dim i As Integer
        For i = 1 To Agrs.Count
            .AddItem Agrs.masterId
        Next
    End With
End Sub

Private Sub CommandButton1_Click()
On Error GoTo catch:
    
    With Worksheets("【A4出力】")
        Select Case Int(Me.ComboBox1.Text)
            Case 1 To Agrs.Count
                .Range("B3").value = Me.ComboBox1.Text
                Me.ComboBox1.Text = ""
            Case Else
                GoTo catch:
        End Select
    End With
    Unload Me
Exit Sub
catch:
    MsgBox "正しいIDを入力してください。"
End Sub
