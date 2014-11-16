Sub wordChange()

'①EXCEL置換テーブルを格納する
Dim rBook As Object, rRow As Object
Set rBook = GetObject(activedoument.Path & "\置換テーブル.xlsx")
 
'②ドキュメントを検索置換モードに切り替え
    With ActiveDocument.Content.Find
          
        '③エクセルシートの検索置換範囲を行ごとに取り出す
        For Each rRow In rBook.worksheets(1).Range("検索置換セット").Rows
        
            '④二次元配列の列１が見つからなかったらNext処理
            If rRow.Cells(1).Value = "" Then Exit For
                
                '⑤二次元配列の列１が見つかったら列２に置換
                .Text = rRow.Cells(1).Value
                .Replacement.Text = rRow.Cells(2).Value
                .Execute Replace:=wdReplaceAll
                
         Next rRow
    End With

rBook.Application.Quit

Set rBook = Nothing

End Sub

-------------------------------------------------------------------------------------------------------

Sub contractName()
Dim FileName As String


'①現在の名前が定型かどうか判断
If InStr(ActiveDocument.Name, "【法務") = 0 Then

    '②定型でなければTaskListから名前作成
     FileName = nameGene

Else
    '③定型の場合は日付と改訂歴を更新
    FileName = nameUpdate

End If


ActiveDocument.SaveAs2 FileName

End Sub

Function nameUpdate()

'①更新履歴回数を取得　インクリメントする
Dim cnt As Long
cnt = Mid(ActiveDocument.Name, 5, 1) + 1
nameUpdate = ActiveDocument.Path & "\【法務(" & cnt & ")" & Format(Date, "yymmdd") & "】" & Mid(ActiveDocument.Name, 14)
   
End Function
Function nameGene()

'①TaskListの項番を取得
Dim TaskList As Object
Dim Target As Excel.Range
On Error GoTo myerror
    
    Set TaskList = GetObject("")
    Set Target = InputBox("TaskListから案件番号を選択してください", Type:=8)


'②名前を付ける
Dim A As String, B As String
    A = Target.Value
    B = Target.Offset(0, 1).Value
    

    nameGene = Path & "【法務(1)" & Format(Date, "yymmdd") & "】" & B & "(" & c & ")" & A & ".docx"
    
End Function


Sub protectionChange()

Dim vba As String
vba = InputBox("Password")

    If ActiveDocument.ProtectionType = wdNoProtection Then
       ActiveDocument.Protect Password:=vba, NoReset:=False, Type:=wdAllowOnlyFormFields
    
    Else
        ActiveDocument.Unprotect Password:=vba
    
    End If

End Sub



Sub specialSave()
    '○定型＋契約書タイトル＋（クライアント）で初回命名
    '①名前が定型かを判断 →最初の括弧を削除して新しく追加
    If InStr(ActiveDocument.Name, ")】") > 0 Then
        Dim i As Integer
        i = Mid(ActiveDocument.Name, 11, 1)
        i = i + 1
         ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "\【" & Format(Date, "yymmdd") & "法務(" & i & ")】" & Mid(ActiveDocument.Name, 14)
    End If
    '②同名のファイルがあった場合は右端に②・・・を追加
End Sub
Sub specialSave2()
    '○定型＋契約書タイトル＋（クライアント）で初回命名
    '①名前が定型かを判断 →最初の括弧を削除して新しく追加
    If InStr(ActiveDocument.Name, ")】") > 0 Then

        ActiveDocument.AcceptAllRevisions()
        ActiveDocument.TrackRevisions = False

        Dim i As Integer
        i = Mid(ActiveDocument.Name, 11, 1)
        i = i
        ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "\【履歴・コメントなし(" & i & ")】" & Mid(ActiveDocument.Name, 14)
    End If
    '②同名のファイルがあった場合は右端に②・・・を追加
End Sub

Sub AutoExec()

    With Application.CommandBars("text")
        .Reset()

        With .Controls.Add(Type:=msoControlButton, Before:=1)
            .Caption = "契約書改訂保存"
            .OnAction = "specialSave"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=2)
            .Caption = "検索置換"
            .OnAction = "wordChange"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=3)
            .Caption = "履歴なし"
            .OnAction = "specialSave2"
        End With

    End With

End Sub

