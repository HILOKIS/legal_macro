Attribute VB_Name = "Module1"
Sub replaceWords_fromExcelTable()

    '①EXCEL置換テーブルを格納する
    Dim rBook As Object, rRow As Object
    rBook = GetObject(activedoument.Path & "\置換テーブル.xlsx")

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

    rBook = Nothing

End Sub

Sub ヘッダーにファイル名を入れる3()

    Dim myRange As Range

    myRange = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range

    With myRange

        'ファイル名の挿入（テキスト）
        .Text = ActiveDocument.Name

        '右揃え
        .Paragraphs.Alignment = wdAlignParagraphRight

    End With

   Set myRange = Nothing

End Sub
Sub 空の段落検索()

  Dim myPara As Paragraph
  Dim myText As String

  For Each myPara In ActiveDocument.Paragraphs
 
     myText = Trim(myPara.Range.Text)
    myText = Replace(myText, vbCr, "")
    myText = Replace(myText, vbTab, "")
  
     If Len(myText) = 0 Then
      myPara.Range.Delete
      Exit For
    End If
  
   Next myPara

 End Sub

Sub OutLineLevel_BodyText()

    '一括解除
    Dim myPara As Paragraph

    For Each myPara In ActiveDocument.Paragraphs
        myPara.Format.OutlineLevel = wdOutlineLevel1
        Next

End Sub

Sub protectionChange()

Dim vba As String
vba = InputBox("Password")

    If ActiveDocument.ProtectionType = wdNoProtection Then
       ActiveDocument.Protect Password:=vba, NoReset:=False, Type:=wdAllowOnlyFormFields
    
    Else
        ActiveDocument.Unprotect Password:=vba
    
    End If

End Sub

Sub saveVersionUpdateWithComments()
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

Sub saveWithoutComments()
    '○定型＋契約書タイトル＋（クライアント）で初回命名
    '①名前が定型かを判断 →最初の括弧を削除して新しく追加
    If InStr(ActiveDocument.Name, ")】") > 0 Then

        ActiveDocument.AcceptAllRevisions
        ActiveDocument.TrackRevisions = False

        Dim i As Integer
        i = Mid(ActiveDocument.Name, 11, 1)
        i = i
        ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "\【履歴・コメントなし(" & i & ")】" & Mid(ActiveDocument.Name, 14)
    End If
    '②同名のファイルがあった場合は右端に②・・・を追加
End Sub

Sub rightClickMenu()

    With Application.CommandBars("text")
        .Reset

        With .Controls.Add(Type:=msoControlButton, Before:=1)
            .Caption = "契約書改訂保存"
            .OnAction = "saveVersionUpdateWithComments"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=2)
            .Caption = "検索置換"
            .OnAction = "wordChange"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=3)
            .Caption = "履歴なし"
            .OnAction = "saveWithoutComments"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=3)
            .Caption = "ヘッダーに名称"
            .OnAction = "ヘッダーにファイル名を入れる3"
        End With

    End With

End Sub

Sub 回答書作成()

Dim 案件番号 As String
Dim 前提 As String
Dim コメント As String

    案件番号 = ActiveDocument.Comments(1).Range.Text
        前提 = ActiveDocument.Comments(2).Range.Text
    コメント = ActiveDocument.Comments(3).Range.Text

MsgBox 案件番号 & 前提 & コメント


'ファイル＋履歴なしファイル保存　後から貼り付けるためにフルパスを変数格納
Dim 履歴ありファイル As String: 履歴ありファイル = ActiveDocument.FullName
Dim 履歴なしファイル As String: 履歴なしファイル = ActiveDocument.Path & "\【履歴・コメントなし" & Mid(ActiveDocument.Name, 10)

    With ActiveDocument
        .save
    
        .AcceptAllRevisions
        .TrackRevisions = False
        .DeleteAllComments
        .SaveAs2 FileName:=履歴なしファイル
        
        .Close
    End With

'回答書が存在するかどうかを確認するには日付を特定する必要がある？案件番号だけで検索？
'案件番号の回答書を格納するファイルが存在する→
'案件番号の回答書を格納するファイルが存在しない→
    Dim Targetfile As String: Targetfile = ""
    
    If Dir(Targetfile) = "" Then
    
    Else
    
    End If
    
    
'ここからエクセル処理
Dim 回答書 As Excel.Workbook: Set 回答書 = Excel.Workbooks.Open(FileName:="")
    
    回答書.Sheets("").Activate
    
    With 回答書.Sheets("")
    
        .Range("").Value = 案件番号
        .Shapes(1).TextFrame2.TextRange.Characters.Text = 前提
        .Shapes(2).TextFrame2.TextRange.Characters.Text = コメント
        
        .Range("").Activate
        .OLEObjects.Add(FileName:=履歴ありファイル, Link:=False, DisplayAsIcon:=True, _
            IconFileName:="C:\Windows\Installer\{90150000-000F-0000-0000-0000000FF1CE}\wordicon.exe", _
            IconIndex:=0, IconLabel:=履歴ありファイル).Select
            
        .Range("").Activate
        .OLEObjects.Add(FileName:=履歴なしファイル, Link:=False, DisplayAsIcon:=True, _
            IconFileName:="C:\Windows\Installer\{90150000-000F-0000-0000-0000000FF1CE}\wordicon.exe", _
            IconIndex:=0, IconLabel:=履歴なしファイル).Select
    
    End With
    
End Sub

