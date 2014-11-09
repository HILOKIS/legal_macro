Attribute VB_Name = "Module2"

Const user = "鈴木" '名前を記入してくださいね

Sub tes()
UserForm1.Show vbModeless
MsgBox user
End Sub


Sub Sample2()
    Const RefFile As String = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    ActiveWorkbook.VBProject.References.AddFromFile RefFile
End Sub
Sub makeMail()

'メール本文
Const stmsg As String = "" & vbCrLf & ""


Dim strFileName As String
Dim stAttachment As String

    On Error GoTo errorhandling
    Application.ScreenUpdating = False
  
Dim stsubject As String
Dim a As String, B As String, C As String, D As String
  
'エクセルファイル操作
Dim WB As Workbook
Dim WS As Worksheet
Dim InLastRow As Long

    Set WB = ActiveWorkbook
    With WB
    
        With .Sheets(1)
    
            'ロックをかけて
            .Sheets(1).Protect Password:="7397mkj"
            .EnableSelection = xlUnlockedCells
            
            '件名を作成
            D = .Range("X6") & "-" & .Range("AA6")
            a = "【契" & D & "回答" & Format(Date, "yymmdd")
            B = .Range("H12")
            C = "(" & .Range("G10") & ")"
            stsubject = a & B & C
            
        
        End With
        
        '契約データフォルダにフォルダがなければ作る
        
        
        '現在のファイル名に契約データ保存場所のパスに名前変更
        .SaveAs Filename:="\\Ts-xhl2cc\法務部共有\【001】契約書データ\55期（2014年度）\" & D & WB.Name
        
        '契約フォルダにファイルコピー
       ' FileCopy WB, "C:\Work\Sub\Book1.xls"
        
        .Close
    End With
      
'アウトルック操作
Dim olApp As Outlook.Application
Dim olNameSpace As Outlook.Namespace
Dim olInbox As Outlook.mapifolder
Dim olNewMail As Outlook.MailItem
Dim incounter As Long
    
    'アウトルックを格納ＯＲ起動
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = New Outlook.Application
        Set olNameSpace = olApp.GetNamespace("MAPI")
        Set olInbox = olNameSpace.GetDefaultFolder(olFolderInbox)
        olInbox.display
    End If
    
    
    Set olNewMail = olApp.CreateItem(olMailItem)
    With olNewMail
        
        '.importance
        
        .Subject = stsubject
       
       ' .Recipients.Add '宛先追加　シートの依頼人を参照
        
        .body = stmsg
        
        
        With .Attachments
            .Add WB.FullName
            .Item(1).DisplayName = WB.Name
        End With
        
        .display
        '.Save
    End With
    
    'Kill stAttachment
    
exitsub:
    Set olNewMail = Nothing
    Set olInbox = Nothing
    Set olNameSpace = Nothing
    Set olApp = Nothing
    Exit Sub
    
errorhandling:
    
    If Err.NUMBER = 429 Then
        Resume Next
    Else
        MsgBox "" & Err.NUMBER & vbNewLine & Err.Description, vbOKOnly
        Resume exitsub
    End If

End Sub

Dim cControl As CommandBarButton

Private Sub WB_ADINSTALL()

On Error Resume Next

Application.CommandBars("worksheet menu bar").Controls("はじめてのアドイン").Delete

Set cControl = Application.CommandBars("worksheet menu bar").Controls.Add

With cControl
    .Caption = "はじめてのアドイン"
    .Style = msoButtonCaption
    .OnAction = "mymacro"
End With
End Sub

Sub addsubmenu()

With CommandBars("Cell").Controls.Add(before:=1, Type:=msoControlPopup)
.Caption = "マクロ"
    With .Controls.Add
    .Caption = "command1"
    .OnAction = "hello"
    End With
    
    With .Controls.Add
    .Caption = "command2"
'    .onaction = ""
    End With
End With

End Sub
Sub hello()

MsgBox "hello"
End Sub
Sub test3()
ActiveWorkbook.Save
ActiveWorkbook.Close
End Sub

Sub Sample1()
    Dim Ref, buf As String
    For Each Ref In ActiveWorkbook.VBProject.References
        buf = buf & Ref.Name & vbTab & Ref.Description & vbCrLf
    Next Ref
    MsgBox buf
End Sub



Sub Samp1()
    ''フルパスを指定する方法
                    ''1
    MkDir "C:\Work\資料"                                            ''2
    Name "C:\Work\Sub\Book1.xls" As "C:\Work\資料\2007年度.xls"     ''3
    RmDir "C:\Work\Sub"                                             ''4
End Sub

Sub Sampe2()
    ''カレントフォルダを指定する方法
    ChDrive "C"                                     ''カレントドライブをCにする
    ChDir "C:\Work"                                 ''カレントフォルダをC:\Workにする
    FileCopy "C:\Book1.xls", "Sub\Book1.xls"        ''1
    MkDir "資料"                                    ''2
    Name "Sub\Book1.xls" As "資料\2007年度.xls"     ''3
    RmDir "Sub"                                     ''4
End Sub

Sub pasteobj()
Dim Target
Dim str As String

Target = Application.GetOpenFilename("WORDブック,*.do??")
str = Target

MsgBox mid(str, InStrRev(str, "\") + 1) '文字列の後方から検索
End Sub
