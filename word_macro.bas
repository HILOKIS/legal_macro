Sub replaceWords_fromExcelTable()

    '�@EXCEL�u���e�[�u�����i�[����
    Dim rBook As Object, rRow As Object
    rBook = GetObject(activedoument.Path & "\�u���e�[�u��.xlsx")

    '�A�h�L�������g�������u�����[�h�ɐ؂�ւ�
    With ActiveDocument.Content.Find

        '�B�G�N�Z���V�[�g�̌����u���͈͂��s���ƂɎ��o��
        For Each rRow In rBook.worksheets(1).Range("�����u���Z�b�g").Rows

            '�C�񎟌��z��̗�P��������Ȃ�������Next����
            If rRow.Cells(1).Value = "" Then Exit For

            '�D�񎟌��z��̗�P�������������Q�ɒu��
            .Text = rRow.Cells(1).Value
            .Replacement.Text = rRow.Cells(2).Value
            .Execute Replace:=wdReplaceAll

        Next rRow
    End With

    rBook.Application.Quit()

    rBook = Nothing

End Sub

Sub �w�b�_�[�Ƀt�@�C����������3()

    Dim myRange As Range

    myRange = ActiveDocument.Sections(1). _
          Headers(wdHeaderFooterPrimary).Range

    With myRange

        '�t�@�C�����̑}���i�e�L�X�g�j
        .Text = ActiveDocument.Name

        '�E����
        .Paragraphs.Alignment = wdAlignParagraphRight

    End With

    myRange = Nothing

End Sub

Sub OutLineLevel_BodyText()

    '�ꊇ����
    Dim myPara As Paragraph

    For Each myPara In ActiveDocument.Paragraphs
        myPara.Format.OutlineLevel = wdOutlineLevelBodyText
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
    '����^�{�_�񏑃^�C�g���{�i�N���C�A���g�j�ŏ��񖽖�
    '�@���O����^���𔻒f ���ŏ��̊��ʂ��폜���ĐV�����ǉ�
    If InStr(ActiveDocument.Name, ")�z") > 0 Then
        Dim i As Integer
        i = Mid(ActiveDocument.Name, 11, 1)
        i = i + 1
         ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "\�y" & Format(Date, "yymmdd") & "�@��(" & i & ")�z" & Mid(ActiveDocument.Name, 14)
    End If
    '�A�����̃t�@�C�����������ꍇ�͉E�[�ɇA�E�E�E��ǉ�
End Sub

Sub saveWithoutComments()
    '����^�{�_�񏑃^�C�g���{�i�N���C�A���g�j�ŏ��񖽖�
    '�@���O����^���𔻒f ���ŏ��̊��ʂ��폜���ĐV�����ǉ�
    If InStr(ActiveDocument.Name, ")�z") > 0 Then

        ActiveDocument.AcceptAllRevisions()
        ActiveDocument.TrackRevisions = False

        Dim i As Integer
        i = Mid(ActiveDocument.Name, 11, 1)
        i = i
        ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path & "\�y�����E�R�����g�Ȃ�(" & i & ")�z" & Mid(ActiveDocument.Name, 14)
    End If
    '�A�����̃t�@�C�����������ꍇ�͉E�[�ɇA�E�E�E��ǉ�
End Sub

Sub rightClickMenu()

    With Application.CommandBars("text")
        .Reset()

        With .Controls.Add(Type:=msoControlButton, Before:=1)
            .Caption = "�_�񏑉����ۑ�"
            .OnAction = "saveVersionUpdateWithComments"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=2)
            .Caption = "�����u��"
            .OnAction = "wordChange"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=3)
            .Caption = "�����Ȃ�"
            .OnAction = "saveWithoutComments"
        End With

        With .Controls.Add(Type:=msoControlButton, Before:=3)
            .Caption = "�w�b�_�[�ɖ���"
            .OnAction = "�w�b�_�[�Ƀt�@�C����������3"
        End With

    End With

End Sub

