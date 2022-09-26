Attribute VB_Name = "transfer_ordersheet"
'�}�N�����s�V�[�g �e��ԍ�
Enum wsImport_Col
    store = 1
    sheetName
End Enum

'�}�N�����s�V�[�g �e�s�ԍ�
Enum wsImport_Row
    Data = 3
End Enum
    

'������ �e�s�ԍ�
'�������Ӂ��� ���ڂ��L�ڂ���Ă���s��12�s�ڂƑz��B
'�����ύX������Ή��L�R�[�h��item = 12�̉ӏ����Y���s�ԍ��֏C�������肢���܂��B
Enum wsInv_Row
    items = 12
    Data
End Enum

'���Y�җl�ʔ�����(�^�[�Q�b�g�V�[�g) �e�s�ԍ�
Enum wsTarget_Row
    Date = 4
    Data = Date + 2
End Enum


'���Y�җl�ʔ�����(�^�[�Q�b�g�V�[�g) �e��ԍ�
Enum wsTarget_Col
    Product = 4
    JAN = 5
    Date = 9
    Sum = 16
End Enum

Sub transfer_ordersheet()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False


'�X�ܕʂɔ�������ǂݍ���
Dim ws As Worksheet
Dim wsImport As Worksheet
Dim wsTarget As Worksheet
Dim wsInv As Worksheet

Dim wsImport_sheetName As String
Dim wsImport_store As String
Dim wsImport_MaxRow As Double

Set wsImport = ThisWorkbook.Worksheets("�}�N�����s�V�[�g")

'�}�N�����s�V�[�g����ΏۂƂȂ锭�����̓X�܂ƃV�[�g���𒊏o

'�}�N�����s�V�[�g�̍ŏI�s�ԍ����擾
wsImport_MaxRow = wsImport.Cells(Rows.count, wsImport_Col.store).End(xlUp).Row

For wsImport_count = 0 To wsImport_MaxRow - wsImport_Row.Data

    wsImport_sheetName = wsImport.Cells(wsImport_Row.Data + wsImport_count, wsImport_Col.sheetName)
    wsImport_store = wsImport.Cells(wsImport_Row.Data + wsImport_count, wsImport_Col.store)
    
    '�}�N�����s�V�[�g���璊�o�����V�[�g�����������A�Y���V�[�g������Γǂݍ��ݑΏۂ̔������V�[�g�Ƃ��Đݒ肷��B
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, wsImport_sheetName) > 0 Then
            Set wsInv = ThisWorkbook.Worksheets(ws.Name)
        
            '���������[�i����ԍ��AJAN�R�[�h��ԍ��A���i����ԍ��A���ʗ�ԍ��A���Y�җ�ԍ������ꂼ��擾�B
            Dim wsInv_Col_deliDate As Long
            Dim wsInv_Col_JAN As Long
            Dim wsInv_Col_product As Long
            Dim wsInv_Col_quant As Long
            Dim wsInv_Col_maker As Long
            Dim count As Long
            
            '�����������Ƀt�B���^�[���|�����Ă����ꍇ�͉�������B
            If wsInv.FilterMode = True Then
                wsInv.ShowAllData
            End If
            
            For count = 1 To wsInv.Cells(wsInv_Row.items, Columns.count).End(xlToLeft).Column
                Select Case wsInv.Cells(wsInv_Row.items, count)
                    Case "�[�i��":
                        wsInv_Col_deliDate = count
                    Case "JAN�R�[�h":
                        wsInv_Col_JAN = count
                    Case "����揤�iCD":
                        wsInv_Col_maker = count
                    Case "���i��"
                        wsInv_Col_product = count
                    Case "����":
                        wsInv_Col_quant = count
                End Select
            Next
            
            
            '��������JAN�R�[�h�ƃ^�[�Q�b�g�V�[�g��JAN�R�[�h���r���A��v������[�i���ɂ��̐��ʂ�]�L
            'JAN�R�[�h����v���Ȃ������ꍇ�͐V���ȍs��ǉ������e��]�L
            Dim wsTarget_JAN As Variant
            Dim wsTarget_deliDate As Date
            Dim wsTarget_sheetName As Variant
            Dim wsTarget_sheetTitle As Variant
            
            Dim wsInv_deliDate As Date
            Dim wsInv_JAN As Variant
            Dim wsInv_product As Variant
            Dim wsInv_quant As Variant
            Dim wsInv_maker As Variant
            
            Dim wsInv_count As Long
            Dim wsTarget_count As Long
            Dim wsTarget_count_Col As Long
            
            Dim wsInv_MaxRow As Long
            Dim wsTarget_MaxRow As Long
            
            Dim transfer_flag As Boolean: transfer_flag = False
            Dim wsTarget_flag As Boolean: wsTarget_flag = False
            Dim tmp As Variant
            Dim n As Long: n = 0
            Dim m As Integer: m = 0
            Dim ws_2 As Worksheet
            
            '�������̍ŏI�s�ԍ����擾
            wsInv_MaxRow = wsInv.Cells(Rows.count, wsInv_Col_JAN).End(xlUp).Row
            
            For wsInv_count = wsInv_Row.Data To wsInv_MaxRow
                
                '���������[�i����ԍ��AJAN�R�[�h��ԍ��A���i����ԍ��A���ʗ�ԍ��A���Y�җ�ԍ����擾
                wsInv_deliDate = wsInv.Cells(wsInv_count, wsInv_Col_deliDate).Value
                wsInv_JAN = wsInv.Cells(wsInv_count, wsInv_Col_JAN).Value
                wsInv_product = wsInv.Cells(wsInv_count, wsInv_Col_product).Value
                wsInv_quant = wsInv.Cells(wsInv_count, wsInv_Col_quant).Value
                                
                '���Y�Җ��̑S�p���p�X�y�[�X���폜����
                tmp = wsInv.Cells(wsInv_count, wsInv_Col_maker).Value
                tmp = Replace(tmp, " ", "")
                tmp = Replace(tmp, "�@", "")
                
                '���Y�Җ��́i�j�̎�O�܂ł��擾����
                '���Y�Җ��̒��� ( �̈ʒu���擾����
                '�i ���S�p�������ꍇ
                If InStr(tmp, "�i") > 0 Then
                    n = InStr(tmp, "�i")
                
                '( �����p�������ꍇ
                ElseIf InStr(tmp, "(") > 0 Then
                    n = InStr(tmp, "(")
                End If
                    
                '���Y�Җ���( �̎�O�ŕ������A���̍������擾����B
                If n <> 0 Then
                    wsInv_maker = Left(tmp, n - 1)
                
                '( ��������΂��̂܂ܔ�������萶�Y�Җ����擾
                Else
                    wsInv_maker = tmp
                End If
                
                'n������������
                n = 0
                
                '�������̐��ʂ�0�A�󗓁A�G���[�ł���ΌJ��Ԃ��I��
                If wsInv_quant = 0 Or wsInv_quant = "" Or IsError(wsInv_quant) Then
                    GoTo Continue
                End If
                
                '�^�[�Q�b�g�V�[�g���w�肷��B
                '�������ɋL�ڂ̐��Y�Җ�+�i�X�ܖ��j�V�[�g������΂��̃V�[�g���w��
                '������ΐV���ɃV�[�g���쐬����B
                
                '�^�[�Q�b�g�V�[�g�̃V�[�g���𐶐�
                wsTarget_sheetName = wsInv_maker + "�i" + wsImport_store + "�j"
                
                For Each ws_2 In ThisWorkbook.Worksheets
                    
                    '���Y�Җ�+(�X�ܖ�)�̃V�[�g���������ꍇ
                    If InStr(ws_2.Name, wsTarget_sheetName) > 0 Then
                        Set wsTarget = ThisWorkbook.Worksheets(wsTarget_sheetName)
                        wsTarget_flag = True
                        Exit For
                    End If
                    
                Next ws_2
                    
                '���Y�Җ�+(�X�ܖ�)�̃V�[�g���Ȃ������ꍇ
                If wsTarget_flag <> True Then
                    ThisWorkbook.Worksheets("Template").Copy Before:=Worksheets(1)
                    ThisWorkbook.Worksheets(1).Name = wsTarget_sheetName
                    Set wsTarget = ThisWorkbook.Worksheets(wsTarget_sheetName)
                End If
                
                wsTarget_flag = False
                
                '�^�[�Q�b�g�V�[�g�̃^�C�g�����L��
                m = InStr(wsTarget.Name, "�i")
                tmp = Mid(wsTarget.Name, m + 1, Len(wsTarget.Name) - m - 1)
                wsTarget_sheetTitle = "���������������" & tmp & "�X�i��������)"
                wsTarget.Cells(2, 4).Value = wsTarget_sheetTitle
                
                '�^�[�Q�b�g�V�[�g�̍ŏI�s�ԍ����擾
                wsTarget_MaxRow = wsTarget.Cells(Rows.count, wsTarget_Col.JAN).End(xlUp).Row
                
                '�^�[�Q�b�g�V�[�g�̍ŏI�s�ԍ����f�[�^�s�ԍ�����O�̏ꍇ�A�f�[�^�s�ԍ�-1���ŏI�s�Ƃ��Ďw�肷��B
                If wsTarget_Row.Data > wsTarget_MaxRow Then
                    wsTarget_MaxRow = wsTarget_Row.Data - 1
                End If
                
                '��������JAN�R�[�h�ƃ^�[�Q�b�g�V�[�g��JAN�R�[�h���r
                For wsTarget_count = wsTarget_Row.Data To wsTarget_MaxRow
                    wsTarget_JAN = wsTarget.Cells(wsTarget_count, wsTarget_Col.JAN)
                    
                    'JAN�R�[�h����v���Ă����ꍇ�A�������ɋL�ڂ̓��t���ɍ��킹�ă^�[�Q�b�g�V�[�g�֐��ʂ��L�ځB
                    If wsTarget_JAN = wsInv_JAN Then
                        
                        '�������̓��t�ƃ^�[�Q�b�g�V�[�g�̓��t���r
                        For wsTarget_count_Col = 0 To 6
                            wsTarget_deliDate = wsTarget.Cells(wsTarget_Row.Date, wsTarget_Col.Date + wsTarget_count_Col)
                            
                            '���t����v������ԍ��̃Z���ցA�������̐��ʂ�]�L
                            If wsTarget_deliDate = wsInv_deliDate Then
                                
                                '���Ƀ^�[�Q�b�g�V�[�g���̊Y���Z���֐��ʂ������Ă����ꍇ�̓X�L�b�v
                                If wsTarget.Cells(wsTarget_count, wsTarget_Col.Date + wsTarget_count_Col).Value = 0 Then
                                    wsTarget.Cells(wsTarget_count, wsTarget_Col.Date + wsTarget_count_Col).Value = wsInv_quant
                                
                                    '�V���ɒǋL���������ꍇ�́Aworksheet�i'P2')�Z�����ɍX�V�L��̕�����ǉ�
                                    wsTarget.Range("P2").Value = "�X�V�L��"
                                    wsTarget.Range("P2").Font.ColorIndex = 2
                                End If
                            
                            End If
                        Next
                        
                        '�]�L������������J��Ԃ��I���B���̔������L��JAN�R�[�h�𒲂ׂ�B
                        transfer_flag = True
                        Exit For
                             
                    End If
                Next
                
                '�������ɋL�ڂ�JAN�R�[�h���^�[�Q�b�g�V�[�g��JAN�R�[�h�̂�����Ƃ���v���Ȃ������ꍇ�A
                '�^�[�Q�b�g�V�[�g�֐V���ȍs��ǉ����]�L����B
                If transfer_flag <> True Then
                    wsTarget_MaxRow = wsTarget_MaxRow + 1
                    
                    wsTarget.Cells(wsTarget_MaxRow, wsTarget_Col.Product) = wsInv_product
                    wsTarget.Cells(wsTarget_MaxRow, wsTarget_Col.JAN) = wsInv_JAN
                    
                    '�������̓��t�ƃ^�[�Q�b�g�V�[�g�̓��t���r
                        For wsTarget_count_Col = 0 To 6
                            wsTarget_deliDate = wsTarget.Cells(wsTarget_Row.Date, wsTarget_Col.Date + wsTarget_count_Col)
                            
                            '���t����v������ԍ��̃Z���ցA�������̐��ʂ�]�L
                            If wsTarget_deliDate = wsInv_deliDate Then
                                
                                '���Ƀ^�[�Q�b�g�V�[�g���̊Y���Z���֐��ʂ������Ă����ꍇ�̓X�L�b�v
                                If wsTarget.Cells(wsTarget_count, wsTarget_Col.Date + wsTarget_count_Col).Value = 0 Then
                                    wsTarget.Cells(wsTarget_MaxRow, wsTarget_Col.Date + wsTarget_count_Col).Value = wsInv_quant
                                
                                    '�V���ɒǋL���������ꍇ�́Aworksheet�i'P2')�Z�����ɍX�V�L��̕�����ǉ�
                                    wsTarget.Range("P2").Value = "�X�V�L��"
                                    wsTarget.Range("P2").Font.ColorIndex = 2
                                End If
                            End If
                        Next
                End If
                
                '�]�L�t���O��������
                transfer_flag = False
Continue:
            Next
            
    '��L�̓]�L������������A���̑ΏۂƂȂ锭�����V�[�g����������
        End If
    Next ws

'��L�̓]�L���S�Ċ���������A���̑ΏۂƂȂ�X�܂̔���������������
Next


'�]�L������������A�V�[�g�����������������ɕ��ёւ���
Dim count_sort As Long
Dim t As Long: t = 1

'�_�~�[�V�[�g��}������
With Worksheets.Add
    '���[�N�V�[�g�����Z���ɏ����o��
    For count_sort = 1 To Worksheets.count
        If InStr(Worksheets(count_sort).Name, "�i") <> 0 Then
            .Cells(t, 1).Value = Worksheets(count_sort).Name
            t = t + 1
        End If
    Next count_sort
        
    '���[�N�V�[�g�����\�[�g����
    .Range("A1").CurrentRegion.Sort .Range("A1")
        
    '���[�N�V�[�g�̈ʒu����בւ���
    Worksheets(.Cells(1, 1).Value).Move Before:=Worksheets(1)
    For count_sort = 2 To .Cells(Rows.count, 1).End(xlUp).Row
        Worksheets(.Cells(count_sort, 1).Value).Move After:=Worksheets(count_sort - 1)
    Next count_sort
        
    '�_�~�[�V�[�g���폜����
    .Delete
    
End With
         
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "�]�L�I�� "
 
End Sub


