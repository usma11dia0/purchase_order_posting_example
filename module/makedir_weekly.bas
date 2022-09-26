Attribute VB_Name = "makedir_weekly"
Sub makedir_weekly()

Dim fold_path As String
Dim weeklyDate As Variant
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Template�V�[�g���T�����߂̓��t���擾����
weeklyDate = ThisWorkbook.Worksheets("Template").Range("I4").Value
    
'���t�𕶎���ϊ����A/���폜����
weeklyDate = Replace(CStr(weeklyDate), "/", "")
    
fold_path = ThisWorkbook.Path + "\" + weeklyDate + "�T"
    
'�����̏T���t�H���_������Έ�U�폜
If Dir(fold_path, vbDirectory) <> "" Then
    FSO.DeleteFolder fold_path
End If

'�T���t�H���_���쐬
MkDir fold_path
 
 
'����͈͂��Đݒ肷��B
'�ŏI�s+1 �` ���v��-1�̊Ԃ̋󔒍s���\���ɂ���B
'���������̗�ԍ���wsTarget����擾

Dim ws As Worksheet
Dim ws_inner As Worksheet
Dim MaxRow_JAN As Double
Dim MaxRow_Sum As Double

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'PDF�o�͑Ώۂ̃��[�N�V�[�g��ԗ����邽�߂̌J�ԃ��[�v
For Each ws In ThisWorkbook.Worksheets
    If InStr(ws.Name, "�i") <> 0 Then
    
        MaxRow_JAN = ws.Cells(Rows.count, wsTarget_Col.JAN).End(xlUp).Row
        MaxRow_Sum = ws.Cells(Rows.count, wsTarget_Col.Sum).End(xlUp).Row
        
        '���v�s�̎�O�܂Ő��l�����܂��Ă���ꍇ�́A����͈͂̍Đݒ�͕K�v�Ȃ��B
        If MaxRow_JAN + 1 <> MaxRow_Sum Then
            ws.Range("B" & CStr(MaxRow_JAN + 1) + ":" + "B" & CStr(MaxRow_Sum - 1)).EntireRow.Delete
        End If
        
        '�e�����ڍאݒ�
        With ws.PageSetup
            .Orientation = xlLandscape
            .Zoom = 90
            .FitToPagesWide = 1
            .PrintArea = "A:P"
        End With
        
    End If
Next ws


'�V�[�g���̐��Y�Җ�����v���Ă�����̂�z��ɂ܂Ƃ߁APDF�t�@�C���֏o�͂���
Dim ws_check As Worksheet
Dim wsName_array() As Variant
Dim strTarget As String
Dim count_array As Integer
Dim add_flag As Boolean: add_flag = False

Dim buf As String
Dim count As Integer
Dim file() As String
Dim mkpdf_flag As Boolean: mkpdf_flag = True

For Each ws In ThisWorkbook.Worksheets
    If InStr(ws.Name, "�i") <> 0 Then
    
        '�z��ƃJ�E���^��������
        ReDim wsName_array(0)
        ReDim file(0)
        wsName_array(0) = ws.Name
        count_array = 0
        count = 0
        
        '�V�[�g���� ( �̍������擾
        strTarget = Left(ws.Name, InStr(ws.Name, "�i") - 1)
        
        
        'strTarget�Ɠ��l����PDF�t�@�C�������ɍ쐬����Ă����ꍇ�̓X�L�b�v�B
        
        buf = Dir(fold_path & "\" & "*")
        
        Do While buf <> ""
            count = count + 1
            ReDim Preserve file(count)
            file(count) = CStr(buf)
            buf = Dir()
        Loop
        
        For count = LBound(file) To UBound(file)
            If (InStr(file(count), strTarget) <> 0) Then
                mkpdf_flag = False
            End If
        Next
        
        
        '�ΏۃV�[�g����PDF�t�@�C�����܂��쐬����Ă��Ȃ����̂݁A���L�����s����B
        If mkpdf_flag = True Then
            For Each ws_check In ThisWorkbook.Worksheets
                
                '�ΏۃV�[�g�̐��Y�Җ��Ɠ��l(����̃V�[�g����)�������ꍇ�́AwsName_array�ɒǉ�
                If ws.Name <> ws_check.Name And InStr(ws_check.Name, strTarget) <> 0 Then
                    count_array = count_array + 1
                    ReDim Preserve wsName_array(count_array)
                    wsName_array(count_array) = ws_check.Name
                End If
                
            Next ws_check
            
            '�ΏۃV�[�g��PDF�ŕۑ�
            '�ۑ���͏T�����ߓ��t�̃t�H���_��
            
            '�Ώۃ��[�N�V�[�g�����i�[���ꂽ�z�����������o���A�Ώۃ��[�N�V�[�g��"P2"�Z�����m�F�B�X�V�t���O�L�����m�F�B
            For Each i In wsName_array
                If ThisWorkbook.Worksheets(i).Range("P2") <> "" Then
                    add_flag = True
                End If
                
                '"P2"�Z���̕������������AP2�Z����������
                ThisWorkbook.Worksheets(i).Range("P2").Value = ""
            Next
            
            '���[�N�V�[�g���O���[�v��
            ThisWorkbook.Worksheets(wsName_array).Select
            
            '�X�V�t���O�̗L���ɂ��t�@�C������ύX
            If add_flag = True Then
                ActiveSheet.ExportAsFixedFormat 0, fold_path & "\" + CStr(strTarget) + "_" + weeklyDate + "�T" + "_" + "�ǋL����" + ".pdf"
            Else
                ActiveSheet.ExportAsFixedFormat 0, fold_path & "\" + CStr(strTarget) + "_" + weeklyDate + "�T" + ".pdf"
            End If
            
            '���[�N�V�[�g�̃O���[�v��������
            ThisWorkbook.Worksheets(wsName_array).Select
            
            '�X�V�t���O��������
            add_flag = False
               
        End If
    
        '�t�@�C���쐬�t���O��������
        mkpdf_flag = True
    
    End If
Next ws

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
MsgBox "PDF�t�@�C���쐬�I�� "

End Sub
