Attribute VB_Name = "makedir_weekly"
Sub makedir_weekly()

Dim fold_path As String
Dim weeklyDate As Variant
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Templateシートより週次初めの日付を取得する
weeklyDate = ThisWorkbook.Worksheets("Template").Range("I4").Value
    
'日付を文字列変換し、/を削除する
weeklyDate = Replace(CStr(weeklyDate), "/", "")
    
fold_path = ThisWorkbook.Path + "\" + weeklyDate + "週"
    
'既存の週次フォルダがあれば一旦削除
If Dir(fold_path, vbDirectory) <> "" Then
    FSO.DeleteFolder fold_path
End If

'週次フォルダを作成
MkDir fold_path
 
 
'印刷範囲を再設定する。
'最終行+1 ～ 合計列-1の間の空白行を非表示にする。
'※発注書の列番号はwsTargetから取得

Dim ws As Worksheet
Dim ws_inner As Worksheet
Dim MaxRow_JAN As Double
Dim MaxRow_Sum As Double

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'PDF出力対象のワークシートを網羅するための繰返ループ
For Each ws In ThisWorkbook.Worksheets
    If InStr(ws.Name, "（") <> 0 Then
    
        MaxRow_JAN = ws.Cells(Rows.count, wsTarget_Col.JAN).End(xlUp).Row
        MaxRow_Sum = ws.Cells(Rows.count, wsTarget_Col.Sum).End(xlUp).Row
        
        '合計行の手前まで数値が埋まっている場合は、印刷範囲の再設定は必要なし。
        If MaxRow_JAN + 1 <> MaxRow_Sum Then
            ws.Range("B" & CStr(MaxRow_JAN + 1) + ":" + "B" & CStr(MaxRow_Sum - 1)).EntireRow.Delete
        End If
        
        '各種印刷詳細設定
        With ws.PageSetup
            .Orientation = xlLandscape
            .Zoom = 90
            .FitToPagesWide = 1
            .PrintArea = "A:P"
        End With
        
    End If
Next ws


'シート名の生産者名が一致しているものを配列にまとめ、PDFファイルへ出力する
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
    If InStr(ws.Name, "（") <> 0 Then
    
        '配列とカウンタを初期化
        ReDim wsName_array(0)
        ReDim file(0)
        wsName_array(0) = ws.Name
        count_array = 0
        count = 0
        
        'シート名の ( の左側を取得
        strTarget = Left(ws.Name, InStr(ws.Name, "（") - 1)
        
        
        'strTargetと同様名のPDFファイルが既に作成されていた場合はスキップ。
        
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
        
        
        '対象シート名のPDFファイルがまだ作成されていない時のみ、下記を実行する。
        If mkpdf_flag = True Then
            For Each ws_check In ThisWorkbook.Worksheets
                
                '対象シートの生産者名と同様(自らのシート除く)だった場合は、wsName_arrayに追加
                If ws.Name <> ws_check.Name And InStr(ws_check.Name, strTarget) <> 0 Then
                    count_array = count_array + 1
                    ReDim Preserve wsName_array(count_array)
                    wsName_array(count_array) = ws_check.Name
                End If
                
            Next ws_check
            
            '対象シートをPDFで保存
            '保存先は週次初め日付のフォルダ内
            
            '対象ワークシート名が格納された配列内を一つずつ取り出し、対象ワークシートの"P2"セルを確認。更新フラグ有無を確認。
            For Each i In wsName_array
                If ThisWorkbook.Worksheets(i).Range("P2") <> "" Then
                    add_flag = True
                End If
                
                '"P2"セルの文言を消去し、P2セルを初期化
                ThisWorkbook.Worksheets(i).Range("P2").Value = ""
            Next
            
            'ワークシートをグループ化
            ThisWorkbook.Worksheets(wsName_array).Select
            
            '更新フラグの有無によりファイル名を変更
            If add_flag = True Then
                ActiveSheet.ExportAsFixedFormat 0, fold_path & "\" + CStr(strTarget) + "_" + weeklyDate + "週" + "_" + "追記あり" + ".pdf"
            Else
                ActiveSheet.ExportAsFixedFormat 0, fold_path & "\" + CStr(strTarget) + "_" + weeklyDate + "週" + ".pdf"
            End If
            
            'ワークシートのグループ化を解除
            ThisWorkbook.Worksheets(wsName_array).Select
            
            '更新フラグを初期化
            add_flag = False
               
        End If
    
        'ファイル作成フラグを初期化
        mkpdf_flag = True
    
    End If
Next ws

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
MsgBox "PDFファイル作成終了 "

End Sub
