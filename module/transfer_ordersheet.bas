Attribute VB_Name = "transfer_ordersheet"
'マクロ実行シート 各列番号
Enum wsImport_Col
    store = 1
    sheetName
End Enum

'マクロ実行シート 各行番号
Enum wsImport_Row
    Data = 3
End Enum
    

'発注書 各行番号
'※※注意※※ 項目が記載されている行は12行目と想定。
'もし変更があれば下記コードのitem = 12の箇所を該当行番号へ修正をお願いします。
Enum wsInv_Row
    items = 12
    Data
End Enum

'生産者様別発注書(ターゲットシート) 各行番号
Enum wsTarget_Row
    Date = 4
    Data = Date + 2
End Enum


'生産者様別発注書(ターゲットシート) 各列番号
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


'店舗別に発注書を読み込む
Dim ws As Worksheet
Dim wsImport As Worksheet
Dim wsTarget As Worksheet
Dim wsInv As Worksheet

Dim wsImport_sheetName As String
Dim wsImport_store As String
Dim wsImport_MaxRow As Double

Set wsImport = ThisWorkbook.Worksheets("マクロ実行シート")

'マクロ実行シートから対象となる発注書の店舗とシート名を抽出

'マクロ実行シートの最終行番号を取得
wsImport_MaxRow = wsImport.Cells(Rows.count, wsImport_Col.store).End(xlUp).Row

For wsImport_count = 0 To wsImport_MaxRow - wsImport_Row.Data

    wsImport_sheetName = wsImport.Cells(wsImport_Row.Data + wsImport_count, wsImport_Col.sheetName)
    wsImport_store = wsImport.Cells(wsImport_Row.Data + wsImport_count, wsImport_Col.store)
    
    'マクロ実行シートから抽出したシート名を検索し、該当シートがあれば読み込み対象の発注書シートとして設定する。
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, wsImport_sheetName) > 0 Then
            Set wsInv = ThisWorkbook.Worksheets(ws.Name)
        
            '発注書より納品日列番号、JANコード列番号、商品名列番号、数量列番号、生産者列番号をそれぞれ取得。
            Dim wsInv_Col_deliDate As Long
            Dim wsInv_Col_JAN As Long
            Dim wsInv_Col_product As Long
            Dim wsInv_Col_quant As Long
            Dim wsInv_Col_maker As Long
            Dim count As Long
            
            'もし発注書にフィルターが掛かっていた場合は解除する。
            If wsInv.FilterMode = True Then
                wsInv.ShowAllData
            End If
            
            For count = 1 To wsInv.Cells(wsInv_Row.items, Columns.count).End(xlToLeft).Column
                Select Case wsInv.Cells(wsInv_Row.items, count)
                    Case "納品日":
                        wsInv_Col_deliDate = count
                    Case "JANコード":
                        wsInv_Col_JAN = count
                    Case "取引先商品CD":
                        wsInv_Col_maker = count
                    Case "商品名"
                        wsInv_Col_product = count
                    Case "数量":
                        wsInv_Col_quant = count
                End Select
            Next
            
            
            '発注書のJANコードとターゲットシートのJANコードを比較し、一致したら納品日にその数量を転記
            'JANコードが一致しなかった場合は新たな行を追加し内容を転記
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
            
            '発注書の最終行番号を取得
            wsInv_MaxRow = wsInv.Cells(Rows.count, wsInv_Col_JAN).End(xlUp).Row
            
            For wsInv_count = wsInv_Row.Data To wsInv_MaxRow
                
                '発注書より納品日列番号、JANコード列番号、商品名列番号、数量列番号、生産者列番号を取得
                wsInv_deliDate = wsInv.Cells(wsInv_count, wsInv_Col_deliDate).Value
                wsInv_JAN = wsInv.Cells(wsInv_count, wsInv_Col_JAN).Value
                wsInv_product = wsInv.Cells(wsInv_count, wsInv_Col_product).Value
                wsInv_quant = wsInv.Cells(wsInv_count, wsInv_Col_quant).Value
                                
                '生産者名の全角半角スペースを削除する
                tmp = wsInv.Cells(wsInv_count, wsInv_Col_maker).Value
                tmp = Replace(tmp, " ", "")
                tmp = Replace(tmp, "　", "")
                
                '生産者名は（）の手前までを取得する
                '生産者名の中の ( の位置を取得する
                '（ が全角だった場合
                If InStr(tmp, "（") > 0 Then
                    n = InStr(tmp, "（")
                
                '( が半角だった場合
                ElseIf InStr(tmp, "(") > 0 Then
                    n = InStr(tmp, "(")
                End If
                    
                '生産者名を( の手前で分割し、その左側を取得する。
                If n <> 0 Then
                    wsInv_maker = Left(tmp, n - 1)
                
                '( が無ければそのまま発注書より生産者名を取得
                Else
                    wsInv_maker = tmp
                End If
                
                'nを初期化する
                n = 0
                
                '発注書の数量が0、空欄、エラーであれば繰り返し終了
                If wsInv_quant = 0 Or wsInv_quant = "" Or IsError(wsInv_quant) Then
                    GoTo Continue
                End If
                
                'ターゲットシートを指定する。
                '発注書に記載の生産者名+（店舗名）シートがあればそのシートを指定
                '無ければ新たにシートを作成する。
                
                'ターゲットシートのシート名を生成
                wsTarget_sheetName = wsInv_maker + "（" + wsImport_store + "）"
                
                For Each ws_2 In ThisWorkbook.Worksheets
                    
                    '生産者名+(店舗名)のシートがあった場合
                    If InStr(ws_2.Name, wsTarget_sheetName) > 0 Then
                        Set wsTarget = ThisWorkbook.Worksheets(wsTarget_sheetName)
                        wsTarget_flag = True
                        Exit For
                    End If
                    
                Next ws_2
                    
                '生産者名+(店舗名)のシートがなかった場合
                If wsTarget_flag <> True Then
                    ThisWorkbook.Worksheets("Template").Copy Before:=Worksheets(1)
                    ThisWorkbook.Worksheets(1).Name = wsTarget_sheetName
                    Set wsTarget = ThisWorkbook.Worksheets(wsTarget_sheetName)
                End If
                
                wsTarget_flag = False
                
                'ターゲットシートのタイトルを記載
                m = InStr(wsTarget.Name, "（")
                tmp = Mid(wsTarget.Name, m + 1, Len(wsTarget.Name) - m - 1)
                wsTarget_sheetTitle = "●●●●株式会社" & tmp & "店（△△△△)"
                wsTarget.Cells(2, 4).Value = wsTarget_sheetTitle
                
                'ターゲットシートの最終行番号を取得
                wsTarget_MaxRow = wsTarget.Cells(Rows.count, wsTarget_Col.JAN).End(xlUp).Row
                
                'ターゲットシートの最終行番号がデータ行番号より手前の場合、データ行番号-1を最終行として指定する。
                If wsTarget_Row.Data > wsTarget_MaxRow Then
                    wsTarget_MaxRow = wsTarget_Row.Data - 1
                End If
                
                '発注書のJANコードとターゲットシートのJANコードを比較
                For wsTarget_count = wsTarget_Row.Data To wsTarget_MaxRow
                    wsTarget_JAN = wsTarget.Cells(wsTarget_count, wsTarget_Col.JAN)
                    
                    'JANコードが一致していた場合、発注書に記載の日付欄に合わせてターゲットシートへ数量を記載。
                    If wsTarget_JAN = wsInv_JAN Then
                        
                        '発注書の日付とターゲットシートの日付を比較
                        For wsTarget_count_Col = 0 To 6
                            wsTarget_deliDate = wsTarget.Cells(wsTarget_Row.Date, wsTarget_Col.Date + wsTarget_count_Col)
                            
                            '日付が一致した列番号のセルへ、発注書の数量を転記
                            If wsTarget_deliDate = wsInv_deliDate Then
                                
                                '既にターゲットシート内の該当セルへ数量が入っていた場合はスキップ
                                If wsTarget.Cells(wsTarget_count, wsTarget_Col.Date + wsTarget_count_Col).Value = 0 Then
                                    wsTarget.Cells(wsTarget_count, wsTarget_Col.Date + wsTarget_count_Col).Value = wsInv_quant
                                
                                    '新たに追記があった場合は、worksheet（'P2')セル内に更新有りの文言を追加
                                    wsTarget.Range("P2").Value = "更新有り"
                                    wsTarget.Range("P2").Font.ColorIndex = 2
                                End If
                            
                            End If
                        Next
                        
                        '転記が完了したら繰り返し終了。次の発注書記載JANコードを調べる。
                        transfer_flag = True
                        Exit For
                             
                    End If
                Next
                
                '発注書に記載のJANコードがターゲットシートのJANコードのいずれとも一致しなかった場合、
                'ターゲットシートへ新たな行を追加し転記する。
                If transfer_flag <> True Then
                    wsTarget_MaxRow = wsTarget_MaxRow + 1
                    
                    wsTarget.Cells(wsTarget_MaxRow, wsTarget_Col.Product) = wsInv_product
                    wsTarget.Cells(wsTarget_MaxRow, wsTarget_Col.JAN) = wsInv_JAN
                    
                    '発注書の日付とターゲットシートの日付を比較
                        For wsTarget_count_Col = 0 To 6
                            wsTarget_deliDate = wsTarget.Cells(wsTarget_Row.Date, wsTarget_Col.Date + wsTarget_count_Col)
                            
                            '日付が一致した列番号のセルへ、発注書の数量を転記
                            If wsTarget_deliDate = wsInv_deliDate Then
                                
                                '既にターゲットシート内の該当セルへ数量が入っていた場合はスキップ
                                If wsTarget.Cells(wsTarget_count, wsTarget_Col.Date + wsTarget_count_Col).Value = 0 Then
                                    wsTarget.Cells(wsTarget_MaxRow, wsTarget_Col.Date + wsTarget_count_Col).Value = wsInv_quant
                                
                                    '新たに追記があった場合は、worksheet（'P2')セル内に更新有りの文言を追加
                                    wsTarget.Range("P2").Value = "更新有り"
                                    wsTarget.Range("P2").Font.ColorIndex = 2
                                End If
                            End If
                        Next
                End If
                
                '転記フラグを初期化
                transfer_flag = False
Continue:
            Next
            
    '上記の転記が完了したら、次の対象となる発注書シートを検索する
        End If
    Next ws

'上記の転記が全て完了したら、次の対象となる店舗の発注書を検索する
Next


'転記が完了したら、シート名をあいうえお順に並び替える
Dim count_sort As Long
Dim t As Long: t = 1

'ダミーシートを挿入する
With Worksheets.Add
    'ワークシート名をセルに書き出す
    For count_sort = 1 To Worksheets.count
        If InStr(Worksheets(count_sort).Name, "（") <> 0 Then
            .Cells(t, 1).Value = Worksheets(count_sort).Name
            t = t + 1
        End If
    Next count_sort
        
    'ワークシート名をソートする
    .Range("A1").CurrentRegion.Sort .Range("A1")
        
    'ワークシートの位置を並べ替える
    Worksheets(.Cells(1, 1).Value).Move Before:=Worksheets(1)
    For count_sort = 2 To .Cells(Rows.count, 1).End(xlUp).Row
        Worksheets(.Cells(count_sort, 1).Value).Move After:=Worksheets(count_sort - 1)
    Next count_sort
        
    'ダミーシートを削除する
    .Delete
    
End With
         
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "転記終了 "
 
End Sub


