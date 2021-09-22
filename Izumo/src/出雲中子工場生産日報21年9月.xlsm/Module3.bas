Attribute VB_Name = "Module3"
'Option Explicit

Public Sub 当月実績追加処理()

   Dim sagyohyo_sheet As String, mst_machine As String
   Dim nippo_nyuryoku_sheet As String, nippo_syukei_sheet As String
   Dim first_cell_of_sagyohyo, first_cell_of_target_summary, first_cell_of_machine As Object
   Dim nippo_nyuryoku_cell As Object, nippo_syukei_cell As Object
   Dim i As Integer, InM As Integer, Lcnt As Integer
   Dim Com1, Com2, Com3, Com5, Com6, Com7, Com8, Com9, Com10 As Long
   Dim Com11, Com12, Com13, Com14, Com15, Com16, Com17, Com18, Com19 As Long
   Dim Com20, Com21, Com22, Com23, Com24, Com28, Com29, Com30, Com31, Com32 As Long
   Dim Com4, Com25, Com26, Com27 As Single
   Dim SVtime, count As Long
   Dim WkCom As Double
   Dim myBtn As Integer
   Dim machine_code As Integer
   Dim nakago_name As String, nakago_code As String
   Dim update_target As String
   Dim M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11, M12 As String
   Dim S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12 As String
   'Dim logger As New Log
   Dim blank_row As Integer, bk_machine_code As Integer

   '初期設定
   Application.ScreenUpdating = False
   '德永専用デバッグ
   'Call logger.Init("D:\Daily_Production_Report_Summary\bin\test\debug.log")
   mst_machine = "マシン名"
   nippo_syukei_sheet = "日報集計"
   nippo_nyuryoku_sheet = "日報入力"
   sagyohyo_sheet = "作業表"
   '処理開始
   myBtn = MsgBox("当月実績追加処理を開始します", vbYesNo + vbExclamation, "当月実績追加処理")
   If myBtn = vbNo Then
      Exit Sub
   End If
   'Call logger.WriteLog("処理開始")

   '作業領域クリア（作業表）
   Worksheets(sagyohyo_sheet).Activate
   Range("A5:AM2000").Select
   Selection.ClearContents
   Range("A5").Select

   '処理開始位置の設定
   Set nippo_syukei_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_syukei_sheet).Range("A5")
   Set nippo_nyuryoku_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_nyuryoku_sheet).Range("G5")
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '日報集計シートの更新
   Call NippouShuukei_Update(nippo_nyuryoku_cell, nippo_syukei_cell)

   '処理開始位置の設定
   Set nippo_syukei_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_syukei_sheet).Range("A5")
   Set nippo_nyuryoku_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_nyuryoku_sheet).Range("G5")

   '実績データ確認
   Do Until nippo_syukei_cell.Value = ""
      With nippo_syukei_cell
      'データ移行
         For i = 0 To 39
            first_cell_of_sagyohyo.Offset(0, i).Value = .Offset(0, i).Value
         Next i
      End With
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
      Set nippo_syukei_cell = nippo_syukei_cell.Offset(1, 0)
   Loop

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'旧マシン別集計作業開始
   '作業用ワークシートアクティブ化（作業表）
   Worksheets(sagyohyo_sheet).Activate
   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   'インデックス初期化
   i = 4
   '実データ領域確認
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   'マシン別に並び替え
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("B")

   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '作業領域初期化
   Com1 = 0    'ショット
   Com2 = 0    '稼動時間
   Com3 = 0    '生産時間
   Com4 = 0    'ＯＰ作業時間
   Com5 = 0    '始業時間
   Com6 = 0    '金型交換
   Com7 = 0    '昇温待ち
   Com8 = 0    '金型調整
   Com9 = 0    'マシン故障停止
   Com10 = 0   '終業時間
   Com11 = 0   '型清掃
   Com12 = 0   'Ｒｂ教示
   Com13 = 0   '他機対応待ち
   Com14 = 0   '離型剤
   Com15 = 0   '中子割れ処理
   Com16 = 0   'その他
   Com17 = 0   '手直不良（良品に含まれる）
   Com18 = 0   '造型不良（廃棄不良）
   Com19 = 0   'ボス割れ表
   Com20 = 0   'ボス割れ裏
   Com21 = 0   '幅木割れ
   Com22 = 0   'フィン割れ
   Com23 = 0   '幅木充填
   Com24 = 0   'フィン充填
   Com25 = 0   'キャンドル残
   Com26 = 0   'その他
   Com27 = 0   '砂総量
   Com28 = 0   '砂良品
   Com29 = 0   '砂不良
   Com30 = 0   '生産金額
   Com31 = 0   '不良金額
   Com32 = 0   '良品数
   SVtime = 0  '出勤総時間
   count = 0   '金型交換回数
'
   nakago_code = first_cell_of_sagyohyo.Offset(0, 1).Value
   SVtime = first_cell_of_sagyohyo.Offset(-4, 0).Value
'
   update_target = "マシン別集計"

'追加先シート初期化
   '作業用ワークシートアクティブ化（マシン別－該当月）
   Worksheets(update_target).Activate
   '処理開始位置の設定
   Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
   'インデックス初期値
   i = 7
   '実データ領域確認
   Do Until first_cell_of_target_summary.Value = ""
      i = i + 1
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   Loop
   'クリア範囲指定
   Range(Cells(7, 1), Cells(i, 32)).Select
   Selection.ClearContents

'マシン名取り込み
   Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
   Set first_cell_of_machine = Worksheets(mst_machine).Range("B4")
   Do Until first_cell_of_machine.Value = ""
      If first_cell_of_machine.Offset(0, 1).Value <> "" Then
         first_cell_of_target_summary.Offset(0, 0).Value = first_cell_of_machine.Offset(0, 0).Value
         first_cell_of_target_summary.Offset(0, 1).Value = first_cell_of_machine.Offset(0, 1).Value
         Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      End If
      Set first_cell_of_machine = first_cell_of_machine.Offset(1, 0)
   Loop

'実績追加処理－マシン別
   'マシン別集計
   Do Until first_cell_of_sagyohyo.Value = ""
      '追加先シート処理開始位置指定
      Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
      Do Until nakago_code <> first_cell_of_sagyohyo.Offset(0, 1).Value
         Com1 = Com1 + first_cell_of_sagyohyo.Offset(0, 4).Value
         Com2 = Com2 + first_cell_of_sagyohyo.Offset(0, 5).Value
         Com3 = Com3 + first_cell_of_sagyohyo.Offset(0, 6).Value
         Com4 = Com4 + first_cell_of_sagyohyo.Offset(0, 7).Value
         Com5 = Com5 + first_cell_of_sagyohyo.Offset(0, 8).Value
         Com6 = Com6 + first_cell_of_sagyohyo.Offset(0, 9).Value
         If first_cell_of_sagyohyo.Offset(0, 9).Value > 0 Then
            count = count + 1
         End If
         Com7 = Com7 + first_cell_of_sagyohyo.Offset(0, 10).Value
         Com8 = Com8 + first_cell_of_sagyohyo.Offset(0, 11).Value
         Com9 = Com9 + first_cell_of_sagyohyo.Offset(0, 12).Value
         Com10 = Com10 + first_cell_of_sagyohyo.Offset(0, 13).Value
         Com11 = Com11 + first_cell_of_sagyohyo.Offset(0, 14).Value
         Com12 = Com12 + first_cell_of_sagyohyo.Offset(0, 15).Value
         Com13 = Com13 + first_cell_of_sagyohyo.Offset(0, 16).Value
         Com14 = Com14 + first_cell_of_sagyohyo.Offset(0, 17).Value
         Com15 = Com15 + first_cell_of_sagyohyo.Offset(0, 18).Value
         Com16 = Com16 + first_cell_of_sagyohyo.Offset(0, 19).Value
         Com17 = Com17 + first_cell_of_sagyohyo.Offset(0, 20).Value
         Com18 = Com18 + first_cell_of_sagyohyo.Offset(0, 21).Value
         Com32 = Com32 + first_cell_of_sagyohyo.Offset(0, 30).Value
         Com27 = Com27 + first_cell_of_sagyohyo.Offset(0, 34).Value
         Com28 = Com28 + first_cell_of_sagyohyo.Offset(0, 35).Value
         Com29 = Com29 + first_cell_of_sagyohyo.Offset(0, 36).Value
         Com30 = Com30 + first_cell_of_sagyohyo.Offset(0, 37).Value
         Com31 = Com31 + first_cell_of_sagyohyo.Offset(0, 38).Value
         Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
      Loop
      'マシンコード位置設定
      Do Until nakago_code = first_cell_of_target_summary.Offset(0, 0).Value
         Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      Loop
      With first_cell_of_target_summary
         .Offset(0, 2).Value = Com1           'ショット数
         .Offset(0, 3).Value = Com32          '良品数
         .Offset(0, 4).Value = Com18          '不良数
         .Offset(0, 5).Value = Com2 / 60      'マシン稼働時間
         .Offset(0, 6).Value = Com3 / 60      'マシン生産時間
         .Offset(0, 7).Value = Com4 / 60      'ＯＰ作業時間
         .Offset(0, 8).Value = Com5 / 60      '始業作業
         .Offset(0, 9).Value = Com6 / 60      '金型交換
         .Offset(0, 10).Value = Com7 / 60     '昇温待ち
         .Offset(0, 11).Value = count         '型交換回数（どこから？）
         .Offset(0, 12).Value = Com8 / 60     '型調整
         .Offset(0, 13).Value = Com9 / 60     '故障停止
         .Offset(0, 14).Value = Com11 / 60    '金型清掃
         .Offset(0, 15).Value = Com10 / 60    '終了作業
         .Offset(0, 16).Value = Com12 / 60    'Ｒｂ教示
         .Offset(0, 17).Value = Com13 / 60    '他機対応待ち
         .Offset(0, 18).Value = Com14 / 60    '離型剤
         .Offset(0, 19).Value = Com15 / 60    '中子割れ処理
         .Offset(0, 20).Value = Com16 / 60    'その他
         .Offset(0, 21).Value = Com27 / 1000  '使用量
         .Offset(0, 22).Value = Com28 / 1000  '良品使用量
         .Offset(0, 23).Value = Com29 / 1000  '不良使用量
         .Offset(0, 24).Value = Com30 / 1000  '生産金額
         .Offset(0, 25).Value = Com31 / 1000  '不良金額
         .Offset(0, 27).Value = (Com2 / 60) / SVtime '設備負荷率
         .Offset(0, 28).Value = Com3 / Com2   '設備稼働率
         .Offset(0, 29).Value = Com30 / (Com2 / 60)  '労働生産性（マシン）
         .Offset(0, 30).Value = Com30 / (Com4 / 60)  '労働生産性（人）

         If Com18 <> 0 Then
            WkCom = Com18 / (Com18 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 26).Value = WkCom     '不良率
      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      nakago_code = first_cell_of_sagyohyo.Offset(0, 1).Value
      '作業エリア初期化
      Com1 = 0    'ショット
      Com2 = 0    '稼動時間
      Com3 = 0    '生産時間
      Com4 = 0    'ＯＰ作業時間
      Com5 = 0    '始業時間
      Com6 = 0    '金型交換
      Com7 = 0    '昇温待ち
      Com8 = 0    '金型調整
      Com9 = 0    'マシン故障停止
      Com10 = 0   '終業時間
      Com11 = 0   '型清掃
      Com12 = 0   'Ｒｂ教示
      Com13 = 0   '他機対応待ち
      Com14 = 0   '離型剤
      Com15 = 0   '中子割れ処理
      Com16 = 0   'その他
      Com17 = 0   '手直不良（良品に含まれる）
      Com18 = 0   '造型不良（廃棄不良）
      Com27 = 0   '砂総量
      Com28 = 0   '砂良品
      Com29 = 0   '砂不良
      Com30 = 0   '生産金額
      Com31 = 0   '不良金額
      Com32 = 0   '良品数
      count = 0   '金型交換回数
   Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '新マシン別集計作業開始
   '作業用ワークシートアクティブ化（作業表）
   Worksheets(sagyohyo_sheet).Activate
   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   'インデックス初期化
   i = 4
   '実データ領域確認
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop
   'マシン順と中子順と複数条件でソート
   With ActiveSheet
      .Sort.SortFields.Clear
      'マシン順
      .Sort.SortFields.Add _
         Key:=ActiveSheet.Range("B5")
      '中子順
      .Sort.SortFields.Add _
         Key:=ActiveSheet.Range("D5")
      With .Sort
         .SetRange Range(Cells(5, 1), Cells(i, 41))
         .Apply
      End With
   End With
   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   SVtime = first_cell_of_sagyohyo.Offset(-4, 0).Value  '出勤総時間
   count = 0   '金型交換回数
   update_target = "新マシン別集計"
   '追加先シート初期化
   '作業用ワークシートアクティブ化（マシン別－該当月）
   Worksheets(update_target).Activate
   '処理開始位置の設定
   Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
   last_row = Range("B7").End(xlDown).Row
   'クリア範囲指定
   Range(first_cell_of_target_summary, Range("AF" & last_row)).Select
   Selection.ClearContents
   '実績追加処理－マシン別
   'マシン別集計
   Dim read_index As Variant
   read_index = Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 30, 34, 35, 36, 37, 38)
   '[0] = 4 // ショット数
   '[1] = 5 // 稼働時間
   '[2] = 6 // 生産時間
   '[3] = 7 // OP作業時間
   '[4] = 8 // 始業作業
   '[5] = 9 // 金型交換
   '[6] = 10 // 昇温待ち
   '[7] = 11 // 金型調整
   '[8] = 12 // マシン故障停止
   '[9] = 13 // 終業作業
   '[10] = 14 // 型清掃
   '[11] = 15 // Rb教示
   '[12] = 16 // 他機対応待ち
   '[13] = 17 // 離散材
   '[14] = 18 // 中子割れ処理
   '[15] = 19 // その他
   '[16] = 20 // 手直不良
   '[17] = 21 // 造形不良数
   '[18] = 30 // 良品数
   '[19] = 34 // 総量
   '[20] = 35 // 良品数
   '[21] = 36 // 不良数
   '[22] = 37 // 生産金額
   '[23] = 38 // 不良金額

   blank_row = 7
   Do Until first_cell_of_sagyohyo.Value = ""
      Dim nippo_by_nakago(23) As Long
      Erase nippo_by_nakago
      nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value
      machine_code = first_cell_of_sagyohyo.Offset(0, 1).Value
      bk_machine_code = 0
      nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value
      'ループ条件：中子コードが変わるまで。
      Do Until nakago_code <> first_cell_of_sagyohyo.Offset(0, 3).Value
         Dim k As Integer
         k = 0
         For Each index In read_index
            If first_cell_of_sagyohyo.Offset(0, index) <> "" Then
               'Call logger.WriteLog("machine_code = " & machine_code & ", nakago_code = " & nakago_code & ", k = " & k & ", index = " & index & " : " & first_cell_of_sagyohyo.Offset(0, index))
               nippo_by_nakago(k) = nippo_by_nakago(k) + first_cell_of_sagyohyo.Offset(0, index)
               'Call logger.WriteLog("NAKAGO_SUMMARY : " & nippo_by_nakago(k))
               If i = 9 Then
                  If first_cell_of_sagyohyo.Offset(0, i) > 0 Then
                     count = count + 1
                  End If
               End If
            End If
            k = k + 1
         Next index
         '1行読み終わったら次行へ
         Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
      Loop
      'マシンコードが前回と同じかどうか
      If bk_machine_code = machine_code Then
         Cells(blank_row, 1).EntireRow.Insert
      End If
      bk_machine_code = machine_code
      blank_row = blank_row + 1
      With first_cell_of_target_summary
         .Offset(0, 0).Value = machine_code
         .Offset(0, 1).Value = WorksheetFunction.VLookup(machine_code, Workbooks(ActiveWorkbook.Name).Worksheets("マシン名").Range("B:C"), 2)
         .Offset(0, 2).Value = nakago_name
         .Offset(0, 3).Value = nippo_by_nakago(0)      'ショット数
         .Offset(0, 4).Value = nippo_by_nakago(18)     '良品数
         .Offset(0, 5).Value = nippo_by_nakago(17)     '不良数
         .Offset(0, 6).Value = nippo_by_nakago(1) / 60     'マシン稼働時間
         .Offset(0, 7).Value = nippo_by_nakago(2) / 60     'マシン生産時間
         .Offset(0, 8).Value = nippo_by_nakago(3) / 60     'ＯＰ作業時間
         .Offset(0, 9).Value = nippo_by_nakago(4) / 60     '始業作業
         .Offset(0, 10).Value = nippo_by_nakago(5) / 60     '金型交換
         .Offset(0, 11).Value = nippo_by_nakago(6) / 60    '昇温待ち
         .Offset(0, 12).Value = count      '型交換回数（どこから？）
         .Offset(0, 13).Value = nippo_by_nakago(7) / 60    '型調整
         .Offset(0, 14).Value = nippo_by_nakago(8) / 60    '故障停止
         .Offset(0, 15).Value = nippo_by_nakago(10) / 60   '金型清掃
         .Offset(0, 16).Value = nippo_by_nakago(9) / 60   '終了作業
         .Offset(0, 17).Value = nippo_by_nakago(11) / 60   'Ｒｂ教示
         .Offset(0, 18).Value = nippo_by_nakago(12) / 60   '他機対応待ち
         .Offset(0, 19).Value = nippo_by_nakago(13) / 60   '離型剤
         .Offset(0, 20).Value = nippo_by_nakago(14) / 60   '中子割れ処理
         .Offset(0, 21).Value = nippo_by_nakago(15) / 60   'その他
         .Offset(0, 22).Value = nippo_by_nakago(19) / 1000  '使用量
         .Offset(0, 23).Value = nippo_by_nakago(20) / 1000  '良品使用量
         .Offset(0, 24).Value = nippo_by_nakago(21) / 1000  '不良使用量
         .Offset(0, 25).Value = nippo_by_nakago(22) / 1000  '生産金額
         .Offset(0, 26).Value = nippo_by_nakago(23) / 1000  '不良金額
         If nippo_by_nakago(17) <> 0 Then
            WkCom = nippo_by_nakago(17) / (nippo_by_nakago(17) + nippo_by_nakago(18))
         Else
            WkCom = 0
         End If
         .Offset(0, 27).Value = WkCom    '不良率
         .Offset(0, 28).Value = (nippo_by_nakago(1) / 60) / SVtime '設備負荷率
         .Offset(0, 29).Value = nippo_by_nakago(2) / nippo_by_nakago(1)   '設備稼働率
         .Offset(0, 30).Value = nippo_by_nakago(22) / (nippo_by_nakago(1) / 60)  '労働生産性（マシン）
         .Offset(0, 31).Value = nippo_by_nakago(22) / (nippo_by_nakago(3) / 60)  '労働生産性（人）
      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      '作業エリア初期化
      count = 0   '金型交換回数
   Loop
   '最終行追加
   last_row = Range("B7").End(xlDown).Row + 1
   Range("B" & last_row) = "合計"
   With Range("D" & last_row)
      .Formula = "=SUM(D7:D" & (last_row - 1) & " )"
      .AutoFill Destination:=.Resize(1, 24)
   End With
   Range("AB" & last_row) = Range("F" & last_row).Value / ( Range("E" & last_row).Value + Range("F" & last_row).Value )
   Range("AC" & last_row).Formula = "=AVERAGE(AC7:AC" & (last_row - 1) & " )"
   Range("AD" & last_row) = Range("H" & last_row).Value /  Range("G" & last_row).Value
   Range("AE" & last_row) = Range("Z" & last_row).Value * 1000 /  Range("H" & last_row).Value
   Range("AF" & last_row) = Range("Z" & last_row).Value * 1000 /  Range("I" & last_row).Value
   Range("AG" & last_row) = Range("H" & last_row).Value * 3600 /  Range("D" & last_row).Value

   '品名別集計作業開始
   '作業用ワークシートアクティブ化（作業表）
   Worksheets(sagyohyo_sheet).Activate
   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   'インデックス初期化
   i = 4

   '実データ領域確認
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   '品名別に並び替え
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("D")

   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '作業領域初期化
   Com1 = 0   'ショット
   Com2 = 0   '稼動時間
   Com3 = 0   '生産時間
   Com4 = 0   'ＯＰ作業時間
   Com5 = 0   '始業時間
   Com6 = 0   '金型交換
   Com7 = 0   '昇温待ち
   Com8 = 0   '金型調整
   Com9 = 0   'マシン故障停止
   Com10 = 0   '終業時間
   Com11 = 0   '型清掃
   Com12 = 0   'Ｒｂ教示
   Com13 = 0   '他機対応待ち
   Com14 = 0   '離型剤
   Com15 = 0   '中子割れ処理
   Com16 = 0   'その他
   Com17 = 0   '手直不良（良品に含まれる）
   Com18 = 0   '造型不良（廃棄不良）
   Com19 = 0   'ボス割れ表
   Com20 = 0   'ボス割れ裏
   Com21 = 0   '幅木割れ
   Com22 = 0   'フィン割れ
   Com23 = 0   '幅木充填
   Com24 = 0   'フィン充填
   Com25 = 0   'キャンドル残
   Com26 = 0   'その他
   Com27 = 0   '砂総量
   Com28 = 0   '砂良品
   Com29 = 0   '砂不良
   Com30 = 0   '生産金額
   Com31 = 0   '不良金額
   Com32 = 0   '良品数
   count = 0   '金型交換回数

   nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value      '中子コード
   nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value      '中子名

   update_target = "品名別集計"

   '追加先シート初期化
   '作業用ワークシートアクティブ化（マシン別－該当月）
   Worksheets(update_target).Activate
   '処理開始位置の設定
   Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
   last_row = Range("B7").End(xlDown).Row
   'クリア範囲指定
   Range(first_cell_of_target_summary, Range("AF" & last_row)).Select
   Selection.ClearContents
   '実績追加処理－品名別
   '追加先シート処理開始位置指定
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A7")
   '品名別集計
   Do Until first_cell_of_sagyohyo.Value = ""
      Do Until nakago_code <> first_cell_of_sagyohyo.Offset(0, 3).Value
         Com1 = Com1 + first_cell_of_sagyohyo.Offset(0, 4).Value
         Com2 = Com2 + first_cell_of_sagyohyo.Offset(0, 5).Value
         Com3 = Com3 + first_cell_of_sagyohyo.Offset(0, 6).Value
         Com4 = Com4 + first_cell_of_sagyohyo.Offset(0, 7).Value
         Com5 = Com5 + first_cell_of_sagyohyo.Offset(0, 8).Value
         Com6 = Com6 + first_cell_of_sagyohyo.Offset(0, 9).Value
         If first_cell_of_sagyohyo.Offset(0, 9).Value > 0 Then
            count = count + 1
         End If
         Com7 = Com7 + first_cell_of_sagyohyo.Offset(0, 10).Value
         Com8 = Com8 + first_cell_of_sagyohyo.Offset(0, 11).Value
         Com9 = Com9 + first_cell_of_sagyohyo.Offset(0, 12).Value
         Com10 = Com10 + first_cell_of_sagyohyo.Offset(0, 13).Value
         Com11 = Com11 + first_cell_of_sagyohyo.Offset(0, 14).Value
         Com12 = Com12 + first_cell_of_sagyohyo.Offset(0, 15).Value
         Com13 = Com13 + first_cell_of_sagyohyo.Offset(0, 16).Value
         Com14 = Com14 + first_cell_of_sagyohyo.Offset(0, 17).Value
         Com15 = Com15 + first_cell_of_sagyohyo.Offset(0, 18).Value
         Com16 = Com16 + first_cell_of_sagyohyo.Offset(0, 19).Value
         Com17 = Com17 + first_cell_of_sagyohyo.Offset(0, 20).Value
         Com18 = Com18 + first_cell_of_sagyohyo.Offset(0, 21).Value
         Com32 = Com32 + first_cell_of_sagyohyo.Offset(0, 30).Value
         Com27 = Com27 + first_cell_of_sagyohyo.Offset(0, 34).Value
         Com28 = Com28 + first_cell_of_sagyohyo.Offset(0, 35).Value
         Com29 = Com29 + first_cell_of_sagyohyo.Offset(0, 36).Value
         Com30 = Com30 + first_cell_of_sagyohyo.Offset(0, 37).Value
         Com31 = Com31 + first_cell_of_sagyohyo.Offset(0, 38).Value
         Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
      Loop

      With first_cell_of_target_summary  '20140408kometani  中子コードを記入するセルを追加したことで右に1個ずつずらした
         .Offset(0, 1).Value = nakago_name      '中子名
         .Offset(0, 2).Value = nakago_code      '中子コード　'20140408kometani　追加
         .Offset(0, 3).Value = Com1      'ショット数
         .Offset(0, 4).Value = Com32     '良品数
         .Offset(0, 5).Value = Com18     '不良数
         .Offset(0, 6).Value = Com2 / 60     'マシン稼働時間
         .Offset(0, 7).Value = Com3 / 60     'マシン生産時間
         .Offset(0, 8).Value = Com4 / 60     'ＯＰ作業時間
         .Offset(0, 9).Value = Com5 / 60     '始業作業
         .Offset(0, 10).Value = Com6 / 60    '金型交換
         .Offset(0, 11).Value = Com7 / 60    '昇温待ち
         .Offset(0, 12).Value = count      '型交換回数
         .Offset(0, 13).Value = Com8 / 60    '型調整
         .Offset(0, 14).Value = Com9 / 60    '故障停止
         .Offset(0, 15).Value = Com11 / 60   '金型清掃
         .Offset(0, 16).Value = Com10 / 60   '終了作業
         .Offset(0, 17).Value = Com12 / 60   'Ｒｂ教示
         .Offset(0, 18).Value = Com13 / 60   '他機対応待ち
         .Offset(0, 19).Value = Com14 / 60   '離型剤
         .Offset(0, 20).Value = Com15 / 60   '中子割れ処理
         .Offset(0, 21).Value = Com16 / 60   'その他
         .Offset(0, 22).Value = Com27      '使用量
         .Offset(0, 23).Value = Com28      '良品使用量
         .Offset(0, 24).Value = Com29      '不良使用量
         .Offset(0, 25).Value = Com30      '生産金額
         .Offset(0, 26).Value = Com31      '不良金額
         .Offset(0, 28).Value = (Com2 / 60) / SVtime '設備負荷率
         If Com2 <> 0 Then
            .Offset(0, 29).Value = Com3 / Com2   '設備稼働率
         Else
            Com2 = 0
         End If
         If Com18 <> 0 Then
            WkCom = Com18 / (Com18 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 27).Value = WkCom

         If Com30 <> 0 Then
            .Offset(0, 30).Value = Com30 / (Com2 / 60)  '労働生産性（マシン）
            .Offset(0, 31).Value = Com30 / (Com4 / 60)  '労働生産性（人）
         Else
            .Offset(0, 30).Value = 0
            .Offset(0, 31).Value = 0
         End If
      End With

      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value
      nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value

      '作業エリア初期化
      Com1 = 0   'ショット
      Com2 = 0   '稼動時間
      Com3 = 0   '生産時間
      Com4 = 0   'ＯＰ作業時間
      Com5 = 0   '始業時間
      Com6 = 0   '金型交換
      Com7 = 0   '昇温待ち
      Com8 = 0   '金型調整
      Com9 = 0   'マシン故障停止
      Com10 = 0   '終業時間
      Com11 = 0   '型清掃
      Com12 = 0   'Ｒｂ教示
      Com13 = 0   '他機対応待ち
      Com14 = 0   '離型剤
      Com15 = 0   '中子割れ処理
      Com16 = 0   'その他
      Com17 = 0   '手直不良（良品に含まれる）
      Com18 = 0   '造型不良（廃棄不良）
      Com27 = 0   '砂総量
      Com28 = 0   '砂良品
      Com29 = 0   '砂不良
      Com30 = 0   '生産金額
      Com31 = 0   '不良金額
      Com32 = 0   '良品数
      count = 0   '金型交換回数
   Loop

   '最終行追加
   last_row = Range("B7").End(xlDown).Row + 1
   
   Range("B" & last_row) = "合計"
   With Range("D" & last_row)
      .Formula = "=SUM(D7:D" & (last_row - 1) & " )"
      .AutoFill Destination:=.Resize(1, 24)
   End With
   Range("AB" & last_row) = Range("F" & last_row).Value / ( Range("E" & last_row).Value + Range("F" & last_row).Value )
   Range("AC" & last_row).Formula = "=AVERAGE(AC7:AC" & (last_row - 1) & " )"
   Range("AD" & last_row) = Range("H" & last_row).Value /  Range("G" & last_row).Value
   Range("AE" & last_row) = Range("Z" & last_row).Value * 1000 /  Range("H" & last_row).Value
   Range("AF" & last_row) = Range("Z" & last_row).Value * 1000 /  Range("I" & last_row).Value
   Range("AG" & last_row) = Range("H" & last_row).Value * 3600 /  Range("D" & last_row).Value

   '作業用ワークシートアクティブ化（品名別－該当月）
   Worksheets(update_target).Activate

   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("B7")

   'インデックス初期化
   i = 7

   '実データ領域確認
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   '生産金額順（降順）に並び替え
   Range(Cells(7, 1), Cells(i, 32)).Sort _
   Key1:=Columns("Z"), Order1:=xlDescending

   '品名に通番付与（生産金額順）
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("B7")
   'カウント初期化
   Lcnt = 1
   '実行
   Do Until first_cell_of_target_summary.Value = ""
      first_cell_of_target_summary.Offset(0, -1).Value = Lcnt   '通番
      Lcnt = Lcnt + 1
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   Loop

   '20091120追加不良別集計
   'マシン別不良集計作業開始
   '作業用ワークシートアクティブ化（作業表）
   Worksheets(sagyohyo_sheet).Activate
   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   'インデックス初期化
   i = 4
   '実データ領域確認
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   'マシン別に並び替え
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("B")

   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '作業領域初期化
   Com17 = 0   '手直不良（良品に含まれる）
   Com18 = 0   '廃棄不良
   Com19 = 0   'ボス割れ表
   Com20 = 0   'ボス割れ裏
   Com21 = 0   '幅木割れ
   Com22 = 0   'フィン割れ
   Com23 = 0   '幅木充填
   Com24 = 0   'フィン充填
   Com25 = 0   'キャンドル残
   Com26 = 0   'その他
   Com32 = 0   '良品数

   update_target = "不良集計【マシン】"

   '追加先シート初期化
   '作業用ワークシートアクティブ化
   Worksheets(update_target).Activate
   '処理開始位置の設定
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")
   'インデックス初期値
   i = 5
   '実データ領域確認
   Do Until first_cell_of_target_summary.Value = ""
      i = i + 1
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   Loop
   'クリア範囲指定
   Range(Cells(6, 1), Cells(i, 15)).Select
   Selection.ClearContents

   'マシン名取り込み
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")
   Set first_cell_of_machine = Workbooks(ActiveWorkbook.Name).Worksheets(mst_machine).Range("B4")
   Do Until first_cell_of_machine.Value = ""
      If first_cell_of_machine.Offset(0, 1).Value <> "" Then
         first_cell_of_target_summary.Offset(0, 0).Value = first_cell_of_machine.Offset(0, 0).Value
         first_cell_of_target_summary.Offset(0, 1).Value = first_cell_of_machine.Offset(0, 1).Value
         Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      End If
      Set first_cell_of_machine = first_cell_of_machine.Offset(1, 0)
   Loop

   '実績追加処理－マシン別
   '追加先シート処理開始位置指定
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")

   machine_code = first_cell_of_sagyohyo.Offset(0, 1).Value
   'マシン別集計
   Do Until first_cell_of_sagyohyo.Value = ""
      Do Until machine_code <> first_cell_of_sagyohyo.Offset(0, 1).Value
         Com17 = Com17 + first_cell_of_sagyohyo.Offset(0, 20).Value
         Com18 = Com18 + first_cell_of_sagyohyo.Offset(0, 21).Value
         Com19 = Com19 + first_cell_of_sagyohyo.Offset(0, 22).Value
         Com20 = Com20 + first_cell_of_sagyohyo.Offset(0, 23).Value
         Com21 = Com21 + first_cell_of_sagyohyo.Offset(0, 24).Value
         Com22 = Com22 + first_cell_of_sagyohyo.Offset(0, 25).Value
         Com23 = Com23 + first_cell_of_sagyohyo.Offset(0, 26).Value
         Com24 = Com24 + first_cell_of_sagyohyo.Offset(0, 27).Value
         Com25 = Com25 + first_cell_of_sagyohyo.Offset(0, 28).Value
         Com26 = Com26 + first_cell_of_sagyohyo.Offset(0, 29).Value
         Com32 = Com32 + first_cell_of_sagyohyo.Offset(0, 30).Value
         Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)

      Loop
      'マシンコード位置設定
      Do Until machine_code = first_cell_of_target_summary.Offset(0, 0).Value
         Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)

      Loop
      With first_cell_of_target_summary
         .Offset(0, 2).Value = Com32     '良品数
         .Offset(0, 3).Value = Com18     '不良数
         .Offset(0, 4).Value = Com19     'ボス割れ表
         .Offset(0, 5).Value = Com20     'ボス割れ裏
         .Offset(0, 6).Value = Com21     '幅木割れ
         .Offset(0, 7).Value = Com22     'フィン割れ
         .Offset(0, 8).Value = Com23     '幅木充填
         .Offset(0, 9).Value = Com24     'フィン充填
         .Offset(0, 10).Value = Com25    'キャンドル残
         .Offset(0, 11).Value = Com26    'その他
         .Offset(0, 12).Value = Com17    '手直不良
         If Com18 <> 0 Then
            WkCom = Com18 / (Com18 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 13).Value = WkCom    '廃棄不良率

         If Com17 <> 0 Then
            WkCom = Com17 / (Com17 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 14).Value = WkCom    '手直不良率

      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      machine_code = first_cell_of_sagyohyo.Offset(0, 1).Value
      '作業エリア初期化
      Com17 = 0   '手直不良（良品に含まれる）
      Com18 = 0   '廃棄不良
      Com19 = 0   'ボス割れ表
      Com20 = 0   'ボス割れ裏
      Com21 = 0   '幅木割れ
      Com22 = 0   'フィン割れ
      Com23 = 0   '幅木充填
      Com24 = 0   'フィン充填
      Com25 = 0   'キャンドル残
      Com26 = 0   'その他
      Com32 = 0   '良品数
   Loop

   '位置の設定
   Range("A1").Select

   '品名別不良集計作業開始
   '作業用ワークシートアクティブ化（作業表）
   Worksheets(sagyohyo_sheet).Activate
   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   'インデックス初期化
   i = 4

   '実データ領域確認
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   '品名別に並び替え
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("D")

   '処理開始位置の設定
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '作業領域初期化
   Com17 = 0   '手直不良（良品に含まれる）
   Com18 = 0   '廃棄不良
   Com19 = 0   'ボス割れ表
   Com20 = 0   'ボス割れ裏
   Com21 = 0   '幅木割れ
   Com22 = 0   'フィン割れ
   Com23 = 0   '幅木充填
   Com24 = 0   'フィン充填
   Com25 = 0   'キャンドル残
   Com26 = 0   'その他
   Com32 = 0   '良品数

   '追加先シート初期化
   '作業用ワークシートアクティブ化（品名別－該当月）
   update_target = "不良集計【品名】"
   Worksheets(update_target).Activate
   '処理開始位置の設定
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")
   'インデックス初期値
   i = 5
   '実データ領域確認
   Do Until first_cell_of_target_summary.Value = ""
      i = i + 1
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   Loop

   'クリア範囲指定
   Range(Cells(6, 1), Cells(i, 14)).Select
   Selection.ClearContents

   '実績追加処理－品名別
   '追加先シート処理開始位置指定
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")

   '品名別集計
   Do Until first_cell_of_sagyohyo.Value = ""
   nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value      '中子コード
   nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value      '中子名
      Do Until nakago_code <> first_cell_of_sagyohyo.Offset(0, 3).Value
         Com17 = Com17 + first_cell_of_sagyohyo.Offset(0, 20).Value
         Com18 = Com18 + first_cell_of_sagyohyo.Offset(0, 21).Value
         Com19 = Com19 + first_cell_of_sagyohyo.Offset(0, 22).Value
         Com20 = Com20 + first_cell_of_sagyohyo.Offset(0, 23).Value
         Com21 = Com21 + first_cell_of_sagyohyo.Offset(0, 24).Value
         Com22 = Com22 + first_cell_of_sagyohyo.Offset(0, 25).Value
         Com23 = Com23 + first_cell_of_sagyohyo.Offset(0, 26).Value
         Com24 = Com24 + first_cell_of_sagyohyo.Offset(0, 27).Value
         Com25 = Com25 + first_cell_of_sagyohyo.Offset(0, 28).Value
         Com26 = Com26 + first_cell_of_sagyohyo.Offset(0, 29).Value
         Com32 = Com32 + first_cell_of_sagyohyo.Offset(0, 30).Value
         Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
      Loop

      With first_cell_of_target_summary
         .Offset(0, 0).Value = nakago_code      '中子コード
         .Offset(0, 1).Value = nakago_name      '中子名
         .Offset(0, 2).Value = Com32     '良品数
         .Offset(0, 3).Value = Com18     '不良数
         .Offset(0, 4).Value = Com19     'ボス割れ表
         .Offset(0, 5).Value = Com20     'ボス割れ裏
         .Offset(0, 6).Value = Com21     '幅木割れ
         .Offset(0, 7).Value = Com22     'フィン割れ
         .Offset(0, 8).Value = Com23     '幅木充填
         .Offset(0, 9).Value = Com24     'フィン充填
         .Offset(0, 10).Value = Com25    'キャンドル残
         .Offset(0, 11).Value = Com26    'その他
         .Offset(0, 12).Value = Com17    '手直不良
         If Com18 <> 0 Then
            WkCom = Com18 / (Com18 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 13).Value = WkCom    '廃棄不良率
         If Com17 <> 0 Then
            WkCom = Com17 / (Com17 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 14).Value = WkCom    '手直不良率
      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   '作業エリア初期化
      Com17 = 0   '手直不良（良品に含まれる）
      Com18 = 0   '廃棄不良
      Com19 = 0   'ボス割れ表
      Com20 = 0   'ボス割れ裏
      Com21 = 0   '幅木割れ
      Com22 = 0   'フィン割れ
      Com23 = 0   '幅木充填
      Com24 = 0   'フィン充填
      Com25 = 0   'キャンドル残
      Com26 = 0   'その他
      Com32 = 0   '良品数
   Loop

   MsgBox "処理を終わりました。", vbOKOnly + vbInformation, "通知"
End Sub
