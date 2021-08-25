Attribute VB_Name = "Module2"

Public Sub NippouShuukei_Update(nippo_nyuryoku_cell As Object, nippo_syukei_cell As Object)
    'セル初期化
    ThisWorkbook.Worksheets("日報集計").Range("A5:AN600").ClearContents
    Range("A5").Select

    '作業表作成
    Do Until nippo_nyuryoku_cell.Value = ""
        nippo_syukei_cell.Offset(0, 0).Value = nippo_nyuryoku_cell.Offset(0, 0).Value '生産日
        nippo_syukei_cell.Offset(0, 1).Value = nippo_nyuryoku_cell.Offset(0, 1).Value 'マシン
        nippo_syukei_cell.Offset(0, 2).Value = nippo_nyuryoku_cell.Offset(0, 2).Value '作業者
        nippo_syukei_cell.Offset(0, 3).Value = nippo_nyuryoku_cell.Offset(0, 3).Value '中子
        nippo_syukei_cell.Offset(0, 4).Value = nippo_nyuryoku_cell.Offset(0, 4).Value 'ショット
        nippo_syukei_cell.Offset(0, 5).Value = nippo_nyuryoku_cell.Offset(0, 5).Value '稼働時間
        nippo_syukei_cell.Offset(0, 6).Value = nippo_nyuryoku_cell.Offset(0, 7).Value '生産時間
        nippo_syukei_cell.Offset(0, 7).Value = nippo_nyuryoku_cell.Offset(0, 5).Value * nippo_nyuryoku_cell.Offset(0, 6) 'OP作業時間
        nippo_syukei_cell.Offset(0, 8).Value = nippo_nyuryoku_cell.Offset(0, 8).Value '始業作業
        nippo_syukei_cell.Offset(0, 9).Value = nippo_nyuryoku_cell.Offset(0, 9).Value '金型交換
        nippo_syukei_cell.Offset(0, 10).Value = nippo_nyuryoku_cell.Offset(0, 10).Value '昇温待ち
        nippo_syukei_cell.Offset(0, 11).Value = nippo_nyuryoku_cell.Offset(0, 11).Value '金型調整
        nippo_syukei_cell.Offset(0, 12).Value = nippo_nyuryoku_cell.Offset(0, 12).Value 'マシン故障停止
        nippo_syukei_cell.Offset(0, 13).Value = nippo_nyuryoku_cell.Offset(0, 13).Value '型清掃
        nippo_syukei_cell.Offset(0, 14).Value = nippo_nyuryoku_cell.Offset(0, 14).Value '終業作業
        nippo_syukei_cell.Offset(0, 15).Value = nippo_nyuryoku_cell.Offset(0, 15).Value 'Rb教示
        nippo_syukei_cell.Offset(0, 16).Value = nippo_nyuryoku_cell.Offset(0, 16).Value '他機対応待ち
        nippo_syukei_cell.Offset(0, 17).Value = nippo_nyuryoku_cell.Offset(0, 17).Value '離型剤
        nippo_syukei_cell.Offset(0, 18).Value = nippo_nyuryoku_cell.Offset(0, 18).Value '中子割れ処理
        nippo_syukei_cell.Offset(0, 19).Value = nippo_nyuryoku_cell.Offset(0, 19).Value 'その他
        nippo_syukei_cell.Offset(0, 20).Value = nippo_nyuryoku_cell.Offset(0, 20).Value '手直し不良
        nippo_syukei_cell.Offset(0, 21).Value = nippo_nyuryoku_cell.Offset(0, 21).Value '造形不良数
        nippo_syukei_cell.Offset(0, 22).Value = nippo_nyuryoku_cell.Offset(0, 22).Value 'ヒビ・カケ・スレ
        nippo_syukei_cell.Offset(0, 23).Value = nippo_nyuryoku_cell.Offset(0, 23).Value 'アカ不良
        nippo_syukei_cell.Offset(0, 24).Value = nippo_nyuryoku_cell.Offset(0, 24).Value '砂落ち不良
        nippo_syukei_cell.Offset(0, 25).Value = nippo_nyuryoku_cell.Offset(0, 25).Value '充填不良
        nippo_syukei_cell.Offset(0, 26).Value = nippo_nyuryoku_cell.Offset(0, 26).Value '焼成不良
        nippo_syukei_cell.Offset(0, 27).Value = nippo_nyuryoku_cell.Offset(0, 27).Value '型ズレ不良
        nippo_syukei_cell.Offset(0, 28).Value = nippo_nyuryoku_cell.Offset(0, 28).Value '作業中落下
        nippo_syukei_cell.Offset(0, 29).Value = nippo_nyuryoku_cell.Offset(0, 29).Value 'その他
        nippo_syukei_cell.Offset(0, 30).Value = nippo_nyuryoku_cell.Offset(0, -2).Value '良品数
        nippo_syukei_cell.Offset(0, 31).Value = nippo_nyuryoku_cell.Offset(0, 30).Value '原料砂
        nippo_syukei_cell.Offset(0, 32).Value = nippo_nyuryoku_cell.Offset(0, 31).Value '単重
        nippo_syukei_cell.Offset(0, 33).Value = nippo_nyuryoku_cell.Offset(0, 32).Value '単価
        nippo_syukei_cell.Offset(0, 34).Value = nippo_nyuryoku_cell.Offset(0, -3).Value * nippo_nyuryoku_cell.Offset(0, 4).Value * nippo_nyuryoku_cell.Offset(0, 31).Value '総量（使用量）
        nippo_syukei_cell.Offset(0, 35).Value = nippo_nyuryoku_cell.Offset(0, -2).Value * nippo_nyuryoku_cell.Offset(0, 31).Value '良品数（使用量）
        nippo_syukei_cell.Offset(0, 36).Value = nippo_syukei_cell.Offset(0, 34).Value - nippo_syukei_cell.Offset(0, 35).Value '不良数（使用量）
        nippo_syukei_cell.Offset(0, 37).Value = nippo_nyuryoku_cell.Offset(0, -2).Value * nippo_nyuryoku_cell.Offset(0, 32).Value '生産金額
        nippo_syukei_cell.Offset(0, 38).Value = nippo_nyuryoku_cell.Offset(0, 21).Value * nippo_nyuryoku_cell.Offset(0, 32).Value '不良金額
        nippo_syukei_cell.Offset(0, 39).Value = nippo_nyuryoku_cell.Offset(0, -4).Value '中子名
        Set nippo_syukei_cell = nippo_syukei_cell.Offset(1, 0)
        Set nippo_nyuryoku_cell = nippo_nyuryoku_cell.Offset(1, 0)
    Loop

End Sub
