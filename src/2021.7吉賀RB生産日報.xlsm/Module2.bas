Attribute VB_Name = "Module2"

Public Sub NippouShuukei_Update(NNCl As Object, NSCl As Object)
    
    Dim n As Long
    
    'セル初期化
    ThisWorkbook.Worksheets("日報集計").Range("A5:AN600").ClearContents
    Range("A5").Select
    
    '作業表作成
    n = 1
    Do Until NNCl.Value = ""
        'Application.StatusBar = "日報入力から日報集計を作成中・・・　" & n & "レコード目"
        NSCl.Offset(0, 0).Value = NNCl.Offset(0, 0).Value '生産日
        NSCl.Offset(0, 1).Value = NNCl.Offset(0, 1).Value 'マシン
        NSCl.Offset(0, 2).Value = NNCl.Offset(0, 2).Value '作業者
        NSCl.Offset(0, 3).Value = NNCl.Offset(0, 3).Value '中子
        NSCl.Offset(0, 4).Value = NNCl.Offset(0, 4).Value 'ショット
        NSCl.Offset(0, 5).Value = NNCl.Offset(0, 5).Value '稼働時間
        NSCl.Offset(0, 6).Value = NNCl.Offset(0, 7).Value '生産時間
        NSCl.Offset(0, 7).Value = NNCl.Offset(0, 5).Value * NNCl.Offset(0, 6) 'OP作業時間
        NSCl.Offset(0, 8).Value = NNCl.Offset(0, 8).Value '始業作業
        NSCl.Offset(0, 9).Value = NNCl.Offset(0, 9).Value '金型交換
        NSCl.Offset(0, 10).Value = NNCl.Offset(0, 10).Value '昇温待ち
        NSCl.Offset(0, 11).Value = NNCl.Offset(0, 11).Value '金型調整
        NSCl.Offset(0, 12).Value = NNCl.Offset(0, 12).Value 'マシン故障停止
        NSCl.Offset(0, 13).Value = NNCl.Offset(0, 13).Value '型清掃
        NSCl.Offset(0, 14).Value = NNCl.Offset(0, 14).Value '終業作業
        NSCl.Offset(0, 15).Value = NNCl.Offset(0, 15).Value 'Rb教示
        NSCl.Offset(0, 16).Value = NNCl.Offset(0, 16).Value '他機対応待ち
        NSCl.Offset(0, 17).Value = NNCl.Offset(0, 17).Value '離型剤
        NSCl.Offset(0, 18).Value = NNCl.Offset(0, 18).Value '中子割れ処理
        NSCl.Offset(0, 19).Value = NNCl.Offset(0, 19).Value 'その他
        NSCl.Offset(0, 20).Value = NNCl.Offset(0, 20).Value '手直し不良
        NSCl.Offset(0, 21).Value = NNCl.Offset(0, 21).Value '造形不良数
        NSCl.Offset(0, 22).Value = NNCl.Offset(0, 22).Value 'ヒビ・カケ・スレ
        NSCl.Offset(0, 23).Value = NNCl.Offset(0, 23).Value 'アカ不良
        NSCl.Offset(0, 24).Value = NNCl.Offset(0, 24).Value '砂落ち不良
        NSCl.Offset(0, 25).Value = NNCl.Offset(0, 25).Value '充填不良
        NSCl.Offset(0, 26).Value = NNCl.Offset(0, 26).Value '焼成不良
        NSCl.Offset(0, 27).Value = NNCl.Offset(0, 27).Value '型ズレ不良
        NSCl.Offset(0, 28).Value = NNCl.Offset(0, 28).Value '作業中落下
        NSCl.Offset(0, 29).Value = NNCl.Offset(0, 29).Value 'その他
        NSCl.Offset(0, 30).Value = NNCl.Offset(0, -2).Value '良品数
        NSCl.Offset(0, 31).Value = NNCl.Offset(0, 30).Value '原料砂
        NSCl.Offset(0, 32).Value = NNCl.Offset(0, 31).Value '単重
        NSCl.Offset(0, 33).Value = NNCl.Offset(0, 32).Value '単価
        NSCl.Offset(0, 34).Value = NNCl.Offset(0, -3).Value * NNCl.Offset(0, 4).Value * NNCl.Offset(0, 31).Value '総量（使用量）
        NSCl.Offset(0, 35).Value = NNCl.Offset(0, -2).Value * NNCl.Offset(0, 31).Value '良品数（使用量）
        NSCl.Offset(0, 36).Value = NSCl.Offset(0, 34).Value - NSCl.Offset(0, 35).Value '不良数（使用量）
        NSCl.Offset(0, 37).Value = NNCl.Offset(0, -2).Value * NNCl.Offset(0, 32).Value '生産金額
        NSCl.Offset(0, 38).Value = NNCl.Offset(0, 21).Value * NNCl.Offset(0, 32).Value '不良金額
        NSCl.Offset(0, 39).Value = NNCl.Offset(0, -4).Value '中子名
        Set NSCl = NSCl.Offset(1, 0)
        Set NNCl = NNCl.Offset(1, 0)
        n = n + 1
    Loop
    
    NNC = 0
    NSU = 1
    
    Application.StatusBar = False
    
End Sub