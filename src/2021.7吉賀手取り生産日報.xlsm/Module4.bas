Attribute VB_Name = "Module4"
'Option Explicit

Public NNC As Long 'NippouNyuuryokuChangeFlug
Public NSU As Long 'NippouShuukeiUpdateFlug


Public Sub 当月実績追加処理()

Dim MBk As String, MSt1 As String, MSt2 As String, MSt3 As String
Dim ABk As String, NNSt As String, NSSt As String
Dim MCl1, MCl2, MCl3 As Object
Dim NNCl As Object, NSCl As Object
Dim i As Integer, InM As Integer, Lcnt As Integer
Dim Com1, Com2, Com3, Com5, Com6, Com7, Com8, Com9, Com10 As Long
Dim Com11, Com12, Com13, Com14, Com15, Com16, Com17, Com18, Com19, ComWK As Long
Dim Com20, Com21, Com22, Com23, Com24, Com28, Com29, Com30, Com31, Com32 As Long
Dim Com4, Com25, Com26, Com27 As Single
Dim SVtime, count As Long
Dim WkCom As Double
Dim myBtn As Integer
Dim myMsg As String
Dim myTitle As String
Dim BKcd As String
Dim BKmn As String
Dim GetMM As String
Dim M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11, M12 As String
Dim S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12 As String

'初期設定
Application.ScreenUpdating = False

'20100221改訂 s.tanaka
'20130313改訂 k.kometani

MSt1 = "作業表"
MSt2 = "マシン名"
ABk = ActiveWorkbook.Name
NSSt = "日報集計"
NNSt = "日報入力"


'処理開始
    myMsg = "当月実績追加処理を開始しますか？"
    myTitle = "当月実績追加処理"
    
    myBtn = MsgBox(myMsg, vbYesNo + vbExclamation, myTitle)
     
    If myBtn = vbNo Then
       Exit Sub
    End If
   
    '作業領域クリア（作業表）
    Worksheets(MSt1).Activate
    Range("A5:AM2000").Select
    Selection.ClearContents
    Range("A5").Select
    
    '処理開始位置の設定
    Set NSCl = Workbooks(ABk).Worksheets(NSSt).Range("A5")
    Set NNCl = Workbooks(ABk).Worksheets(NNSt).Range("G5")
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")
    
    '日報集計シートの更新
    Call NippouShuukei_Update(NNCl, NSCl)

    '処理開始位置の設定
    Set NSCl = Workbooks(ABk).Worksheets(NSSt).Range("A5")
    Set NNCl = Workbooks(ABk).Worksheets(NNSt).Range("G5")

    '実績データ確認
    n = 1
    Do Until NSCl.Value = ""
       Application.StatusBar = "日報集計から作業表を作成中・・・　" & n & "レコード目"
       With NSCl
         'データ移行
          For i = 0 To 39
              MCl1.Offset(0, i).Value = .Offset(0, i).Value
          Next i
       End With
       Set MCl1 = MCl1.Offset(1, 0)
       Set NSCl = NSCl.Offset(1, 0)
    Loop


'マシン別集計作業開始
    Application.StatusBar = "マシン別集計中・・・　"
   '作業用ワークシートアクティブ化（作業表）
    Worksheets(MSt1).Activate
   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")
   'インデックス初期化
    i = 4
   '実データ領域確認
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   'マシン別に並び替え
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("B")

   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

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
    ComWK = 0   '計算ワーク
    SVtime = 0  '出勤総時間
    count = 0   '金型交換回数
'
    BKcd = MCl1.Offset(0, 1).Value
    BKmn = MCl1.Offset(0, 2).Value
    SVtime = MCl1.Offset(-4, 0).Value
'
   GetMM = "マシン別集計"

'追加先シート初期化
   '作業用ワークシートアクティブ化（マシン別－該当月）
    Worksheets(GetMM).Activate
   '処理開始位置の設定
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")
   'インデックス初期値
    i = 7
   '実データ領域確認
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop
   'クリア範囲指定
    Range(Cells(7, 1), Cells(i, 32)).Select
    Selection.ClearContents

'マシン名取り込み
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")
    Set MCl3 = Workbooks(ABk).Worksheets(MSt2).Range("B4")
    Do Until MCl3.Value = ""
       If MCl3.Offset(0, 1).Value <> "" Then
          MCl2.Offset(0, 0).Value = MCl3.Offset(0, 0).Value
          MCl2.Offset(0, 1).Value = MCl3.Offset(0, 1).Value
          Set MCl2 = MCl2.Offset(1, 0)
       End If
       Set MCl3 = MCl3.Offset(1, 0)
    Loop

'実績追加処理－マシン別
   'マシン別集計
    Do Until MCl1.Value = ""
       '追加先シート処理開始位置指定
       Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")
       Do Until BKcd <> MCl1.Offset(0, 1).Value
          Com1 = Com1 + MCl1.Offset(0, 4).Value
          Com2 = Com2 + MCl1.Offset(0, 5).Value
          Com3 = Com3 + MCl1.Offset(0, 6).Value
          Com4 = Com4 + MCl1.Offset(0, 7).Value
          Com5 = Com5 + MCl1.Offset(0, 8).Value
          Com6 = Com6 + MCl1.Offset(0, 9).Value
          If MCl1.Offset(0, 9).Value > 0 Then
             count = count + 1
          End If
          Com7 = Com7 + MCl1.Offset(0, 10).Value
          Com8 = Com8 + MCl1.Offset(0, 11).Value
          Com9 = Com9 + MCl1.Offset(0, 12).Value
          Com10 = Com10 + MCl1.Offset(0, 13).Value
          Com11 = Com11 + MCl1.Offset(0, 14).Value
          Com12 = Com12 + MCl1.Offset(0, 15).Value
          Com13 = Com13 + MCl1.Offset(0, 16).Value
          Com14 = Com14 + MCl1.Offset(0, 17).Value
          Com15 = Com15 + MCl1.Offset(0, 18).Value
          Com16 = Com16 + MCl1.Offset(0, 19).Value
          Com17 = Com17 + MCl1.Offset(0, 20).Value
          Com18 = Com18 + MCl1.Offset(0, 21).Value
          Com32 = Com32 + MCl1.Offset(0, 30).Value
          Com27 = Com27 + MCl1.Offset(0, 34).Value
          Com28 = Com28 + MCl1.Offset(0, 35).Value
          Com29 = Com29 + MCl1.Offset(0, 36).Value
          Com30 = Com30 + MCl1.Offset(0, 37).Value
          Com31 = Com31 + MCl1.Offset(0, 38).Value
          Set MCl1 = MCl1.Offset(1, 0)
       Loop
      '生産時間算出
      'ComWK = Com2 - Com3 - Com4 - Com5 - Com6 - Com7 - Com8 - Com9 - Com10 - Com11 - Com12
      'マシンコード位置設定
       Do Until BKcd = MCl2.Offset(0, 0).Value
          Set MCl2 = MCl2.Offset(1, 0)
       Loop
       With MCl2
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
          '.Offset(0, 26).Value = Com18 / Com32 * 100  '不良率
          .Offset(0, 27).Value = (Com2 / 60) / SVtime '設備負荷率
          .Offset(0, 28).Value = Com3 / Com2   '設備稼働率
          .Offset(0, 29).Value = Com30 / (Com2 / 60)  '労働生産性（マシン）
          .Offset(0, 30).Value = Com30 / (Com4 / 60)  '労働生産性（人）
         '
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 26).Value = WkCom     '不良率
         ' If Com2 <> 0 Then
         '    'WkCom = Com2 / Com2 * 100
         '    WkCom = ComWK / Com2
         '   Else
         '    WkCom = 0
         ' End If
         ' .Offset(0, 16).Value = WkCom     '稼働率
         ' If Com25 <> 0 Then
         '    WkCom = Com25 / (ComWK / 60)
         '   Else
         '    WkCom = 0
         ' End If
         ' .Offset(0, 17).Value = WkCom     '労働生産性
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 1).Value
       BKmn = MCl1.Offset(0, 2).Value
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
       ComWK = 0   '計算ワーク
       count = 0   '金型交換回数
    Loop

   '位置の設定
    Range("A1").Select

'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************


'品名別集計作業開始
    Application.StatusBar = "品名別集計中・・・　"
   '作業用ワークシートアクティブ化（作業表）
    Worksheets(MSt1).Activate
   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   'インデックス初期化
    i = 4

   '実データ領域確認
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '品名別に並び替え
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("D")

   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

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
    ComWK = 0   '計算ワーク
    count = 0   '金型交換回数
'
    BKcd = MCl1.Offset(0, 3).Value        '中子コード
    BKmn = MCl1.Offset(0, 39).Value        '中子名


   GetMM = "品名別集計"

'追加先シート初期化
   '作業用ワークシートアクティブ化（マシン別－該当月）
    Worksheets(GetMM).Activate
   '処理開始位置の設定
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")
   'インデックス初期値
    i = 7
   '実データ領域確認
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop
   'クリア範囲指定
    Range(Cells(7, 1), Cells(i, 32)).Select
    Selection.ClearContents
'
'実績追加処理－品名別
   '追加先シート処理開始位置指定
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")

   '品名別集計
    Do Until MCl1.Value = ""
       Do Until BKcd <> MCl1.Offset(0, 3).Value
          Com1 = Com1 + MCl1.Offset(0, 4).Value
          Com2 = Com2 + MCl1.Offset(0, 5).Value
          Com3 = Com3 + MCl1.Offset(0, 6).Value
          Com4 = Com4 + MCl1.Offset(0, 7).Value
          Com5 = Com5 + MCl1.Offset(0, 8).Value
          Com6 = Com6 + MCl1.Offset(0, 9).Value
          If MCl1.Offset(0, 9).Value > 0 Then
             count = count + 1
          End If
          Com7 = Com7 + MCl1.Offset(0, 10).Value
          Com8 = Com8 + MCl1.Offset(0, 11).Value
          Com9 = Com9 + MCl1.Offset(0, 12).Value
          Com10 = Com10 + MCl1.Offset(0, 13).Value
          Com11 = Com11 + MCl1.Offset(0, 14).Value
          Com12 = Com12 + MCl1.Offset(0, 15).Value
          Com13 = Com13 + MCl1.Offset(0, 16).Value
          Com14 = Com14 + MCl1.Offset(0, 17).Value
          Com15 = Com15 + MCl1.Offset(0, 18).Value
          Com16 = Com16 + MCl1.Offset(0, 19).Value
          Com17 = Com17 + MCl1.Offset(0, 20).Value
          Com18 = Com18 + MCl1.Offset(0, 21).Value
          Com32 = Com32 + MCl1.Offset(0, 30).Value
          Com27 = Com27 + MCl1.Offset(0, 34).Value
          Com28 = Com28 + MCl1.Offset(0, 35).Value
          Com29 = Com29 + MCl1.Offset(0, 36).Value
          Com30 = Com30 + MCl1.Offset(0, 37).Value
          Com31 = Com31 + MCl1.Offset(0, 38).Value
          Set MCl1 = MCl1.Offset(1, 0)
       Loop
      '生産時間算出
      'ComWK = Com2 - Com3 - Com4 - Com5 - Com6 - Com7 - Com8 - Com9 - Com10 - Com11 - Com12
      '
      With MCl2  '20140408kometani  中子コードを記入するセルを追加したことで右に1個ずつずらした
          .Offset(0, 1).Value = BKmn           '中子名
          .Offset(0, 2).Value = BKcd           '中子コード　'20140408kometani　追加
          .Offset(0, 3).Value = Com1           'ショット数
          .Offset(0, 4).Value = Com32          '良品数
          .Offset(0, 5).Value = Com18          '不良数
          .Offset(0, 6).Value = Com2 / 60      'マシン稼働時間
          .Offset(0, 7).Value = Com3 / 60      'マシン生産時間
          .Offset(0, 8).Value = Com4 / 60      'ＯＰ作業時間
          .Offset(0, 9).Value = Com5 / 60      '始業作業
          .Offset(0, 10).Value = Com6 / 60     '金型交換
          .Offset(0, 11).Value = Com7 / 60     '昇温待ち
          .Offset(0, 12).Value = count         '型交換回数
          .Offset(0, 13).Value = Com8 / 60     '型調整
          .Offset(0, 14).Value = Com9 / 60     '故障停止
          .Offset(0, 15).Value = Com11 / 60    '金型清掃
          .Offset(0, 16).Value = Com10 / 60    '終了作業
          .Offset(0, 17).Value = Com12 / 60    'Ｒｂ教示
          .Offset(0, 18).Value = Com13 / 60    '他機対応待ち
          .Offset(0, 19).Value = Com14 / 60    '離型剤
          .Offset(0, 20).Value = Com15 / 60    '中子割れ処理
          .Offset(0, 21).Value = Com16 / 60    'その他
          .Offset(0, 22).Value = Com27         '使用量
          .Offset(0, 23).Value = Com28         '良品使用量
          .Offset(0, 24).Value = Com29         '不良使用量
          .Offset(0, 25).Value = Com30         '生産金額
          .Offset(0, 26).Value = Com31         '不良金額
          '.Offset(0, 27).Value = Com18 / Com32 * 100  '不良率
          .Offset(0, 28).Value = (Com2 / 60) / SVtime '設備負荷率
          .Offset(0, 29).Value = Com3 / Com2   '設備稼働率
          '.Offset(0, 30).Value = Com30 / (Com3 / 60)  '労働生産性（マシン）
          '.Offset(0, 31).Value = Com30 / (Com4 / 60)  '労働生産性（人）
         '
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 27).Value = WkCom
'''''''''''''
          If Com30 <> 0 Then
             .Offset(0, 30).Value = Com30 / (Com2 / 60)  '労働生産性（マシン）
             .Offset(0, 31).Value = Com30 / (Com4 / 60)  '労働生産性（人）
            Else
             .Offset(0, 30).Value = 0
             .Offset(0, 31).Value = 0
          End If
'''''''''''''
          'If Com25 <> 0 Then
          '   WkCom = Com25 / (ComWK / 60)
          '  Else
          '   WkCom = 0
          'End If
          '.Offset(0, 20).Value = WkCom     '労働生産性
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 3).Value
       BKmn = MCl1.Offset(0, 39).Value

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
       ComWK = 0   '計算ワーク
       count = 0   '金型交換回数
    Loop

   '作業用ワークシートアクティブ化（品名別－該当月）
    Worksheets(GetMM).Activate

   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(GetMM).Range("B7")

   'インデックス初期化
    i = 7

   '実データ領域確認
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '生産金額順（降順）に並び替え
    Range(Cells(7, 1), Cells(i, 32)).Sort _
    Key1:=Columns("Z"), Order1:=xlDescending

'品名に通番付与（生産金額順）
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("B7")
   'カウント初期化
    Lcnt = 1
   '実行
    Do Until MCl2.Value = ""
       MCl2.Offset(0, -1).Value = Lcnt   '通番
       Lcnt = Lcnt + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop

'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************

'20091120追加不良別集計
'マシン別不良集計作業開始
    Application.StatusBar = "マシン別不良集計中・・・　"
   '作業用ワークシートアクティブ化（作業表）
    Worksheets(MSt1).Activate
   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")
   'インデックス初期化
    i = 4
   '実データ領域確認
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   'マシン別に並び替え
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("B")

   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

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
    ComWK = 0   '計算ワーク
    BKcd = MCl1.Offset(0, 1).Value
    BKmn = MCl1.Offset(0, 2).Value

   GetMM = "不良集計【マシン】"

'追加先シート初期化
   '作業用ワークシートアクティブ化
    Worksheets(GetMM).Activate
   '処理開始位置の設定
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")
   'インデックス初期値
    i = 5
   '実データ領域確認
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop
   'クリア範囲指定
    Range(Cells(6, 1), Cells(i, 15)).Select
    Selection.ClearContents

'マシン名取り込み
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")
    Set MCl3 = Workbooks(ABk).Worksheets(MSt2).Range("B4")
    Do Until MCl3.Value = ""
       If MCl3.Offset(0, 1).Value <> "" Then
          MCl2.Offset(0, 0).Value = MCl3.Offset(0, 0).Value
          MCl2.Offset(0, 1).Value = MCl3.Offset(0, 1).Value
          Set MCl2 = MCl2.Offset(1, 0)
       End If
       Set MCl3 = MCl3.Offset(1, 0)
    Loop

'実績追加処理－マシン別
   '追加先シート処理開始位置指定
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")

   'マシン別集計
    Do Until MCl1.Value = ""
       Do Until BKcd <> MCl1.Offset(0, 1).Value
          Com17 = Com17 + MCl1.Offset(0, 20).Value
          Com18 = Com18 + MCl1.Offset(0, 21).Value
          Com19 = Com19 + MCl1.Offset(0, 22).Value
          Com20 = Com20 + MCl1.Offset(0, 23).Value
          Com21 = Com21 + MCl1.Offset(0, 24).Value
          Com22 = Com22 + MCl1.Offset(0, 25).Value
          Com23 = Com23 + MCl1.Offset(0, 26).Value
          Com24 = Com24 + MCl1.Offset(0, 27).Value
          Com25 = Com25 + MCl1.Offset(0, 28).Value
          Com26 = Com26 + MCl1.Offset(0, 29).Value
          Com32 = Com32 + MCl1.Offset(0, 30).Value
          Set MCl1 = MCl1.Offset(1, 0)
       Loop
      'マシンコード位置設定
       Do Until BKcd = MCl2.Offset(0, 0).Value
          Set MCl2 = MCl2.Offset(1, 0)
       Loop
       With MCl2
'         .Offset(0, 0).Value = BKcd       'マシンコード
'         .Offset(0, 1).Value = BKmn       'マシン名
          .Offset(0, 2).Value = Com32      '良品数
          .Offset(0, 3).Value = Com18      '不良数
          .Offset(0, 4).Value = Com19      'ボス割れ表
          .Offset(0, 5).Value = Com20      'ボス割れ裏
          .Offset(0, 6).Value = Com21      '幅木割れ
          .Offset(0, 7).Value = Com22      'フィン割れ
          .Offset(0, 8).Value = Com23      '幅木充填
          .Offset(0, 9).Value = Com24      'フィン充填
          .Offset(0, 10).Value = Com25     'キャンドル残
          .Offset(0, 11).Value = Com26     'その他
          .Offset(0, 12).Value = Com17     '手直不良
          '.Offset(0, 13).Value = Com24 / Com32     '廃棄不良率
          '.Offset(0, 14).Value = Com17 / Com32     '手直不良率
'
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 13).Value = WkCom     '廃棄不良率
'
          If Com17 <> 0 Then
             WkCom = Com17 / (Com17 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 14).Value = WkCom     '手直不良率
'
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 1).Value
       BKmn = MCl1.Offset(0, 2).Value
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
       ComWK = 0   '計算ワーク
      Loop

   '位置の設定
    Range("A1").Select

'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************

'品名別不良集計作業開始
    Application.StatusBar = "品名別不良集計中・・・　"
   '作業用ワークシートアクティブ化（作業表）
    Worksheets(MSt1).Activate
   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   'インデックス初期化
    i = 4

   '実データ領域確認
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '品名別に並び替え
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("D")

   '処理開始位置の設定
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

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
    ComWK = 0   '計算ワーク
    BKcd = MCl1.Offset(0, 3).Value        '中子コード
    BKmn = MCl1.Offset(0, 39).Value        '中子名

   GetMM = "不良集計【品名】"

'追加先シート初期化
   '作業用ワークシートアクティブ化（品名別－該当月）
    Worksheets(GetMM).Activate
   '処理開始位置の設定
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")
   'インデックス初期値
    i = 5
   '実データ領域確認
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop

   'クリア範囲指定
    Range(Cells(6, 1), Cells(i, 14)).Select
    Selection.ClearContents

'実績追加処理－品名別
   '追加先シート処理開始位置指定
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")

   '品名別集計
    Do Until MCl1.Value = ""
       Do Until BKcd <> MCl1.Offset(0, 3).Value
          Com17 = Com17 + MCl1.Offset(0, 20).Value
          Com18 = Com18 + MCl1.Offset(0, 21).Value
          Com19 = Com19 + MCl1.Offset(0, 22).Value
          Com20 = Com20 + MCl1.Offset(0, 23).Value
          Com21 = Com21 + MCl1.Offset(0, 24).Value
          Com22 = Com22 + MCl1.Offset(0, 25).Value
          Com23 = Com23 + MCl1.Offset(0, 26).Value
          Com24 = Com24 + MCl1.Offset(0, 27).Value
          Com25 = Com25 + MCl1.Offset(0, 28).Value
          Com26 = Com26 + MCl1.Offset(0, 29).Value
          Com32 = Com32 + MCl1.Offset(0, 30).Value
          Set MCl1 = MCl1.Offset(1, 0)
       Loop
       With MCl2
          .Offset(0, 0).Value = BKcd       '中子コード
          .Offset(0, 1).Value = BKmn       '中子名
          .Offset(0, 2).Value = Com32      '良品数
          .Offset(0, 3).Value = Com18      '不良数
          .Offset(0, 4).Value = Com19      'ボス割れ表
          .Offset(0, 5).Value = Com20      'ボス割れ裏
          .Offset(0, 6).Value = Com21      '幅木割れ
          .Offset(0, 7).Value = Com22      'フィン割れ
          .Offset(0, 8).Value = Com23      '幅木充填
          .Offset(0, 9).Value = Com24      'フィン充填
          .Offset(0, 10).Value = Com25     'キャンドル残
          .Offset(0, 11).Value = Com26     'その他
          .Offset(0, 12).Value = Com17     '手直不良
          '.Offset(0, 13).Value = Com24 / Com32     '廃棄不良率
          '.Offset(0, 14).Value = Com17 / Com32     '手直不良率
'
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 13).Value = WkCom     '廃棄不良率
'
          If Com17 <> 0 Then
             WkCom = Com17 / (Com17 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 14).Value = WkCom     '手直不良率
'
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 3).Value
       BKmn = MCl1.Offset(0, 39).Value

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
       ComWK = 0   '計算ワーク
    Loop
          















'*********************************************************************************
'************ここから　　　　20130313kometani追加　　　　ここから*****************
'*********************************************************************************
         
'品名別ショット数集計開始

    Dim wb As Workbook
    Dim 元品番 As Object
    Dim 先品番 As Object
    Dim 生産日 As Variant
    Dim YandM As Variant
    Dim temp As Object
    
    'ショット数集計ファイルを読み出し
    Set wb = Workbooks.Open(Filename:=ThisWorkbook.Path & "\..\..\ショット管理表\【吉賀】ショット数集計.xls ")
    
    '集計月の算出
    Set 生産日 = ThisWorkbook.Worksheets("日報入力").Range("G5")
    If Month(生産日.Value) <> 12 Then
        YandM = Year(生産日.Value) & "年" & (Month(生産日.Value) + 1) & "月度"
    Else
        YandM = (Year(生産日.Value) + 1) & "年" & "1月度"
    End If
    'ショット数を入力していく列を検索
    Set temp = wb.Worksheets("吉賀中子工場").Range("J3")
    Do While temp.Value <> ""
        If YandM <> temp.Value Then
            Set temp = temp.Offset(0, 1)
        Else
            'ショット数を入力していく列を確定
            temp.Activate
            temp.Font.ColorIndex = 1
            Exit Do
        End If
    Loop
    
    '日報集計ファイルの「品名別集計」シートのセル色初期化(白色)
    '   ショット数書き換え後にチェック用として
    '   元品番のセル色を赤にする処理を追加したため
    ThisWorkbook.Worksheets("品名別集計").Range("B7:B41").Interior.ColorIndex = 2
    
    Set 元品番 = ThisWorkbook.Worksheets("品名別集計").Range("C7")
    Do While 元品番.Value <> ""
        '先品番の初期化
        Set 先品番 = wb.Worksheets("吉賀中子工場").Range("D6")
        '先品番の検索
        Do While 元品番.Value <> 先品番.Value
            Set 先品番 = 先品番.Offset(1, 0)
            '見つからなかった場合(ループを抜ける)
            If 先品番.Value = "" Then
                GoTo rt1
            End If
        Loop
        '値の書き換え
        If 元品番.Value = 8 Then 'BP4Yの場合
            With wb.Worksheets("吉賀中子工場")
                'AB型に対して
                .Cells(先品番.Row, ActiveCell.Column).Value = ThisWorkbook.Worksheets("品名別集計").Cells(元品番.Row, 4).Value / 4
                .Cells(先品番.Row, ActiveCell.Column).Font.ColorIndex = 1
                'CD型に対して
                .Cells(先品番.Row + 5, ActiveCell.Column).Value = ThisWorkbook.Worksheets("品名別集計").Cells(元品番.Row, 4).Value / 4
                .Cells(先品番.Row + 5, ActiveCell.Column).Font.ColorIndex = 1
                'EF型に対して
                .Cells(先品番.Row + 10, ActiveCell.Column).Value = ThisWorkbook.Worksheets("品名別集計").Cells(元品番.Row, 4).Value / 4
                .Cells(先品番.Row + 10, ActiveCell.Column).Font.ColorIndex = 1
                'GH型に対して
                .Cells(先品番.Row + 15, ActiveCell.Column).Value = ThisWorkbook.Worksheets("品名別集計").Cells(元品番.Row, 4).Value / 4
                .Cells(先品番.Row + 15, ActiveCell.Column).Font.ColorIndex = 1
            End With
        ElseIf 元品番.Value = 12 Then 'DF71の場合
            With wb.Worksheets("吉賀中子工場")
                '１番型に対して
                .Cells(先品番.Row, ActiveCell.Column).Value = ThisWorkbook.Worksheets("品名別集計").Cells(元品番.Row, 4).Value / 2
                .Cells(先品番.Row, ActiveCell.Column).Font.ColorIndex = 1
                '２番型に対して
                .Cells(先品番.Row + 5, ActiveCell.Column).Value = ThisWorkbook.Worksheets("品名別集計").Cells(元品番.Row, 4).Value / 2
                .Cells(先品番.Row + 5, ActiveCell.Column).Font.ColorIndex = 1
            End With
        Else '金型が１型しかないその他の品種
            With wb.Worksheets("吉賀中子工場")
                .Cells(先品番.Row, ActiveCell.Column).Value = ThisWorkbook.Worksheets("品名別集計").Cells(元品番.Row, 4).Value
                .Cells(先品番.Row, ActiveCell.Column).Font.ColorIndex = 1
            End With
        End If
        'チェック用
        元品番.Offset(0, -1).Interior.ColorIndex = 3
        
rt1:
    
        Set 元品番 = 元品番.Offset(1, 0)
        'MsgBox 元品番.Value
        
        
        
        '新規中子などへの対応
        '
        '
        '
        '
        '
    Loop

    
    
    'ゼロ自動入力
    Dim NowCell As Object '現在参照中セル

    Set NowCell = ActiveCell
    '参照中のセルが2000行にいくまでループ(出雲で1100行しかないため)
    Do While NowCell.Row < 2000
        If NowCell.Font.ColorIndex = 3 Then 'セル内の文字が赤ならば
            NowCell.Value = 0               '内容を「０」にして
            NowCell.Font.ColorIndex = 1     '文字色を黒にする
        End If
        Set NowCell = NowCell.Offset(1, 0)  '参照中のセルを下に１つずらす
    Loop
    
    
    
    '平均ショット数（過去6ヶ月）更新
    Set 先品番 = wb.Worksheets("吉賀中子工場").Range("D6")
    Do Until 先品番 = ""
        If 先品番 <> 先品番.Offset(-1, 0) Then
            With wb.Worksheets("吉賀中子工場")
                '平均ショット数（過去6ヶ月）
                .Cells(先品番.Row, 7).FormulaR1C1 = "=sum(RC[" & temp.Column - 7 - 5 & "]:RC[" & temp.Column - 7 - 0 & "])/6"
            End With
        End If
        Set 先品番 = 先品番.Offset(1, 0)
    Loop
    
    
    
    '先６ヶ月分の列追加
    Dim fy, fm, s, t As Integer
    Dim temp2 As Object
    
    Set temp2 = temp
    t = 0
    For s = 1 To 6  '6ヶ月（半年）分処理を繰り返す
        fm = (Month(生産日.Value) + 1) + s
        fy = Year(生産日.Value)
        '年越し処理
        If fm > 12 Then
            fm = fm Mod 12
            fy = fy + 1
            t = t + 1
        End If
        '列追加の要・不要を判断
        If temp2.Offset(0, 1).Value = fy & "年" & fm & "月度" Then
            '列追加不要
            Set temp2 = temp2.Offset(0, 1)
        Else
            '列追加必要
            Columns(temp2.Column).Copy
            Columns(temp2.Offset(0, 1).Column).Insert
            Set temp2 = temp2.Offset(0, 1)
            temp2.Value = fy & "年" & fm & "月度"
        End If
    Next
    
    '修正チェック
    Application.Run "【吉賀】ショット数集計.xls!修正check"
    
    Application.DisplayAlerts = False
    wb.Close (True)
    Application.DisplayAlerts = True
    
'*********************************************************************************
'************ここまで　　　　20130313kometani追加　　　　ここまで*****************
'*********************************************************************************








   '位置の設定
    Range("A1").Select
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "処理を終わりました。", vbOKOnly + vbInformation, "通知"
End Sub




Sub セル色初期化()

    ThisWorkbook.Worksheets("品名別集計").Range("B7:B41").Interior.ColorIndex = 2

End Sub















