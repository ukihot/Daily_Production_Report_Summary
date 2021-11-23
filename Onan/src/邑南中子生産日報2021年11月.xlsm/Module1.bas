Attribute VB_Name = "Module1"
Dim myBtn As Integer
Dim myMsg As String
Dim myTitle As String

Sub クリア()
'
   myMsg = "入力データを削除しますか？"
   myTitle = "データの削除確認"

   myBtn = MsgBox(myMsg, vbYesNo + vbExclamation, myTitle)

   If myBtn = vbYes Then
      Range("G5:AJ2000").Select
      Selection.ClearContents
      Range("G5").Select
      '作業領域クリア（作業表）
      Worksheets("作業表").Activate
      Range("A5:AN2000").Select
      Selection.ClearContents
      Range("A5").Select
      '作業領域クリア（"日報集計"）
      Worksheets("日報集計").Activate
      Range("A5:AW1478").Select
      Selection.ClearContents
      Range("A5").Select
   End If
End Sub


