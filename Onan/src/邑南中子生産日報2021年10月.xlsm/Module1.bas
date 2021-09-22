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
     End If
End Sub


