Attribute VB_Name = "Module1"
Dim myBtn As Integer
Dim myMsg As String
Dim myTitle As String

Sub �N���A()
'
   myMsg = "���̓f�[�^���폜���܂����H"
   myTitle = "�f�[�^�̍폜�m�F"

   myBtn = MsgBox(myMsg, vbYesNo + vbExclamation, myTitle)

   If myBtn = vbYes Then
      Range("G5:AJ2000").Select
      Selection.ClearContents
      Range("G5").Select
      '��Ɨ̈�N���A�i��ƕ\�j
      Worksheets("��ƕ\").Activate
      Range("A5:AN2000").Select
      Selection.ClearContents
      Range("A5").Select
      '��Ɨ̈�N���A�i"����W�v"�j
      Worksheets("����W�v").Activate
      Range("A5:AW1478").Select
      Selection.ClearContents
      Range("A5").Select
   End If
End Sub


