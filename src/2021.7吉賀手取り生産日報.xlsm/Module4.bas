Attribute VB_Name = "Module4"
'Option Explicit

Public NNC As Long 'NippouNyuuryokuChangeFlug
Public NSU As Long 'NippouShuukeiUpdateFlug


Public Sub �������ђǉ�����()

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

'�����ݒ�
Application.ScreenUpdating = False

'20100221���� s.tanaka
'20130313���� k.kometani

MSt1 = "��ƕ\"
MSt2 = "�}�V����"
ABk = ActiveWorkbook.Name
NSSt = "����W�v"
NNSt = "�������"


'�����J�n
    myMsg = "�������ђǉ��������J�n���܂����H"
    myTitle = "�������ђǉ�����"
    
    myBtn = MsgBox(myMsg, vbYesNo + vbExclamation, myTitle)
     
    If myBtn = vbNo Then
       Exit Sub
    End If
   
    '��Ɨ̈�N���A�i��ƕ\�j
    Worksheets(MSt1).Activate
    Range("A5:AM2000").Select
    Selection.ClearContents
    Range("A5").Select
    
    '�����J�n�ʒu�̐ݒ�
    Set NSCl = Workbooks(ABk).Worksheets(NSSt).Range("A5")
    Set NNCl = Workbooks(ABk).Worksheets(NNSt).Range("G5")
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")
    
    '����W�v�V�[�g�̍X�V
    Call NippouShuukei_Update(NNCl, NSCl)

    '�����J�n�ʒu�̐ݒ�
    Set NSCl = Workbooks(ABk).Worksheets(NSSt).Range("A5")
    Set NNCl = Workbooks(ABk).Worksheets(NNSt).Range("G5")

    '���уf�[�^�m�F
    n = 1
    Do Until NSCl.Value = ""
       Application.StatusBar = "����W�v�����ƕ\���쐬���E�E�E�@" & n & "���R�[�h��"
       With NSCl
         '�f�[�^�ڍs
          For i = 0 To 39
              MCl1.Offset(0, i).Value = .Offset(0, i).Value
          Next i
       End With
       Set MCl1 = MCl1.Offset(1, 0)
       Set NSCl = NSCl.Offset(1, 0)
    Loop


'�}�V���ʏW�v��ƊJ�n
    Application.StatusBar = "�}�V���ʏW�v���E�E�E�@"
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
    Worksheets(MSt1).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")
   '�C���f�b�N�X������
    i = 4
   '���f�[�^�̈�m�F
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '�}�V���ʂɕ��ёւ�
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("B")

   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   '��Ɨ̈揉����
    Com1 = 0    '�V���b�g
    Com2 = 0    '�ғ�����
    Com3 = 0    '���Y����
    Com4 = 0    '�n�o��Ǝ���
    Com5 = 0    '�n�Ǝ���
    Com6 = 0    '���^����
    Com7 = 0    '�����҂�
    Com8 = 0    '���^����
    Com9 = 0    '�}�V���̏��~
    Com10 = 0   '�I�Ǝ���
    Com11 = 0   '�^���|
    Com12 = 0   '�q������
    Com13 = 0   '���@�Ή��҂�
    Com14 = 0   '���^��
    Com15 = 0   '���q���ꏈ��
    Com16 = 0   '���̑�
    Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
    Com18 = 0   '���^�s�ǁi�p���s�ǁj
    Com19 = 0   '�{�X����\
    Com20 = 0   '�{�X���ꗠ
    Com21 = 0   '���؊���
    Com22 = 0   '�t�B������
    Com23 = 0   '���؏[�U
    Com24 = 0   '�t�B���[�U
    Com25 = 0   '�L�����h���c
    Com26 = 0   '���̑�
    Com27 = 0   '������
    Com28 = 0   '���Ǖi
    Com29 = 0   '���s��
    Com30 = 0   '���Y���z
    Com31 = 0   '�s�ǋ��z
    Com32 = 0   '�Ǖi��
    ComWK = 0   '�v�Z���[�N
    SVtime = 0  '�o�Α�����
    count = 0   '���^������
'
    BKcd = MCl1.Offset(0, 1).Value
    BKmn = MCl1.Offset(0, 2).Value
    SVtime = MCl1.Offset(-4, 0).Value
'
   GetMM = "�}�V���ʏW�v"

'�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�}�V���ʁ|�Y�����j
    Worksheets(GetMM).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")
   '�C���f�b�N�X�����l
    i = 7
   '���f�[�^�̈�m�F
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop
   '�N���A�͈͎w��
    Range(Cells(7, 1), Cells(i, 32)).Select
    Selection.ClearContents

'�}�V������荞��
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

'���ђǉ������|�}�V����
   '�}�V���ʏW�v
    Do Until MCl1.Value = ""
       '�ǉ���V�[�g�����J�n�ʒu�w��
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
      '���Y���ԎZ�o
      'ComWK = Com2 - Com3 - Com4 - Com5 - Com6 - Com7 - Com8 - Com9 - Com10 - Com11 - Com12
      '�}�V���R�[�h�ʒu�ݒ�
       Do Until BKcd = MCl2.Offset(0, 0).Value
          Set MCl2 = MCl2.Offset(1, 0)
       Loop
       With MCl2
          .Offset(0, 2).Value = Com1           '�V���b�g��
          .Offset(0, 3).Value = Com32          '�Ǖi��
          .Offset(0, 4).Value = Com18          '�s�ǐ�
          .Offset(0, 5).Value = Com2 / 60      '�}�V���ғ�����
          .Offset(0, 6).Value = Com3 / 60      '�}�V�����Y����
          .Offset(0, 7).Value = Com4 / 60      '�n�o��Ǝ���
          .Offset(0, 8).Value = Com5 / 60      '�n�ƍ��
          .Offset(0, 9).Value = Com6 / 60      '���^����
          .Offset(0, 10).Value = Com7 / 60     '�����҂�
          .Offset(0, 11).Value = count         '�^�����񐔁i�ǂ�����H�j
          .Offset(0, 12).Value = Com8 / 60     '�^����
          .Offset(0, 13).Value = Com9 / 60     '�̏��~
          .Offset(0, 14).Value = Com11 / 60    '���^���|
          .Offset(0, 15).Value = Com10 / 60    '�I�����
          .Offset(0, 16).Value = Com12 / 60    '�q������
          .Offset(0, 17).Value = Com13 / 60    '���@�Ή��҂�
          .Offset(0, 18).Value = Com14 / 60    '���^��
          .Offset(0, 19).Value = Com15 / 60    '���q���ꏈ��
          .Offset(0, 20).Value = Com16 / 60    '���̑�
          .Offset(0, 21).Value = Com27 / 1000  '�g�p��
          .Offset(0, 22).Value = Com28 / 1000  '�Ǖi�g�p��
          .Offset(0, 23).Value = Com29 / 1000  '�s�ǎg�p��
          .Offset(0, 24).Value = Com30 / 1000  '���Y���z
          .Offset(0, 25).Value = Com31 / 1000  '�s�ǋ��z
          '.Offset(0, 26).Value = Com18 / Com32 * 100  '�s�Ǘ�
          .Offset(0, 27).Value = (Com2 / 60) / SVtime '�ݔ����ח�
          .Offset(0, 28).Value = Com3 / Com2   '�ݔ��ғ���
          .Offset(0, 29).Value = Com30 / (Com2 / 60)  '�J�����Y���i�}�V���j
          .Offset(0, 30).Value = Com30 / (Com4 / 60)  '�J�����Y���i�l�j
         '
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 26).Value = WkCom     '�s�Ǘ�
         ' If Com2 <> 0 Then
         '    'WkCom = Com2 / Com2 * 100
         '    WkCom = ComWK / Com2
         '   Else
         '    WkCom = 0
         ' End If
         ' .Offset(0, 16).Value = WkCom     '�ғ���
         ' If Com25 <> 0 Then
         '    WkCom = Com25 / (ComWK / 60)
         '   Else
         '    WkCom = 0
         ' End If
         ' .Offset(0, 17).Value = WkCom     '�J�����Y��
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 1).Value
       BKmn = MCl1.Offset(0, 2).Value
      '��ƃG���A������
       Com1 = 0    '�V���b�g
       Com2 = 0    '�ғ�����
       Com3 = 0    '���Y����
       Com4 = 0    '�n�o��Ǝ���
       Com5 = 0    '�n�Ǝ���
       Com6 = 0    '���^����
       Com7 = 0    '�����҂�
       Com8 = 0    '���^����
       Com9 = 0    '�}�V���̏��~
       Com10 = 0   '�I�Ǝ���
       Com11 = 0   '�^���|
       Com12 = 0   '�q������
       Com13 = 0   '���@�Ή��҂�
       Com14 = 0   '���^��
       Com15 = 0   '���q���ꏈ��
       Com16 = 0   '���̑�
       Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
       Com18 = 0   '���^�s�ǁi�p���s�ǁj
       Com27 = 0   '������
       Com28 = 0   '���Ǖi
       Com29 = 0   '���s��
       Com30 = 0   '���Y���z
       Com31 = 0   '�s�ǋ��z
       Com32 = 0   '�Ǖi��
       ComWK = 0   '�v�Z���[�N
       count = 0   '���^������
    Loop

   '�ʒu�̐ݒ�
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


'�i���ʏW�v��ƊJ�n
    Application.StatusBar = "�i���ʏW�v���E�E�E�@"
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
    Worksheets(MSt1).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   '�C���f�b�N�X������
    i = 4

   '���f�[�^�̈�m�F
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '�i���ʂɕ��ёւ�
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("D")

   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   '��Ɨ̈揉����
    Com1 = 0    '�V���b�g
    Com2 = 0    '�ғ�����
    Com3 = 0    '���Y����
    Com4 = 0    '�n�o��Ǝ���
    Com5 = 0    '�n�Ǝ���
    Com6 = 0    '���^����
    Com7 = 0    '�����҂�
    Com8 = 0    '���^����
    Com9 = 0    '�}�V���̏��~
    Com10 = 0   '�I�Ǝ���
    Com11 = 0   '�^���|
    Com12 = 0   '�q������
    Com13 = 0   '���@�Ή��҂�
    Com14 = 0   '���^��
    Com15 = 0   '���q���ꏈ��
    Com16 = 0   '���̑�
    Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
    Com18 = 0   '���^�s�ǁi�p���s�ǁj
    Com19 = 0   '�{�X����\
    Com20 = 0   '�{�X���ꗠ
    Com21 = 0   '���؊���
    Com22 = 0   '�t�B������
    Com23 = 0   '���؏[�U
    Com24 = 0   '�t�B���[�U
    Com25 = 0   '�L�����h���c
    Com26 = 0   '���̑�
    Com27 = 0   '������
    Com28 = 0   '���Ǖi
    Com29 = 0   '���s��
    Com30 = 0   '���Y���z
    Com31 = 0   '�s�ǋ��z
    Com32 = 0   '�Ǖi��
    ComWK = 0   '�v�Z���[�N
    count = 0   '���^������
'
    BKcd = MCl1.Offset(0, 3).Value        '���q�R�[�h
    BKmn = MCl1.Offset(0, 39).Value        '���q��


   GetMM = "�i���ʏW�v"

'�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�}�V���ʁ|�Y�����j
    Worksheets(GetMM).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")
   '�C���f�b�N�X�����l
    i = 7
   '���f�[�^�̈�m�F
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop
   '�N���A�͈͎w��
    Range(Cells(7, 1), Cells(i, 32)).Select
    Selection.ClearContents
'
'���ђǉ������|�i����
   '�ǉ���V�[�g�����J�n�ʒu�w��
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A7")

   '�i���ʏW�v
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
      '���Y���ԎZ�o
      'ComWK = Com2 - Com3 - Com4 - Com5 - Com6 - Com7 - Com8 - Com9 - Com10 - Com11 - Com12
      '
      With MCl2  '20140408kometani  ���q�R�[�h���L������Z����ǉ��������ƂŉE��1�����炵��
          .Offset(0, 1).Value = BKmn           '���q��
          .Offset(0, 2).Value = BKcd           '���q�R�[�h�@'20140408kometani�@�ǉ�
          .Offset(0, 3).Value = Com1           '�V���b�g��
          .Offset(0, 4).Value = Com32          '�Ǖi��
          .Offset(0, 5).Value = Com18          '�s�ǐ�
          .Offset(0, 6).Value = Com2 / 60      '�}�V���ғ�����
          .Offset(0, 7).Value = Com3 / 60      '�}�V�����Y����
          .Offset(0, 8).Value = Com4 / 60      '�n�o��Ǝ���
          .Offset(0, 9).Value = Com5 / 60      '�n�ƍ��
          .Offset(0, 10).Value = Com6 / 60     '���^����
          .Offset(0, 11).Value = Com7 / 60     '�����҂�
          .Offset(0, 12).Value = count         '�^������
          .Offset(0, 13).Value = Com8 / 60     '�^����
          .Offset(0, 14).Value = Com9 / 60     '�̏��~
          .Offset(0, 15).Value = Com11 / 60    '���^���|
          .Offset(0, 16).Value = Com10 / 60    '�I�����
          .Offset(0, 17).Value = Com12 / 60    '�q������
          .Offset(0, 18).Value = Com13 / 60    '���@�Ή��҂�
          .Offset(0, 19).Value = Com14 / 60    '���^��
          .Offset(0, 20).Value = Com15 / 60    '���q���ꏈ��
          .Offset(0, 21).Value = Com16 / 60    '���̑�
          .Offset(0, 22).Value = Com27         '�g�p��
          .Offset(0, 23).Value = Com28         '�Ǖi�g�p��
          .Offset(0, 24).Value = Com29         '�s�ǎg�p��
          .Offset(0, 25).Value = Com30         '���Y���z
          .Offset(0, 26).Value = Com31         '�s�ǋ��z
          '.Offset(0, 27).Value = Com18 / Com32 * 100  '�s�Ǘ�
          .Offset(0, 28).Value = (Com2 / 60) / SVtime '�ݔ����ח�
          .Offset(0, 29).Value = Com3 / Com2   '�ݔ��ғ���
          '.Offset(0, 30).Value = Com30 / (Com3 / 60)  '�J�����Y���i�}�V���j
          '.Offset(0, 31).Value = Com30 / (Com4 / 60)  '�J�����Y���i�l�j
         '
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 27).Value = WkCom
'''''''''''''
          If Com30 <> 0 Then
             .Offset(0, 30).Value = Com30 / (Com2 / 60)  '�J�����Y���i�}�V���j
             .Offset(0, 31).Value = Com30 / (Com4 / 60)  '�J�����Y���i�l�j
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
          '.Offset(0, 20).Value = WkCom     '�J�����Y��
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 3).Value
       BKmn = MCl1.Offset(0, 39).Value

   '��ƃG���A������
       Com1 = 0    '�V���b�g
       Com2 = 0    '�ғ�����
       Com3 = 0    '���Y����
       Com4 = 0    '�n�o��Ǝ���
       Com5 = 0    '�n�Ǝ���
       Com6 = 0    '���^����
       Com7 = 0    '�����҂�
       Com8 = 0    '���^����
       Com9 = 0    '�}�V���̏��~
       Com10 = 0   '�I�Ǝ���
       Com11 = 0   '�^���|
       Com12 = 0   '�q������
       Com13 = 0   '���@�Ή��҂�
       Com14 = 0   '���^��
       Com15 = 0   '���q���ꏈ��
       Com16 = 0   '���̑�
       Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
       Com18 = 0   '���^�s�ǁi�p���s�ǁj
       Com27 = 0   '������
       Com28 = 0   '���Ǖi
       Com29 = 0   '���s��
       Com30 = 0   '���Y���z
       Com31 = 0   '�s�ǋ��z
       Com32 = 0   '�Ǖi��
       ComWK = 0   '�v�Z���[�N
       count = 0   '���^������
    Loop

   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�i���ʁ|�Y�����j
    Worksheets(GetMM).Activate

   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(GetMM).Range("B7")

   '�C���f�b�N�X������
    i = 7

   '���f�[�^�̈�m�F
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '���Y���z���i�~���j�ɕ��ёւ�
    Range(Cells(7, 1), Cells(i, 32)).Sort _
    Key1:=Columns("Z"), Order1:=xlDescending

'�i���ɒʔԕt�^�i���Y���z���j
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("B7")
   '�J�E���g������
    Lcnt = 1
   '���s
    Do Until MCl2.Value = ""
       MCl2.Offset(0, -1).Value = Lcnt   '�ʔ�
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

'20091120�ǉ��s�ǕʏW�v
'�}�V���ʕs�ǏW�v��ƊJ�n
    Application.StatusBar = "�}�V���ʕs�ǏW�v���E�E�E�@"
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
    Worksheets(MSt1).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")
   '�C���f�b�N�X������
    i = 4
   '���f�[�^�̈�m�F
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '�}�V���ʂɕ��ёւ�
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("B")

   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   '��Ɨ̈揉����
    Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
    Com18 = 0   '�p���s��
    Com19 = 0   '�{�X����\
    Com20 = 0   '�{�X���ꗠ
    Com21 = 0   '���؊���
    Com22 = 0   '�t�B������
    Com23 = 0   '���؏[�U
    Com24 = 0   '�t�B���[�U
    Com25 = 0   '�L�����h���c
    Com26 = 0   '���̑�
    Com32 = 0   '�Ǖi��
    ComWK = 0   '�v�Z���[�N
    BKcd = MCl1.Offset(0, 1).Value
    BKmn = MCl1.Offset(0, 2).Value

   GetMM = "�s�ǏW�v�y�}�V���z"

'�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u��
    Worksheets(GetMM).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")
   '�C���f�b�N�X�����l
    i = 5
   '���f�[�^�̈�m�F
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop
   '�N���A�͈͎w��
    Range(Cells(6, 1), Cells(i, 15)).Select
    Selection.ClearContents

'�}�V������荞��
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

'���ђǉ������|�}�V����
   '�ǉ���V�[�g�����J�n�ʒu�w��
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")

   '�}�V���ʏW�v
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
      '�}�V���R�[�h�ʒu�ݒ�
       Do Until BKcd = MCl2.Offset(0, 0).Value
          Set MCl2 = MCl2.Offset(1, 0)
       Loop
       With MCl2
'         .Offset(0, 0).Value = BKcd       '�}�V���R�[�h
'         .Offset(0, 1).Value = BKmn       '�}�V����
          .Offset(0, 2).Value = Com32      '�Ǖi��
          .Offset(0, 3).Value = Com18      '�s�ǐ�
          .Offset(0, 4).Value = Com19      '�{�X����\
          .Offset(0, 5).Value = Com20      '�{�X���ꗠ
          .Offset(0, 6).Value = Com21      '���؊���
          .Offset(0, 7).Value = Com22      '�t�B������
          .Offset(0, 8).Value = Com23      '���؏[�U
          .Offset(0, 9).Value = Com24      '�t�B���[�U
          .Offset(0, 10).Value = Com25     '�L�����h���c
          .Offset(0, 11).Value = Com26     '���̑�
          .Offset(0, 12).Value = Com17     '�蒼�s��
          '.Offset(0, 13).Value = Com24 / Com32     '�p���s�Ǘ�
          '.Offset(0, 14).Value = Com17 / Com32     '�蒼�s�Ǘ�
'
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 13).Value = WkCom     '�p���s�Ǘ�
'
          If Com17 <> 0 Then
             WkCom = Com17 / (Com17 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 14).Value = WkCom     '�蒼�s�Ǘ�
'
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 1).Value
       BKmn = MCl1.Offset(0, 2).Value
      '��ƃG���A������
       Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
       Com18 = 0   '�p���s��
       Com19 = 0   '�{�X����\
       Com20 = 0   '�{�X���ꗠ
       Com21 = 0   '���؊���
       Com22 = 0   '�t�B������
       Com23 = 0   '���؏[�U
       Com24 = 0   '�t�B���[�U
       Com25 = 0   '�L�����h���c
       Com26 = 0   '���̑�
       Com32 = 0   '�Ǖi��
       ComWK = 0   '�v�Z���[�N
      Loop

   '�ʒu�̐ݒ�
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

'�i���ʕs�ǏW�v��ƊJ�n
    Application.StatusBar = "�i���ʕs�ǏW�v���E�E�E�@"
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
    Worksheets(MSt1).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   '�C���f�b�N�X������
    i = 4

   '���f�[�^�̈�m�F
    Do Until MCl1.Value = ""
       i = i + 1
       Set MCl1 = MCl1.Offset(1, 0)
    Loop

   '�i���ʂɕ��ёւ�
    Range(Cells(5, 1), Cells(i, 41)).Sort _
    Key1:=Columns("D")

   '�����J�n�ʒu�̐ݒ�
    Set MCl1 = Workbooks(ABk).Worksheets(MSt1).Range("A5")

   '��Ɨ̈揉����
    Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
    Com18 = 0   '�p���s��
    Com19 = 0   '�{�X����\
    Com20 = 0   '�{�X���ꗠ
    Com21 = 0   '���؊���
    Com22 = 0   '�t�B������
    Com23 = 0   '���؏[�U
    Com24 = 0   '�t�B���[�U
    Com25 = 0   '�L�����h���c
    Com26 = 0   '���̑�
    Com32 = 0   '�Ǖi��
    ComWK = 0   '�v�Z���[�N
    BKcd = MCl1.Offset(0, 3).Value        '���q�R�[�h
    BKmn = MCl1.Offset(0, 39).Value        '���q��

   GetMM = "�s�ǏW�v�y�i���z"

'�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�i���ʁ|�Y�����j
    Worksheets(GetMM).Activate
   '�����J�n�ʒu�̐ݒ�
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")
   '�C���f�b�N�X�����l
    i = 5
   '���f�[�^�̈�m�F
    Do Until MCl2.Value = ""
       i = i + 1
       Set MCl2 = MCl2.Offset(1, 0)
    Loop

   '�N���A�͈͎w��
    Range(Cells(6, 1), Cells(i, 14)).Select
    Selection.ClearContents

'���ђǉ������|�i����
   '�ǉ���V�[�g�����J�n�ʒu�w��
    Set MCl2 = Workbooks(ABk).Worksheets(GetMM).Range("A6")

   '�i���ʏW�v
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
          .Offset(0, 0).Value = BKcd       '���q�R�[�h
          .Offset(0, 1).Value = BKmn       '���q��
          .Offset(0, 2).Value = Com32      '�Ǖi��
          .Offset(0, 3).Value = Com18      '�s�ǐ�
          .Offset(0, 4).Value = Com19      '�{�X����\
          .Offset(0, 5).Value = Com20      '�{�X���ꗠ
          .Offset(0, 6).Value = Com21      '���؊���
          .Offset(0, 7).Value = Com22      '�t�B������
          .Offset(0, 8).Value = Com23      '���؏[�U
          .Offset(0, 9).Value = Com24      '�t�B���[�U
          .Offset(0, 10).Value = Com25     '�L�����h���c
          .Offset(0, 11).Value = Com26     '���̑�
          .Offset(0, 12).Value = Com17     '�蒼�s��
          '.Offset(0, 13).Value = Com24 / Com32     '�p���s�Ǘ�
          '.Offset(0, 14).Value = Com17 / Com32     '�蒼�s�Ǘ�
'
          If Com18 <> 0 Then
             WkCom = Com18 / (Com18 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 13).Value = WkCom     '�p���s�Ǘ�
'
          If Com17 <> 0 Then
             WkCom = Com17 / (Com17 + Com32)
            Else
             WkCom = 0
          End If
          .Offset(0, 14).Value = WkCom     '�蒼�s�Ǘ�
'
       End With
       Set MCl2 = MCl2.Offset(1, 0)
       BKcd = MCl1.Offset(0, 3).Value
       BKmn = MCl1.Offset(0, 39).Value

   '��ƃG���A������
       Com17 = 0   '�蒼�s�ǁi�Ǖi�Ɋ܂܂��j
       Com18 = 0   '�p���s��
       Com19 = 0   '�{�X����\
       Com20 = 0   '�{�X���ꗠ
       Com21 = 0   '���؊���
       Com22 = 0   '�t�B������
       Com23 = 0   '���؏[�U
       Com24 = 0   '�t�B���[�U
       Com25 = 0   '�L�����h���c
       Com26 = 0   '���̑�
       Com32 = 0   '�Ǖi��
       ComWK = 0   '�v�Z���[�N
    Loop
          















'*********************************************************************************
'************��������@�@�@�@20130313kometani�ǉ��@�@�@�@��������*****************
'*********************************************************************************
         
'�i���ʃV���b�g���W�v�J�n

    Dim wb As Workbook
    Dim ���i�� As Object
    Dim ��i�� As Object
    Dim ���Y�� As Variant
    Dim YandM As Variant
    Dim temp As Object
    
    '�V���b�g���W�v�t�@�C����ǂݏo��
    Set wb = Workbooks.Open(Filename:=ThisWorkbook.Path & "\..\..\�V���b�g�Ǘ��\\�y�g��z�V���b�g���W�v.xls ")
    
    '�W�v���̎Z�o
    Set ���Y�� = ThisWorkbook.Worksheets("�������").Range("G5")
    If Month(���Y��.Value) <> 12 Then
        YandM = Year(���Y��.Value) & "�N" & (Month(���Y��.Value) + 1) & "���x"
    Else
        YandM = (Year(���Y��.Value) + 1) & "�N" & "1���x"
    End If
    '�V���b�g������͂��Ă����������
    Set temp = wb.Worksheets("�g�ꒆ�q�H��").Range("J3")
    Do While temp.Value <> ""
        If YandM <> temp.Value Then
            Set temp = temp.Offset(0, 1)
        Else
            '�V���b�g������͂��Ă�������m��
            temp.Activate
            temp.Font.ColorIndex = 1
            Exit Do
        End If
    Loop
    
    '����W�v�t�@�C���́u�i���ʏW�v�v�V�[�g�̃Z���F������(���F)
    '   �V���b�g������������Ƀ`�F�b�N�p�Ƃ���
    '   ���i�Ԃ̃Z���F��Ԃɂ��鏈����ǉ���������
    ThisWorkbook.Worksheets("�i���ʏW�v").Range("B7:B41").Interior.ColorIndex = 2
    
    Set ���i�� = ThisWorkbook.Worksheets("�i���ʏW�v").Range("C7")
    Do While ���i��.Value <> ""
        '��i�Ԃ̏�����
        Set ��i�� = wb.Worksheets("�g�ꒆ�q�H��").Range("D6")
        '��i�Ԃ̌���
        Do While ���i��.Value <> ��i��.Value
            Set ��i�� = ��i��.Offset(1, 0)
            '������Ȃ������ꍇ(���[�v�𔲂���)
            If ��i��.Value = "" Then
                GoTo rt1
            End If
        Loop
        '�l�̏�������
        If ���i��.Value = 8 Then 'BP4Y�̏ꍇ
            With wb.Worksheets("�g�ꒆ�q�H��")
                'AB�^�ɑ΂���
                .Cells(��i��.Row, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 4).Value / 4
                .Cells(��i��.Row, ActiveCell.Column).Font.ColorIndex = 1
                'CD�^�ɑ΂���
                .Cells(��i��.Row + 5, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 4).Value / 4
                .Cells(��i��.Row + 5, ActiveCell.Column).Font.ColorIndex = 1
                'EF�^�ɑ΂���
                .Cells(��i��.Row + 10, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 4).Value / 4
                .Cells(��i��.Row + 10, ActiveCell.Column).Font.ColorIndex = 1
                'GH�^�ɑ΂���
                .Cells(��i��.Row + 15, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 4).Value / 4
                .Cells(��i��.Row + 15, ActiveCell.Column).Font.ColorIndex = 1
            End With
        ElseIf ���i��.Value = 12 Then 'DF71�̏ꍇ
            With wb.Worksheets("�g�ꒆ�q�H��")
                '�P�Ԍ^�ɑ΂���
                .Cells(��i��.Row, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 4).Value / 2
                .Cells(��i��.Row, ActiveCell.Column).Font.ColorIndex = 1
                '�Q�Ԍ^�ɑ΂���
                .Cells(��i��.Row + 5, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 4).Value / 2
                .Cells(��i��.Row + 5, ActiveCell.Column).Font.ColorIndex = 1
            End With
        Else '���^���P�^�����Ȃ����̑��̕i��
            With wb.Worksheets("�g�ꒆ�q�H��")
                .Cells(��i��.Row, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 4).Value
                .Cells(��i��.Row, ActiveCell.Column).Font.ColorIndex = 1
            End With
        End If
        '�`�F�b�N�p
        ���i��.Offset(0, -1).Interior.ColorIndex = 3
        
rt1:
    
        Set ���i�� = ���i��.Offset(1, 0)
        'MsgBox ���i��.Value
        
        
        
        '�V�K���q�Ȃǂւ̑Ή�
        '
        '
        '
        '
        '
    Loop

    
    
    '�[����������
    Dim NowCell As Object '���ݎQ�ƒ��Z��

    Set NowCell = ActiveCell
    '�Q�ƒ��̃Z����2000�s�ɂ����܂Ń��[�v(�o�_��1100�s�����Ȃ�����)
    Do While NowCell.Row < 2000
        If NowCell.Font.ColorIndex = 3 Then '�Z�����̕������ԂȂ��
            NowCell.Value = 0               '���e���u�O�v�ɂ���
            NowCell.Font.ColorIndex = 1     '�����F�����ɂ���
        End If
        Set NowCell = NowCell.Offset(1, 0)  '�Q�ƒ��̃Z�������ɂP���炷
    Loop
    
    
    
    '���σV���b�g���i�ߋ�6�����j�X�V
    Set ��i�� = wb.Worksheets("�g�ꒆ�q�H��").Range("D6")
    Do Until ��i�� = ""
        If ��i�� <> ��i��.Offset(-1, 0) Then
            With wb.Worksheets("�g�ꒆ�q�H��")
                '���σV���b�g���i�ߋ�6�����j
                .Cells(��i��.Row, 7).FormulaR1C1 = "=sum(RC[" & temp.Column - 7 - 5 & "]:RC[" & temp.Column - 7 - 0 & "])/6"
            End With
        End If
        Set ��i�� = ��i��.Offset(1, 0)
    Loop
    
    
    
    '��U�������̗�ǉ�
    Dim fy, fm, s, t As Integer
    Dim temp2 As Object
    
    Set temp2 = temp
    t = 0
    For s = 1 To 6  '6�����i���N�j���������J��Ԃ�
        fm = (Month(���Y��.Value) + 1) + s
        fy = Year(���Y��.Value)
        '�N�z������
        If fm > 12 Then
            fm = fm Mod 12
            fy = fy + 1
            t = t + 1
        End If
        '��ǉ��̗v�E�s�v�𔻒f
        If temp2.Offset(0, 1).Value = fy & "�N" & fm & "���x" Then
            '��ǉ��s�v
            Set temp2 = temp2.Offset(0, 1)
        Else
            '��ǉ��K�v
            Columns(temp2.Column).Copy
            Columns(temp2.Offset(0, 1).Column).Insert
            Set temp2 = temp2.Offset(0, 1)
            temp2.Value = fy & "�N" & fm & "���x"
        End If
    Next
    
    '�C���`�F�b�N
    Application.Run "�y�g��z�V���b�g���W�v.xls!�C��check"
    
    Application.DisplayAlerts = False
    wb.Close (True)
    Application.DisplayAlerts = True
    
'*********************************************************************************
'************�����܂Ł@�@�@�@20130313kometani�ǉ��@�@�@�@�����܂�*****************
'*********************************************************************************








   '�ʒu�̐ݒ�
    Range("A1").Select
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "�������I���܂����B", vbOKOnly + vbInformation, "�ʒm"
End Sub




Sub �Z���F������()

    ThisWorkbook.Worksheets("�i���ʏW�v").Range("B7:B41").Interior.ColorIndex = 2

End Sub















