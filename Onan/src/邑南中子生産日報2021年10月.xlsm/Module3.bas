Attribute VB_Name = "Module3"
'Option Explicit

Public Sub �������ђǉ�����()

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

   '�����ݒ�
   Application.ScreenUpdating = False
   '���i��p�f�o�b�O
   'Call logger.Init("D:\Daily_Production_Report_Summary\bin\test\debug.log")
   mst_machine = "�}�V����"
   nippo_syukei_sheet = "����W�v"
   nippo_nyuryoku_sheet = "�������"
   sagyohyo_sheet = "��ƕ\"
   '�����J�n
   myBtn = MsgBox("�������ђǉ��������J�n���܂�", vbYesNo + vbExclamation, "�������ђǉ�����")
   If myBtn = vbNo Then
      Exit Sub
   End If
   'Call logger.WriteLog("�����J�n")

   '��Ɨ̈�N���A�i��ƕ\�j
   Worksheets(sagyohyo_sheet).Activate
   Range("A5:AM2000").Select
   Selection.ClearContents
   Range("A5").Select

   '�����J�n�ʒu�̐ݒ�
   Set nippo_syukei_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_syukei_sheet).Range("A5")
   Set nippo_nyuryoku_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_nyuryoku_sheet).Range("G5")
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '����W�v�V�[�g�̍X�V
   Call NippouShuukei_Update(nippo_nyuryoku_cell, nippo_syukei_cell)

   '�����J�n�ʒu�̐ݒ�
   Set nippo_syukei_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_syukei_sheet).Range("A5")
   Set nippo_nyuryoku_cell = Workbooks(ActiveWorkbook.Name).Worksheets(nippo_nyuryoku_sheet).Range("G5")

   '���уf�[�^�m�F
   Do Until nippo_syukei_cell.Value = ""
      With nippo_syukei_cell
      '�f�[�^�ڍs
         For i = 0 To 39
            first_cell_of_sagyohyo.Offset(0, i).Value = .Offset(0, i).Value
         Next i
      End With
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
      Set nippo_syukei_cell = nippo_syukei_cell.Offset(1, 0)
   Loop

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���}�V���ʏW�v��ƊJ�n
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
   Worksheets(sagyohyo_sheet).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   '�C���f�b�N�X������
   i = 4
   '���f�[�^�̈�m�F
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   '�}�V���ʂɕ��ёւ�
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("B")

   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

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
   SVtime = 0  '�o�Α�����
   count = 0   '���^������
'
   nakago_code = first_cell_of_sagyohyo.Offset(0, 1).Value
   SVtime = first_cell_of_sagyohyo.Offset(-4, 0).Value
'
   update_target = "�}�V���ʏW�v"

'�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�}�V���ʁ|�Y�����j
   Worksheets(update_target).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
   '�C���f�b�N�X�����l
   i = 7
   '���f�[�^�̈�m�F
   Do Until first_cell_of_target_summary.Value = ""
      i = i + 1
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   Loop
   '�N���A�͈͎w��
   Range(Cells(7, 1), Cells(i, 32)).Select
   Selection.ClearContents

'�}�V������荞��
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

'���ђǉ������|�}�V����
   '�}�V���ʏW�v
   Do Until first_cell_of_sagyohyo.Value = ""
      '�ǉ���V�[�g�����J�n�ʒu�w��
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
      '�}�V���R�[�h�ʒu�ݒ�
      Do Until nakago_code = first_cell_of_target_summary.Offset(0, 0).Value
         Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      Loop
      With first_cell_of_target_summary
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
         .Offset(0, 27).Value = (Com2 / 60) / SVtime '�ݔ����ח�
         .Offset(0, 28).Value = Com3 / Com2   '�ݔ��ғ���
         .Offset(0, 29).Value = Com30 / (Com2 / 60)  '�J�����Y���i�}�V���j
         .Offset(0, 30).Value = Com30 / (Com4 / 60)  '�J�����Y���i�l�j

         If Com18 <> 0 Then
            WkCom = Com18 / (Com18 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 26).Value = WkCom     '�s�Ǘ�
      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      nakago_code = first_cell_of_sagyohyo.Offset(0, 1).Value
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
      count = 0   '���^������
   Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '�V�}�V���ʏW�v��ƊJ�n
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
   Worksheets(sagyohyo_sheet).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   '�C���f�b�N�X������
   i = 4
   '���f�[�^�̈�m�F
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop
   '�}�V�����ƒ��q���ƕ��������Ń\�[�g
   With ActiveSheet
      .Sort.SortFields.Clear
      '�}�V����
      .Sort.SortFields.Add _
         Key:=ActiveSheet.Range("B5")
      '���q��
      .Sort.SortFields.Add _
         Key:=ActiveSheet.Range("D5")
      With .Sort
         .SetRange Range(Cells(5, 1), Cells(i, 41))
         .Apply
      End With
   End With
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   SVtime = first_cell_of_sagyohyo.Offset(-4, 0).Value  '�o�Α�����
   count = 0   '���^������
   update_target = "�V�}�V���ʏW�v"
   '�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�}�V���ʁ|�Y�����j
   Worksheets(update_target).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
   last_row = Range("B7").End(xlDown).Row
   '�N���A�͈͎w��
   Range(first_cell_of_target_summary, Range("AF" & last_row)).Select
   Selection.ClearContents
   '���ђǉ������|�}�V����
   '�}�V���ʏW�v
   Dim read_index As Variant
   read_index = Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 30, 34, 35, 36, 37, 38)
   '[0] = 4 // �V���b�g��
   '[1] = 5 // �ғ�����
   '[2] = 6 // ���Y����
   '[3] = 7 // OP��Ǝ���
   '[4] = 8 // �n�ƍ��
   '[5] = 9 // ���^����
   '[6] = 10 // �����҂�
   '[7] = 11 // ���^����
   '[8] = 12 // �}�V���̏��~
   '[9] = 13 // �I�ƍ��
   '[10] = 14 // �^���|
   '[11] = 15 // Rb����
   '[12] = 16 // ���@�Ή��҂�
   '[13] = 17 // ���U��
   '[14] = 18 // ���q���ꏈ��
   '[15] = 19 // ���̑�
   '[16] = 20 // �蒼�s��
   '[17] = 21 // ���`�s�ǐ�
   '[18] = 30 // �Ǖi��
   '[19] = 34 // ����
   '[20] = 35 // �Ǖi��
   '[21] = 36 // �s�ǐ�
   '[22] = 37 // ���Y���z
   '[23] = 38 // �s�ǋ��z

   blank_row = 7
   Do Until first_cell_of_sagyohyo.Value = ""
      Dim nippo_by_nakago(23) As Long
      Erase nippo_by_nakago
      nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value
      machine_code = first_cell_of_sagyohyo.Offset(0, 1).Value
      bk_machine_code = 0
      nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value
      '���[�v�����F���q�R�[�h���ς��܂ŁB
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
         '1�s�ǂݏI������玟�s��
         Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
      Loop
      '�}�V���R�[�h���O��Ɠ������ǂ���
      If bk_machine_code = machine_code Then
         Cells(blank_row, 1).EntireRow.Insert
      End If
      bk_machine_code = machine_code
      blank_row = blank_row + 1
      With first_cell_of_target_summary
         .Offset(0, 0).Value = machine_code
         .Offset(0, 1).Value = WorksheetFunction.VLookup(machine_code, Workbooks(ActiveWorkbook.Name).Worksheets("�}�V����").Range("B:C"), 2)
         .Offset(0, 2).Value = nakago_name
         .Offset(0, 3).Value = nippo_by_nakago(0)      '�V���b�g��
         .Offset(0, 4).Value = nippo_by_nakago(18)     '�Ǖi��
         .Offset(0, 5).Value = nippo_by_nakago(17)     '�s�ǐ�
         .Offset(0, 6).Value = nippo_by_nakago(1) / 60     '�}�V���ғ�����
         .Offset(0, 7).Value = nippo_by_nakago(2) / 60     '�}�V�����Y����
         .Offset(0, 8).Value = nippo_by_nakago(3) / 60     '�n�o��Ǝ���
         .Offset(0, 9).Value = nippo_by_nakago(4) / 60     '�n�ƍ��
         .Offset(0, 10).Value = nippo_by_nakago(5) / 60     '���^����
         .Offset(0, 11).Value = nippo_by_nakago(6) / 60    '�����҂�
         .Offset(0, 12).Value = count      '�^�����񐔁i�ǂ�����H�j
         .Offset(0, 13).Value = nippo_by_nakago(7) / 60    '�^����
         .Offset(0, 14).Value = nippo_by_nakago(8) / 60    '�̏��~
         .Offset(0, 15).Value = nippo_by_nakago(10) / 60   '���^���|
         .Offset(0, 16).Value = nippo_by_nakago(9) / 60   '�I�����
         .Offset(0, 17).Value = nippo_by_nakago(11) / 60   '�q������
         .Offset(0, 18).Value = nippo_by_nakago(12) / 60   '���@�Ή��҂�
         .Offset(0, 19).Value = nippo_by_nakago(13) / 60   '���^��
         .Offset(0, 20).Value = nippo_by_nakago(14) / 60   '���q���ꏈ��
         .Offset(0, 21).Value = nippo_by_nakago(15) / 60   '���̑�
         .Offset(0, 22).Value = nippo_by_nakago(19) / 1000  '�g�p��
         .Offset(0, 23).Value = nippo_by_nakago(20) / 1000  '�Ǖi�g�p��
         .Offset(0, 24).Value = nippo_by_nakago(21) / 1000  '�s�ǎg�p��
         .Offset(0, 25).Value = nippo_by_nakago(22) / 1000  '���Y���z
         .Offset(0, 26).Value = nippo_by_nakago(23) / 1000  '�s�ǋ��z
         If nippo_by_nakago(17) <> 0 Then
            WkCom = nippo_by_nakago(17) / (nippo_by_nakago(17) + nippo_by_nakago(18))
         Else
            WkCom = 0
         End If
         .Offset(0, 27).Value = WkCom    '�s�Ǘ�
         .Offset(0, 28).Value = (nippo_by_nakago(1) / 60) / SVtime '�ݔ����ח�
         .Offset(0, 29).Value = nippo_by_nakago(2) / nippo_by_nakago(1)   '�ݔ��ғ���
         .Offset(0, 30).Value = nippo_by_nakago(22) / (nippo_by_nakago(1) / 60)  '�J�����Y���i�}�V���j
         .Offset(0, 31).Value = nippo_by_nakago(22) / (nippo_by_nakago(3) / 60)  '�J�����Y���i�l�j
      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      '��ƃG���A������
      count = 0   '���^������
   Loop
   '�ŏI�s�ǉ�
   last_row = Range("B7").End(xlDown).Row + 1
   If last_row > 100000 Then
      last_row = 8
   End If
   Range("B" & last_row) = "���v"
   With Range("D" & last_row)
      .Formula = "=SUM(D7:D" & (last_row - 1) & " )"
      .AutoFill Destination:=.Resize(1, 24)
   End With
   Range("AB" & last_row) = Range("F" & last_row).Value / (Range("E" & last_row).Value + Range("F" & last_row).Value)
   Range("AC" & last_row).Formula = "=AVERAGE(AC7:AC" & (last_row - 1) & " )"
   Range("AD" & last_row) = Range("H" & last_row).Value / Range("G" & last_row).Value
   Range("AE" & last_row) = Range("Z" & last_row).Value * 1000 / Range("H" & last_row).Value
   Range("AF" & last_row) = Range("Z" & last_row).Value * 1000 / Range("I" & last_row).Value

   '�ŏI�s�F�t
   Range("A" & 7 & ":AF" & last_row).Interior.ColorIndex = 0
   Range("A" & last_row & ":AF" & last_row).Interior.ColorIndex = 20

   '�i���ʏW�v��ƊJ�n
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
   Worksheets(sagyohyo_sheet).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   '�C���f�b�N�X������
   i = 4

   '���f�[�^�̈�m�F
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   '�i���ʂɕ��ёւ�
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("D")

   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '��Ɨ̈揉����
   Com1 = 0   '�V���b�g
   Com2 = 0   '�ғ�����
   Com3 = 0   '���Y����
   Com4 = 0   '�n�o��Ǝ���
   Com5 = 0   '�n�Ǝ���
   Com6 = 0   '���^����
   Com7 = 0   '�����҂�
   Com8 = 0   '���^����
   Com9 = 0   '�}�V���̏��~
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
   count = 0   '���^������

   nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value      '���q�R�[�h
   nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value      '���q��

   update_target = "�i���ʏW�v"

   '�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�}�V���ʁ|�Y�����j
   Worksheets(update_target).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_target_summary = Worksheets(update_target).Range("A7")
   last_row = Range("B7").End(xlDown).Row
   '�N���A�͈͎w��
   Range(first_cell_of_target_summary, Range("AJ" & last_row)).Select
   Selection.ClearContents
   '���ђǉ������|�i����
   '�ǉ���V�[�g�����J�n�ʒu�w��
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A7")
   '�i���ʏW�v
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

      With first_cell_of_target_summary
         .Offset(0, 0).Formula = "=Row()-6"
         .Offset(0, 1).Value = nakago_name      '���q��
         .Offset(0, 2).Value = nakago_code      '���q�R�[�h�@'20140408kometani�@�ǉ�
         .Offset(0, 3).Value = Com1      '�V���b�g��
         .Offset(0, 4).Value = Com32     '�Ǖi��
         .Offset(0, 5).Value = Com18     '�s�ǐ�
         .Offset(0, 6).Value = Com2 / 60     '�}�V���ғ�����
         .Offset(0, 7).Value = Com3 / 60     '�}�V�����Y����
         .Offset(0, 8).Value = Com4 / 60     '�n�o��Ǝ���
         .Offset(0, 9).Value = Com5 / 60     '�n�ƍ��
         .Offset(0, 10).Value = Com6 / 60    '���^����
         .Offset(0, 11).Value = Com7 / 60    '�����҂�
         .Offset(0, 12).Value = count      '�^������
         .Offset(0, 13).Value = Com8 / 60    '�^����
         .Offset(0, 14).Value = Com9 / 60    '�̏��~
         .Offset(0, 15).Value = Com11 / 60   '���^���|
         .Offset(0, 16).Value = Com10 / 60   '�I�����
         .Offset(0, 17).Value = Com12 / 60   '�q������
         .Offset(0, 18).Value = Com13 / 60   '���@�Ή��҂�
         .Offset(0, 19).Value = Com14 / 60   '���^��
         .Offset(0, 20).Value = Com15 / 60   '���q���ꏈ��
         .Offset(0, 21).Value = Com16 / 60   '���̑�
         .Offset(0, 22).Value = Com27      '�g�p��
         .Offset(0, 23).Value = Com28      '�Ǖi�g�p��
         .Offset(0, 24).Value = Com29      '�s�ǎg�p��
         .Offset(0, 25).Value = Com30  '���Y���z
         .Offset(0, 26).Value = Com31  '�s�ǋ��z
         .Offset(0, 28).Value = (Com2 / 60) / SVtime '�ݔ����ח�
         If Com2 <> 0 Then
            .Offset(0, 29).Value = Com3 / Com2   '�ݔ��ғ���
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
            .Offset(0, 30).Value = Com30 / (Com2 / 60)  '�J�����Y���i�}�V���j
            .Offset(0, 31).Value = Com30 / (Com4 / 60)  '�J�����Y���i�l�j
         Else
            .Offset(0, 30).Value = 0
            .Offset(0, 31).Value = 0
         End If
         .Offset(0, 33).Formula = "=VLOOKUP(C" & first_cell_of_target_summary.Row & " , ���q�f�[�^!A4:J800,9)"  '�ݒ�T�C�N��
         If Com1 <> 0 Then
            .Offset(0, 32).Value = (Com3 / 60) / Com1 * 3600 '���уT�C�N��
            .Offset(0, 34).Value = .Offset(0, 33).Value / .Offset(0, 32).Value  '���\�ғ���
         Else
            .Offset(0, 32).Value = 0
            .Offset(0, 34).Value = 0
         End If
         .Offset(0, 35).Value = .Offset(0, 29).Value * .Offset(0, 34).Value * (1 - .Offset(0, 27).Value)  '�ݔ���������
      End With

      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value
      nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value

      '��ƃG���A������
      Com1 = 0   '�V���b�g
      Com2 = 0   '�ғ�����
      Com3 = 0   '���Y����
      Com4 = 0   '�n�o��Ǝ���
      Com5 = 0   '�n�Ǝ���
      Com6 = 0   '���^����
      Com7 = 0   '�����҂�
      Com8 = 0   '���^����
      Com9 = 0   '�}�V���̏��~
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
      count = 0   '���^������
   Loop

   '�ŏI�s�ǉ�
   last_row = Range("B7").End(xlDown).Row + 1
   If last_row > 100000 Then
      last_row = 8
   End If

   '���Y���z���Ƀ\�[�g
   Range("A7:AJ" & last_row - 1).Sort _
      Key1:=Range("Z7"), Order1:=xlDescending

   With Worksheets(update_target)
      .Range("B" & last_row) = "���v"
      With .Range("D" & last_row)
         .Formula = "=SUM(D7:D" & (last_row - 1) & " )"
         .AutoFill Destination:=.Resize(1, 24)
      End With
      .Range("AB" & last_row) = .Range("F" & last_row).Value / (.Range("E" & last_row).Value + .Range("F" & last_row).Value)
      .Range("AC" & last_row).Formula = "=AVERAGE(AC7:AC" & (last_row - 1) & " )"
      .Range("AD" & last_row) = .Range("H" & last_row).Value / .Range("G" & last_row).Value
      .Range("AE" & last_row) = .Range("Z" & last_row).Value / .Range("H" & last_row).Value
      .Range("AF" & last_row) = .Range("Z" & last_row).Value / .Range("I" & last_row).Value
      .Range("AG" & last_row) = .Range("H" & last_row).Value * 3600 / .Range("D" & last_row).Value
      .Range("AI" & last_row).Formula = "=SUMPRODUCT(D7:D" & (last_row - 1) & " ,AG7:AG" & (last_row - 1) & ") / (H" & last_row & " * 3600)"
      .Range("AJ" & last_row) = .Range("AI" & last_row).Value * .Range("AD" & last_row).Value * (1 - .Range("AB" & last_row).Value)
   End With

   '�ŏI�s�F�t
   Range("A" & 7 & ":AJ" & last_row).Interior.ColorIndex = 0
   Range("A" & last_row & ":AJ" & last_row).Interior.ColorIndex = 20

   '�}�V���ʕs�ǏW�v��ƊJ�n
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
   Worksheets(sagyohyo_sheet).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")
   '�C���f�b�N�X������
   i = 4
   '���f�[�^�̈�m�F
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   '�}�V���ʂɕ��ёւ�
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("B")

   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

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

   update_target = "�s�ǏW�v�y�}�V���z"

   '�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u��
   Worksheets(update_target).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")
   '�C���f�b�N�X�����l
   i = 5
   '���f�[�^�̈�m�F
   Do Until first_cell_of_target_summary.Value = ""
      i = i + 1
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   Loop
   '�N���A�͈͎w��
   Range(Cells(6, 1), Cells(i, 15)).Select
   Selection.ClearContents

   '�}�V������荞��
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

   '���ђǉ������|�}�V����
   '�ǉ���V�[�g�����J�n�ʒu�w��
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")

   machine_code = first_cell_of_sagyohyo.Offset(0, 1).Value
   '�}�V���ʏW�v
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
      '�}�V���R�[�h�ʒu�ݒ�
      Do Until machine_code = first_cell_of_target_summary.Offset(0, 0).Value
         Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)

      Loop
      With first_cell_of_target_summary
         .Offset(0, 2).Value = Com32     '�Ǖi��
         .Offset(0, 3).Value = Com18     '�s�ǐ�
         .Offset(0, 4).Value = Com19     '�{�X����\
         .Offset(0, 5).Value = Com20     '�{�X���ꗠ
         .Offset(0, 6).Value = Com21     '���؊���
         .Offset(0, 7).Value = Com22     '�t�B������
         .Offset(0, 8).Value = Com23     '���؏[�U
         .Offset(0, 9).Value = Com24     '�t�B���[�U
         .Offset(0, 10).Value = Com25    '�L�����h���c
         .Offset(0, 11).Value = Com26    '���̑�
         .Offset(0, 12).Value = Com17    '�蒼�s��
         If Com18 <> 0 Then
            WkCom = Com18 / (Com18 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 13).Value = WkCom    '�p���s�Ǘ�

         If Com17 <> 0 Then
            WkCom = Com17 / (Com17 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 14).Value = WkCom    '�蒼�s�Ǘ�

      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
      machine_code = first_cell_of_sagyohyo.Offset(0, 1).Value
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
   Loop

   '�ʒu�̐ݒ�
   Range("A1").Select

   '�i���ʕs�ǏW�v��ƊJ�n
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i��ƕ\�j
   Worksheets(sagyohyo_sheet).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

   '�C���f�b�N�X������
   i = 4

   '���f�[�^�̈�m�F
   Do Until first_cell_of_sagyohyo.Value = ""
      i = i + 1
      Set first_cell_of_sagyohyo = first_cell_of_sagyohyo.Offset(1, 0)
   Loop

   '�i���ʂɕ��ёւ�
   Range(Cells(5, 1), Cells(i, 41)).Sort _
   Key1:=Columns("D")

   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_sagyohyo = Workbooks(ActiveWorkbook.Name).Worksheets(sagyohyo_sheet).Range("A5")

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

   '�ǉ���V�[�g������
   '��Ɨp���[�N�V�[�g�A�N�e�B�u���i�i���ʁ|�Y�����j
   update_target = "�s�ǏW�v�y�i���z"
   Worksheets(update_target).Activate
   '�����J�n�ʒu�̐ݒ�
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")
   '�C���f�b�N�X�����l
   i = 5
   '���f�[�^�̈�m�F
   Do Until first_cell_of_target_summary.Value = ""
      i = i + 1
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
   Loop

   '�N���A�͈͎w��
   Range(Cells(6, 1), Cells(i, 14)).Select
   Selection.ClearContents

   '���ђǉ������|�i����
   '�ǉ���V�[�g�����J�n�ʒu�w��
   Set first_cell_of_target_summary = Workbooks(ActiveWorkbook.Name).Worksheets(update_target).Range("A6")

   '�i���ʏW�v
   Do Until first_cell_of_sagyohyo.Value = ""
   nakago_code = first_cell_of_sagyohyo.Offset(0, 3).Value      '���q�R�[�h
   nakago_name = first_cell_of_sagyohyo.Offset(0, 39).Value      '���q��
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
         .Offset(0, 0).Value = nakago_code      '���q�R�[�h
         .Offset(0, 1).Value = nakago_name      '���q��
         .Offset(0, 2).Value = Com32     '�Ǖi��
         .Offset(0, 3).Value = Com18     '�s�ǐ�
         .Offset(0, 4).Value = Com19     '�{�X����\
         .Offset(0, 5).Value = Com20     '�{�X���ꗠ
         .Offset(0, 6).Value = Com21     '���؊���
         .Offset(0, 7).Value = Com22     '�t�B������
         .Offset(0, 8).Value = Com23     '���؏[�U
         .Offset(0, 9).Value = Com24     '�t�B���[�U
         .Offset(0, 10).Value = Com25    '�L�����h���c
         .Offset(0, 11).Value = Com26    '���̑�
         .Offset(0, 12).Value = Com17    '�蒼�s��
         If Com18 <> 0 Then
            WkCom = Com18 / (Com18 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 13).Value = WkCom    '�p���s�Ǘ�
         If Com17 <> 0 Then
            WkCom = Com17 / (Com17 + Com32)
         Else
            WkCom = 0
         End If
         .Offset(0, 14).Value = WkCom    '�蒼�s�Ǘ�
      End With
      Set first_cell_of_target_summary = first_cell_of_target_summary.Offset(1, 0)
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
   Loop
      '�i���ʃV���b�g���W�v�J�n

    Dim wb As Workbook
    Dim ���i�� As Object
    Dim ��i�� As Object
    Dim ���Y�� As Variant
    Dim YandM As Variant
    Dim temp As Object
    
    '�V���b�g���W�v�t�@�C����ǂݏo��
    Set wb = Workbooks.Open(Filename:=ThisWorkbook.Path & "\..\�V���b�g�Ǘ��\\�y�W��z�V���b�g���W�v.xls ")
    
    '�W�v���̎Z�o
    Set ���Y�� = ThisWorkbook.Worksheets("�������").Range("G5")
    If Month(���Y��.Value) <> 12 Then
        YandM = Year(���Y��.Value) & "�N" & (Month(���Y��.Value) + 1) & "���x"
    Else
        YandM = (Year(���Y��.Value) + 1) & "�N" & "1���x"
    End If
    '�V���b�g������͂��Ă����������
    Set temp = wb.Worksheets("�W�쒆�q�H��").Range("J3")
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
    ThisWorkbook.Worksheets("�i���ʏW�v").Range("B7:B99").Interior.ColorIndex = 2
    
    Set ���i�� = ThisWorkbook.Worksheets("�i���ʏW�v").Range("B7")
    Do While ���i��.Value <> ""
        '��i�Ԃ̏�����
        Set ��i�� = wb.Worksheets("�W�쒆�q�H��").Range("B6")
        '��i�Ԃ̌���
        Do While ���i��.Value <> ��i��.Value
            Set ��i�� = ��i��.Offset(1, 0)
            '������Ȃ������ꍇ(���[�v�𔲂���)
            If ��i��.Value = "" Then
                GoTo rt1
            End If
        Loop
        '�l�̏�������
        With wb.Worksheets("�W�쒆�q�H��")
            .Cells(��i��.Row, ActiveCell.Column).Value = ThisWorkbook.Worksheets("�i���ʏW�v").Cells(���i��.Row, 3).Value
            .Cells(��i��.Row, ActiveCell.Column).Font.ColorIndex = 1
        End With
        '�`�F�b�N�p
        ���i��.Interior.ColorIndex = 3
        
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
    
    
    
    
    '���σV���b�g���i�ߋ�6�����j�X�V  '20121001�ǉ�
    Set ��i�� = wb.Worksheets("�W�쒆�q�H��").Range("B6")
    Do Until ��i�� = ""
        If ��i�� <> ��i��.Offset(-1, 0) Then
            With wb.Worksheets("�W�쒆�q�H��")
                '���σV���b�g���i�ߋ�6�����j
                .Cells(��i��.Row, 6).FormulaR1C1 = "=sum(RC[" & temp.Column - 6 - 5 & "]:RC[" & temp.Column - 6 - 0 & "])/6"
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
            'If t = 0 Then  '20121203�C���@�N�x���������\�����ꂸ�A���̂��߂�if����������Ȃ�
                fy = fy + 1
                t = t + 1
            'End If
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
    Application.Run "�y�W��z�V���b�g���W�v.xls!�C��check"
    
    
    
    
    
    
    Application.DisplayAlerts = False
    wb.Close (True)
    Application.DisplayAlerts = True
    
   '�ʒu�̐ݒ�
    Range("A1").Select
         
    Application.ScreenUpdating = True
    MsgBox "�������I���܂����B", vbOKOnly + vbInformation, "�ʒm"
End Sub



Sub �Z���F������()

    ThisWorkbook.Worksheets("�i���ʏW�v").Range("B7:B99").Interior.ColorIndex = 2

End Sub

