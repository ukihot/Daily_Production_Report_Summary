Attribute VB_Name = "Module2"

Public Sub NippouShuukei_Update(nippo_nyuryoku_cell As Object, nippo_syukei_cell As Object)
    '�Z��������
    ThisWorkbook.Worksheets("����W�v").Range("A5:AN600").ClearContents
    Range("A5").Select

    '��ƕ\�쐬
    Do Until nippo_nyuryoku_cell.Value = ""
        nippo_syukei_cell.Offset(0, 0).Value = nippo_nyuryoku_cell.Offset(0, 0).Value '���Y��
        nippo_syukei_cell.Offset(0, 1).Value = nippo_nyuryoku_cell.Offset(0, 1).Value '�}�V��
        nippo_syukei_cell.Offset(0, 2).Value = nippo_nyuryoku_cell.Offset(0, 2).Value '��Ǝ�
        nippo_syukei_cell.Offset(0, 3).Value = nippo_nyuryoku_cell.Offset(0, 3).Value '���q
        nippo_syukei_cell.Offset(0, 4).Value = nippo_nyuryoku_cell.Offset(0, 4).Value '�V���b�g
        nippo_syukei_cell.Offset(0, 5).Value = nippo_nyuryoku_cell.Offset(0, 5).Value '�ғ�����
        nippo_syukei_cell.Offset(0, 6).Value = nippo_nyuryoku_cell.Offset(0, 7).Value '���Y����
        nippo_syukei_cell.Offset(0, 7).Value = nippo_nyuryoku_cell.Offset(0, 5).Value * nippo_nyuryoku_cell.Offset(0, 6) 'OP��Ǝ���
        nippo_syukei_cell.Offset(0, 8).Value = nippo_nyuryoku_cell.Offset(0, 8).Value '�n�ƍ��
        nippo_syukei_cell.Offset(0, 9).Value = nippo_nyuryoku_cell.Offset(0, 9).Value '���^����
        nippo_syukei_cell.Offset(0, 10).Value = nippo_nyuryoku_cell.Offset(0, 10).Value '�����҂�
        nippo_syukei_cell.Offset(0, 11).Value = nippo_nyuryoku_cell.Offset(0, 11).Value '���^����
        nippo_syukei_cell.Offset(0, 12).Value = nippo_nyuryoku_cell.Offset(0, 12).Value '�}�V���̏��~
        nippo_syukei_cell.Offset(0, 13).Value = nippo_nyuryoku_cell.Offset(0, 13).Value '�^���|
        nippo_syukei_cell.Offset(0, 14).Value = nippo_nyuryoku_cell.Offset(0, 14).Value '�I�ƍ��
        nippo_syukei_cell.Offset(0, 15).Value = nippo_nyuryoku_cell.Offset(0, 15).Value 'Rb����
        nippo_syukei_cell.Offset(0, 16).Value = nippo_nyuryoku_cell.Offset(0, 16).Value '���@�Ή��҂�
        nippo_syukei_cell.Offset(0, 17).Value = nippo_nyuryoku_cell.Offset(0, 17).Value '���^��
        nippo_syukei_cell.Offset(0, 18).Value = nippo_nyuryoku_cell.Offset(0, 18).Value '���q���ꏈ��
        nippo_syukei_cell.Offset(0, 19).Value = nippo_nyuryoku_cell.Offset(0, 19).Value '���̑�
        nippo_syukei_cell.Offset(0, 20).Value = nippo_nyuryoku_cell.Offset(0, 20).Value '�蒼���s��
        nippo_syukei_cell.Offset(0, 21).Value = nippo_nyuryoku_cell.Offset(0, 21).Value '���`�s�ǐ�
        nippo_syukei_cell.Offset(0, 22).Value = nippo_nyuryoku_cell.Offset(0, 22).Value '�q�r�E�J�P�E�X��
        nippo_syukei_cell.Offset(0, 23).Value = nippo_nyuryoku_cell.Offset(0, 23).Value '�A�J�s��
        nippo_syukei_cell.Offset(0, 24).Value = nippo_nyuryoku_cell.Offset(0, 24).Value '�������s��
        nippo_syukei_cell.Offset(0, 25).Value = nippo_nyuryoku_cell.Offset(0, 25).Value '�[�U�s��
        nippo_syukei_cell.Offset(0, 26).Value = nippo_nyuryoku_cell.Offset(0, 26).Value '�Đ��s��
        nippo_syukei_cell.Offset(0, 27).Value = nippo_nyuryoku_cell.Offset(0, 27).Value '�^�Y���s��
        nippo_syukei_cell.Offset(0, 28).Value = nippo_nyuryoku_cell.Offset(0, 28).Value '��ƒ�����
        nippo_syukei_cell.Offset(0, 29).Value = nippo_nyuryoku_cell.Offset(0, 29).Value '���̑�
        nippo_syukei_cell.Offset(0, 30).Value = nippo_nyuryoku_cell.Offset(0, -2).Value '�Ǖi��
        nippo_syukei_cell.Offset(0, 31).Value = nippo_nyuryoku_cell.Offset(0, 30).Value '������
        nippo_syukei_cell.Offset(0, 32).Value = nippo_nyuryoku_cell.Offset(0, 31).Value '�P�d
        nippo_syukei_cell.Offset(0, 33).Value = nippo_nyuryoku_cell.Offset(0, 32).Value '�P��
        nippo_syukei_cell.Offset(0, 34).Value = nippo_nyuryoku_cell.Offset(0, -3).Value * nippo_nyuryoku_cell.Offset(0, 4).Value * nippo_nyuryoku_cell.Offset(0, 31).Value '���ʁi�g�p�ʁj
        nippo_syukei_cell.Offset(0, 35).Value = nippo_nyuryoku_cell.Offset(0, -2).Value * nippo_nyuryoku_cell.Offset(0, 31).Value '�Ǖi���i�g�p�ʁj
        nippo_syukei_cell.Offset(0, 36).Value = nippo_syukei_cell.Offset(0, 34).Value - nippo_syukei_cell.Offset(0, 35).Value '�s�ǐ��i�g�p�ʁj
        nippo_syukei_cell.Offset(0, 37).Value = nippo_nyuryoku_cell.Offset(0, -2).Value * nippo_nyuryoku_cell.Offset(0, 32).Value '���Y���z
        nippo_syukei_cell.Offset(0, 38).Value = nippo_nyuryoku_cell.Offset(0, 21).Value * nippo_nyuryoku_cell.Offset(0, 32).Value '�s�ǋ��z
        nippo_syukei_cell.Offset(0, 39).Value = nippo_nyuryoku_cell.Offset(0, -4).Value '���q��
        Set nippo_syukei_cell = nippo_syukei_cell.Offset(1, 0)
        Set nippo_nyuryoku_cell = nippo_nyuryoku_cell.Offset(1, 0)
    Loop

End Sub
