Attribute VB_Name = "Module2"

Public Sub NippouShuukei_Update(NNCl As Object, NSCl As Object)
    
    Dim n As Long
    
    '�Z��������
    ThisWorkbook.Worksheets("����W�v").Range("A5:AN600").ClearContents
    Range("A5").Select
    
    '��ƕ\�쐬
    n = 1
    Do Until NNCl.Value = ""
        'Application.StatusBar = "������͂������W�v���쐬���E�E�E�@" & n & "���R�[�h��"
        NSCl.Offset(0, 0).Value = NNCl.Offset(0, 0).Value '���Y��
        NSCl.Offset(0, 1).Value = NNCl.Offset(0, 1).Value '�}�V��
        NSCl.Offset(0, 2).Value = NNCl.Offset(0, 2).Value '��Ǝ�
        NSCl.Offset(0, 3).Value = NNCl.Offset(0, 3).Value '���q
        NSCl.Offset(0, 4).Value = NNCl.Offset(0, 4).Value '�V���b�g
        NSCl.Offset(0, 5).Value = NNCl.Offset(0, 5).Value '�ғ�����
        NSCl.Offset(0, 6).Value = NNCl.Offset(0, 7).Value '���Y����
        NSCl.Offset(0, 7).Value = NNCl.Offset(0, 5).Value * NNCl.Offset(0, 6) 'OP��Ǝ���
        NSCl.Offset(0, 8).Value = NNCl.Offset(0, 8).Value '�n�ƍ��
        NSCl.Offset(0, 9).Value = NNCl.Offset(0, 9).Value '���^����
        NSCl.Offset(0, 10).Value = NNCl.Offset(0, 10).Value '�����҂�
        NSCl.Offset(0, 11).Value = NNCl.Offset(0, 11).Value '���^����
        NSCl.Offset(0, 12).Value = NNCl.Offset(0, 12).Value '�}�V���̏��~
        NSCl.Offset(0, 13).Value = NNCl.Offset(0, 13).Value '�^���|
        NSCl.Offset(0, 14).Value = NNCl.Offset(0, 14).Value '�I�ƍ��
        NSCl.Offset(0, 15).Value = NNCl.Offset(0, 15).Value 'Rb����
        NSCl.Offset(0, 16).Value = NNCl.Offset(0, 16).Value '���@�Ή��҂�
        NSCl.Offset(0, 17).Value = NNCl.Offset(0, 17).Value '���^��
        NSCl.Offset(0, 18).Value = NNCl.Offset(0, 18).Value '���q���ꏈ��
        NSCl.Offset(0, 19).Value = NNCl.Offset(0, 19).Value '���̑�
        NSCl.Offset(0, 20).Value = NNCl.Offset(0, 20).Value '�蒼���s��
        NSCl.Offset(0, 21).Value = NNCl.Offset(0, 21).Value '���`�s�ǐ�
        NSCl.Offset(0, 22).Value = NNCl.Offset(0, 22).Value '�q�r�E�J�P�E�X��
        NSCl.Offset(0, 23).Value = NNCl.Offset(0, 23).Value '�A�J�s��
        NSCl.Offset(0, 24).Value = NNCl.Offset(0, 24).Value '�������s��
        NSCl.Offset(0, 25).Value = NNCl.Offset(0, 25).Value '�[�U�s��
        NSCl.Offset(0, 26).Value = NNCl.Offset(0, 26).Value '�Đ��s��
        NSCl.Offset(0, 27).Value = NNCl.Offset(0, 27).Value '�^�Y���s��
        NSCl.Offset(0, 28).Value = NNCl.Offset(0, 28).Value '��ƒ�����
        NSCl.Offset(0, 29).Value = NNCl.Offset(0, 29).Value '���̑�
        NSCl.Offset(0, 30).Value = NNCl.Offset(0, -2).Value '�Ǖi��
        NSCl.Offset(0, 31).Value = NNCl.Offset(0, 30).Value '������
        NSCl.Offset(0, 32).Value = NNCl.Offset(0, 31).Value '�P�d
        NSCl.Offset(0, 33).Value = NNCl.Offset(0, 32).Value '�P��
        NSCl.Offset(0, 34).Value = NNCl.Offset(0, -3).Value * NNCl.Offset(0, 4).Value * NNCl.Offset(0, 31).Value '���ʁi�g�p�ʁj
        NSCl.Offset(0, 35).Value = NNCl.Offset(0, -2).Value * NNCl.Offset(0, 31).Value '�Ǖi���i�g�p�ʁj
        NSCl.Offset(0, 36).Value = NSCl.Offset(0, 34).Value - NSCl.Offset(0, 35).Value '�s�ǐ��i�g�p�ʁj
        NSCl.Offset(0, 37).Value = NNCl.Offset(0, -2).Value * NNCl.Offset(0, 32).Value '���Y���z
        NSCl.Offset(0, 38).Value = NNCl.Offset(0, 21).Value * NNCl.Offset(0, 32).Value '�s�ǋ��z
        NSCl.Offset(0, 39).Value = NNCl.Offset(0, -4).Value '���q��
        Set NSCl = NSCl.Offset(1, 0)
        Set NNCl = NNCl.Offset(1, 0)
        n = n + 1
    Loop
    
    NNC = 0
    NSU = 1
    
    Application.StatusBar = False
    
End Sub