Attribute VB_Name = "AI�ޯ�"
Public atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1 To 2) As Integer  '�ޯ�-AI-��̬d�w-���t���C�p��C�ƭȬ����Ȯɼ�(1.�ثe�p��ƭ�/2.(�o��))
Public atking_AI_������_�����Ҧ����A��(1 To 2) As Integer 'AI-�����ڦ����Ҧ����A�ˬd��(1.���A���涥�q/2.���A�Ұ��ˬd��)
Public atking_AI_����_Jackpot������(1 To 2) As Integer '�ޯ�-AI-����-Jackpot��P������(1.�`�@��/2.�ثe��)
Public atking_AI_���[_�O�d���Ų�_tot(1 To 2) As Integer  '�ޯ�-���[-�O�d���Ų���l�q�����Ȯ��ܼ�(1.�ƭ�/2.�O�_�Ұ�)
Public atking_AI_���_�Q�T����_tot(1 To 2) As Integer '�ޯ�-AI-���-�Q�T������l�q�����Ȯ��ܼ�(1.�ƭ�/2.�O�_�Ұ�)
Public atking_AI_�Ǧh_�]�G���ۻ�q������(1 To 3) As Integer '�ޯ�-AI-�Ǧh-�]�G�����Y��q������(1.��1��(����)���G/2.��2��(�ޯ�)���G/3.���R�ᵲ�G)
Public atking_AI_�Ǧh_�]�G����O����(1 To 108) As Integer '�ޯ�-AI-�Ǧh-�]�G����������X�P�s����(1~106.�O���P�s��/107.�`�@�^�i��/108.�ثe��)
Public atking_AI_���_�i���h_���@�g��_�j�ƭȬ����� As Integer '�ޯ�-AI-���_�i���h-���@�g���ثe�֭p�[�j�ƭȬ�����
Public atking_AI_��_�u�@�Ҧ����A�Ұʭ� As Boolean '�ޯ�-��-AI-Ex-�󫵦�-�[�ʯP���u�@�K�����˼Ҧ��Ұʭ�
Public atking_AI_��B�����S_�p��������(1 To 2) As Integer '�ޯ�-AI-��B�����S-�p�������P������(1.�`�@��/2.�ثe��)
Public atking_AI_��B�����S_���������� As Integer '�ޯ�-AI_��B�����S-������P�ثe��
Public atking_AI_�Q��_�������T�Ϭ�����(1 To 2) As Integer '�ޯ�-AI-�Q��-�������T�ϩ�P�ثe��(1.�`�@��/2.�ثe��)
Public atking_AI_������S_���������(0 To 107) As Integer '�ޯ�-AI-������S-����������P�s���Ȯɼ�(0.�ثe���i�ƭ�/1~106�P�s����ܭ�/107.�`�@����i�ƭ�)
Public atking_AI_�w�ǥ���_�ƨg���۬����� As Integer '�ޯ�-AI-�w�ǥ���-�ƨg���ۥ����P�����ثe��
Public atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ As Integer '�ޯ�-AI-�����g-�g�����b�P�ݦ大�j�m�P�ثe��
Public atking_AI_�v��L_�����Ҧ����A��(1 To 5) As Integer 'AI-�v��L�����Ҧ����A�ˬd��(1.���A���涥�q/2.���A�Ұ��ˬd��/3.�����ƭ�(��l)/4.�����ƭ�(�ܧ��)/5.�ƭȬ����O�_�Ұ�)
Public atking_AI_�L���S_�j�t���q������(1 To 3) As Integer '�ޯ�-AI-�L���S-�j�t���Y��q������(1.��1��(����)���G/2.��2��(�ޯ�)���G/3.���R�ᵲ�G)
Public atking_AI_�����i_���y�����p��X�P�i�Ƭ����� As Integer  '�ޯ�-AI-�����i-���y�����p��X�P�i�ƭȬ����Ȯɼ�
Public atking_AI_�����i_�t���¥�������(1 To 2) As Integer  '�ޯ�-AI-�����i-�t���¥������Ȯɼ�(1.����^�X���m�O/2.����^�X�X�P��/3.�ϥΪ̷�^�X�����O)
Public atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� As Integer  '�ޯ�-AI-�S�{��-���M�C�{�p��C�d�i�ƭȬ����Ȯɼ�
Public atking_AI_����_���Ϥ۹�_��P������(1 To 2) As Integer '�ޯ�-AI-����-���Ϥ۹ک�P�ثe��(1.�`�@��/2.�ثe��)
Public atking_AI_�j�|�˺��h_�믫�O�l��������(0 To 106) As Integer '�ޯ�-AI-�j�|�˺��h-�믫�O�l���������P�s���Ȯɼ�(0.�`�@�i�ƭ�/1~106�P�s����ܭ�)
Public atking_AI_��ܵY_��k���Ӫ������(0 To 2) As Integer '�ޯ�-AI-��ܵY-��k���Ӫ�������P�s���Ȯɼ�(0.�`�@�i�ƭ�/1~2�P�s��)
Public atking_AI_��ܵY_�����ۺh���q������(0 To 106, 1 To 4) As Integer '�ޯ�-AI-��ܵY-�����ۺh�����ĪG�ζ��q�Ȯɼ�(0.(1).��e�ĪG/(2).��e�ĪG���q/(3)�`�@��P�ƶq/(4)�ثe��/��P�ƶq,1~106.(1)�P����w������)
Public atking_AI_����_�o�����c������ As Integer '�ޯ�-AI-����-�o�����c��P�ثe��
Public atking_AI_���Y�F_����_��P������(1 To 2) As Integer '�ޯ�-AI-���Y�F-������P�ثe��(1.�`�@��/2.�ثe��)
Public atking_AI_���Y�F_��������������A��(1 To 106) As Boolean '�ޯ�-AI-���Y�F-��������������X�P�s����
Public atking_AI_���Y�F_����B_�����O�[��������(1 To 2) As Boolean '�ޯ�-AI-���Y�F-����B�����O�[���Ȯɬ�����(1.�O�_10�i�w+10/2.�O�_15�i�w+15)
Public atking_AI_��_�צ�_�L�ɽ��j���׵������� As Integer  '�ޯ�-AI-��-Ex-�צ�-�L�ɽ��j���׵�������⤧���m�P�ȼȮɼ�
Public atking_AI_ù��Y_�����ۼv�������A��(1 To 106) As Boolean '�ޯ�-AI-ù��Y-�����ۼv(���BEX)�������X�P�s����
Public atking_AI_�����g_�f��ԧ����j�T_��P������(1 To 2) As Integer '�ޯ�-AI-�����g-�f��ԧ����j�T��P�ثe��(1.�`�@��/2.�ثe��)
Public atking_AI_�J�y_�Ѩ����_�ܵP������(1 To 2) As Integer  '�ޯ�-AI-�J�y-�Ѩ���ƹܨ����X�P�P��������(1.�ܵP�s��/2.�ܵP���X�P����)
Public atking_AI_�J�y_�����g����q������(1 To 3) As Integer '�ޯ�-AI-�J�y-�����g���Y��q������(1.��1��(����)���G/2.��2��(�ޯ�)���G/3.���R���`���G)
Public atking_AI_�J�y_�c�N����������(0 To 106) As Integer '�ޯ�-AI-�J�y-�c�N�����������P�s���Ȯɼ�(0.�ثe���q/1~106�P�s����ܭ�)
Public atking_AI_�ײ��d_�W���ثe���q������(1 To 4)  As Integer  '�ޯ�-AI-�ײ��d-�W������ثe���q�ƭȬ����Ȯɼ�(1.�����ƭ�(��l)/2.�����ƭ�(�ܧ��)/3.�ثe���涥�q(�`)/4.�W��3�ɧ𨾻�q�[���O�_�Ұ�)

Sub �j�|�˺��h_�r��()
Dim rrr As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(1).Caption = "�r��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(3, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�j�|�˺��h" Then
   Select Case atkingckai(3, 1)
      Case 1
          If movecp = 1 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          End If
          If rrr >= 2 And atkingckai(3, 2) = 0 Then
             �������m��l�`��(2) = �������m��l�`��(2) + 6
             atkingckai(3, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 2 And atkingckai(3, 2) = 1 Then
             �������m��l�`��(2) = �������m��l�`��(2) - 6
             atkingckai(3, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
      Case 2
             atkingckai(3, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�j�|�˺��h\�j�|�˺��h_�r��_2.jpeg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 10305
                   atkingno(i, 6) = 8925
                   atkingno(i, 7) = 8
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub �j�|�˺��h_�大����()
Dim bloodtot As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "�大����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(62, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�j�|�˺��h" Then
   Select Case atkingckai(62, 1)
        Case 1
             If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 2 And atkingckai(62, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(62, 2) = 0 Then
               atkingckai(62, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 6
            ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 2) And atkingckai(62, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(62, 2) = 1 Then
               atkingckai(62, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 6
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�j�|�˺��h\Grunwaldatking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6915
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 62
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(FormMainMode.��ܦC1.goi1) <= 0 Then
                atkingckai(62, 2) = 0
            End If
        Case 4
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) <= 0 Then
                bloodtot = Abs(Val(�Y���淾�q�Ȯ��ܼ�(2)))
                �԰��t����.�^�_����_�q�� bloodtot, 1
            End If
            '=============
            atkingckai(62, 2) = 0
   End Select
End If
End Sub
Sub �Ϩ��~2012_�P�R���()
If FormMainMode.comaiatk(2).Caption = "�P�R���" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(15, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Ϩ��~2012" Then
   Select Case atkingckai(15, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 1 And atkingckai(15, 2) = 0 Then
               atkingckai(15, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(1)
          End If
      Case 2
             atkingckai(15, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\�Ϩ��~2012\�Ϩ��~2012_�P�R���_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 47
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub �Ϩ��~2012_�P�R�ļ�()
If FormMainMode.comaiatk(1).Caption = "�P�R�ļ�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(14, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Ϩ��~2012" Then
   Select Case atkingckai(14, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 1 And atkingckai(14, 2) = 0 Then
               atkingckai(14, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(1) + 10
          End If
      Case 2
             atkingckai(14, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\�Ϩ��~2012\�Ϩ��~2012_�P�R�ļ�_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 46
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             �������m��l�`��(2) = �������m��l�`��(1) + 10
             �԰��t����.�����g�J��ܦC�ƭ� 2, �������m��l�`��(2)
   End Select
End If
End Sub
Sub �l��V���̶�_�l��()
If FormMainMode.comaiatk(1).Caption = "�l��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(16, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�l��V���̶�" Then
   Select Case atkingckai(16, 1)
      Case 1
         If movecp = 1 Then
            If atkingpagetot(2, 1) >= 6 And atkingckai(16, 2) = 0 Then
               atkingckai(16, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 7
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(16, 2) = 1 Then
               atkingckai(16, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 7
            End If
        End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\�l��V���̶�\VampireLAMIAatking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 48
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(16, 1) = 3
       Case 3
            atkingckai(16, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                �^�_����_�q�� 1, 1
            End If
   End Select
End If
End Sub
Sub �l��V���̶�_���Q�����\()
If FormMainMode.comaiatk(2).Caption = "���Q�����\" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(17, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�l��V���̶�" Then
   Select Case atkingckai(17, 1)
      Case 1
         If movecp > 1 Then
            If atkingpagetot(2, 5) >= 4 And atkingckai(17, 2) = 0 Then
               atkingckai(17, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 5) < 4 And atkingckai(17, 2) = 1 Then
               atkingckai(17, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\�l��V���̶�\VampireLAMIAatking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 17
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(17, 1) = 3
       Case 3
            atkingckai(17, 2) = 0
            �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 1, 1
   End Select
End If
End Sub
Sub �l��V���̶�_����()
If FormMainMode.comaiatk(3).Caption = "����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(18, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�l��V���̶�" Then
   Select Case atkingckai(18, 1)
      Case 1
            If atkingpagetot(2, 2) >= 3 And atkingckai(18, 2) = 0 Then
               atkingckai(18, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 2) < 3 And atkingckai(18, 2) = 1 Then
               atkingckai(18, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\�l��V���̶�\VampireLAMIAatking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 50
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(18, 1) = 3
       Case 3
            atkingckai(18, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                �Y���淾�q�Ȯ��ܼ�(2) = Val(�Y���淾�q�Ȯ��ܼ�(2)) - 1
                �Y����ˮ`�� = �Y���淾�q�Ȯ��ܼ�(2)
            End If
   End Select
End If
End Sub
Sub ������m_�B�����l()
If FormMainMode.comaiatk(1).Caption = "�B�����l" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(8, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������m" Then
   Select Case atkingckai(8, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 2 And atkingckai(8, 2) = 0 Then
               atkingckai(8, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\������m\������m_�B�����l_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -840
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 8
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(8, 1) = 3
       Case 3
            atkingckai(8, 2) = 0
                Do
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                     If �H�����`���A��Ʈw(1, i, 3) = 10 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                         FormMainMode.personusspe(i).person_turn = 3
                         �H�����`���A��Ʈw(1, i, 2) = 3
                         Exit Do
                     End If
                   Next
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                         �԰��t����.�H�����`���A��]�w_��] 1, i, 10, app_path & "gif\���`���A\atkdown.gif", 5, 3
                         ���`���A�ˬd��(10, 1) = 1
                         ���`���A�ˬd��(10, 2) = 1
                         Exit Do
                      End If
                   Next
                Loop
   End Select
End If
End Sub
Sub ������m_�Һ����l()
If FormMainMode.comaiatk(2).Caption = "�Һ����l" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(9, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������m" Then
   Select Case atkingckai(9, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 2 And atkingckai(9, 2) = 0 Then
               atkingckai(9, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\������m\������m_�Һ����l_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -840
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 9
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(9, 1) = 3
       Case 3
            atkingckai(9, 2) = 0
                Do
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                     If �H�����`���A��Ʈw(1, i, 3) = 11 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                         FormMainMode.personusspe(i).person_turn = 3
                         �H�����`���A��Ʈw(1, i, 2) = 3
                         Exit Do
                     End If
                   Next
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                         �԰��t����.�H�����`���A��]�w_��] 1, i, 11, app_path & "gif\���`���A\defdown.gif", 5, 3
                         ���`���A�ˬd��(11, 1) = 1
                         ���`���A�ˬd��(11, 2) = 1
                         Exit Do
                      End If
                   Next
                Loop
   End Select
End If
End Sub
Sub ������m_�V�P���l()
If FormMainMode.comaiatk(3).Caption = "�V�P���l" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(10, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������m" Then
   Select Case atkingckai(10, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 2 And atkingckai(10, 2) = 0 Then
               atkingckai(10, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\������m\������m_�V�P���l_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6180
                   atkingno(i, 6) = 9630
                   atkingno(i, 7) = 10
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(10, 1) = 3
       Case 3
            atkingckai(10, 2) = 0
                Do
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                     If �H�����`���A��Ʈw(1, i, 3) = 12 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                         FormMainMode.personusspe(i).person_turn = 3
                         �H�����`���A��Ʈw(1, i, 2) = 3
                         Exit Do
                     End If
                   Next
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                         �԰��t����.�H�����`���A��]�w_��] 1, i, 12, app_path & "gif\���`���A\movdown.gif", 1, 3
                         ���`���A�ˬd��(12, 1) = 1
                         ���`���A�ˬd��(12, 2) = 1
                         Exit Do
                      End If
                   Next
                Loop
   End Select
End If
End Sub
Sub �n�ʤ�_����()
If FormMainMode.comaiatk(1).Caption = "����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(7, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�n�ʤ�" Then
   Select Case atkingckai(7, 1)
      Case 1
          If Val(FormMainMode.pagecomqlead) >= 3 And atkingckai(7, 2) = 0 Then
               atkingckai(7, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\�n�ʤ�\�n�ʤ�_����_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1080
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 28
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(7, 1) = 3
       Case 3
            atkingckai(7, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                Do
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                     If �H�����`���A��Ʈw(1, i, 3) = 16 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                         FormMainMode.personusspe(i).person_turn = 3
                         �H�����`���A��Ʈw(1, i, 2) = 3
                         Exit Do
                     End If
                   Next
                   For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                         �԰��t����.�H�����`���A��]�w_��] 1, i, 16, app_path & "gif\���`���A\moveerr.gif", 0, 3
                         ���`���A�ˬd��(16, 1) = 1
                         ���`���A�ˬd��(16, 2) = 1
                         Exit Do
                      End If
                   Next
                Loop
            End If
   End Select
End If
End Sub
Sub �n�ʤ�_�W�A��()
If FormMainMode.comaiatk(2).Caption = "�W�A��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(6, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�n�ʤ�" Then
   Select Case atkingckai(6, 1)
      Case 1
          If atkingpagetot(2, 3) >= 1 And atkingckai(6, 2) = 0 Then
               atkingckai(6, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
          ElseIf atkingpagetot(2, 3) < 1 And atkingckai(6, 2) = 1 Then
               atkingckai(6, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǫ��d\�n�ʤ�\�n�ʤ�_�W�A��_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = -2040
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 6
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
            �^�_����_�q�� 3, 1
            atkingckai(6, 2) = 0
   End Select
End If
End Sub
Sub ����_�۱��ɦV(ByVal Index As Integer)
Dim atkingtotai As Integer '�S�ƶq�Ȯɲέp�ܼ�
Dim a As Integer, i As Integer, j As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(1).Caption = "�۱��ɦV" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(1, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
 Select Case atkingckai(1, 1)
   Case 1
        For i = 1 To 106
            If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) <> 1 Then
                 If pagecardnum(i, 1) = a4a And pagecardnum(i, 3) = a4a Then
                     atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                 ElseIf pagecardnum(i, 1) = a4a Then
                     atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                 ElseIf pagecardnum(i, 3) = a4a Then
                     atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
                 End If
            End If
        Next
      
      If atkingtotai < livecom(����H����ԤH��(2, 2)) And atkingtotai > 1 Then
         For a = 1 To 106
             �԰��t����.comatk_AI_����_�۱��ɦV_�S a
         Next
      ElseIf atkingtotai >= livecom(����H����ԤH��(2, 2)) Then
         atkingtotai = 0
            If livecom(����H����ԤH��(2, 2)) >= (livecom(����H����ԤH��(2, 2)) \ 4) * 3 Then '�p�G��q�j��4����3���ܡA�S3�H�W�d�u��
                For i = 106 To 55 Step -1
                    If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) <> 1 Then
                       If Val(pagecardnum(i, 2)) >= 3 And pagecardnum(i, 1) = a4a Then
                            atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                            If atkingtotai >= livecom(����H����ԤH��(2, 2)) Then Exit For
                            �԰��t����.comatk_AI_����_�۱��ɦV_�S i
                       ElseIf Val(pagecardnum(i, 4)) >= 3 And pagecardnum(i, 3) = a4a Then
                            atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
                            If atkingtotai >= livecom(����H����ԤH��(2, 2)) Then Exit For
                            �԰��t����.comatk_AI_����_�۱��ɦV_�S i
                       End If
                    End If
                    
'                    If atkingtotai >= livecom(����H����ԤH��(2, 2)) Then Exit For
'
'                    �԰��t����.comatk_AI_����_�۱��ɦV_�S a
                Next
            End If
            If atkingtotai < livecom(����H����ԤH��(2, 2)) Then
               a = 1
               Do While a <= 106
                  If livecom(����H����ԤH��(2, 2)) <= 4 Then
                      If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 And (Val(pagecardnum(i, 2)) <> 3 And pagecardnum(i, 1) = a4a) Then
                            If pagecardnum(a, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                            If pagecardnum(a, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                            If atkingtotai >= livecom(����H����ԤH��(2, 2)) Then Exit Do
                            '===========================
                            �԰��t����.comatk_AI_����_�۱��ɦV_�S a
                      End If
                  Else
                      If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 Then
                            If pagecardnum(a, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                            If pagecardnum(a, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                            If atkingtotai >= livecom(����H����ԤH��(2, 2)) Then Exit Do
                            '===========================
                            �԰��t����.comatk_AI_����_�۱��ɦV_�S a
                      End If
                  End If
                    
'                    If atkingtotai >= livecom(����H����ԤH��(2, 2)) Then Exit Do
'
'                    �԰��t����.comatk_AI_����_�۱��ɦV_�S a
'                  End If
                  a = a + 1
               Loop
            End If
      End If
   Case 2
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(1, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(1, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2)) * 5
               If atkingckai(1, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(1, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
    Case 3
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(1, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(1, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 3) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4)) * 5
               If atkingckai(1, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(1, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
   Case 4
        For i = �H���ޯ�Ʀr���� To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\����\����_�۱��ɦV_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 0
                atkingno(i, 6) = 0
                atkingno(i, 7) = 1
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
       '-------------
    Case 5
        �԰��t����.�ˮ`����_�ޯઽ��_�q�� Val(atkingpagetot(2, 4)), 1
        atkingckai(1, 2) = 0
   End Select
End If
End Sub
Sub ����_�����()
Dim atkingtotai As Integer '�S�ƶq�Ȯɲέp�ܼ�
Dim a As Integer, i As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(2).Caption = "�����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(12, 2) = 1) _
    And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
 Select Case atkingckai(12, 1)
   Case 1
      atkingckai(12, 1) = 2
      For i = 55 To 106
         If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And ((Val(pagecardnum(i, 2)) = 3 And pagecardnum(i, 1) = a4a) Or (Val(pagecardnum(i, 4)) = 3 And pagecardnum(i, 3) = a4a)) Then
            atkingtotai = Val(atkingtotai) + 1
         End If
      Next
      If atkingtotai >= 1 Then
         Select Case livecom(����H����ԤH��(2, 2))
            Case Is < 3
                If Val(FormMainMode.��ܦC1.goi1) - Val(FormMainMode.��ܦC1.goi2) >= livecom(����H����ԤH��(2, 2)) Then
                    GoTo AI�ޯ�_����_�����_�X�P���q�G
                End If
            Case 3
                If Val(FormMainMode.��ܦC1.goi1) - Val(FormMainMode.��ܦC1.goi2) >= 9 Then
                    GoTo AI�ޯ�_����_�����_�X�P���q�G
                End If
            Case Is > 3
                If Int(Val(FormMainMode.��ܦC1.goi1) / 3 + 0.9) - Int(Val(FormMainMode.��ܦC1.goi2) / 3 + 0.9) >= livecom(����H����ԤH��(2, 2)) Then
                    GoTo AI�ޯ�_����_�����_�X�P���q�G
                End If
         End Select
      End If
      '==========�p�G���ŦX��������
      Exit Sub
    '================================
AI�ޯ�_����_�����_�X�P���q�G:
      For a = 55 To 106
            If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) = 3 Then
                    �԰��t����.comatk_AI_����_�h�g�H_�����_�S a
                    Exit For
                ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) = 3 Then
                    �԰��t����.comatk_AI_����_�h�g�H_�����_�S a
                    Exit For
                End If
             End If
      Next
    Case 2
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
'                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) >= 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
             If rrr >= 1 And atkingckai(12, 2) = 0 Then
                atkingckai(12, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
             If rrr < 1 And atkingckai(12, 2) = 1 Then
                atkingckai(12, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
   Case 3
        For i = �H���ޯ�Ʀr���� To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\����\����_�����_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 10110
                atkingno(i, 7) = 12
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
          atkingckai(12, 2) = 0
          If Val(�Y���淾�q�Ȯ��ܼ�(2)) >= livecom(����H����ԤH��(2, 2)) And ���`���A�ˬd��(18, 2) = 0 Then
                Do
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 6
                              FormMainMode.personcomspe(j).person_turn = 3
                              �H�����`���A��Ʈw(2, j, 1) = 6
                              �H�����`���A��Ʈw(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 6, 3
                          ���`���A�ˬd��(1, 1) = 1
                          ���`���A�ˬd��(1, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '==================================
                Do
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, j, 3) = 18 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              �H�����`���A��Ʈw(2, j, 1) = 0
                              �H�����`���A��Ʈw(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 18, app_path & "gif\���`���A\����.gif", 0, 3
                          ���`���A�ˬd��(18, 1) = 1
                          ���`���A�ˬd��(18, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '===============================
                Do
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, j, 3) = 19 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              �H�����`���A��Ʈw(2, j, 1) = 0
                              �H�����`���A��Ʈw(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 19, app_path & "gif\���`���A\���a.gif", 0, 3
                          ���`���A�ˬd��(19, 1) = 1
                          ���`���A�ˬd��(19, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
         End If
   End Select
End If
End Sub
Sub ����_���j�¤�()
Dim a As Integer, i As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "���j�¤�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(2, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
  Select Case atkingckai(2, 1)
   Case 1
       For i = 1 To 106
          If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) >= 3 And Val(pagecardnum(i, 5)) = 2 And movecp < 3 Then
             �԰��t����.comatk_AI_����_���j�¤�_�C i
             Exit For
          ElseIf pagecardnum(i, 3) = a1a And Val(pagecardnum(i, 4)) >= 3 And Val(pagecardnum(i, 5)) = 2 And movecp < 3 Then
             �԰��t����.comatk_AI_����_���j�¤�_�C i
             Exit For
          End If
       Next
       atkingckai(2, 1) = 2
    Case 2
          If movecp < 3 Then
            If atkingpagetot(2, 1) >= 3 And atkingckai(2, 2) = 0 Then
               atkingckai(2, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 1) < 3 And atkingckai(2, 2) = 1 Then
               atkingckai(2, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
   Case 3
       For i = �H���ޯ�Ʀr���� To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\����\����_���j�¤�_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 9810
                atkingno(i, 6) = 8940
                atkingno(i, 7) = 2
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
        Do
           atkingckai(2, 2) = 0
           For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
             If �H�����`���A��Ʈw(1, i, 3) = 11 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                 FormMainMode.personusspe(i).person_num = 4
                 FormMainMode.personusspe(i).person_turn = 3
                 �H�����`���A��Ʈw(1, i, 1) = 4
                 �H�����`���A��Ʈw(1, i, 2) = 3
                 Exit Do
             End If
           Next
           For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
              If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                 �԰��t����.�H�����`���A��]�w_��] 1, i, 11, app_path & "gif\���`���A\defdown.gif", 4, 3
                 ���`���A�ˬd��(11, 1) = 1
                 ���`���A�ˬd��(11, 2) = 1
                 Exit Do
             End If
           Next
        Loop
  End Select
End If
End Sub
Sub ����_���b�B()
Dim atkingtotai As Integer '�S�ƶq�Ȯɲέp�ܼ�
Dim ak As Integer, j As Integer, ui As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "���b�B" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(5, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
 Select Case atkingckai(5, 1)
   Case 1
      If movecp = 3 Then
          For j = 49 To 54   '��1��1�d�u��
              If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                   If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                      �԰��t����.comatk_AI_����_���b�B_�� j
                      ak = 1
                      Exit For
                   ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                      �԰��t����.comatk_AI_����_���b�B_�� j
                      ak = 1
                      Exit For
                   End If
              End If
          Next
          If ak = 0 Then
             For j = 39 To 44   '�j1��1�d�䦸�u��
                If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                   If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                      �԰��t����.comatk_AI_����_���b�B_�� j
                      ak = 1
                      Exit For
                   ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                      �԰��t����.comatk_AI_����_���b�B_�� j
                      ak = 1
                      Exit For
                   End If
                End If
             Next
          End If
          If ak = 0 Then
             For j = 1 To 106
                If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                   If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                      �԰��t����.comatk_AI_����_���b�B_�� j
                      ak = 1
                      Exit For
                   ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                      �԰��t����.comatk_AI_����_���b�B_�� j
                      ak = 1
                      Exit For
                   End If
                End If
             Next
          End If
          If ak = 1 Then
             atkingckai(5, 2) = 1
    '         atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
       End If
   Case 2
      atkingckai(5, 1) = 3
      atkingckai(5, 2) = 0 '���AI�X�P�����P�_��h
      If moveturn = 2 Then
'        If livecom(����H����ԤH��(2, 2)) <= 5 Then  '�����p�J���S3��2�d
'            ui = 54
'        Else
'            ui = 57
'        End If
        For j = 1 To 106
          If (livecom(����H����ԤH��(2, 2)) <= 5 And Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 And ((Val(pagecardnum(j, 2)) <> 3 And pagecardnum(j, 1) = a4a) Or (Val(pagecardnum(j, 4)) <> 3 And pagecardnum(j, 3) = a4a))) Or _
              livecom(����H����ԤH��(2, 2)) > 5 And Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
            pagecardnum(j, 11) = 1
            If pagecardnum(j, 1) = a4a And pagecardnum(j, 3) = a4a Then
                pagecardnum(j, 11) = 0
            ElseIf pagecardnum(j, 1) = a4a Then
                 cspce = pagecardnum(j, 1)
                 cspme = pagecardnum(j, 2)
                 pagecardnum(j, 1) = pagecardnum(j, 3)
                 pagecardnum(j, 2) = pagecardnum(j, 4)
                 pagecardnum(j, 3) = cspce
                 pagecardnum(j, 4) = cspme
                 If pageonin(j) = 2 Then
                    pageonin(j) = 1
                 Else
                   pageonin(j) = 2
                 End If
            End If
         End If
       Next
    End If
   Case 3
          If atkingckai(5, 2) = 0 And movecp = 3 Then
             For i = 1 To 106
                  If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                     atkingckai(5, 2) = 1
                     atkingckai(5, 1) = 4
                     atkingtrn(2) = Val(atkingtrn(2)) + 1
                     �������m��l�`��(2) = �������m��l�`��(2) + Val(FormMainMode.pagecomqlead) * 2
                     atking_sheri_4_tot_ai = Val(FormMainMode.pagecomqlead)
                     Exit For
                  End If
             Next
          End If
    Case 4
            If atkingpagetot(2, 3) = 0 Then
               atkingckai(5, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               atkingckai(5, 1) = 3
               If Val(FormMainMode.pagecomqlead) = atking_sheri_4_tot_ai Then
                  �������m��l�`��(2) = �������m��l�`��(2) - Val(FormMainMode.pagecomqlead) * 2
               Else
                  �������m��l�`��(2) = �������m��l�`��(2) - Val(FormMainMode.pagecomqlead) * 2 - 2
               End If
               atking_sheri_4_tot_ai = 0
            ElseIf atkingpagetot(2, 3) > 1 Then
               For i = 1 To 106
                 If pagecardnum(i, 1) = a3a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                    ttt = ttt + 1
                 End If
               Next
               If ttt = 0 Then
                 atkingckai(5, 2) = 0
                 atkingtrn(2) = Val(atkingtrn(2)) - 1
                 atkingckai(5, 1) = 3
                 If Val(FormMainMode.pagecomqlead) = atking_sheri_4_tot_ai Then
                    �������m��l�`��(2) = �������m��l�`��(2) - Val(FormMainMode.pagecomqlead) * 2
                 Else
                    �������m��l�`��(2) = �������m��l�`��(2) - Val(FormMainMode.pagecomqlead) * 2 - 2
                 End If
                 atking_sheri_4_tot_ai = 0
               End If
            End If
            If atkingckai(5, 2) = 1 Then
               �������m��l�`��(2) = �������m��l�`��(2) + (Val(FormMainMode.pagecomqlead) - Val(atking_sheri_4_tot_ai)) * 2
               atking_sheri_4_tot_ai = Val(FormMainMode.pagecomqlead)
            End If
   Case 5
        For i = �H���ޯ�Ʀr���� To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\����\����_���b�B_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 9690
                atkingno(i, 7) = 22
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
        atkingckai(5, 2) = 0
   End Select
End If
End Sub
Sub ��_���ۦ�_�[���⪺�L��()
If FormMainMode.comaiatk(1).Caption = "���ۦ�-�[���⪺�L��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(4, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
   Select Case atkingckai(4, 1)
      Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 4 And atkingckai(4, 2) = 0 Then
               atkingckai(4, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 6
'               �������m��l�`��(1) = �������m��l�`��(1) - 3
            ElseIf atkingpagetot(2, 1) < 4 And atkingckai(4, 2) = 1 Then
               atkingckai(4, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 6
'               �������m��l�`��(1) = �������m��l�`��(1) + 3
            End If
          End If
      Case 2
             atkingckai(4, 2) = 0
             �԰��t����.�۰ʱ��b����
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��\��-���ۦ�-�[���⪺�L��_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7485
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 20
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '================
             �԰��t����.�����g�J��ܦC�ƭ� 1, Val(FormMainMode.��ܦC1.goi1) - 3
   End Select
End If
End Sub
Sub ��_EX_���ۦ�_�[���⪺�L��()
If FormMainMode.comaiatk(1).Caption = "Ex���ۦ�-�[���⪺�L��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(13, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
   Select Case atkingckai(13, 1)
      Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 5 And atkingckai(13, 2) = 0 Then
               atkingckai(13, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 8
'               �������m��l�`��(1) = �������m��l�`��(1) - 6
             ElseIf atkingpagetot(2, 1) < 5 And atkingckai(13, 2) = 1 Then
               atkingckai(13, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 8
'               �������m��l�`��(1) = �������m��l�`��(1) + 6
             End If
          End If
      Case 2
             atkingckai(13, 2) = 0
             �԰��t����.�۰ʱ��b����
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��\��-EX-���ۦ�-�[���⪺�L��2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7320
                   atkingno(i, 6) = 9000
                   atkingno(i, 7) = 41
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '================
             �԰��t����.�����g�J��ܦC�ƭ� 1, Val(FormMainMode.��ܦC1.goi1) - 6
   End Select
End If
End Sub
Sub ��_EX_�󫵦�_�[�ʯP���u�@()
If FormMainMode.comaiatk(2).Caption = "Ex�󫵦�-�[�ʯP���u�@" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(58, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
   Select Case atkingckai(58, 1)
        Case 1
            If atkingpagetot(2, 4) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(58, 2) = 0 Then
'            If atkingpagetot(2, 3) >= 1 And atkingckai(58, 2) = 0 Then
               atkingckai(58, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
'               �������m��l�`��(1) = �������m��l�`��(1) + 5
            ElseIf (atkingpagetot(2, 4) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(58, 2) = 1 Then
'            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(58, 2) = 1 Then
               atkingckai(58, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
'               �������m��l�`��(1) = �������m��l�`��(1) - 5
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��\��-EX-�󫵦�-�[�ʯP���u�@_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6540
                   atkingno(i, 6) = 9420
                   atkingno(i, 7) = 38
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
          '==========================
          atking_AI_��_�u�@�Ҧ����A�Ұʭ� = True
    Case 3
          atking_AI_��_�u�@�Ҧ����A�Ұʭ� = False
          atkingckai(58, 2) = 0
    Case 4
          �Y���淾�q�Ȯ��ܼ�(2) = Val(�Y���淾�q�Ȯ��ܼ�(2)) - 5
          �Y����ˮ`�� = �Y����ˮ`�� - 5
   End Select
End If
End Sub
Sub ��_EX_�w�_���������q()
Dim rrr As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "Ex�w�-���������q" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(63, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
   Select Case atkingckai(63, 1)
        Case 1
            For i = 1 To 106
               If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                  rrr = rrr + 1
               End If
            Next
          If rrr >= 1 And atkingckai(63, 2) = 0 Then
             atkingckai(63, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 1 And atkingckai(63, 2) = 1 Then
             atkingckai(63, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��\��-EX-�w�-���������q_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9675
                   atkingno(i, 6) = 10155
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(63, 2) = 0
             If livecom(����H����ԤH��(2, 2)) <= 0 Then
                 For i = 2 To 3
                     If livecom(����ݾ��H��������(2, i)) > 0 Then
                        Do
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                                  If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 9
                                      FormMainMode.personcomspe(j).person_turn = 2
                                      �H�����`���A��Ʈw(2, j, 1) = 9
                                      �H�����`���A��Ʈw(2, j, 2) = 2
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                               If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 9, 2
                                  ���`���A�ˬd��(1, 1) = 1
                                  ���`���A�ˬd��(1, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        Do
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                                  If �H�����`���A��Ʈw(2, j, 3) = 2 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 9
                                      FormMainMode.personcomspe(j).person_turn = 2
                                      �H�����`���A��Ʈw(2, j, 1) = 9
                                      �H�����`���A��Ʈw(2, j, 2) = 2
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                               If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 2, j, 2, app_path & "gif\���`���A\defup.gif", 9, 2
                                  ���`���A�ˬd��(2, 1) = 1
                                  ���`���A�ˬd��(2, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                     End If
                Next
            End If
   End Select
End If
End Sub
Sub ��_�צ�_�L�ɽ��j���׵�()
Dim atkingtotai As Integer '�S�ƶq�Ȯɲέp�ܼ�
Dim pagene(1 To 106) As Integer '��ܵP�Ȯ��ܼ�
Dim a As Integer, i As Integer '�Ȯ��ܼ�
Dim k As String '�Ȯ��ܼ�
Dim num(1 To 2) As Integer
If FormMainMode.comaiatk(4).Caption = "�צ�-�L�ɽ��j���׵�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(11, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
 Select Case atkingckai(11, 1)
   Case 1
      atkingckai(11, 1) = 2
       If movecp < 3 Then
            For i = 1 To 106
               If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
                 If pagecardnum(i, 1) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 2))
                 If pagecardnum(i, 3) = a4a Then atkingtotai = Val(atkingtotai) + Val(pagecardnum(i, 4))
               End If
            Next
            
            If atkingtotai >= 4 Then
               atkingtotai = 0
               Select Case movecp
                      Case 1
                         k = a1a
                      Case Is > 1
                         k = a5a
                End Select
               '====================1���q-����ܲĤ@�i�P
               Do
                    '===========(�D�P��U���q����)
                    For a = 106 To 1 Step -1
                       If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                           If pagecardnum(a, 1) = a4a And pagecardnum(a, 3) <> k And pagene(a) = 0 Then
                               atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                               pagene(a) = 1
                               Exit Do
                           ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 1) <> k And pagene(a) = 0 Then
                               atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                               pagene(a) = 1
                               Exit Do
                           End If
                        End If
                    Next
                    If atkingtotai = 0 Then
                        '===========(��ܩҦ�)
                        For a = 106 To 1 Step -1
                           If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                               If pagecardnum(a, 1) = a4a And pagene(a) = 0 Then
                                   atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                                   pagene(a) = 1
                                   Exit Do
                               ElseIf pagecardnum(a, 3) = a4a And pagene(a) = 0 Then
                                   atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                                   pagene(a) = 1
                                   Exit Do
                               End If
                            End If
                        Next
                    End If
               Loop
               If atkingtotai < 4 Then
                   '==============2���q-�̳ѤU�S�ƭȿ�D�P��U���q�����j/�C����2�i�P(�����S�ƭ�)
                  For a = 106 To 1 Step -1
                     If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                         If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) >= 4 - atkingtotai And pagecardnum(a, 3) <> k And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                             pagene(a) = 1
                             Exit For
                         ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) >= 4 - atkingtotai And pagecardnum(a, 1) <> k And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                             pagene(a) = 1
                             Exit For
                         End If
                      End If
                      If atkingtotai >= 4 Then Exit For
                  Next
              End If
              If atkingtotai < 4 Then
                   '====================3���q-�̳ѤU�S�ƭȿ�s�쪺��2�i�P
                  For a = 106 To 1 Step -1
                     If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                         If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) = 4 - atkingtotai And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                             pagene(a) = 1
                             Exit For
                         ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) = 4 - atkingtotai And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                             pagene(a) = 1
                             Exit For
                         End If
                      End If
                      If atkingtotai >= 4 Then Exit For
                  Next
              End If
              If atkingtotai < 4 Then
                 '====================4���q-��Ҧ��ѤU���P
                 For a = 106 To 1 Step -1
                     If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                         If pagecardnum(a, 1) = a4a And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 2))
                             pagene(a) = 1
                         ElseIf pagecardnum(a, 3) = a4a And pagene(a) = 0 Then
                             atkingtotai = Val(atkingtotai) + Val(pagecardnum(a, 4))
                             pagene(a) = 1
                         End If
                      End If
                      If atkingtotai >= 4 Then Exit For
                  Next
              End If
           End If
           If atkingtotai >= 4 Then
               '===========�i���ڥX�P�{��
               For a = 1 To 106
                   If pagene(a) = 1 Then
                       �԰��t����.comatk_AI_��_�צ�_�L�ɽ��j���׵�_�S a
                   End If
               Next
           End If
       End If
   Case 2
          If movecp < 3 Then
            If atkingpagetot(2, 4) >= 4 And atkingckai(11, 2) = 0 Then
               atkingckai(11, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 16
            ElseIf atkingpagetot(2, 4) < 4 And atkingckai(11, 2) = 1 Then
               atkingckai(11, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 16
            End If
          End If
   Case 3
        For i = �H���ޯ�Ʀr���� To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\��\��-�צ�-�L�ɽ��j���׵�_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 8655
                atkingno(i, 6) = 0
                atkingno(i, 7) = 39
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
             atkingckai(11, 2) = 0
             If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                 num(1) = 1
                 num(2) = FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption
                 For i = 2 To 3
                    If FormMainMode.usbi1(����ݾ��H��������(1, i)).Caption > 0 And FormMainMode.usbi1(����ݾ��H��������(1, i)).Caption < num(2) Then
                        num(1) = i
                        num(2) = FormMainMode.usbi1(����ݾ��H��������(1, i)).Caption
                    End If
                Next
                �ˮ`����_�ޯઽ��_�ϥΪ� Val(�Y���淾�q�Ȯ��ܼ�(2)), num(1)
            End If
            �Y���淾�q�Ȯ��ܼ�(2) = 0
            �Y����ˮ`�� = 0
   End Select
End If
End Sub
Sub ��B�����S_��K�g��()
If FormMainMode.comaiatk(1).Caption = "��K�g��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(19, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��B�����S" Then
   Select Case atkingckai(19, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 2 And atkingckai(19, 2) = 0 Then
                   atkingckai(19, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   �������m��l�`��(2) = �������m��l�`��(2) + 4
                End If
                If atkingpagetot(2, 5) < 2 And atkingckai(19, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 4
                   atkingckai(19, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(19, 2) = 0
             �԰��t����.�۰ʱ��b����
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��B�����S\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6525
                   atkingno(i, 6) = 10110
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub ��B�����S_�p��()
If FormMainMode.comaiatk(2).Caption = "�p��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(66, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��B�����S" Then
   Select Case atkingckai(66, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 2 And atkingckai(66, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 4
                   atkingckai(66, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 2 And atkingckai(66, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 4
                   atkingckai(66, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��B�����S\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -240
                   atkingno(i, 5) = 9795
                   atkingno(i, 6) = 10215
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 And Val(FormMainMode.pageusglead.Caption) > 0 Then
                 atking_AI_��B�����S_�p��������(1) = Val(�Y���淾�q�Ȯ��ܼ�(2))
                 atking_AI_��B�����S_�p��������(2) = 1
                 '==========================
                  Do Until atking_AI_��B�����S_�p��������(2) > atking_AI_��B�����S_�p��������(1) Or Val(FormMainMode.pageusglead.Caption) <= 0
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                            �ثe��(21) = 4
                            �ثe��(20) = m
                            atking_AI_��B�����S_�p��������(2) = atking_AI_��B�����S_�p��������(2) + 1
                            FormMainMode.tr�ϥΪ�_��P.Enabled = True
                            Exit Sub
                        End If
                   Loop
             Else
                 atkingckai(66, 1) = 5
                 FormMainMode.��l���槹�Ұ�.Enabled = True
             End If
        Case 4
             Do Until atking_AI_��B�����S_�p��������(2) > atking_AI_��B�����S_�p��������(1) Or Val(FormMainMode.pageusglead.Caption) <= 0
                 Randomize
                 m = Int(Rnd() * 106) + 1
                 If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                     �ثe��(21) = 4
                     �ثe��(20) = m
                     atking_AI_��B�����S_�p��������(2) = atking_AI_��B�����S_�p��������(2) + 1
                     FormMainMode.tr�ϥΪ�_��P.Enabled = True
                     Exit Sub
                 End If
            Loop
            If atking_AI_��B�����S_�p��������(2) > atking_AI_��B�����S_�p��������(1) Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
                atkingckai(66, 1) = 5
                �ثe��(24) = 22
                FormMainMode.���ݮɶ�_2.Enabled = True
            End If
        Case 5
            atkingckai(66, 2) = 0
            Erase atking_AI_��B�����S_�p��������
   End Select
End If
End Sub
Sub ��B�����S_���L()
If FormMainMode.comaiatk(3).Caption = "���L" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(67, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��B�����S" Then
   Select Case atkingckai(67, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 2 And atkingpagetot(2, 2) >= 2 And atkingckai(67, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 7
                   atkingckai(67, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 4) < 2 Or atkingpagetot(2, 2) < 2) And atkingckai(67, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 7
                   atkingckai(67, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��B�����S\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1200
                   atkingno(i, 5) = 6705
                   atkingno(i, 6) = 10245
                   atkingno(i, 7) = 67
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\��B�����S\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(FormMainMode.��ܦC1.goi1) <= 0 Then
                atkingckai(67, 2) = 0
            End If
        Case 4
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) < 0 Then
                �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� Abs(�Y���淾�q�Ȯ��ܼ�(2)), 1
            End If
            atkingckai(67, 2) = 0
   End Select
End If
End Sub
Sub ��B�����S_����()
If FormMainMode.comaiatk(4).Caption = "����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(68, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��B�����S" Then
   Select Case atkingckai(68, 1)
      Case 1
            If pageqlead(2) >= 3 And atkingckai(68, 2) = 0 Then
               atkingckai(68, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If pageqlead(2) < 3 And atkingckai(68, 2) = 1 Then
               atkingckai(68, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��B�����S\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6060
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 68
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(FormMainMode.pageul.Caption) < 2 And atking_AI_��B�����S_���������� = 0 Then
               �԰��t����.����ʧ@_�~�P
            End If
            atking_AI_��B�����S_���������� = atking_AI_��B�����S_���������� + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_��B�����S_���������� > 2
                    �ثe��(15) = 23
                    FormMainMode.tr�P��_��P_�q��.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_��B�����S_���������� > 2 Or Val(FormMainMode.pageul.Caption) <= 0 Then
               atking_AI_��B�����S_���������� = 0
               atkingckai(68, 2) = 0
            End If
   End Select
End If
End Sub
Sub �v��L_�������x()
If FormMainMode.comaiatk(1).Caption = "�������x" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(88, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�v��L" Then
   Select Case atkingckai(88, 1)
        Case 1
            If pageqlead(2) >= 3 And atkingckai(88, 2) = 0 Then
               atkingckai(88, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf pageqlead(2) < 3 And atkingckai(88, 2) = 1 Then
               atkingckai(88, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�v��L\�v��L_�������x_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -360
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 88
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.�p�H���Ϥ� = app_path & "gif\�v��L\����\Staciamini2.png"
            FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = app_path & "gif\�v��L\����\Staciaminidown2.png"
            FormMainMode.personcomminijpg.�p�H���v�lLeft = 90
            FormMainMode.personcomminijpg.�p�H���v�ltop�t = -60
            FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\�v��L\����\Staciaperson2.png"
            FormMainMode.��ܦC1.�q����p�H���Ϥ� = app_path & "gif\�v��L\����\Staciaf2.png"
            atking_AI_�v��L_�����Ҧ����A��(2) = 1
            atkingckai(88, 2) = 0
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personcomminijpg.Visible = True
   End Select
End If
End Sub

Sub �v��L_�M�̤���()
Dim apn As Integer
If FormMainMode.comaiatk(2).Caption = "�M�̤���" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(20, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�v��L" Then
   Select Case atkingckai(20, 1)
        Case 1
            If movecp < 3 Then
             For i = 1 To 3
                 If liveus(i) > 0 Then
                     apn = apn + 1
                 End If
             Next
             If atkingpagetot(2, 1) >= 6 And atkingckai(20, 2) = 0 Then
               atkingckai(20, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + apn * 4
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(20, 2) = 1 Then
               atkingckai(20, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - apn * 4
            End If
          End If
        Case 2
             �԰��t����.�۰ʱ��b����
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�v��L\�v��L_�M�̤���_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 76
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(20, 1) = 3
        Case 3
            For i = 1 To 3
                 If liveus(i) > 0 Then
                     apn = apn + 1
                 End If
            Next
            If atking_AI_�v��L_�����Ҧ����A��(2) = 1 Then
                �԰��t����.�ˮ`����_�ޯઽ��_�q�� apn, 1
            End If
            atkingckai(20, 2) = 0
   End Select
End If
End Sub
Sub �v��L_�ɶ��ؤl()
Dim bloodtot As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "�ɶ��ؤl" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(55, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�v��L" Then
   Select Case atkingckai(55, 1)
        Case 1
            If movecp < 3 Then
             If atkingpagetot(2, 2) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(55, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(55, 2) = 0 Then
               atkingckai(55, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 2) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(55, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(55, 2) = 1 Then
               atkingckai(55, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�v��L\�v��L_�ɶ��ؤl_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 55
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            bloodtot = Val(FormMainMode.turni) \ 2
            If bloodtot > 4 Then bloodtot = 4
            '=============
            If Val(livecom(����H����ԤH��(2, 2))) < Val(livecommax(����H����ԤH��(2, 2))) Then
               Select Case atking_AI_�v��L_�����Ҧ����A��(2)
                   Case 0
                        �^�_����_�q�� bloodtot, 1
                   Case 1
                        �^�_����_�q�� bloodtot \ 2, 1
                  End Select
            End If
            atkingckai(55, 2) = 0
   End Select
End If
End Sub

Sub �v��L_�R�B���K��()
Dim num(1 To 2, 1 To 2) As Integer '��ܤH���Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "�R�B���K��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(21, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�v��L" Then
   Select Case atkingckai(21, 1)
        Case 1
         If movecp = 3 Then
             If atkingpagetot(2, 1) >= 9 And atkingckai(21, 2) = 0 Then
               atkingckai(21, 2) = 1
               atkingtrn(1) = Val(atkingtrn(1)) + 1
            ElseIf atkingpagetot(2, 1) < 9 And atkingckai(21, 2) = 1 Then
               atkingckai(21, 2) = 0
               atkingtrn(1) = Val(atkingtrn(1)) - 1
            End If
         End If
        Case 2
             �԰��t����.�۰ʱ��b����
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�v��L\�v��L_�R�B���K��_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6720
                   atkingno(i, 6) = 9435
                   atkingno(i, 7) = 21
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(livecom(����H����ԤH��(2, 2))) < Val(livecommax(����H����ԤH��(2, 2))) Then
               If atking_AI_�v��L_�����Ҧ����A��(2) = 1 Then
                  �^�_����_�q�� 3, 1
               End If
            End If
            atkingckai(21, 1) = 4
        Case 4
           num(1, 2) = 999 '�ت����̧CHP��
           num(2, 2) = 999
           For i = 2 To 3
               If livecom(����ݾ��H��������(2, i)) < num(2, 2) And livecom(����ݾ��H��������(2, i)) > 0 Then
                   num(2, 1) = i
                   num(2, 2) = livecom(����ݾ��H��������(2, i))
               End If
            Next
            For i = 1 To 3
               If Val(FormMainMode.usbi1(����ݾ��H��������(1, i)).Caption) < num(1, 2) And Val(FormMainMode.usbi1(����ݾ��H��������(1, i)).Caption) > 0 Then
                   num(1, 1) = i
                   num(1, 2) = FormMainMode.usbi1(����ݾ��H��������(1, i)).Caption
               End If
           Next
           If num(2, 2) < num(1, 2) Or num(1, 2) = num(2, 2) Then
               �԰��t����.�ˮ`����_�ߧY���`_�q�� num(2, 1)
           Else
               �԰��t����.�ˮ`����_�ߧY���`_�ϥΪ� num(1, 1)
           End If
           atkingckai(21, 2) = 0
   End Select
End If
End Sub
Sub ������_�Q���{��()
If FormMainMode.comaiatk(1).Caption = "�Q���{��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(22, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
   Select Case atkingckai(22, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingckai(22, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 5
                   atkingckai(22, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 1) < 3 And atkingckai(22, 2) = 1 Then
                   �������m��l�`��(1) = �������m��l�`��(1) - 5
                   atkingckai(22, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(22, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 10290
                   atkingno(i, 6) = 8490
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub ������_�{�q�ۭ���()
If FormMainMode.comaiatk(2).Caption = "�{�q�ۭ���" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(71, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
   Select Case atkingckai(71, 1)
      Case 1
           If movecp = 2 Then
                If atkingpagetot(2, 3) >= 1 And atkingckai(71, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 4
                   atkingckai(71, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 3) < 1 And atkingckai(71, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 4
                   atkingckai(71, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6555
                   atkingno(i, 6) = 8625
                   atkingno(i, 7) = 71
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(71, 2) = 0
             If movecp > 1 Then
                 �԰��t����.����ʧ@_�Z���ܧ� movecp - 1
             End If
   End Select
End If
End Sub
Sub ������_�ۼv�C�R()
Dim rrr(1 To 3) As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "�ۼv�C�R" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(23, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
   Select Case atkingckai(23, 1)
      Case 1
            If movecp = 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                        If pagecardnum(i, 1) = a1a And pagecardnum(i, 2) = 1 Then
                           rrr(1) = rrr(1) + 1
                        End If
                        If pagecardnum(i, 1) = a1a And pagecardnum(i, 2) = 2 Then
                           rrr(2) = rrr(2) + 1
                        End If
                        If pagecardnum(i, 1) = a1a And pagecardnum(i, 2) = 3 Then
                           rrr(3) = rrr(3) + 1
                        End If
                    End If
                 Next
            End If
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(23, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(23, 2) = 0 Then
                atkingckai(23, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
                �������m��l�`��(2) = �������m��l�`��(2) + 9
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(23, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(23, 2) = 1 Then
                atkingckai(23, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
                �������m��l�`��(2) = �������m��l�`��(2) - 9
              End If
      Case 2
             atkingckai(23, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8520
                   atkingno(i, 6) = 8280
                   atkingno(i, 7) = 99
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\������\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub ����_�ɶ��z�u()
Dim tn As Integer
If FormMainMode.comaiatk(3).Caption = "�ɶ��z�u" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(24, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(24, 1)
        Case 1
             If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(24, 2) = 0 Then
               atkingckai(24, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 7
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(24, 2) = 1 Then
               atkingckai(24, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 7
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7125
                   atkingno(i, 6) = 9330
                   atkingno(i, 7) = 24
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\����\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(24, 2) = 0
            tn = Val(FormMainMode.turni)
            If tn = 2 Or tn = 3 Or tn = 5 Or tn = 7 Or tn = 11 Or tn = 13 Or tn = 17 Then
               �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 3, 1
            End If
   End Select
End If
End Sub
Sub ����_�ɶ��l�y()
If FormMainMode.comaiatk(2).Caption = "�ɶ��l�y" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(70, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(70, 1)
        Case 1
             If movecp < 3 Then
                 If atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(70, 2) = 0 Then
                   atkingckai(70, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(70, 2) = 1 Then
                   atkingckai(70, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6510
                   atkingno(i, 6) = 9690
                   atkingno(i, 7) = 83
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '=====================
            atkingckai(70, 2) = 0
            �԰��t����.�����g�J��ܦC�ƭ� 1, Val(FormMainMode.��ܦC1.goi1) - Val(FormMainMode.turni)
'            �������m��l�`��(1) = FormMainMode.��ܦC1.goi1
   End Select
End If
End Sub

Sub ��̬d�w_�s�g()
If FormMainMode.comaiatk(1).Caption = "�s�g" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(25, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��̬d�w" Then
   Select Case atkingckai(25, 1)
      Case 1
           If movecp > 1 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          End If
          If rrr >= 2 And atkingckai(25, 2) = 0 Then
             �������m��l�`��(2) = �������m��l�`��(2) + 6
             atkingckai(25, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 2 And atkingckai(25, 2) = 1 Then
             �������m��l�`��(2) = �������m��l�`��(2) - 6
             atkingckai(25, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
      Case 2
             atkingckai(25, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��̬d�w\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6780
                   atkingno(i, 6) = 10185
                   atkingno(i, 7) = 25
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\��̬d�w\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub ��̬d�w_����@��()
Dim ape As Integer
If FormMainMode.comaiatk(3).Caption = "����@��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(69, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��̬d�w" Then
   Select Case atkingckai(69, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(69, 2) = 0 Then
                   ape = (livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2))) * 2
                   If ape > 16 Then ape = 16
                   atkingckai(69, 2) = 1
                   �������m��l�`��(2) = �������m��l�`��(2) + ape
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(69, 2) = 1 Then
                   ape = (livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2))) * 2
                   If ape > 16 Then ape = 16
                   �������m��l�`��(2) = �������m��l�`��(2) - ape
                   atkingckai(69, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(69, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��̬d�w\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6615
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\��̬d�w\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub

Sub ��̬d�w_���t���C(ByVal Index As Integer)
Dim aw As Integer
If FormMainMode.comaiatk(2).Caption = "���t���C" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(26, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��̬d�w" Then
   Select Case atkingckai(26, 1)
      Case 1
             If movecp > 1 Then
                 If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 1) >= 2 And atkingckai(26, 2) = 0 Then
                     aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                     �������m��l�`��(2) = �������m��l�`��(2) + (aw - atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1))
                     atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1) = aw
                     atkingckai(26, 2) = 1
                     atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
            End If
      Case 2
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 2 And atkingckai(26, 2) = 1 Then
                   aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                   �������m��l�`��(2) = �������m��l�`��(2) + (aw - atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1))
                   atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1) = aw
            End If
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 2 And atkingckai(26, 2) = 1 Then
                   If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 1) >= 2 And atkingckai(26, 2) = 1 Then
                        aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                        �������m��l�`��(2) = �������m��l�`��(2) - (atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1) - aw)
                        atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1) = aw
                   ElseIf (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 1) < 2) And atkingckai(26, 2) = 1 Then
                        �������m��l�`��(2) = �������m��l�`��(2) - atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1)
                        atkingckai(26, 2) = 0
                        atkingckai(26, 1) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) - 1
                        Erase atking_AI_��̬d�w_���t���C�p��ƭȬ�����
                    End If
            End If
'            formmainmode.trgoi2.Enabled = True
    Case 3
        If Val(pagecardnum(Index, 5)) = 2 And atkingckai(26, 2) = 1 Then
               If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 1) >= 2 Then
                    aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                    �������m��l�`��(2) = �������m��l�`��(2) + (aw - atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1))
                    atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1) = aw
               ElseIf (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 1) < 2) Then
                    �������m��l�`��(2) = �������m��l�`��(2) - atking_AI_��̬d�w_���t���C�p��ƭȬ�����(1)
                    atkingckai(26, 2) = 0
                    atkingckai(26, 1) = 1
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
                    Erase atking_AI_��̬d�w_���t���C�p��ƭȬ�����
                End If
        End If
'        formmainmode.trgoi2.Enabled = True
      Case 4
             atkingckai(26, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��̬d�w\atking2-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 9165
                   atkingno(i, 6) = 10350
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\��̬d�w\atking2-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             Erase atking_AI_��̬d�w_���t���C�p��ƭȬ�����
   End Select
End If
End Sub
Sub ��̬d�w_���}����()
If FormMainMode.comaiatk(4).Caption = "���}����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(27, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��̬d�w" Then
   Select Case atkingckai(27, 1)
      Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a2a And Val(pagecardnum(i, 2)) = 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          If rrr >= 2 And atkingckai(27, 2) = 0 Then
'          If pageqlead(2) >= 1 And atkingckai(27, 2) = 0 Then
             atkingckai(27, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 2 And atkingckai(27, 2) = 1 Then
'          If pageqlead(2) < 1 And atkingckai(27, 2) = 1 Then
             atkingckai(27, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��̬d�w\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5760
                   atkingno(i, 6) = 9450
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(27, 2) = 0
             If Val(�Y���淾�q�Ȯ��ܼ�(2)) >= livecom(����H����ԤH��(2, 2)) Then
                 �Y���淾�q�Ȯ��ܼ�(2) = livecom(����H����ԤH��(2, 2)) - 1
                 �Y����ˮ`�� = �Y���淾�q�Ȯ��ܼ�(2)
             End If
   End Select
End If
End Sub
Sub ������_�r�֩��(ByVal Index As Integer)
Dim n(1 To 2) As Integer
If FormMainMode.comaiatk(3).Caption = "�r�֩��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(111, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
 Select Case atkingckai(111, 1)
    Case 1
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(111, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(111, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2)) * 5
               If atkingckai(111, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(111, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
    Case 2
        If pagecardnum(Index, 1) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2)) * 5
               If atkingckai(111, 2) = 0 And atkingpagetot(2, 4) > 0 Then
                  atkingckai(111, 2) = 1
                  atkingtrn(2) = Val(atkingtrn(2)) + 1
               End If
        End If
        If pagecardnum(Index, 3) = a4a And Val(pagecardnum(Index, 5)) = 2 Then
               �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4)) * 5
               If atkingckai(111, 2) = 1 And atkingpagetot(2, 4) = 0 Then
                  atkingckai(111, 2) = 0
                  atkingtrn(2) = Val(atkingtrn(2)) - 1
               End If
        End If
    Case 3
        For i = �H���ޯ�Ʀr���� To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\������\atking3_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6645
                atkingno(i, 6) = 9555
                atkingno(i, 7) = 111
                atkingno(i, 8) = 1
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
       '-------------
    Case 4
       atkingckai(111, 2) = 0
        n(1) = 999 '���̤pHP��
        n(2) = 0
        For i = 2 To 3
            If livecom(����ݾ��H��������(2, i)) > 0 And livecom(����ݾ��H��������(2, i)) < n(1) Then
                n(1) = livecom(����ݾ��H��������(2, i))
                n(2) = i
            End If
        Next
        If n(2) > 0 Then
            �԰��t����.�ˮ`����_�ޯઽ��_�q�� Val(atkingpagetot(2, 4)), n(2)
        Else
            �԰��t����.�ˮ`����_�ޯઽ��_�q�� Val(atkingpagetot(2, 4)), 1
        End If
  End Select
End If
End Sub
Sub ������_�ŬX�`�g()
Dim n(1 To 2) As Integer
If FormMainMode.comaiatk(2).Caption = "�ŬX�`�g" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(28, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
   Select Case atkingckai(28, 1)
        Case 1
            If movecp < 3 Then
             If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 2) >= 2 And atkingckai(28, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(28, 2) = 0 Then
               atkingckai(28, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 5
            ElseIf (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 2) < 2) And atkingckai(28, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(28, 2) = 1 Then
               atkingckai(28, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 5
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6165
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 28
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(28, 2) = 0
            '=======================
            n(1) = 999 '���̤pHP��
            n(2) = 0
            For i = 2 To 3
                If livecom(����ݾ��H��������(2, i)) > 0 And livecom(����ݾ��H��������(2, i)) < n(1) Then
                    n(1) = livecom(����ݾ��H��������(2, i))
                    n(2) = i
                End If
            Next
            If n(2) > 0 Then
                If livecom(����H����ԤH��(2, 2)) >= n(1) Then
                    �԰��t����.�^�_����_�q�� livecom(����H����ԤH��(2, 2)) - n(1), n(2)
                Else
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� n(1) - livecom(����H����ԤH��(2, 2)), n(2)
                End If
            End If
   End Select
End If
End Sub
Sub ������_��������()
If FormMainMode.comaiatk(1).Caption = "��������" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(29, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
   Select Case atkingckai(29, 1)
        Case 1
            If pageqlead(2) >= 2 And atkingckai(29, 2) = 0 Then
               atkingckai(29, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf pageqlead(2) < 2 And atkingckai(29, 2) = 1 Then
               atkingckai(29, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6360
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 29
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\������\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(29, 2) = 0
            '==============================
            For k = 2 To 3
                �ˮ`����_�ޯઽ��_�q�� 1, k
            Next
            '==============================
            atking_AI_������_�����Ҧ����A��(2) = 1
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.�p�H���Ϥ� = app_path & "gif\������\����\Nenemmini2.png"
            FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = app_path & "gif\������\����\Nenemminidown2.png"
            FormMainMode.personcomminijpg.�p�H���v�lLeft = 20
            FormMainMode.personcomminijpg.�p�H���v�ltop�t = -90
            FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\������\����\Nenemperson2.png"
            FormMainMode.��ܦC1.�q����p�H���Ϥ� = app_path & "gif\������\����\Nenemf2.png"
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personcomminijpg.Visible = True
   End Select
End If
End Sub
Sub ������_���K�W��()
If FormMainMode.comaiatk(4).Caption = "���K�W��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(112, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
   Select Case atkingckai(112, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(112, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(112, 2) = 0 Then
               atkingckai(112, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(112, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(112, 2) = 1 Then
               atkingckai(112, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6750
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 112
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(112, 2) = 0
            '=======================
            For i = 2 To 3
                �԰��t����.�^�_����_�q�� 10, i
            Next
            �԰��t����.�ˮ`����_�ߧY���`_�q�� 1
            '=======================
            If atking_AI_������_�����Ҧ����A��(2) = 1 Then
                �P�`���q��(2) = �P�`���q��(2) + 1
            End If
   End Select
End If
End Sub
Sub ����_High_hand()
If FormMainMode.comaiatk(1).Caption = "High hand" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(64, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(64, 1)
        Case 1
             If atkingpagetot(2, 4) >= 2 And atkingckai(64, 2) = 0 Then
               atkingckai(64, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + pageqlead(1) * 2
            ElseIf atkingpagetot(2, 4) < 2 And atkingckai(64, 2) = 1 Then
               atkingckai(64, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - pageqlead(1) * 2
            End If
        Case 2
             atkingckai(64, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\High hand_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7770
                   atkingno(i, 6) = 10020
                   atkingno(i, 7) = 63
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub ����_Lowball()
If FormMainMode.comaiatk(3).Caption = "Lowball" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(65, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(65, 1)
        Case 1
             If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 _
                And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 5) >= 1 And atkingckai(65, 2) = 0 Then
'            If atkingpagetot(2, 1) >= 1 And atkingckai(65, 2) = 0 Then
                    atkingckai(65, 2) = 1
                    atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1 _
               Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 5) < 1) And atkingckai(65, 2) = 1 Then
'            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(65, 2) = 1 Then
                    atkingckai(65, 2) = 0
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\Lowball_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7020
                   atkingno(i, 6) = 9555
                   atkingno(i, 7) = 65
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             �Y���淾�q�Ȯ��ܼ�(2) = Val(�Y���淾�q�Ȯ��ܼ�(2)) + 5
             �Y����ˮ`�� = �Y���淾�q�Ȯ��ܼ�(2)
             atkingckai(65, 2) = 0
   End Select
End If
End Sub
Sub ����_Gamble()
If FormMainMode.comaiatk(4).Caption = "Gamble" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(30, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(30, 1)
        Case 1
            If movecp = 1 Then
                 If pageqlead(2) >= 3 And atkingckai(30, 2) = 0 Then
                   atkingckai(30, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf pageqlead(2) < 3 And atkingckai(30, 2) = 1 Then
                   atkingckai(30, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\Gamble_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6960
                   atkingno(i, 6) = 9780
                   atkingno(i, 7) = 106
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(30, 2) = 0
             If Val(�Y���淾�q�Ȯ��ܼ�(2)) = 1 Then
                 �԰��t����.�ˮ`����_�ߧY���`_�ϥΪ� 1
             End If
   End Select
End If
End Sub
Sub ����_Jackpot()
Dim m As Integer
If FormMainMode.comaiatk(2).Caption = "Jackpot" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(31, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(31, 1)
        Case 1
            If movecp = 2 Then
                If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(31, 2) = 0 Then
'                If atkingpagetot(2, 2) >= 1 And atkingckai(31, 2) = 0 Then
                   atkingckai(31, 2) = 1
                ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(31, 2) = 1 Then
'                ElseIf atkingpagetot(2, 2) < 1 And atkingckai(31, 2) = 1 Then
                   atkingckai(31, 2) = 0
                End If
            End If
        Case 2
             atking_AI_����_Jackpot������(1) = pageqlead(2) * 2
             atking_AI_����_Jackpot������(2) = 1
        Case 3
             atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 4
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\Jackpot_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6480
                   atkingno(i, 6) = 10020
                   atkingno(i, 7) = 31
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
             If Val(FormMainMode.pageul.Caption) < atking_AI_����_Jackpot������(1) And atking_AI_����_Jackpot������(2) = 1 Then
               �԰��t����.����ʧ@_�~�P
             End If
             If atking_AI_����_Jackpot������(2) > atking_AI_����_Jackpot������(1) Or Val(FormMainMode.pageul.Caption) <= 0 Then
                 atkingckai(31, 2) = 0
                 �԰��t����.����ʧ@_�ޯ��ʵ���
            Else
                �ثe��(15) = 22
                FormMainMode.tr�P��_��P_�q��.Enabled = True
                atking_AI_����_Jackpot������(2) = atking_AI_����_Jackpot������(2) + 1
            End If
   End Select
End If
End Sub
Sub ù��Y_�V�大�b()
If FormMainMode.comaiatk(2).Caption = "�V�大�b" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(32, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "ù��Y" Then
   Select Case atkingckai(32, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(32, 2) = 0 Then
               atkingckai(32, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 5
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(32, 2) = 1 Then
               atkingckai(32, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 5
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\ù��Y\ù��Y_�V�大�b_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 32
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            �^�_����_�q�� 1, 1
        Case 4
            atkingckai(32, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                �^�_����_�q�� 1, 1
            End If
   End Select
End If
End Sub
Sub ù��Y_��������¶()
Dim m As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "��������¶" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(59, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "ù��Y" Then
   Select Case atkingckai(59, 1)
        Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(59, 2) = 0 Then
               atkingckai(59, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 4
            ElseIf atkingpagetot(2, 4) < 2 And atkingckai(59, 2) = 1 Then
               atkingckai(59, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 4
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\ù��Y\ù��Y_��������¶_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -240
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6390
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 52
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(59, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                       Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 20 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  �H�����`���A��Ʈw(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 20, app_path & "gif\���`���A\damage.gif", 0, 2
                                  ���`���A�ˬd��(20, 1) = 1
                                  ���`���A�ˬd��(20, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 16 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  �H�����`���A��Ʈw(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 16, app_path & "gif\���`���A\moveerr.gif", 0, 2
                                  ���`���A�ˬd��(16, 1) = 1
                                  ���`���A�ˬd��(16, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 22 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  �H�����`���A��Ʈw(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 22, app_path & "gif\���`���A\atkingerr.gif", 0, 2
                                  ���`���A�ˬd��(22, 1) = 1
                                  ���`���A�ˬd��(22, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
            End If
   End Select
End If
End Sub
Sub ù��Y_�C�G����L()
If FormMainMode.comaiatk(3).Caption = "�C�G����L" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(60, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "ù��Y" Then
   Select Case atkingckai(60, 1)
        Case 1
            If movecp > 1 Then
                If atkingpagetot(2, 2) >= 5 And atkingpagetot(2, 4) >= 1 And atkingckai(60, 2) = 0 Then
    '             If atkingpagetot(2, 2) >= 1 And atkingck(24, 2) = 0 Then
                   atkingckai(60, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 2) < 5 Or atkingpagetot(2, 4) < 1) And atkingckai(60, 2) = 1 Then
    '            ElseIf atkingpagetot(2, 2) < 1 And atkingck(24, 2) = 1 Then
                   atkingckai(60, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             atkingckai(60, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\ù��Y\ù��Y_�C�G����L_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6975
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 53
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===============
             If atkingpagetot(1, 4) >= 1 Then
                 �԰��t����.�����g�J��ܦC�ƭ� 1, Int(Val(FormMainMode.��ܦC1.goi1) / 3 + 0.9)
             Else
                 �԰��t����.�����g�J��ܦC�ƭ� 1, Int(Val(FormMainMode.��ܦC1.goi1) / 2 + 0.9)
             End If
'             �������m��l�`��(1) = FormMainMode.��ܦC1.goi1
   End Select
End If
End Sub
Sub CC_���ߪŶ�()
If FormMainMode.comaiatk(1).Caption = "���ߪŶ�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(103, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "C.C." Then
   Select Case atkingckai(103, 1)
        Case 1
             If atkingpagetot(2, 4) >= 1 And atkingckai(103, 2) = 0 Then
               atkingckai(103, 2) = 1
            ElseIf atkingpagetot(2, 4) < 1 And atkingckai(103, 2) = 1 Then
               atkingckai(103, 2) = 0
            End If
        Case 2
            atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_���ߪŶ�_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7275
                   atkingno(i, 6) = 9480
                   atkingno(i, 7) = 103
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            For i = 1 To 3
                �^�_����_�q�� 1, i
            Next
            atkingckai(103, 2) = 0
            '======================
               �԰��t����.����ʧ@_�M���Ҧ����`���A_�q��
           '======================
   End Select
End If
End Sub
Sub CC_�ջȾԾ�()
Dim bloodntot As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(2).Caption = "�ջȾԾ�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(33, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "C.C." Then
   Select Case atkingckai(33, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 5) >= 2 And atkingckai(33, 2) = 0 Then
'             If atkingpagetot(2, 1) >= 1 And atkingckai(33, 2) = 0 Then
               atkingckai(33, 2) = 1
               �������m��l�`��(2) = �������m��l�`��(2) + 4
            ElseIf (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 5) < 2) And atkingckai(33, 2) = 1 Then
'            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(33, 2) = 1 Then
               atkingckai(33, 2) = 0
               �������m��l�`��(2) = �������m��l�`��(2) - 4
            End If
          End If
        Case 2
            atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_�ջȾԾ�_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -720
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 33
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            atkingckai(33, 2) = 0
            For i = 1 To 3
                If i = 1 Then
                    Randomize
                    bloodntot = Int(Rnd() * 3) + 0
                    If Val(FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption) > 1 And bloodntot < Val(FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption) Then
                       �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� bloodntot, 1
                    ElseIf Val(FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption) = 2 And bloodntot = 2 Then
                       bloodntot = 1
                       �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� bloodntot, 1
                    End If
                Else
                    Randomize
                    bloodntot = Int(Rnd() * 3) + 0
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� bloodntot, i
                End If
            Next
   End Select
End If
End Sub
Sub CC_��l����()
If FormMainMode.comaiatk(3).Caption = "��l����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(57, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "C.C." Then
   Select Case atkingckai(57, 1)
        Case 1
             If atkingpagetot(2, 2) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(57, 2) = 0 Then
               atkingckai(57, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 2
            ElseIf (atkingpagetot(2, 2) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(57, 2) = 1 Then
               atkingckai(57, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 2
            End If
        Case 2
            '===========�N�Ҧ��ޯ�L�Ĥ�-�ϥΪ̤�(���q1)
            atkingtrn(1) = 0
            For i = 1 To UBound(atkingck)
                 atkingck(i, 2) = 0
            Next
        Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_��l����_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6945
                   atkingno(i, 6) = 10050
                   atkingno(i, 7) = 57
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
              '=================���ƭȬ����ƭ�
              FormMainMode.��ܦC1.goi1 = �������m��l�`��(3)
              FormMainMode.��ܦC1.goi2 = �������m��l�`��(4) + 2
              '===================
                For i = 1 To 4
                    �԰��t����.�H���ޯ���O�}�� False, i
                Next
                '==================
                atking_��_�u�@�Ҧ����A�Ұʭ� = False
                Erase atking_�v��L_�����Ҧ����A��
        Case 4
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 1)
            FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 2)
            FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 4)
            FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�width
            FormMainMode.personusminijpg.�p�H���v�lLeft = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 5))
            FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 6))
            FormMainMode.personusminijpg.Visible = True
            FormDice.jpgus.�j�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 3)
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            atkingckai(57, 2) = 0
   End Select
End If
End Sub

Sub CC_���W�q�Ϥ�N�M()
If FormMainMode.comaiatk(4).Caption = "���W�q�Ϥ�N�M" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(50, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "C.C." Then
   Select Case atkingckai(50, 1)
        Case 1
            If movecp = 1 Then
                If atkingpagetot(2, 4) >= 6 And atkingckai(50, 2) = 0 Then
                   atkingckai(50, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   �������m��l�`��(2) = �������m��l�`��(2) + 24
                ElseIf atkingpagetot(2, 4) < 6 And atkingckai(50, 2) = 1 Then
                   atkingckai(50, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   �������m��l�`��(2) = �������m��l�`��(2) - 24
                End If
            End If
        Case 2
             atkingckai(50, 1) = 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\CC\CC_���W�q�Ϥ�N�M_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9630
                   atkingno(i, 6) = 8940
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
          '===========���褤���`���A
            Do
                For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                  If �H�����`���A��Ʈw(1, i, 3) = 16 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                      FormMainMode.personusspe(i).person_turn = 3
                      �H�����`���A��Ʈw(1, i, 2) = 3
                      Exit Do
                  End If
                Next
                For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                   If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                      �԰��t����.�H�����`���A��]�w_��] 1, i, 16, app_path & "gif\���`���A\moveerr.gif", 0, 3
                      ���`���A�ˬd��(16, 1) = 1
                      ���`���A�ˬd��(16, 2) = 1
                      Exit Do
                   End If
                Next
            Loop
            '===============
            Do
                For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                  If �H�����`���A��Ʈw(2, i, 3) = 17 And �H�����`���A��Ʈw(2, i, 2) > 0 Then
                      FormMainMode.personcomspe(i).person_turn = 3
                      �H�����`���A��Ʈw(2, i, 2) = 3
                      Exit Do
                  End If
                Next
                For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                   If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                      �԰��t����.�H�����`���A��]�w_��] 2, i, 17, app_path & "gif\���`���A\moveerr.gif", 0, 3
                      ���`���A�ˬd��(17, 1) = 1
                      ���`���A�ˬd��(17, 2) = 1
                      Exit Do
                   End If
                Next
            Loop
            atkingckai(50, 2) = 0
   End Select
End If
End Sub

Sub ���[_�ԷX���T��()
Dim rrr As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "�ԷX���T��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(34, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���[" Then
   Select Case atkingckai(34, 1)
      Case 1
         If movecp = 1 Then
            If atkingpagetot(2, 1) >= 6 And atkingckai(34, 2) = 0 Then
               atkingckai(34, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(34, 2) = 1 Then
               atkingckai(34, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���[\���[_�ԷX���T��_1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6645
                   atkingno(i, 6) = 9330
                   atkingno(i, 7) = 34
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
     Case 3
           For rrr = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
             If �H�����`���A��Ʈw(2, rrr, 3) = 26 Then
                �^�_����_�q�� �H�����`���A��Ʈw(2, rrr, 2), 1
                �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� �H�����`���A��Ʈw(2, rrr, 2), 1
                Exit For
             End If
           Next
            '=====================
               ����ʧ@_�M���Ҧ����`���A_�ϥΪ�
               ����ʧ@_�M���Ҧ����`���A_�q��
           '======================
           atkingckai(34, 2) = 0
   End Select
End If
End Sub
Sub ���[_�O�d���Ų�()
If FormMainMode.comaiatk(3).Caption = "�O�d���Ų�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(35, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���[" Then
   Select Case atkingckai(35, 1)
      Case 1
          If movecp > 1 Then
             If atkingpagetot(2, 1) >= 6 And atkingckai(35, 2) = 0 Then
               atkingckai(35, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 1) < 6 And atkingckai(35, 2) = 1 Then
               atkingckai(35, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
          End If
      Case 2
          atking_AI_���[_�O�d���Ų�_tot(1) = atking_AI_���[_�O�d���Ų�_tot(1) + �������m��l�`��(2)
          �������m��l�`��(2) = 0
          atking_AI_���[_�O�d���Ų�_tot(2) = 1
          atkingckai(35, 1) = 1
      Case 3
          atking_AI_���[_�O�d���Ų�_tot(1) = atking_AI_���[_�O�d���Ų�_tot(1) + �������m��l�`��(2)
          �������m��l�`��(2) = atking_AI_���[_�O�d���Ų�_tot(1)
          atking_AI_���[_�O�d���Ų�_tot(1) = 0
          atking_AI_���[_�O�d���Ų�_tot(2) = 0
          atkingckai(35, 1) = 1
      Case 4
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���[\���[_�O�d���Ų�_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6945
                   atkingno(i, 6) = 9870
                   atkingno(i, 7) = 35
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 5
            Do
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, i, 2) >= 9 And �H�����`���A��Ʈw(2, i, 3) = 26 Then
                    Exit Do
                 End If
                 If �H�����`���A��Ʈw(2, i, 3) = 26 And �H�����`���A��Ʈw(2, i, 2) = 8 Then
                     FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2) + 1
                     �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + 1
                     Exit Do
                 ElseIf �H�����`���A��Ʈw(2, i, 3) = 26 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) <= 7 Then
'                 If �H�����`���A��Ʈw(2, i, 3) = 26 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) <= 97 Then
                     FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2) + 2
                     �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + 2
                     Exit Do
                 End If
               Next
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                  If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                     �԰��t����.�H�����`���A��]�w_��] 2, i, 26, app_path & "gif\���`���A\�t��.gif", 0, 2
                     ���`���A�ˬd��(26, 1) = 1
                     ���`���A�ˬd��(26, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
            '================
            �^�_����_�q�� 2, 1
            '================
            atkingckai(35, 2) = 0
            atkingckai(35, 1) = 0
            Erase atking_AI_���[_�O�d���Ų�_tot
   End Select
End If
End Sub
Sub ���[_�R�Ĥ��I()
If FormMainMode.comaiatk(2).Caption = "�R�Ĥ��I" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(36, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���[" Then
   Select Case atkingckai(36, 1)
      Case 1
         If movecp < 3 Then
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 2 And atkingckai(36, 2) = 0 Then
'            If atkingpagetot(2, 1) >= 1 And atkingckai(36, 2) = 0 Then
               atkingckai(36, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 2) And atkingckai(36, 2) = 1 Then
'            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(36, 2) = 1 Then
               atkingckai(36, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        End If
      Case 2
             atkingckai(36, 1) = 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���[\���[_�R�Ĥ��I_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8340
                   atkingno(i, 6) = 8520
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
      Case 3
            For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, i, 3) = 26 And �H�����`���A��Ʈw(2, i, 2) >= 1 Then
                     �Y���淾�q�Ȯ��ܼ�(2) = Val(�Y���淾�q�Ȯ��ܼ�(2)) + �H�����`���A��Ʈw(2, i, 2)
                     �Y����ˮ`�� = �Y���淾�q�Ȯ��ܼ�(2)
                     Exit For
                 End If
            Next
           atkingckai(36, 2) = 0
           For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
             If �H�����`���A��Ʈw(2, i, 3) = 26 And �H�����`���A��Ʈw(2, i, 2) >= 1 Then
                 FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2) - 1
                 �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
                 If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                     '===�~�ӤU�@���A���
                     �԰��t����.���`���A�~��_�q��
                 End If
                 Exit For
             End If
           Next
   End Select
End If
End Sub
Sub ���_�Q�T����()
Dim rrr(1 To 2) As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "�Q�T����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(37, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���" Then
   Select Case atkingckai(37, 1)
        Case 1
           If movecp < 3 Then
             For i = 1 To 106
                If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr(1) = rrr(1) + 1
                End If
                If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr(2) = rrr(2) + 1
                End If
             Next
           End If
          If rrr(1) >= 1 And rrr(2) >= 1 And atkingckai(37, 2) = 0 Then
'          If rrr(1) >= 1 And atkingckai(37, 2) = 0 Then
             atkingckai(37, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If (rrr(1) < 1 Or rrr(2) < 1) And atkingckai(37, 2) = 1 Then
'          If rrr(1) < 1 And atkingckai(37, 2) = 1 Then
             atkingckai(37, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
        Case 2
            If atking_AI_���_�Q�T����_tot(2) = 0 Then
                atking_AI_���_�Q�T����_tot(1) = �������m��l�`��(2)
                atking_AI_���_�Q�T����_tot(2) = 1
                �������m��l�`��(2) = 13
                �������m��l�`��(1) = 0
                atkingckai(37, 1) = 1
            ElseIf atking_AI_���_�Q�T����_tot(2) = 1 Then
                atking_AI_���_�Q�T����_tot(1) = atking_AI_���_�Q�T����_tot(1) + (�������m��l�`��(2) - 13)
                �������m��l�`��(2) = 13
                �������m��l�`��(1) = 0
                atkingckai(37, 1) = 1
            End If
        Case 3
           atking_AI_���_�Q�T����_tot(1) = atking_AI_���_�Q�T����_tot(1) + (�������m��l�`��(2) - 13)
           �������m��l�`��(2) = atking_AI_���_�Q�T����_tot(1)
           atking_AI_���_�Q�T����_tot(1) = 0
           atking_AI_���_�Q�T����_tot(2) = 0
           atkingckai(37, 1) = 1
        Case 4
             �԰��t����.�۰ʱ��b����
             atkingckai(37, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���\���_�Q�T����_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7980
                   atkingno(i, 6) = 9015
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             Erase atking_AI_���_�Q�T����_tot
             '==============
            �԰��t����.�����g�J��ܦC�ƭ� 2, 13
'            �������m��l�`��(2) = FormMainMode.��ܦC1.goi2
            �԰��t����.�����g�J��ܦC�ƭ� 1, 0
'            �������m��l�`��(1) = FormMainMode.��ܦC1.goi1
        Case 5
            �������m��l�`��(1) = 0
   End Select
End If
End Sub
Sub ���_��Ө���()
Dim bloodtot As Single  '�Ȯ��ܼ�
Dim num As Integer
If FormMainMode.comaiatk(2).Caption = "��Ө���" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(38, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���" Then
   Select Case atkingckai(38, 1)
        Case 1
             If atkingpagetot(2, 3) >= 1 And atkingckai(38, 2) = 0 Then
               atkingckai(38, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(38, 2) = 1 Then
               atkingckai(38, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���\���_��Ө���_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6330
                   atkingno(i, 6) = 9285
                   atkingno(i, 7) = 114
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(38, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                bloodtot = Val(�Y���淾�q�Ȯ��ܼ�(2)) \ Val(2)
                Do
                    Randomize
                    num = Int(Rnd() * 3) + 1
                    If liveus(����ݾ��H��������(1, num)) > 0 Then
                        �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� bloodtot, num
                        Exit Do
                    End If
                Loop
            End If
   End Select
End If
End Sub
Sub ���_�E���F��()
Dim bloodtot As Single  '�Ȯ��ܼ�
Dim pic As Integer 'RND�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "�E���F��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(56, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���" Then
   Select Case atkingckai(56, 1)
        Case 1
            If movecp > 1 Then
             If atkingpagetot(2, 2) >= 5 And atkingpagetot(2, 4) >= 1 And atkingckai(56, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(56, 2) = 0 Then
               atkingckai(56, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 9
            ElseIf (atkingpagetot(2, 2) < 5 Or atkingpagetot(2, 4) < 1) And atkingckai(56, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(56, 2) = 1 Then
               atkingckai(56, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 9
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���\���_�E���F��_2\���_�E���F��main.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6330
                   atkingno(i, 6) = 9510
                   atkingno(i, 7) = 56
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 11) = 0
                   '=================
                   Randomize
                   pic = Int(Rnd() * 8) + 1
                   atkingno(i, 10) = app_path & "gif\���\���_�E���F��_2\���_�E���F��" & pic & ".jpg"
                   Exit For
                 End If
             Next
        Case 3
            bloodtot = Int(atkingpagetot(2, 4) / 2 + 0.5)
            '=============
            If Val(livecom(����H����ԤH��(2, 2))) < Val(livecommax(����H����ԤH��(2, 2))) Then
                �԰��t����.�^�_����_�q�� bloodtot, 1
            End If
            atkingckai(56, 2) = 0
   End Select
End If
End Sub

Sub �L���S_�V����()
If FormMainMode.comaiatk(2).Caption = "�V����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(39, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�L���S" Then
   Select Case atkingckai(39, 1)
      Case 1
            If movecp < 3 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(39, 2) = 0 Then
                   atkingckai(39, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(39, 2) = 1 Then
                   atkingckai(39, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                 End If
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�L���S\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9375
                   atkingno(i, 7) = 39
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(39, 2) = 0
             '========================
             For i = 18 To (turn + 3) Step -1
                  pageeventnum(2, i, 1) = pageeventnum(2, i - 2, 1)
                  pageeventnum(2, i, 2) = pageeventnum(2, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 2)
                  pageeventnum(2, i, 1) = "�C5/�j5"
                  pageeventnum(2, i, 2) = �@��t����.�ƥ�d��Ʈw("�C5/�j5", 2)
             Next
   End Select
End If
End Sub
Sub �L���S_���֪��z��()
If FormMainMode.comaiatk(4).Caption = "���֪��z��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(115, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�L���S" Then
   Select Case atkingckai(115, 1)
      Case 1
            If atkingpagetot(2, 4) >= 3 And atkingckai(115, 2) = 0 Then
               atkingckai(115, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 3 And atkingckai(115, 2) = 1 Then
               atkingckai(115, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�L���S\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 600
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6915
                   atkingno(i, 6) = 9690
                   atkingno(i, 7) = 115
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(115, 2) = 0
             '========================
             If �P�`���q��(2) > 0 Then
                 �P�`���q��(2) = �P�`���q��(2) - 1
             End If
             '========================
             For i = 18 To (turn + 4) Step -1
                  pageeventnum(2, i, 1) = pageeventnum(2, i - 2, 1)
                  pageeventnum(2, i, 2) = pageeventnum(2, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 3)
                  pageeventnum(2, i, 1) = "���|5"
                  pageeventnum(2, i, 2) = �@��t����.�ƥ�d��Ʈw("���|5", 2)
             Next
   End Select
End If
End Sub
Sub �L���S_�j�t��()
Dim p, i, j As Integer
If FormMainMode.comaiatk(1).Caption = "�j�t��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(90, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�L���S" Then
   Select Case atkingckai(90, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(90, 2) = 0 Then
               atkingckai(90, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(90, 2) = 1 Then
               atkingckai(90, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�L���S\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 6705
                   atkingno(i, 6) = 10185
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(90, 1) = 3
        Case 3
             atking_AI_�L���S_�j�t���q������(1) = �Y����ˮ`��
             �Y���淾�q�Ȯ��ܼ�(2) = 0
             �Y���淾�q�Ȯ��ܼ�(3) = 0
             '========================================
                For p = 1 To Val(FormMainMode.��ܦC1.goi1)
                   Randomize Timer
                   i = Int(Rnd() * 6) + 1
                   If i = 1 Or i = 6 Then �Y���淾�q�Ȯ��ܼ�(2) = Val(�Y���淾�q�Ȯ��ܼ�(2)) + 1
                Next
                For p = 1 To Val(FormMainMode.��ܦC1.goi2)
                   Randomize Timer
                   j = Int(Rnd() * 6) + 1
                   If j = 1 Or j = 6 Then �Y���淾�q�Ȯ��ܼ�(3) = Val(�Y���淾�q�Ȯ��ܼ�(3)) + 1
                Next
                '=============================
                �ޯ�ʵe��ܶ��q�� = 1
                atkingckai(90, 1) = 4
                FormMainMode.��l���槹�Ұ�.Enabled = False
                �ثe��(22) = 12
                FormMainMode.���ݮɶ�.Enabled = True
          Case 4
                atking_AI_�L���S_�j�t���q������(2) = �Y����ˮ`��
                '==========================
                If atking_AI_�L���S_�j�t���q������(1) > atking_AI_�L���S_�j�t���q������(2) Then
                    �Y���淾�q�Ȯ��ܼ�(2) = atking_AI_�L���S_�j�t���q������(2)
                Else
                    �Y���淾�q�Ȯ��ܼ�(2) = atking_AI_�L���S_�j�t���q������(1)
                End If
                �Y����ˮ`�� = Val(�Y���淾�q�Ȯ��ܼ�(2))
                atkingckai(90, 2) = 0
                Erase atking_AI_�L���S_�j�t���q������
   End Select
End If
End Sub

Sub ���纸_Rud_913()
If FormMainMode.comaiatk(1).Caption = "Rud-913" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(40, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���纸" Then
   Select Case atkingckai(40, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(40, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                   atkingckai(40, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(40, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                   atkingckai(40, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���纸\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6675
                   atkingno(i, 6) = 9105
                   atkingno(i, 7) = 40
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(40, 2) = 0
            '================
            �԰��t����.����ʧ@_�Z���ܧ� 3
   End Select
End If
End Sub
Sub ���纸_Von_541()
If FormMainMode.comaiatk(2).Caption = "Von-541" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(76, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���纸" Then
   Select Case atkingckai(76, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingckai(76, 2) = 0 Then
               �������m��l�`��(2) = �������m��l�`��(2) + 4
               atkingckai(76, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If (atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1) And atkingckai(76, 2) = 1 Then
               �������m��l�`��(2) = �������m��l�`��(2) - 4
               atkingckai(76, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���纸\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7140
                   atkingno(i, 6) = 9645
                   atkingno(i, 7) = 117
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(76, 2) = 0
            '================
            If �Y����ˮ`�� >= 10 Then
                �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� �Y����ˮ`��, 1
                �Y����ˮ`�� = 0
                �Y���淾�q�Ȯ��ܼ�(2) = 0
            End If
   End Select
End If
End Sub

Sub ���纸_Wil_846()
If FormMainMode.comaiatk(4).Caption = "Wil-846" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(41, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���纸" Then
   Select Case atkingckai(41, 1)
      Case 1
           If movecp = 3 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(41, 2) = 0 Then
                   atkingckai(41, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 5) < 2) And atkingckai(41, 2) = 1 Then
                   atkingckai(41, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���纸\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6720
                   atkingno(i, 6) = 10320
                   atkingno(i, 7) = 41
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(41, 2) = 0
            '================
            �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 2, 1
            '================
                For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, j, 3) = 7 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          �H�����`���A��Ʈw(1, j, 1) = 9
                      End If
                      If �H�����`���A��Ʈw(1, j, 3) = 8 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          �H�����`���A��Ʈw(1, j, 1) = 9
                      End If
                      If �H�����`���A��Ʈw(1, j, 3) = 9 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          �H�����`���A��Ʈw(1, j, 1) = 9
                      End If
                      If �H�����`���A��Ʈw(1, j, 3) = 10 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          �H�����`���A��Ʈw(1, j, 1) = 9
                      End If
                      If �H�����`���A��Ʈw(1, j, 3) = 11 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          �H�����`���A��Ʈw(1, j, 1) = 9
                      End If
                      If �H�����`���A��Ʈw(1, j, 3) = 12 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                          FormMainMode.personusspe(j).person_num = 9
                          �H�����`���A��Ʈw(1, j, 1) = 9
                      End If
                 Next
                 '==========================
                For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                    If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        �H�����`���A��Ʈw(2, j, 1) = 9
                    End If
                    If �H�����`���A��Ʈw(2, j, 3) = 2 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        �H�����`���A��Ʈw(2, j, 1) = 9
                    End If
                    If �H�����`���A��Ʈw(2, j, 3) = 3 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        �H�����`���A��Ʈw(2, j, 1) = 9
                    End If
                    If �H�����`���A��Ʈw(2, j, 3) = 4 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        �H�����`���A��Ʈw(2, j, 1) = 9
                    End If
                    If �H�����`���A��Ʈw(2, j, 3) = 5 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        �H�����`���A��Ʈw(2, j, 1) = 9
                    End If
                    If �H�����`���A��Ʈw(2, j, 3) = 6 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                        FormMainMode.personcomspe(j).person_num = 9
                        �H�����`���A��Ʈw(2, j, 1) = 9
                    End If
                Next
   End Select
End If
End Sub
Sub ���纸_Chr_799()
Dim m As Integer
If FormMainMode.comaiatk(3).Caption = "Chr-799" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(77, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���纸" Then
   Select Case atkingckai(77, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(77, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                   atkingckai(77, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 5) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(77, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                   atkingckai(77, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���纸\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 120
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6750
                   atkingno(i, 6) = 9255
                   atkingno(i, 7) = 77
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(77, 2) = 0
            '================
            m = Int(Rnd() * 3) + 1
            Select Case m
                Case 1
                       Do
                            For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                                  If �H�����`���A��Ʈw(1, j, 3) = 10 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 3
                                      FormMainMode.personusspe(j).person_turn = 5
                                      �H�����`���A��Ʈw(1, j, 1) = 3
                                      �H�����`���A��Ʈw(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, j, 10, app_path & "gif\���`���A\atkdown.gif", 3, 5
                                  ���`���A�ˬd��(10, 1) = 1
                                  ���`���A�ˬd��(10, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                                If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 3
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 �H�����`���A��Ʈw(2, j, 1) = 3
                                 �H�����`���A��Ʈw(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                              If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                 �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 3, 5
                                 ���`���A�ˬd��(1, 1) = 1
                                 ���`���A�ˬd��(1, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
                Case 2
                        Do
                            For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                                  If �H�����`���A��Ʈw(1, j, 3) = 11 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 3
                                      FormMainMode.personusspe(j).person_turn = 5
                                      �H�����`���A��Ʈw(1, j, 1) = 3
                                      �H�����`���A��Ʈw(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, j, 11, app_path & "gif\���`���A\defdown.gif", 3, 5
                                  ���`���A�ˬd��(11, 1) = 1
                                  ���`���A�ˬd��(11, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                                If �H�����`���A��Ʈw(2, j, 3) = 2 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 3
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 �H�����`���A��Ʈw(2, j, 1) = 3
                                 �H�����`���A��Ʈw(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                              If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                 �԰��t����.�H�����`���A��]�w_��] 2, j, 2, app_path & "gif\���`���A\defup.gif", 3, 5
                                 ���`���A�ˬd��(2, 1) = 1
                                 ���`���A�ˬd��(2, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
                Case 3
                        Do
                            For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                                  If �H�����`���A��Ʈw(1, j, 3) = 12 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                                      FormMainMode.personusspe(j).person_num = 1
                                      FormMainMode.personusspe(j).person_turn = 5
                                      �H�����`���A��Ʈw(1, j, 1) = 1
                                      �H�����`���A��Ʈw(1, j, 2) = 5
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, j, 12, app_path & "gif\���`���A\movdown.gif", 1, 5
                                  ���`���A�ˬd��(12, 1) = 1
                                  ���`���A�ˬd��(12, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        '==========================
                        Do
                            For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                                If �H�����`���A��Ʈw(2, j, 3) = 3 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                 FormMainMode.personcomspe(j).person_num = 1
                                 FormMainMode.personcomspe(j).person_turn = 5
                                 �H�����`���A��Ʈw(2, j, 1) = 1
                                 �H�����`���A��Ʈw(2, j, 2) = 5
                                 Exit Do
                                End If
                            Next
                           For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                              If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                 �԰��t����.�H�����`���A��]�w_��] 2, j, 3, app_path & "gif\���`���A\movup.gif", 1, 5
                                 ���`���A�ˬd��(3, 1) = 1
                                 ���`���A�ˬd��(3, 2) = 1
                                 Exit Do
                             End If
                           Next
                        Loop
            End Select
   End Select
End If
End Sub
Sub ������S_���()
Dim m As Integer
If FormMainMode.comaiatk(1).Caption = "���" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(78, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������S" Then
   Select Case atkingckai(78, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 4) >= 1 And atkingckai(78, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 3
                   atkingckai(78, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 1 And atkingckai(78, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 3
                   atkingckai(78, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������S\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -240
                   atkingno(i, 5) = 6195
                   atkingno(i, 6) = 10350
                   atkingno(i, 7) = 78
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
             Erase atking_AI_������S_���������
             '========================
             For i = 1 To 106
                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                    If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                         atking_AI_������S_���������(i) = 1
                         atking_AI_������S_���������(107) = atking_AI_������S_���������(107) + 1
                     End If
                End If
            Next
            If atking_AI_������S_���������(107) > 2 Then
                atking_AI_������S_���������(107) = 2
            End If
            '=========================
            If atking_AI_������S_���������(107) > 0 Then
                Do
                    m = Int(Rnd() * 106) + 1
                    If atking_AI_������S_���������(m) = 1 Then
                        �ثe��(20) = m
                        �ثe��(21) = 5
                        atking_AI_������S_���������(m) = 0
                        atking_AI_������S_���������(0) = atking_AI_������S_���������(0) + 1
                        FormMainMode.tr�ϥΪ�_��P.Enabled = True
                        Exit Sub
                    End If
                Loop
            Else
               �ثe��(22) = 25
               FormMainMode.���ݮɶ�.Enabled = True
            End If
        Case 4
            If atking_AI_������S_���������(107) > 1 And atking_AI_������S_���������(0) < atking_AI_������S_���������(107) Then
                Do
                    m = Int(Rnd() * 106) + 1
                    If atking_AI_������S_���������(m) = 1 Then
                        �ثe��(20) = m
                        �ثe��(21) = 5
                        atking_AI_������S_���������(m) = 0
                        atking_AI_������S_���������(0) = atking_AI_������S_���������(0) + 1
                        FormMainMode.tr�ϥΪ�_��P.Enabled = True
                        Exit Sub
                    End If
                Loop
            ElseIf atking_AI_������S_���������(0) >= 2 Then
               �ثe��(24) = 26
               FormMainMode.���ݮɶ�_2.Enabled = True
            Else
               �ثe��(24) = 25
               FormMainMode.���ݮɶ�_2.Enabled = True
            End If
        Case 5
            If atking_AI_������S_���������(107) = 0 Then
                atking_AI_������S_���������(107) = 99
               �ثe��(22) = 25
               FormMainMode.���ݮɶ�.Enabled = True
            ElseIf atking_AI_������S_���������(107) > 0 And atking_AI_������S_���������(0) = 0 Then
               atkingckai(78, 2) = 0
               �԰��t����.����ʧ@_�ޯ��ʵ���
            ElseIf atking_AI_������S_���������(107) > 0 And atking_AI_������S_���������(0) = 1 Then
               �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� atking_AI_������S_���������(0), 1
               �ثe��(24) = 26
               FormMainMode.���ݮɶ�_2.Enabled = True
            End If
        Case 6
            If atking_AI_������S_���������(107) > 0 And atking_AI_������S_���������(0) = 1 Then
               atkingckai(78, 2) = 0
               �԰��t����.����ʧ@_�ޯ��ʵ���
            ElseIf atking_AI_������S_���������(0) >= 2 Then
               �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� atking_AI_������S_���������(0), 1
               atkingckai(78, 2) = 0
               �԰��t����.����ʧ@_�ޯ��ʵ���
            End If
   End Select
End If
End Sub

Sub ������S_�鱫()
If FormMainMode.comaiatk(2).Caption = "�鱫" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(42, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������S" Then
   Select Case atkingckai(42, 1)
        Case 1
            If movecp = 1 Then
             If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 3) >= 1 And atkingckai(42, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(42, 2) = 0 Then
               atkingckai(42, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 5
            ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 3) < 1) And atkingckai(42, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(42, 2) = 1 Then
               atkingckai(42, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 5
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������S\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5580
                   atkingno(i, 6) = 9465
                   atkingno(i, 7) = 126
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(42, 2) = 0
            '===============
            If �Y����ˮ`�� <= 0 Then
                Do
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 3) = 16 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 2
                          �H�����`���A��Ʈw(1, i, 2) = 2
                          Exit Do
                      End If
                    Next
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                       If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 1, i, 16, app_path & "gif\���`���A\moveerr.gif", 0, 2
                          ���`���A�ˬd��(16, 1) = 1
                          ���`���A�ˬd��(16, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub ������S_�a���y���~()
Dim m As Integer
If FormMainMode.comaiatk(4).Caption = "�a���y���~" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(43, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������S" Then
   Select Case atkingckai(43, 1)
        Case 1
             If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(43, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(43, 2) = 0 Then
               atkingckai(43, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(43, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(43, 2) = 1 Then
               atkingckai(43, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������S\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6630
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 43
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(43, 2) = 0
            '===============
            m = (atkingpagetot(2, 1) + atkingpagetot(2, 5)) \ 5
            �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� m, 1
   End Select
End If
End Sub
Sub �w�ǥ���_�F�z���������¼�()
If FormMainMode.comaiatk(1).Caption = "�F�z���������¼�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(44, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�w�ǥ���" Then
   Select Case atkingckai(44, 1)
      Case 1
           If movecp = 3 Then
                �������m��l�`��(2) = �������m��l�`��(2) + 2
                atkingckai(44, 2) = 1
                atkingtrn(1) = Val(atkingtrn(1)) + 1
          End If
      Case 2
             atkingckai(44, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�w�ǥ���\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 960
                   atkingno(i, 4) = 1560
                   atkingno(i, 5) = 6270
                   atkingno(i, 6) = 9645
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub �w�ǥ���_�ƨg����()
If FormMainMode.comaiatk(2).Caption = "�ƨg����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(79, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�w�ǥ���" Then
   Select Case atkingckai(79, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 1 And atkingckai(79, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 2
                   atkingckai(79, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 1 And atkingckai(79, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 2
                   atkingckai(79, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�w�ǥ���\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -720
                   atkingno(i, 5) = 8505
                   atkingno(i, 6) = 10140
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(FormMainMode.pageusglead.Caption) > 0 Then
                 atking_AI_�w�ǥ���_�ƨg���۬����� = 1
                 '==========================
                  Do Until atking_AI_�w�ǥ���_�ƨg���۬����� > 3 Or Val(FormMainMode.pageusglead.Caption) <= 0
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                            �ثe��(21) = 6
                            �ثe��(20) = m
                            atking_AI_�w�ǥ���_�ƨg���۬����� = atking_AI_�w�ǥ���_�ƨg���۬����� + 1
                            FormMainMode.tr�ϥΪ�_��P.Enabled = True
                            Exit Sub
                        End If
                   Loop
             Else
                 atkingckai(79, 1) = 5
                 FormMainMode.��l���槹�Ұ�.Enabled = True
             End If
        Case 4
             Do Until atking_AI_�w�ǥ���_�ƨg���۬����� > 3 Or Val(FormMainMode.pageusglead.Caption) <= 0
                 Randomize
                 m = Int(Rnd() * 106) + 1
                 If Val(pagecardnum(m, 5)) = 1 And Val(pagecardnum(m, 6)) = 1 Then
                     �ثe��(21) = 6
                     �ثe��(20) = m
                     atking_AI_�w�ǥ���_�ƨg���۬����� = atking_AI_�w�ǥ���_�ƨg���۬����� + 1
                     FormMainMode.tr�ϥΪ�_��P.Enabled = True
                     Exit Sub
                 End If
            Loop
            If atking_AI_�w�ǥ���_�ƨg���۬����� > 3 Or Val(FormMainMode.pageusglead.Caption) <= 0 Then
                atkingckai(79, 1) = 5
                �ثe��(24) = 22
                FormMainMode.���ݮɶ�_2.Enabled = True
            End If
        Case 5
            atkingckai(79, 2) = 0
   End Select
End If
End Sub

Sub �w�ǥ���_�·t�x��()
Dim m As Integer
If FormMainMode.comaiatk(4).Caption = "�·t�x��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(46, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�w�ǥ���" Then
   Select Case atkingckai(46, 1)
        Case 1
             If atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(46, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(46, 2) = 0 Then
               atkingckai(46, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 3
            ElseIf (atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(46, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(46, 2) = 1 Then
               atkingckai(46, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 3
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�w�ǥ���\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6480
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 46
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(46, 2) = 0
            '===============
            m = movecp + 1
            If m > 3 Then m = 3
            �԰��t����.����ʧ@_�Z���ܧ� m
   End Select
End If
End Sub
Sub �w�ǥ���_�`�W()
Dim m As Integer
If FormMainMode.comaiatk(3).Caption = "�`�W" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(45, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�w�ǥ���" Then
   Select Case atkingckai(45, 1)
        Case 1
             If atkingpagetot(2, 4) >= 3 And atkingckai(45, 2) = 0 Then
'             If atkingpagetot(2, 2) >= 1 And atkingckai(45, 2) = 0 Then
               atkingckai(45, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 4) < 3 And atkingckai(45, 2) = 1 Then
'            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(45, 2) = 1 Then
               atkingckai(45, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�w�ǥ���\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8970
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 45
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(45, 2) = 0
            '===============
            m = Int(atkingpagetot(2, 4) / 2 + 0.9)
            �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� m, 1
   End Select
End If
End Sub
Sub ����P��_CTL()
Dim i As Integer
If FormMainMode.comaiatk(1).Caption = "C.T.L" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(80, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����P��" Then
   Select Case atkingckai(80, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(80, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                   atkingckai(80, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   For i = 1 To 106
                       If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                              �������m��l�`��(2) = �������m��l�`��(2) + 6
                              �ثe��(28) = 1
                              Exit For
                          End If
                       End If
                   Next
                End If
                If (atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) And atkingckai(80, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                   If �ثe��(28) = 1 Then
                       �������m��l�`��(2) = �������m��l�`��(2) - 6
                       �ثe��(28) = 0
                   End If
                   atkingckai(80, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             atkingckai(80, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����P��\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6540
                   atkingno(i, 6) = 9990
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\����P��\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             If �ثe��(28) = 1 Then
                 For i = 1 To 106
                       If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                          If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                              Exit For
                          End If
                       End If
                  Next
                  If i = 107 Then
                      �������m��l�`��(2) = �������m��l�`��(2) - 6
                      �����g�J��ܦC�ƭ� 2, Val(FormMainMode.��ܦC1.goi2) - 6
                      For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, i, 3) = 32 Then
                              �������m��l�`��(2) = �������m��l�`��(2) - 6
                              �����g�J��ܦC�ƭ� 2, Val(FormMainMode.��ܦC1.goi2) - 6
                          End If
                      Next
                  End If
                  �ثe��(28) = 0
             End If
   End Select
End If
End Sub
Sub ����P��_BPA()
If FormMainMode.comaiatk(2).Caption = "B.P.A" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(81, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����P��" Then
   Select Case atkingckai(81, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(81, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 3
                   atkingckai(81, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(81, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 3
                   atkingckai(81, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����P��\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6030
                   atkingno(i, 6) = 10530
                   atkingno(i, 7) = 81
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� pageqlead(1), 1
             atkingckai(81, 2) = 0
   End Select
End If
End Sub

Sub ����P��_LAR()
If FormMainMode.comaiatk(3).Caption = "L.A.R" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(47, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����P��" Then
   Select Case atkingckai(47, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 2) >= 2 And atkingckai(47, 2) = 0 Then
                   atkingckai(47, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 2) < 2 And atkingckai(47, 2) = 1 Then
                   atkingckai(47, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����P��\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5400
                   atkingno(i, 6) = 9015
                   atkingno(i, 7) = 47
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             �԰��t����.�^�_����_�q�� 1, 1
        Case 4
             atkingckai(47, 2) = 0
             If �Y����ˮ`�� > 0 Then
                 �԰��t����.�^�_����_�q�� 1, 1
             End If
   End Select
End If
End Sub
Sub �Ǧh_�]�G����()
Dim p, i, j As Integer
Dim ak As Integer
If FormMainMode.comaiatk(4).Caption = "�]�G����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(48, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Ǧh" Then
   Select Case atkingckai(48, 1)
      Case 1
            If atkingpagetot(2, 3) >= 1 And atkingckai(48, 2) = 0 Then
'            If pageqlead(2) >= 1 And atkingckai(48, 2) = 0 Then
               atkingckai(48, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 4
            End If
            If atkingpagetot(2, 3) < 1 And atkingckai(48, 2) = 1 Then
'            If pageqlead(2) < 1 And atkingckai(48, 2) = 1 Then
               atkingckai(48, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 4
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǧh\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7590
                   atkingno(i, 6) = 9420
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(48, 1) = 3
        Case 3
             atking_AI_�Ǧh_�]�G���ۻ�q������(1) = �Y����ˮ`��
             �Y���淾�q�Ȯ��ܼ�(2) = 0
             �Y���淾�q�Ȯ��ܼ�(3) = 0
             '========================================
                For p = 1 To Val(FormMainMode.��ܦC1.goi1)
                   Randomize Timer
                   i = Int(Rnd() * 6) + 1
                   If i = 1 Or i = 6 Then �Y���淾�q�Ȯ��ܼ�(2) = Val(�Y���淾�q�Ȯ��ܼ�(2)) + 1
                Next
                For p = 1 To Val(FormMainMode.��ܦC1.goi2)
                   Randomize Timer
                   j = Int(Rnd() * 6) + 1
                   If j = 1 Or j = 6 Then �Y���淾�q�Ȯ��ܼ�(3) = Val(�Y���淾�q�Ȯ��ܼ�(3)) + 1
                Next
                '=============================
                �ޯ�ʵe��ܶ��q�� = 1
                atkingckai(48, 1) = 4
                FormMainMode.��l���槹�Ұ�.Enabled = False
                �ثe��(22) = 12
                FormMainMode.���ݮɶ�.Enabled = True
          Case 4
                atking_AI_�Ǧh_�]�G���ۻ�q������(2) = �Y����ˮ`��
                '==========================
                If atking_AI_�Ǧh_�]�G���ۻ�q������(1) < atking_AI_�Ǧh_�]�G���ۻ�q������(2) Then
                    �Y���淾�q�Ȯ��ܼ�(2) = atking_AI_�Ǧh_�]�G���ۻ�q������(2)
                Else
                    �Y���淾�q�Ȯ��ܼ�(2) = atking_AI_�Ǧh_�]�G���ۻ�q������(1)
                End If
                �Y����ˮ`�� = Val(�Y���淾�q�Ȯ��ܼ�(2))
                atkingckai(48, 2) = 0
                Erase atking_AI_�Ǧh_�]�G���ۻ�q������
          Case 5
                 atkingckai(48, 1) = 1
                  For j = 49 To 54   '��1��1�d�u��
                      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                              �԰��t����.comatk_AI_����_���b�B_�� j
                              ak = 1
                              Exit For
                           ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                              �԰��t����.comatk_AI_����_���b�B_�� j
                              ak = 1
                              Exit For
                           End If
                      End If
                  Next
                  If ak = 0 Then
                     For j = 39 To 44   '�j1��1�d�䦸�u��
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
                              �԰��t����.comatk_AI_�Ǧh_�]�G����_�� j
                              ak = 1
                              Exit For
                           ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
                              �԰��t����.comatk_AI_�Ǧh_�]�G����_�� j
                              ak = 1
                              Exit For
                           End If
                        End If
                     Next
                  End If
                  If ak = 0 Then
                     For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) >= 1 Then
                              �԰��t����.comatk_AI_�Ǧh_�]�G����_�� j
                              ak = 1
                              Exit For
                           ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) >= 1 Then
                              �԰��t����.comatk_AI_�Ǧh_�]�G����_�� j
                              ak = 1
                              Exit For
                           End If
                        End If
                     Next
                  End If
   End Select
End If
End Sub
Sub ��ܵY_��������()
Dim bloodtot As Integer '�Ȯ��ܼ�
Dim num(1 To 2) As Integer '��ܤH���Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "��������" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(51, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��ܵY" Then
   Select Case atkingckai(51, 1)
        Case 1
            If movecp < 3 Then
                 If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 5) >= 2 And atkingpagetot(2, 4) >= 1 And atkingckai(51, 2) = 0 Then
    '             If atkingpagetot(2, 1) >= 1 And atkingckai(51, 2) = 0 Then
                   atkingckai(51, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   �������m��l�`��(2) = �������m��l�`��(2) + 13
                ElseIf (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 5) < 2 Or atkingpagetot(2, 4) < 1) And atkingckai(51, 2) = 1 Then
    '            ElseIf atkingpagetot(2, 1) < 1 And atkingckai(51, 2) = 1 Then
                   atkingckai(51, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   �������m��l�`��(2) = �������m��l�`��(2) - 13
                End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��ܵY\Evelynatking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 8250
                   atkingno(i, 6) = 10275
                   atkingno(i, 7) = 51
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\��ܵY\Evelynatking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            Do
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, i, 2) >= 9 And �H�����`���A��Ʈw(2, i, 3) = 25 Then
                    Exit Do
                 End If
                 If �H�����`���A��Ʈw(2, i, 3) = 25 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) < 9 Then
                     FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2) + 1
                     �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                  If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                     �԰��t����.�H�����`���A��]�w_��] 2, i, 25, app_path & "gif\���`���A\��O�C�U.gif", 0, 1
                     ���`���A�ˬd��(25, 1) = 1
                     ���`���A�ˬd��(25, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
        Case 4
            bloodtot = Val(FormMainMode.��ܦC1.goi2) \ 10
            num(2) = 999
            For i = 1 To 3
               If liveus(����ݾ��H��������(1, i)) < num(2) And liveus(����ݾ��H��������(1, i)) > 0 Then
                   num(1) = i
                   num(2) = liveus(����ݾ��H��������(1, i))
               End If
            Next
            �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� bloodtot, num(1)
            atkingckai(51, 2) = 0
   End Select
End If
End Sub
Sub �h�g�H_�ߦ���()
If FormMainMode.comaiatk(4).Caption = "�ߦ���" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(52, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�h�g�H" Then
   Select Case atkingckai(52, 1)
      Case 1
            If movecp = 1 Then
                    If atkingpagetot(2, 1) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(52, 2) = 0 Then
                       atkingckai(52, 2) = 1
                       atkingtrn(2) = Val(atkingtrn(2)) + 1
                       �������m��l�`��(2) = �������m��l�`��(2) + 8
                    End If
                    If (atkingpagetot(2, 1) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(52, 2) = 1 Then
                       atkingckai(52, 2) = 0
                       atkingtrn(2) = Val(atkingtrn(2)) - 1
                       �������m��l�`��(2) = �������m��l�`��(2) - 8
                     End If
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�h�g�H\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6450
                   atkingno(i, 6) = 10200
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(52, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                Do
                    For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                          If �H�����`���A��Ʈw(1, j, 3) = 15 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                              FormMainMode.personusspe(j).person_num = 0
                              FormMainMode.personusspe(j).person_turn = 5
                              �H�����`���A��Ʈw(1, j, 1) = 0
                              �H�����`���A��Ʈw(1, j, 2) = 5
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                       If �H�����`���A��Ʈw(1, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 1, j, 15, app_path & "gif\���`���A\���a.gif", 0, 5
                          ���`���A�ˬd��(15, 1) = 1
                          ���`���A�ˬd��(15, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub �h�g�H_�ݭh�ɦV()
If FormMainMode.comaiatk(1).Caption = "�ݭh�ɦV" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(53, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�h�g�H" Then
   Select Case atkingckai(53, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(53, 2) = 0 Then
               atkingckai(53, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(53, 2) = 1 Then
               atkingckai(53, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�h�g�H\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8160
                   atkingno(i, 6) = 9120
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(53, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                    Case 1
                       Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 20 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  �H�����`���A��Ʈw(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 20, app_path & "gif\���`���A\damage.gif", 0, 2
                                  ���`���A�ˬd��(20, 1) = 1
                                  ���`���A�ˬd��(20, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 16 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  �H�����`���A��Ʈw(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 16, app_path & "gif\���`���A\moveerr.gif", 0, 2
                                  ���`���A�ˬd��(16, 1) = 1
                                  ���`���A�ˬd��(16, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 22 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 2
                                  �H�����`���A��Ʈw(1, i, 2) = 2
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 22, app_path & "gif\���`���A\atkingerr.gif", 0, 2
                                  ���`���A�ˬd��(22, 1) = 1
                                  ���`���A�ˬd��(22, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
            End If
   End Select
End If
End Sub
Sub �h�g�H_�����()
Dim atkingtotai As Integer '�S�ƶq�Ȯɲέp�ܼ�
Dim a As Integer, i As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(2).Caption = "�����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(82, 2) = 1) _
    And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�h�g�H" Then
 Select Case atkingckai(82, 1)
   Case 1
      atkingckai(82, 1) = 2
      For i = 55 To 106
         If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And ((Val(pagecardnum(i, 2)) = 3 And pagecardnum(i, 1) = a4a) Or (Val(pagecardnum(i, 4)) = 3 And pagecardnum(i, 3) = a4a)) Then
            atkingtotai = Val(atkingtotai) + 1
         End If
      Next
      If atkingtotai >= 1 Then
         Select Case livecom(����H����ԤH��(2, 2))
            Case Is < 3
                If Val(FormMainMode.��ܦC1.goi1) - Val(FormMainMode.��ܦC1.goi2) >= livecom(����H����ԤH��(2, 2)) Then
                    GoTo AI�ޯ�_�h�g�H_�����_�X�P���q�G
                End If
            Case 3
                If Val(FormMainMode.��ܦC1.goi1) - Val(FormMainMode.��ܦC1.goi2) >= 9 Then
                    GoTo AI�ޯ�_�h�g�H_�����_�X�P���q�G
                End If
            Case Is > 3
                If Int(Val(FormMainMode.��ܦC1.goi1) / 3 + 0.9) - Int(Val(FormMainMode.��ܦC1.goi2) / 3 + 0.9) >= livecom(����H����ԤH��(2, 2)) Then
                    GoTo AI�ޯ�_�h�g�H_�����_�X�P���q�G
                End If
         End Select
      End If
      '==========�p�G���ŦX��������
      Exit Sub
    '================================
AI�ޯ�_�h�g�H_�����_�X�P���q�G:
      For a = 55 To 106
            If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
                If pagecardnum(a, 1) = a4a And pagecardnum(a, 2) = 3 Then
                    �԰��t����.comatk_AI_����_�h�g�H_�����_�S a
                    Exit For
                ElseIf pagecardnum(a, 3) = a4a And pagecardnum(a, 4) = 3 Then
                    �԰��t����.comatk_AI_����_�h�g�H_�����_�S a
                    Exit For
                End If
             End If
      Next
    Case 2
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
'                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) >= 1 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
             If rrr >= 1 And atkingckai(82, 2) = 0 Then
                atkingckai(82, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             End If
             If rrr < 1 And atkingckai(82, 2) = 1 Then
                atkingckai(82, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
   Case 3
        For i = �H���ޯ�Ʀr���� To 1 Step -1
             If atkingno(i, 1) = "" Then
                atkingno(i, 1) = app_path & "gif\�h�g�H\atking2_2.jpg"
                atkingno(i, 2) = 2
                atkingno(i, 3) = 0
                atkingno(i, 4) = 0
                atkingno(i, 5) = 6585
                atkingno(i, 6) = 10110
                atkingno(i, 7) = 82
                atkingno(i, 8) = 0
                atkingno(i, 9) = 0
                atkingno(i, 10) = 0
                atkingno(i, 11) = 0
                Exit For
             End If
        Next
    Case 4
          atkingckai(82, 2) = 0
          If Val(�Y���淾�q�Ȯ��ܼ�(2)) >= livecom(����H����ԤH��(2, 2)) And ���`���A�ˬd��(18, 2) = 0 Then
                Do
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 6
                              FormMainMode.personcomspe(j).person_turn = 3
                              �H�����`���A��Ʈw(2, j, 1) = 6
                              �H�����`���A��Ʈw(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 6, 3
                          ���`���A�ˬd��(1, 1) = 1
                          ���`���A�ˬd��(1, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '==================================
                Do
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, j, 3) = 18 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              �H�����`���A��Ʈw(2, j, 1) = 0
                              �H�����`���A��Ʈw(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 18, app_path & "gif\���`���A\����.gif", 0, 3
                          ���`���A�ˬd��(18, 1) = 1
                          ���`���A�ˬd��(18, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '===============================
                Do
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, j, 3) = 19 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 0
                              FormMainMode.personcomspe(j).person_turn = 3
                              �H�����`���A��Ʈw(2, j, 1) = 0
                              �H�����`���A��Ʈw(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 19, app_path & "gif\���`���A\���a.gif", 0, 3
                          ���`���A�ˬd��(19, 1) = 1
                          ���`���A�ˬd��(19, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
         End If
   End Select
End If
End Sub

Sub ��_�󫵦�_�[�ʯP���u�@()
Dim i As Integer, j As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(2).Caption = "�󫵦�-�[�ʯP���u�@" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(54, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
   Select Case atkingckai(54, 1)
        Case 1
            If atkingpagetot(2, 4) >= 2 And atkingpagetot(2, 3) >= 1 And atkingckai(54, 2) = 0 Then
'            If atkingpagetot(2, 3) >= 1 And atkingckai(54, 2) = 0 Then
               atkingckai(54, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 4) < 2 Or atkingpagetot(2, 3) < 1) And atkingckai(54, 2) = 1 Then
'            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(54, 2) = 1 Then
               atkingckai(54, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��\��-�󫵦�-�[�ʯP���u�@_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6180
                   atkingno(i, 6) = 9000
                   atkingno(i, 7) = 11
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
          Do
            atkingckai(54, 2) = 0
            For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, j, 1) >= 10 And �H�����`���A��Ʈw(2, j, 3) = 2 Then
                     FormMainMode.personcomspe(j).person_turn = 3
                     �H�����`���A��Ʈw(2, j, 2) = 3
                     Exit Do
                 End If
                 If �H�����`���A��Ʈw(2, j, 3) = 2 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                     FormMainMode.personcomspe(j).person_num = �H�����`���A��Ʈw(2, j, 1) + 1
                     FormMainMode.personcomspe(j).person_turn = 3
                     �H�����`���A��Ʈw(2, j, 1) = �H�����`���A��Ʈw(2, j, 1) + 1
                     �H�����`���A��Ʈw(2, j, 2) = 3
                     '========DEF+1�ߧY�ͮ�
'                         �������m��l�`��(2) = �������m��l�`��(2) + 1
                         �԰��t����.�����g�J��ܦC�ƭ� 2, Val(FormMainMode.��ܦC1.goi2) + 1
                    '===============
                     Exit Do
                 End If
            Next
           For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
              If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                 �԰��t����.�H�����`���A��]�w_��] 2, j, 2, app_path & "gif\���`���A\defup.gif", 3, 3
                 ���`���A�ˬd��(2, 1) = 1
                 ���`���A�ˬd��(2, 2) = 1
                  '========DEF+3�ߧY�ͮ�
'                         �������m��l�`��(2) = �������m��l�`��(2) + 3
                         �԰��t����.�����g�J��ܦC�ƭ� 2, Val(FormMainMode.��ܦC1.goi2) + 3
                  '===============
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub ���_�i���h_�j�a�Y�a()
If FormMainMode.comaiatk(1).Caption = "�j�a�Y�a" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(89, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���_�i���h" Then
   Select Case atkingckai(89, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(89, 2) = 0 Then
                   atkingckai(89, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(89, 2) = 1 Then
                   atkingckai(89, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���_�i���h\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8055
                   atkingno(i, 6) = 10620
                   atkingno(i, 7) = 89
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\���_�i���h\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(89, 2) = 0
             '=================
             �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 2, 1
   End Select
End If
End Sub

Sub ���_�i���h_�P�R�j��()
Dim rrr(1 To 3) As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(2).Caption = "�P�R�j��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(83, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���_�i���h" Then
   Select Case atkingckai(83, 1)
      Case 1
            If movecp > 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                        If pagecardnum(i, 1) = a5a And pagecardnum(i, 2) = 1 Then
                           rrr(1) = rrr(1) + 1
                        End If
                        If pagecardnum(i, 1) = a5a And pagecardnum(i, 2) = 2 Then
                           rrr(2) = rrr(2) + 1
                        End If
                        If pagecardnum(i, 1) = a5a And pagecardnum(i, 2) = 3 Then
                           rrr(3) = rrr(3) + 1
                        End If
                    End If
                 Next
            End If
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(83, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(83, 2) = 0 Then
                atkingckai(83, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
                �������m��l�`��(2) = �������m��l�`��(2) + 9
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(83, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(83, 2) = 1 Then
                atkingckai(83, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
                �������m��l�`��(2) = �������m��l�`��(2) - 9
              End If
      Case 2
             atkingckai(83, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���_�i���h\atking2-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8520
                   atkingno(i, 6) = 8280
                   atkingno(i, 7) = 150
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\���_�i���h\atking2-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub ���_�i���h_�T�v����()
If FormMainMode.comaiatk(3).Caption = "�T�v����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(84, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���_�i���h" Then
   Select Case atkingckai(84, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(84, 2) = 0 Then
               atkingckai(84, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 1 And atkingckai(84, 2) = 1 Then
               atkingckai(84, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���_�i���h\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6255
                   atkingno(i, 6) = 10395
                   atkingno(i, 7) = 151
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(84, 2) = 0
             '======================
             If �Y����ˮ`�� > 0 Then
               Do
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 3) = 16 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 2
                          �H�����`���A��Ʈw(1, i, 2) = 2
                          Exit Do
                      End If
                    Next
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                       If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 1, i, 16, app_path & "gif\���`���A\moveerr.gif", 0, 2
                          ���`���A�ˬd��(16, 1) = 1
                          ���`���A�ˬd��(16, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub

Sub ���_�i���h_���@�g��()
If FormMainMode.comaiatk(4).Caption = "���@�g��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(49, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���_�i���h" Then
   Select Case atkingckai(49, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + (atkingpagetot(2, 5) - atking_AI_���_�i���h_���@�g��_�j�ƭȬ�����)
                   atking_AI_���_�i���h_���@�g��_�j�ƭȬ����� = atkingpagetot(2, 5)
                   If atkingckai(49, 2) = 0 Then
                        atkingckai(49, 2) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) + 1
                        �������m��l�`��(2) = �������m��l�`��(2) + 2
                   End If
                End If
                If atkingpagetot(2, 5) < 1 And atkingckai(49, 2) = 1 Then
                   atkingckai(49, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   �������m��l�`��(2) = �������m��l�`��(2) - atking_AI_���_�i���h_���@�g��_�j�ƭȬ����� - 2
                   atking_AI_���_�i���h_���@�g��_�j�ƭȬ����� = 0
                 End If
          End If
      Case 2
             atkingckai(49, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���_�i���h\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7725
                   atkingno(i, 6) = 9345
                   atkingno(i, 7) = 49
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atking_AI_���_�i���h_���@�g��_�j�ƭȬ����� = 0
   End Select
End If
End Sub
Sub ��ܵY_�E�����q()
If FormMainMode.comaiatk(2).Caption = "�E�����q" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(61, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��ܵY" Then
   Select Case atkingckai(61, 1)
        Case 1
            If movecp > 1 Then
                If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 1 And atkingckai(61, 2) = 0 Then
    '             If atkingpagetot(2, 2) >= 1 And atkingckai(61, 2) = 0 Then
                   atkingckai(61, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 1) And atkingckai(61, 2) = 1 Then
    '            ElseIf atkingpagetot(2, 2) < 1 And atkingckai(61, 2) = 1 Then
                   atkingckai(61, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��ܵY\Evelynatking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6195
                   atkingno(i, 6) = 8730
                   atkingno(i, 7) = 61
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===============
             �԰��t����.�����g�J��ܦC�ƭ� 1, Val(FormMainMode.��ܦC1.goi1) \ 2
'             �������m��l�`��(1) = FormMainMode.��ܦC1.goi1
        Case 3
            atkingckai(61, 2) = 0
            Do
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, i, 2) >= 9 And �H�����`���A��Ʈw(2, i, 3) = 25 Then
                    Exit Do
                 End If
                 If �H�����`���A��Ʈw(2, i, 3) = 25 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) < 9 Then
                     FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2) + 1
                     �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                  If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                     �԰��t����.�H�����`���A��]�w_��] 2, i, 25, app_path & "gif\���`���A\��O�C�U.gif", 0, 1
                     ���`���A�ˬd��(25, 1) = 1
                     ���`���A�ˬd��(25, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
   End Select
End If
End Sub
Sub �Q��_�T�v����()
If FormMainMode.comaiatk(1).Caption = "�T�v����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(72, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Q��" Then
   Select Case atkingckai(72, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(72, 2) = 0 Then
               atkingckai(72, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 1 And atkingckai(72, 2) = 1 Then
               atkingckai(72, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Q��\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9705
                   atkingno(i, 6) = 9090
                   atkingno(i, 7) = 90
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(72, 2) = 0
             '======================
             If �Y����ˮ`�� > 0 Then
                    Do
                         For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                           If �H�����`���A��Ʈw(1, i, 3) = 16 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                               FormMainMode.personusspe(i).person_turn = 2
                               �H�����`���A��Ʈw(1, i, 2) = 2
                               Exit Do
                           End If
                         Next
                         For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                            If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                               �԰��t����.�H�����`���A��]�w_��] 1, i, 16, app_path & "gif\���`���A\moveerr.gif", 0, 2
                               ���`���A�ˬd��(16, 1) = 1
                               ���`���A�ˬd��(16, 2) = 1
                               Exit Do
                            End If
                         Next
                     Loop
             End If
   End Select
End If
End Sub
Sub �Q��_�r��()
If FormMainMode.comaiatk(2).Caption = "�r��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(73, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Q��" Then
   Select Case atkingckai(73, 1)
      Case 1
            If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 4) >= 3 And atkingckai(73, 2) = 0 Then
                   atkingckai(73, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   �������m��l�`��(2) = �������m��l�`��(2) + 5
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 4) < 3) And atkingckai(73, 2) = 1 Then
                   atkingckai(73, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   �������m��l�`��(2) = �������m��l�`��(2) - 5
                 End If
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Q��\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6090
                   atkingno(i, 6) = 10125
                   atkingno(i, 7) = 91
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(73, 2) = 0
             '======================
             If �Y����ˮ`�� > 0 Then
               Do
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 3) = 20 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 3
                          �H�����`���A��Ʈw(1, i, 2) = 3
                          Exit Do
                      End If
                    Next
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                       If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 1, i, 20, app_path & "gif\���`���A\damage.gif", 0, 3
                          ���`���A�ˬd��(20, 1) = 1
                          ���`���A�ˬd��(20, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub �Q��_�������T��()
If FormMainMode.comaiatk(3).Caption = "�������T��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(74, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Q��" Then
   Select Case atkingckai(74, 1)
        Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(74, 2) = 0 Then
               atkingckai(74, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 4) < 1 And atkingckai(74, 2) = 1 Then
               atkingckai(74, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Q��\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -840
                   atkingno(i, 5) = 8250
                   atkingno(i, 6) = 10155
                   atkingno(i, 7) = 92
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If �Y����ˮ`�� > 0 And livecom(����H����ԤH��(2, 2)) > 0 Then
                atking_AI_�Q��_�������T�Ϭ�����(1) = �Y����ˮ`�� + 1
                If Val(FormMainMode.pageul.Caption) < atking_AI_�Q��_�������T�Ϭ�����(1) And atking_AI_�Q��_�������T�Ϭ�����(2) = 0 Then
                   �԰��t����.����ʧ@_�~�P
                End If
                atking_AI_�Q��_�������T�Ϭ�����(2) = atking_AI_�Q��_�������T�Ϭ�����(2) + 1
                If Val(FormMainMode.pageul.Caption) > 0 Then
                    Do Until atking_AI_�Q��_�������T�Ϭ�����(2) > atking_AI_�Q��_�������T�Ϭ�����(1)
                        �ثe��(15) = 25
                        FormMainMode.tr�P��_��P_�q��.Enabled = True
                        Exit Sub
                    Loop
                End If
            End If
            If atking_AI_�Q��_�������T�Ϭ�����(2) > atking_AI_�Q��_�������T�Ϭ�����(1) Or �Y����ˮ`�� <= 0 _
                Or Val(FormMainMode.pageul.Caption) <= 0 Or livecom(����H����ԤH��(2, 2)) <= 0 Then
                �ثe��(24) = 22
                atkingckai(74, 1) = 4
                FormMainMode.���ݮɶ�_2.Enabled = True
            End If
        Case 4
            atkingckai(74, 2) = 0
            Erase atking_AI_�Q��_�������T�Ϭ�����
   End Select
End If
End Sub
Sub �Q��_�I��()
If FormMainMode.comaiatk(4).Caption = "�I��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(75, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Q��" Then
   Select Case atkingckai(75, 1)
      Case 1
            If movecp = 3 Then
                If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingckai(75, 2) = 0 Then
                   If ����ʧ@_�ˬd�O�_�����w���`���A(1, 16) = True Then
                        atkingckai(75, 2) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) + 1
                        �������m��l�`��(2) = �������m��l�`��(2) + 12
                    End If
                End If
                If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3) And atkingckai(75, 2) = 1 Then
                   atkingckai(75, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   �������m��l�`��(2) = �������m��l�`��(2) - 12
                 End If
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Q��\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6795
                   atkingno(i, 6) = 9405
                   atkingno(i, 7) = 93
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '=============================
             atkingckai(75, 2) = 0
             If ����ʧ@_�ˬd�O�_�����w���`���A(1, 16) = False Then
                 �����g�J��ܦC�ƭ� 2, Val(FormMainMode.��ܦC1.goi2) - 12
                 �������m��l�`��(2) = Val(FormMainMode.��ܦC1.goi2)
             End If
   End Select
End If
End Sub
Sub �����g_���ɷP��()
If FormMainMode.comaiatk(2).Caption = "���ɷP��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(85, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����g" Then
   Select Case atkingckai(85, 1)
        Case 1
             If atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(85, 2) = 0 Then
               atkingckai(85, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + pageqlead(1) * 2
            ElseIf (atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(85, 2) = 1 Then
               atkingckai(85, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - pageqlead(1) * 2
            End If
        Case 2
             atkingckai(85, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����g\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 840
                   atkingno(i, 5) = 8325
                   atkingno(i, 6) = 9285
                   atkingno(i, 7) = 63
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub �����g_��������()
If FormMainMode.comaiatk(3).Caption = "��������" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(86, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����g" Then
   Select Case atkingckai(86, 1)
        Case 1
             If movecp = 3 Then
                     If atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(86, 2) = 0 Then
                       atkingckai(86, 2) = 1
                       atkingtrn(2) = Val(atkingtrn(2)) + 1
                       �������m��l�`��(2) = �������m��l�`��(2) + 8
                    ElseIf (atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(86, 2) = 1 Then
                       atkingckai(86, 2) = 0
                       atkingtrn(2) = Val(atkingtrn(2)) - 1
                       �������m��l�`��(2) = �������m��l�`��(2) - 8
                    End If
             End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����g\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8430
                   atkingno(i, 6) = 8985
                   atkingno(i, 7) = 63
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(86, 2) = 0
             If �Y����ˮ`�� > 0 Then
                    Do
                         For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                           If �H�����`���A��Ʈw(1, i, 3) = 10 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                FormMainMode.personusspe(j).person_num = 10
                                FormMainMode.personusspe(j).person_turn = 1
                                �H�����`���A��Ʈw(1, j, 1) = 10
                                �H�����`���A��Ʈw(1, j, 2) = 1
                               Exit Do
                           End If
                         Next
                         For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                            If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                               �԰��t����.�H�����`���A��]�w_��] 1, i, 10, app_path & "gif\���`���A\atkdown.gif", 10, 1
                               ���`���A�ˬd��(10, 1) = 1
                               ���`���A�ˬd��(10, 2) = 1
                               Exit Do
                            End If
                         Next
                     Loop
              End If
   End Select
End If
End Sub
Sub �����g_�g�����b�P�ݦ大�j()
If FormMainMode.comaiatk(4).Caption = "�g�����b�P�ݦ大�j" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(87, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����g" Then
   Select Case atkingckai(87, 1)
        Case 1
             If movecp = 1 Then
                     If atkingpagetot(2, 1) >= 5 And atkingpagetot(2, 5) >= 5 And atkingckai(87, 2) = 0 Then
'                     If pageqlead(2) >= 1 And atkingckai(87, 2) = 0 Then
                       atkingckai(87, 2) = 1
                       atkingtrn(2) = Val(atkingtrn(2)) + 1
                       �������m��l�`��(2) = �������m��l�`��(2) + 6
                    ElseIf (atkingpagetot(2, 1) < 5 Or atkingpagetot(2, 5) < 5) And atkingckai(87, 2) = 1 Then
'                    ElseIf pageqlead(2) < 1 And atkingckai(87, 2) = 1 Then
                       atkingckai(87, 2) = 0
                       atkingtrn(2) = Val(atkingtrn(2)) - 1
                       �������m��l�`��(2) = �������m��l�`��(2) - 6
                    End If
             End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����g\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9495
                   atkingno(i, 6) = 9360
                   atkingno(i, 7) = 87
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
             atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 0
        Case 3
             For i = 1 To 106
                 If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                     �ثe��(20) = i
                     �ثe��(21) = 7
                     atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ + 1
                     FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
                     Exit Sub
                 End If
             Next
             If i = 107 Then
                 If atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 0 Then
                     For k = 1 To 3
                         �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 3, k
                     Next
                     �ثe��(22) = 28
                     FormMainMode.���ݮɶ�.Enabled = True
                 ElseIf atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 1 Or atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 2 Then
                     �ثe��(24) = 29
                     FormMainMode.���ݮɶ�_2.Enabled = True
                 ElseIf atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ > 2 Then
                     atkingckai(87, 2) = 0
                     �԰��t����.����ʧ@_�ޯ��ʵ���
                 End If
             End If
        Case 4
             If atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 0 Then
                atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 99
                �ثe��(22) = 28
                FormMainMode.���ݮɶ�.Enabled = True
             ElseIf atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 1 Then
                For k = 1 To 3
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 3, k
                Next
                atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 99
                �ثe��(24) = 29
                 FormMainMode.���ݮɶ�_2.Enabled = True
             ElseIf atking_AI_�����g_�g�����b�P�ݦ大�j_�m�P������ = 2 Then
                For k = 1 To 3
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 3, k
                Next
                atkingckai(87, 2) = 0
                �԰��t����.����ʧ@_�ޯ��ʵ���
             Else
                 atkingckai(87, 2) = 0
                �԰��t����.����ʧ@_�ޯ��ʵ���
             End If
   End Select
End If
End Sub
Sub �����i_���y����()
Dim dge As Integer
If FormMainMode.comaiatk(1).Caption = "���y����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(91, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����i" Then
   Select Case atkingckai(91, 1)
        Case 1
             If atkingpagetot(2, 1) >= 5 And atkingckai(91, 2) = 0 Then
               atkingckai(91, 2) = 1
               atkingckai(91, 1) = 2
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + pageqlead(2) * 3
               atking_AI_�����i_���y�����p��X�P�i�Ƭ����� = pageqlead(2)
            ElseIf atkingpagetot(2, 4) < 2 And atkingckai(91, 2) = 1 Then
               atkingckai(91, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - pageqlead(2) * 3
               atking_AI_�����i_���y�����p��X�P�i�Ƭ����� = 0
            End If
        Case 2
                 If atkingpagetot(2, 1) < 5 Then
                     atkingckai(91, 2) = 0
                     atkingtrn(2) = Val(atkingtrn(2)) - 1
                     If pageqlead(2) = atking_AI_�����i_���y�����p��X�P�i�Ƭ����� Then
                        �������m��l�`��(2) = �������m��l�`��(2) - pageqlead(2) * 3
                     Else
                        �������m��l�`��(2) = �������m��l�`��(2) - pageqlead(2) * 3 - 3
                     End If
                     atking_AI_�����i_���y�����p��X�P�i�Ƭ����� = 0
                     atkingckai(91, 1) = 1
                  End If
                  If atkingckai(91, 2) = 1 Then
                     �������m��l�`��(2) = �������m��l�`��(2) + (pageqlead(2) - Val(atking_AI_�����i_���y�����p��X�P�i�Ƭ�����)) * 3
                     atking_AI_�����i_���y�����p��X�P�i�Ƭ����� = pageqlead(2)
                  End If
        Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����i\atking1-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10725
                   atkingno(i, 7) = 91
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\�����i\atking1-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            dge = Val(FormMainMode.pagecomglead.Caption)
            If dge > 4 Then dge = 4
            �Y����ˮ`�� = Val(�Y����ˮ`��) - dge
            �Y���淾�q�Ȯ��ܼ�(2) = �Y����ˮ`��
            atking_AI_�����i_���y�����p��X�P�i�Ƭ����� = 0
            atkingckai(91, 2) = 0
   End Select
End If
End Sub
Sub �����i_�զʦX()
Dim dge As Integer
If FormMainMode.comaiatk(2).Caption = "�զʦX" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(92, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����i" Then
   Select Case atkingckai(92, 1)
        Case 1
             If movecp < 3 Then
                 If pageqlead(2) >= 2 And atkingckai(92, 2) = 0 Then
                   atkingckai(92, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If pageqlead(2) < 2 And atkingckai(92, 2) = 1 Then
                   atkingckai(92, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
             End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����i\atking2-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 92
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\�����i\atking2-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(92, 2) = 0
            '===================
            If �Y����ˮ`�� > 0 Then
               �԰��t����.����ʧ@_�M���Ҧ����`���A_�ϥΪ�
               '==================
               Do
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 3) = 22 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 1
                          �H�����`���A��Ʈw(1, i, 2) = 1
                          Exit Do
                      End If
                    Next
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                       If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 1, i, 22, app_path & "gif\���`���A\atkingerr.gif", 0, 1
                          ���`���A�ˬd��(22, 1) = 1
                          ���`���A�ˬd��(22, 2) = 1
                          Exit Do
                       End If
                    Next
                Loop
            End If
   End Select
End If
End Sub
Sub �����i_�t���¥�()
Dim dge As Integer
If FormMainMode.comaiatk(3).Caption = "�t���¥�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(93, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����i" Then
   Select Case atkingckai(93, 1)
        Case 1
             If atkingpagetot(2, 4) >= 3 And atkingckai(93, 2) = 0 Then
               atkingckai(93, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 3 And atkingckai(93, 2) = 1 Then
               atkingckai(93, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����i\atking3-1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10230
                   atkingno(i, 7) = 93
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 1
                   atkingno(i, 10) = app_path & "gif\�����i\atking3-2_2.jpg"
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atking_AI_�����i_�t���¥�������(1) = Val(FormMainMode.��ܦC1.goi1)
             atking_AI_�����i_�t���¥�������(2) = pageqlead(1)
        Case 4
            atkingckai(93, 2) = 0
            '===================
            If �Y����ˮ`�� <= 0 Then
               dge = Int(atking_AI_�����i_�t���¥�������(1) / 4 + 0.9)
               �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� dge, 1
            End If
            '===================
            If atking_AI_�����i_�t���¥�������(2) = 0 Then
                �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 2, 1
            End If
            '===================
            Erase atking_AI_�����i_�t���¥�������
   End Select
End If
End Sub
Sub �����i_���٤Ѩ�()
Dim dge As Integer
If FormMainMode.comaiatk(4).Caption = "���٤Ѩ�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(94, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����i" Then
   Select Case atkingckai(94, 1)
        Case 1
             If atkingpagetot(2, 4) >= 5 And atkingckai(94, 2) = 0 Then
               atkingckai(94, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 5 And atkingckai(94, 2) = 1 Then
               atkingckai(94, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����i\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7665
                   atkingno(i, 6) = 10590
                   atkingno(i, 7) = 94
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(94, 2) = 0
            '===================
            If livecom(����ݾ��H��������(2, 2)) = 0 And livecom(����ݾ��H��������(2, 3)) = 0 Then
                Do
                    For j = 14 * (����ݾ��H��������(2, 1) - 1) + 1 To 14 * ����ݾ��H��������(2, 1)
                          If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 7
                              FormMainMode.personcomspe(j).person_turn = 4
                              �H�����`���A��Ʈw(2, j, 1) = 7
                              �H�����`���A��Ʈw(2, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����ݾ��H��������(2, 1) - 1) + 1 To 14 * ����ݾ��H��������(2, 1)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 7, 4
                          ���`���A�ˬd��(1, 1) = 1
                          ���`���A�ˬd��(1, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '=================================
                Do
                    For j = 14 * (����ݾ��H��������(2, 1) - 1) + 1 To 14 * ����ݾ��H��������(2, 1)
                          If �H�����`���A��Ʈw(2, j, 3) = 2 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_num = 7
                              FormMainMode.personcomspe(j).person_turn = 4
                              �H�����`���A��Ʈw(2, j, 1) = 7
                              �H�����`���A��Ʈw(2, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����ݾ��H��������(2, 1) - 1) + 1 To 14 * ����ݾ��H��������(2, 1)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 2, app_path & "gif\���`���A\defup.gif", 7, 4
                          ���`���A�ˬd��(2, 1) = 1
                          ���`���A�ˬd��(2, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '================
                Do
                    For j = 14 * (����ݾ��H��������(2, 1) - 1) + 1 To 14 * ����ݾ��H��������(2, 1)
                          If �H�����`���A��Ʈw(2, j, 3) = 38 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_turn = 4
                              �H�����`���A��Ʈw(2, j, 2) = 4
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����ݾ��H��������(2, 1) - 1) + 1 To 14 * ����ݾ��H��������(2, 1)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 38, app_path & "gif\���`���A\�A��.gif", 0, 4
                          ���`���A�ˬd��(38, 1) = 1
                          ���`���A�ˬd��(38, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
                '================
            Else
                '================
                For i = 2 To 3
                     If livecom(����ݾ��H��������(2, i)) > 0 Then
                        Do
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                                  If �H�����`���A��Ʈw(2, j, 3) = 36 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_turn = 1
                                      �H�����`���A��Ʈw(2, j, 2) = 1
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                               If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 2, j, 36, app_path & "gif\���`���A\���@.png", 0, 1
                                  ���`���A�ˬd��(36, 1) = 1
                                  ���`���A�ˬd��(36, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        �԰��t����.�^�_����_�q�� 1, i
                     End If
                Next
            End If
   End Select
End If
End Sub
Sub �S�{��_�G�����F()
If FormMainMode.comaiatk(1).Caption = "�G�����F" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(95, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�S�{��" Then
   Select Case atkingckai(95, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 3 And atkingckai(95, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                   atkingckai(95, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 1) < 3 And atkingckai(95, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                   atkingckai(95, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�S�{��\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             Do
                For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                  If �H�����`���A��Ʈw(1, i, 3) = 33 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                      FormMainMode.personusspe(i).person_turn = 3
                      �H�����`���A��Ʈw(1, i, 2) = 3
                      Exit Do
                  End If
                Next
                For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                   If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                      �԰��t����.�H�����`���A��]�w_��] 1, i, 33, app_path & "gif\���`���A\�G��.gif", 0, 3
                      ���`���A�ˬd��(33, 1) = 1
                      ���`���A�ˬd��(33, 2) = 1
                      Exit Do
                   End If
                Next
            Loop
            atkingckai(95, 2) = 0
   End Select
End If
End Sub
Sub �S�{��_�a�g���t()
If FormMainMode.comaiatk(2).Caption = "�a�g���t" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(96, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�S�{��" Then
   Select Case atkingckai(96, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(96, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                   atkingckai(96, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(96, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                   atkingckai(96, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�S�{��\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 96
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(96, 2) = 0
            '=================
            �԰��t����.����ʧ@_�Z���ܧ� 1
   End Select
End If
End Sub
Sub �S�{��_�t�v���l()
If FormMainMode.comaiatk(3).Caption = "�t�v���l" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(97, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�S�{��" Then
   Select Case atkingckai(97, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(97, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                   atkingckai(97, 2) = 1
                End If
                If (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(97, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                   atkingckai(97, 2) = 0
                 End If
          End If
      Case 2
           atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�S�{��\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 7830
                   atkingno(i, 6) = 10260
                   atkingno(i, 7) = 97
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 4
            atkingckai(97, 2) = 0
            '=================
            �԰��t����.����ʧ@_�Z���ܧ� 3
            If Val(�Y����ˮ`��) < 0 Then
                �^�_����_�q�� 1, 1
            End If
   End Select
End If
End Sub
Sub �S�{��_���M�C�{(ByVal Index As Integer)
Dim aw As Integer
If FormMainMode.comaiatk(4).Caption = "���M�C�{" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(98, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�S�{��" Then
   Select Case atkingckai(98, 1)
      Case 1
             If movecp = 3 Then
                 If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(98, 2) = 0 Then
'                     aw = Int(atkingpagetot(2, 1) / 2 + 0.5)
                     For i = 1 To 106
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                           aw = aw + 1
                        End If
                     Next
                     �������m��l�`��(2) = �������m��l�`��(2) + (aw - atking_AI_�S�{��_���M�C�{�p��i�Ƭ�����) * 5 + 8
                     atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� = aw
                     atkingckai(98, 2) = 1
                     atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
            End If
      Case 2
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 2 And Val(pagecardnum(Index, 5)) = 1 And atkingckai(98, 2) = 1 Then
                   For i = 1 To 106
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                           aw = aw + 1
                        End If
                   Next
                   �������m��l�`��(2) = �������m��l�`��(2) + (aw - atking_AI_�S�{��_���M�C�{�p��i�Ƭ�����) * 5
                   atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� = aw
            End If
            If pagecardnum(Index, 1) = a1a And Val(pagecardnum(Index, 6)) = 1 And Val(pagecardnum(Index, 5)) = 1 And atkingckai(98, 2) = 1 Then
                   If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(98, 2) = 1 Then
                        For i = 1 To 106
                             If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                                aw = aw + 1
                             End If
                        Next
                        �������m��l�`��(2) = �������m��l�`��(2) + (aw - atking_AI_�S�{��_���M�C�{�p��i�Ƭ�����) * 5
                        atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� = aw
                   ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) And atkingckai(98, 2) = 1 Then
                        �������m��l�`��(2) = �������m��l�`��(2) - (atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� * 5) - 8
                        atkingckai(98, 2) = 0
                        atkingckai(98, 1) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) - 1
                        atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� = 0
                    End If
            End If
'            formmainmode.trgoi2.Enabled = True
    Case 3
        If Val(pagecardnum(Index, 5)) = 1 And atkingckai(98, 2) = 1 Then
               If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 Then
                    For i = 1 To 106
                         If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                            aw = aw + 1
                         End If
                    Next
                    �������m��l�`��(2) = �������m��l�`��(2) + (aw - atking_AI_�S�{��_���M�C�{�p��i�Ƭ�����) * 5
                    atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� = aw
               ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) Then
                    �������m��l�`��(2) = �������m��l�`��(2) - (atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� * 5) - 8
                    atkingckai(98, 2) = 0
                    atkingckai(98, 1) = 1
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
                    atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� = 0
                End If
        End If
'        formmainmode.trgoi2.Enabled = True
      Case 4
             atkingckai(98, 2) = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�S�{��\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8760
                   atkingno(i, 6) = 10530
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atking_AI_�S�{��_���M�C�{�p��i�Ƭ����� = 0
   End Select
End If
End Sub
Sub ����_�ڤ��]��()
Dim m As Integer, n As Integer, bd As Integer
If FormMainMode.comaiatk(1).Caption = "�ڤ��]��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(99, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(99, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 3 And atkingckai(99, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 4
                   atkingckai(99, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 5) < 3 And atkingckai(99, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 4
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   atkingckai(99, 2) = 0
                 End If
          End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9405
                   atkingno(i, 6) = 10245
                   atkingno(i, 7) = 99
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             �ޯ�ʵe��ܶ��q�� = 10
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.���m���q_���q��l.Enabled = False
        Case 3
            Randomize
            m = Int(Rnd() * 100) + 1
            If livecom(����H����ԤH��(2, 2)) <= livecom41(����H����ԤH��(2, 2)) Then
                Randomize
                bd = Int(Rnd() * 2) + 1
            End If
            If m Mod (2 - bd) = 0 Then '===�۷��50~100%���v
                 Randomize
                 n = Int(Rnd() * 100) + 1
                 If livecom(����H����ԤH��(2, 2)) <= livecommax(����H����ԤH��(2, 2)) Then
                     bd = livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2))
                     If bd > 8 Then bd = 8
                 End If
                 If n Mod (10 - bd) = 0 Then '===�۷��10~50%���v
                     �������m��l�`��(2) = �������m��l�`��(2) * 4
                     FormMainMode.messageus.AddItem "�ڤ��]���ĪG�o��!  �����O�ܬ�4��"
                     �԰��t����.�۰ʱ��b����
                Else
                     �������m��l�`��(2) = �������m��l�`��(2) * 2
                     FormMainMode.messageus.AddItem "�ڤ��]���ĪG�o��!  �����O�ܬ�2��"
                     �԰��t����.�۰ʱ��b����
                End If
                FormMainMode.trgoi2_Timer
            End If
            atkingckai(99, 1) = 4
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(99, 1) = 5
        Case 5
             atkingckai(99, 2) = 0
             If Val(�Y����ˮ`��) <= 0 Then
                 Do
                    For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                        If �H�����`���A��Ʈw(1, j, 3) = 11 And �H�����`���A��Ʈw(1, j, 2) > 0 Then
                         FormMainMode.personusspe(j).person_num = 3
                         FormMainMode.personusspe(j).person_turn = 3
                         �H�����`���A��Ʈw(1, j, 1) = 3
                         �H�����`���A��Ʈw(1, j, 2) = 3
                         Exit Do
                        End If
                    Next
                   For j = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, j, 2) = 0 Then
                         �԰��t����.�H�����`���A��]�w_��] 1, j, 11, app_path & "gif\���`���A\defdown.gif", 3, 3
                         ���`���A�ˬd��(11, 1) = 1
                         ���`���A�ˬd��(11, 2) = 1
                         Exit Do
                     End If
                   Next
                Loop
            End If
   End Select
End If
End Sub
Sub ����_�K�a�ڦ�()
Dim m As Integer, n As Integer, bd As Integer
If FormMainMode.comaiatk(2).Caption = "�K�a�ڦ�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(100, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(100, 1)
      Case 1
            If atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 2) >= 3 And atkingckai(100, 2) = 0 Then
               �������m��l�`��(2) = �������m��l�`��(2) + 3
               atkingckai(100, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If (atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 2) < 3) And atkingckai(100, 2) = 1 Then
               �������m��l�`��(2) = �������m��l�`��(2) - 3
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               atkingckai(100, 2) = 0
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8865
                   atkingno(i, 6) = 9210
                   atkingno(i, 7) = 100
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             �ޯ�ʵe��ܶ��q�� = 11
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.�������q_���q2.Enabled = False
        Case 3
            bd = 0
            Randomize
            m = Int(Rnd() * 100) + 1
            If livecom(����H����ԤH��(2, 2)) <= livecom41(����H����ԤH��(2, 2)) Then
                bd = 1
            End If
            If m Mod (3 - bd) = 0 Then '===�۷��33~50%���v
                 Randomize
                 n = Int(Rnd() * 100) + 1
                 If livecom(����H����ԤH��(2, 2)) <= livecommax(����H����ԤH��(2, 2)) Then
                     bd = livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2))
                     If bd > 8 Then bd = 8
                 End If
                 If n Mod (10 - bd) = 0 Then '===�۷��10~50%���v
                     �������m��l�`��(1) = 0
                     FormMainMode.messageus.AddItem "�K�a�ڦЮĪG�o��!  �ڤ�����O�ܬ�0"
                     �԰��t����.�۰ʱ��b����
                Else
                     �������m��l�`��(1) = �������m��l�`��(1) \ 2
                     FormMainMode.messageus.AddItem "�K�a�ڦЮĪG�o��!  �ڤ�����O�ܬ�1/2"
                     �԰��t����.�۰ʱ��b����
                End If
            Else
                �������m��l�`��(1) = Int((�������m��l�`��(1) * 2) / 3)
                FormMainMode.messageus.AddItem "�K�a�ڦЮĪG�o��!  �ڤ�����O�ܬ�2/3"
                �԰��t����.�۰ʱ��b����
            End If
            FormMainMode.trgoi1_Timer
            '=====================
            �԰��t����.�^�_����_�q�� 1, 1
            '=====================
            atkingckai(100, 1) = 4
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(100, 2) = 0
   End Select
End If
End Sub
Sub ����_���Ϥ۹�()
Dim bloodnum As Integer
If FormMainMode.comaiatk(3).Caption = "���Ϥ۹�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(101, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(101, 1)
      Case 1
           If movecp < 3 Then
                If pageqlead(2) >= 2 And atkingckai(101, 2) = 0 Then
                    atkingckai(101, 2) = 1
                 End If
                 If pageqlead(2) < 2 And atkingckai(101, 2) = 1 Then
                    atkingckai(101, 2) = 0
                  End If
            End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6930
                   atkingno(i, 6) = 9540
                   atkingno(i, 7) = 101
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_AI_����_���Ϥ۹�_��P������
                   Select Case livecom(����H����ԤH��(2, 2))
                       Case Is >= 5
                           atking_AI_����_���Ϥ۹�_��P������(2) = 4
                           atkingno(i, 11) = 1
                       Case Else
                           atking_AI_����_���Ϥ۹�_��P������(2) = 2
                           atkingno(i, 11) = 0
                    End Select
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_AI_����_���Ϥ۹�_��P������(2) And atking_AI_����_���Ϥ۹�_��P������(1) = 0 Then
               �԰��t����.����ʧ@_�~�P
            End If
            atking_AI_����_���Ϥ۹�_��P������(1) = atking_AI_����_���Ϥ۹�_��P������(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_����_���Ϥ۹�_��P������(1) > atking_AI_����_���Ϥ۹�_��P������(2)
                    �ثe��(15) = 30
                    FormMainMode.tr�P��_��P_�q��.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_����_���Ϥ۹�_��P������(1) > atking_AI_����_���Ϥ۹�_��P������(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_AI_����_���Ϥ۹�_��P������(2)) <= 2 Then
                   �԰��t����.�ˮ`����_�ޯઽ��_�q�� 1, 1
                   atkingckai(101, 2) = 0
               Else
                   �ثe��(24) = 32
                   FormMainMode.���ݮɶ�_2.Enabled = True
               End If
            End If
        Case 5
            atkingckai(101, 2) = 0
            �԰��t����.�ˮ`����_�ޯઽ��_�q�� 1, 1
            �԰��t����.����ʧ@_�ޯ��ʵ���
   End Select
End If
End Sub
Sub ����_�ڹҷn�x()
If FormMainMode.comaiatk(4).Caption = "�ڹҷn�x" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(102, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(102, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(102, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 3
                   atkingckai(102, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(102, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 3
                   atkingckai(102, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7710
                   atkingno(i, 6) = 9030
                   atkingno(i, 7) = 102
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If livecom(����H����ԤH��(2, 2)) > 2 Then
                 For i = 1 To 3
                     �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 1, i
                 Next
            Else
                 For i = 1 To 3
                     �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 4, i
                 Next
             End If
             atkingckai(102, 2) = 0
   End Select
End If
End Sub
Sub �j�|�˺��h_�����[��()
Dim i As Integer, j As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(2).Caption = "�����[��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(104, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�j�|�˺��h" Then
   Select Case atkingckai(104, 1)
        Case 1
            If atkingpagetot(2, 4) >= 2 And atkingpagetot(2, 3) = 0 And atkingckai(104, 2) = 0 Then
               atkingckai(104, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 4) < 2 Or atkingpagetot(2, 3) <> 0) And atkingckai(104, 2) = 1 Then
               atkingckai(104, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�j�|�˺��h\�j�|�˺��h-�����[��2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 240
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6600
                   atkingno(i, 6) = 9345
                   atkingno(i, 7) = 104
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
         Do
            atkingckai(104, 2) = 0
            For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                     FormMainMode.personcomspe(j).person_num = 5
                     FormMainMode.personcomspe(j).person_turn = 1
                     �H�����`���A��Ʈw(2, j, 1) = 5
                     �H�����`���A��Ʈw(2, j, 2) = 1
                     Exit Do
                 End If
            Next
           For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
              If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                 �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 5, 1
                 ���`���A�ˬd��(1, 1) = 1
                 ���`���A�ˬd��(1, 2) = 1
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub �j�|�˺��h_�믫�O�l��()
Dim rrr(1 To 3) As Integer '�P�P�_�Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "�믫�O�l��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(105, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�j�|�˺��h" Then
   Select Case atkingckai(105, 1)
        Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                    If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(1) = rrr(1) + 1
                    End If
                    If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(2) = rrr(2) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(3) = rrr(3) + 1
                    End If
                End If
             Next
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(105, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(105,  2) = 0 Then
                atkingckai(105, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(105, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(105,  2) = 1 Then
                atkingckai(105, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
              End If
        Case 2
              For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�j�|�˺��h\Grunwaldatking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -1080
                   atkingno(i, 5) = 8025
                   atkingno(i, 6) = 9525
                   atkingno(i, 7) = 105
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
            cardpn = 0
'            Erase cardp
            Erase atking_AI_�j�|�˺��h_�믫�O�l��������
            '=====================
            For i = 1 To 106
                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                    If pagecardnum(i, 1) = a4a Or pagecardnum(i, 3) = a4a Then
                         atking_AI_�j�|�˺��h_�믫�O�l��������(i) = 1
                         atking_AI_�j�|�˺��h_�믫�O�l��������(0) = atking_AI_�j�|�˺��h_�믫�O�l��������(0) + 1
                     End If
                End If
            Next
            If atking_AI_�j�|�˺��h_�믫�O�l��������(0) > 0 Then
                atking_AI_�j�|�˺��h_�믫�O�l��������(0) = 0
                For i = 1 To 106
                    If atking_AI_�j�|�˺��h_�믫�O�l��������(i) = 1 Then
                        atking_AI_�j�|�˺��h_�믫�O�l��������(0) = Val(atking_AI_�j�|�˺��h_�믫�O�l��������(0)) + 1
                        �ثe��(20) = i
                        �ثe��(21) = 8
                        atking_AI_�j�|�˺��h_�믫�O�l��������(i) = 0
                        FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
                        Exit Sub
                    End If
                Next
            Else
               �ثe��(22) = 31
               FormMainMode.���ݮɶ�.Enabled = True
            End If
        Case 4
'            FormMainMode.tr�q���P_���P.Enabled = True
'            �ثe��(17) = 5
        Case 5
            If atking_AI_�j�|�˺��h_�믫�O�l��������(0) > 0 Then
                For i = 1 To 106
                    If atking_AI_�j�|�˺��h_�믫�O�l��������(i) = 1 And atking_AI_�j�|�˺��h_�믫�O�l��������(0) < 3 Then
                        atking_AI_�j�|�˺��h_�믫�O�l��������(0) = Val(atking_AI_�j�|�˺��h_�믫�O�l��������(0)) + 1
                        �ثe��(20) = i
                        �ثe��(21) = 8
                        atking_AI_�j�|�˺��h_�믫�O�l��������(i) = 0
                        FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
                        Exit Sub
                    End If
                Next
                If i = 107 Then
                    atkingckai(105, 2) = 0
                    �԰��t����.����ʧ@_�ޯ��ʵ���
                 End If
            Else
               �ثe��(22) = 31
               FormMainMode.���ݮɶ�.Enabled = True
            End If
        Case 6
            If atking_AI_�j�|�˺��h_�믫�O�l��������(0) = 0 Then
                atking_AI_�j�|�˺��h_�믫�O�l��������(0) = 99
               �ثe��(22) = 31
               FormMainMode.���ݮɶ�.Enabled = True
            ElseIf atking_AI_�j�|�˺��h_�믫�O�l��������(0) > 0 Then
               atkingckai(105, 2) = 0
               �԰��t����.����ʧ@_�ޯ��ʵ���
            End If
   End Select
End If

End Sub
Sub ���[_���㤧��()
If FormMainMode.comaiatk(1).Caption = "���㤧��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(106, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���[" Then
   Select Case atkingckai(106, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(106, 2) = 0 Then
               atkingckai(106, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf atkingpagetot(2, 4) < 1 And atkingckai(106, 2) = 1 Then
               atkingckai(106, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���[\���[_���㤧��_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -600
                   atkingno(i, 4) = 1440
                   atkingno(i, 5) = 7050
                   atkingno(i, 6) = 9090
                   atkingno(i, 7) = 106
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
    Case 3
        Do
           atkingckai(106, 2) = 0
           For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
             If �H�����`���A��Ʈw(2, i, 2) >= 9 And �H�����`���A��Ʈw(2, i, 3) = 26 Then
                Exit Do
             End If
             If �H�����`���A��Ʈw(2, i, 3) = 26 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) < 9 Then
                 FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2) + 1
                 �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + 1
                 Exit Do
             End If
           Next
           For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
              If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                 �԰��t����.�H�����`���A��]�w_��] 2, i, 26, app_path & "gif\���`���A\�t��.gif", 0, 1
                 ���`���A�ˬd��(26, 1) = 1
                 ���`���A�ˬd��(26, 2) = 1
                 Exit Do
             End If
           Next
        Loop
   End Select
End If
End Sub
Sub ��ܵY_��k���Ӫ�()
Dim cardp(1 To 106) As Boolean '�����Ȯ��ܼ�
Dim cardpn As Integer '�����P�`�ƼȮ��ܼ�
If FormMainMode.comaiatk(1).Caption = "��k���Ӫ�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(107, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��ܵY" Then
   Select Case atkingckai(107, 1)
        Case 1
           If movecp < 3 Then
                 If atkingpagetot(2, 4) >= 2 And atkingckai(107, 2) = 0 Then
'                 If pageqlead(2) >= 1 And atkingckai(107, 2) = 0 Then
                   atkingckai(107, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf atkingpagetot(2, 4) < 2 And atkingckai(107, 2) = 1 Then
'                ElseIf pageqlead(2) < 1 And atkingckai(107, 2) = 1 Then
                   atkingckai(107, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
           End If
        Case 2
              For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��ܵY\Evelynatking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 6345
                   atkingno(i, 6) = 9810
                   atkingno(i, 7) = 107
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 3
            Do
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, i, 2) >= 9 And �H�����`���A��Ʈw(2, i, 3) = 25 Then
                    Exit Do
                 End If
                 If �H�����`���A��Ʈw(2, i, 3) = 25 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) < 9 Then
                     FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2) + 1
                     �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + 1
                     Exit Do
                 End If
               Next
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                  If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                     �԰��t����.�H�����`���A��]�w_��] 2, i, 25, app_path & "gif\���`���A\��O�C�U.gif", 0, 1
                     ���`���A�ˬd��(25, 1) = 1
                     ���`���A�ˬd��(25, 2) = 1
                     Exit Do
                 End If
               Next
            Loop
            '=====================
            cardpn = 0
            Erase cardp
            Erase atking_AI_��ܵY_��k���Ӫ������
            '=====================
            Do
               Randomize
               i = Int(Rnd() * 106) + 1
               If cardp(i) = False Then
                    cardp(i) = True
                    cardpn = cardpn + 1
                    If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                      Select Case movecp
                         Case 1
                             If pagecardnum(i, 1) = a1a Or pagecardnum(i, 3) = a1a Then
                                  atking_AI_��ܵY_��k���Ӫ������(atking_AI_��ܵY_��k���Ӫ������(0) + 1) = i
                                  atking_AI_��ܵY_��k���Ӫ������(0) = atking_AI_��ܵY_��k���Ӫ������(0) + 1
                              End If
                         Case Is > 1
                             If pagecardnum(i, 1) = a5a Or pagecardnum(i, 3) = a5a Then
                                  atking_AI_��ܵY_��k���Ӫ������(atking_AI_��ܵY_��k���Ӫ������(0) + 1) = i
                                  atking_AI_��ܵY_��k���Ӫ������(0) = atking_AI_��ܵY_��k���Ӫ������(0) + 1
                              End If
                        End Select
                    End If
               End If
               If atking_AI_��ܵY_��k���Ӫ������(0) >= 2 Then
                   Exit Do
               End If
            Loop While cardpn < 106
            If atking_AI_��ܵY_��k���Ӫ������(0) > 0 Then
                �ثe��(20) = atking_AI_��ܵY_��k���Ӫ������(1)
                atkingckai(107, 1) = 4
                �ثe��(21) = 9
                FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
            Else
'               atkingckai(107, 2) = 0
               atkingckai(107, 1) = 4
               �ثe��(22) = 32
               FormMainMode.���ݮɶ�.Enabled = True
            End If
        Case 4
             If atking_AI_��ܵY_��k���Ӫ������(0) < 2 Then
                �ثe��(22) = 32
                FormMainMode.���ݮɶ�.Enabled = True
            Else
                �ثe��(20) = atking_AI_��ܵY_��k���Ӫ������(2)
                atkingckai(107, 1) = 5
                �ثe��(21) = 9
                FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
            End If
        Case 5
            If atking_AI_��ܵY_��k���Ӫ������(0) = 0 Then
               atking_AI_��ܵY_��k���Ӫ������(0) = 3
               �ثe��(22) = 32
               FormMainMode.���ݮɶ�.Enabled = True
               Exit Sub
            ElseIf atking_AI_��ܵY_��k���Ӫ������(0) = 2 Then
               �ثe��(24) = 34
               FormMainMode.���ݮɶ�_2.Enabled = True
               Exit Sub
            ElseIf atking_AI_��ܵY_��k���Ӫ������(0) > 0 And atking_AI_��ܵY_��k���Ӫ������(0) <> 2 Then
               atkingckai(107, 2) = 0
               ����ʧ@_�ޯ��ʵ���
            End If
        Case 6
            atkingckai(107, 2) = 0
            ����ʧ@_�ޯ��ʵ���
   End Select
End If

End Sub
Sub ��ܵY_�����ۺh()
Dim mkp As Integer '�Ȯ��ܼ�
Dim cardp(1 To 106) As Boolean '�����Ȯ��ܼ�
Dim cardpn(1 To 2) As Integer '�����P�`�ƼȮ��ܼ�(1.�P�����ثe�`��/2.�P��w�ثe�`��)
If FormMainMode.comaiatk(4).Caption = "�����ۺh" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(108, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��ܵY" Then
   If Formsetting.checktest.Value = 1 Then Debug.Print "�g�L�����ۺh�D�W�r�P�_"
   Select Case atkingckai(108, 1)
        Case 1
            If movecp = 3 Then
                 If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 _
                    And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 5) >= 1 And atkingckai(108, 2) = 0 Then
'                 If atkingpagetot(2, 3) >= 1 And atkingckai(108, 2) = 0 Then
                   atkingckai(108, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1 _
                   Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 5) < 1) And atkingckai(108, 2) = 1 Then
'                ElseIf atkingpagetot(2, 3) < 1 And atkingckai(108, 2) = 1 Then
                   atkingckai(108, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
          End If
        Case 2
             '==================================
             Erase atking_AI_��ܵY_�����ۺh���q������
             Randomize
             mkp = Int(Rnd() * 16) + 1
             atking_AI_��ܵY_�����ۺh���q������(0, 1) = mkp
             If Formsetting.checktest.Value = 1 Then Debug.Print "�ޯ� - AI - ��ܵY - �����ۺh�ĪG��" & mkp
             '===================================
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��ܵY\Evelynatking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9330
                   atkingno(i, 6) = 9165
                   atkingno(i, 7) = 108
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                    If atking_AI_��ܵY_�����ۺh���q������(0, 1) <= 9 Or atking_AI_��ܵY_�����ۺh���q������(0, 1) >= 13 Then
                       atkingno(i, 11) = 0
                    Else
                       atkingno(i, 11) = 1
                    End If
                   Exit For
                 End If
             Next
        Case 3
            '======================
               �԰��t����.����ʧ@_�M���Ҧ����`���A_�q��
            '======================
            Select Case atking_AI_��ܵY_�����ۺh���q������(0, 1)
                Case 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 1, 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 1, 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 1, 3
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 1, 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 1, 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 1, 3
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv����P��趤�����1�I�ˮ`�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                Case 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 3, 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 3, 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 3, 3
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 3, 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 3, 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 3, 3
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv����P��趤�����3�I�ˮ`�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                Case 3
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 5, 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 5, 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 5, 3
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 5, 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 5, 2
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� 5, 3
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv����P��趤�����5�I�ˮ`�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                '==============================================
                Case 4
                    �^�_����_�ϥΪ� 1, 1
                    �^�_����_�ϥΪ� 1, 2
                    �^�_����_�ϥΪ� 1, 3
                    �^�_����_�q�� 1, 1
                    �^�_����_�q�� 1, 2
                    �^�_����_�q�� 1, 3
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv����P��趤�����HP�^�_1�I�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                Case 5
                    �^�_����_�ϥΪ� 3, 1
                    �^�_����_�ϥΪ� 3, 2
                    �^�_����_�ϥΪ� 3, 3
                    �^�_����_�q�� 3, 1
                    �^�_����_�q�� 3, 2
                    �^�_����_�q�� 3, 3
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv����P��趤�����HP�^�_3�I�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                Case 6
                    �^�_����_�ϥΪ� 5, 1
                    �^�_����_�ϥΪ� 5, 2
                    �^�_����_�ϥΪ� 5, 3
                    �^�_����_�q�� 5, 1
                    �^�_����_�q�� 5, 2
                    �^�_����_�q�� 5, 3
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv����P��趤�����HP�^�_5�I�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                '===============================================
                Case 7
                    �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� Val(liveus(����H����ԤH��(1, 2))) - 1, 1
                    �԰��t����.�ˮ`����_�ޯઽ��_�q�� Val(livecom(����H����ԤH��(2, 2))) - 1, 1
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv�P��誺HP�ܬ�1�I�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                '============================================
                Case 8
                    If Val(liveus(����H����ԤH��(1, 2))) > 5 Then
                        �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� Val(liveus(����H����ԤH��(1, 1))) - 5, 1
                    Else
                        �^�_����_�ϥΪ� 5 - Val(liveus(����H����ԤH��(1, 1))), 1
                    End If
                    If Val(livecom(����H����ԤH��(2, 2))) > 5 Then
                        �԰��t����.�ˮ`����_�ޯઽ��_�q�� Val(livecom(����H����ԤH��(2, 2))) - 5, 1
                    Else
                        �^�_����_�q�� 5 - Val(livecom(����H����ԤH��(2, 2))), 1
                    End If
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv�P��誺HP�ܬ�5�I�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                '===============================================
                Case 9
                    �^�_����_�ϥΪ� Val(liveusmax(����H����ԤH��(1, 2))) - Val(liveus(����H����ԤH��(1, 2))), 1
                    �^�_����_�q�� Val(livecommax(����H����ԤH��(2, 2))) - Val(livecom(����H����ԤH��(2, 2))), 1
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �ۤv�P��誺HP������_�C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                '===============================================
                Case 10
                    �ثe��(20) = 1
                    atking_AI_��ܵY_�����ۺh���q������(0, 2) = 1
                '==========�ϥΪ̱�P���q
                    Do
                        If Val(pagecardnum(�ثe��(20), 5)) = 1 And Val(pagecardnum(�ثe��(20), 6)) = 1 Then
                            atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                            �ثe��(21) = 10
                            FormMainMode.tr�ϥΪ�_��P.Enabled = True
                            Exit Sub
                        End If
                        �ثe��(20) = �ثe��(20) + 1
                    Loop Until �ثe��(20) > 106
                    If �ثe��(20) > 106 Then
                        GoTo �ĪG10_�ϥΪ̱�P���q�������L
                    End If
                '============================================
                Case 11
                    �ثe��(20) = 1
                    '========�ϥΪ̵P�ƧP�_�ο��
                    If Val(FormMainMode.pageusglead) > 8 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 1
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_��ܵY_�����ۺh���q������(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pageusglead) - 8 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 1 Then
                                atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 0
                                atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                                �ثe��(21) = 10
                                FormMainMode.tr�ϥΪ�_��P.Enabled = True
                                Exit Sub
                            End If
                            �ثe��(20) = �ثe��(20) + 1
                        Loop Until �ثe��(20) > 106
                    ElseIf Val(FormMainMode.pageusglead) < 8 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 3
                        atking_AI_��ܵY_�����ۺh���q������(0, 3) = 8
                        atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                        �ثe��(15) = 33
                        If Val(FormMainMode.pageul) < 8 - Val(FormMainMode.pageusglead) Then
                            �԰��t����.����ʧ@_�~�P
                        End If
                        FormMainMode.tr�P��_��P_�ϥΪ�.Enabled = True
                    Else
                        GoTo �ĪG11_���ܹq���P�_
                    End If
                '============================================
                Case 12
                    �ثe��(20) = 1
                    '========�ϥΪ̵P�ƧP�_�ο��
                    If Val(FormMainMode.pageusglead) > 15 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 1
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_��ܵY_�����ۺh���q������(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pageusglead) - 15 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 1 Then
                                atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 0
                                atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                                �ثe��(21) = 10
                                FormMainMode.tr�ϥΪ�_��P.Enabled = True
                                Exit Sub
                            End If
                            �ثe��(20) = �ثe��(20) + 1
                        Loop Until �ثe��(20) > 106
                    ElseIf Val(FormMainMode.pageusglead) < 15 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 3
                        atking_AI_��ܵY_�����ۺh���q������(0, 3) = 15
                        atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                        �ثe��(15) = 33
                        If Val(FormMainMode.pageul) < 15 - Val(FormMainMode.pageusglead) Then
                            �԰��t����.����ʧ@_�~�P
                        End If
                        FormMainMode.tr�P��_��P_�ϥΪ�.Enabled = True
                    Else
                        GoTo �ĪG12_���ܹq���P�_
                    End If
                '===============================================
                Case 13
                    ����ʧ@_�Z���ܧ� (1)
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �Z���ܬ���Z���C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                Case 14
                    ����ʧ@_�Z���ܧ� (2)
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �Z���ܬ����Z���C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                Case 15
                    ����ʧ@_�Z���ܧ� (3)
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  �Z���ܬ����Z���C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
                Case 16
                    FormMainMode.messageus.AddItem "�����ۺh�o��!  ���򳣨S���o�͡C"
                    �԰��t����.�۰ʱ��b����
                    atkingckai(108, 2) = 0
            End Select
        '=====================================================
       Case 4
             If atking_AI_��ܵY_�����ۺh���q������(0, 1) = 10 Then
                    '==========�ϥΪ̱�P���q2
                    If �ثe��(20) <= 106 Then
                        Do
                            If Val(pagecardnum(�ثe��(20), 5)) = 1 And Val(pagecardnum(�ثe��(20), 6)) = 1 Then
                                atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                                �ثe��(21) = 10
                                FormMainMode.tr�ϥΪ�_��P.Enabled = True
                                Exit Sub
                            End If
                            �ثe��(20) = �ثe��(20) + 1
                        Loop Until �ثe��(20) > 106
                    End If
�ĪG10_�ϥΪ̱�P���q�������L:
                    If �ثe��(20) > 106 And atking_AI_��ܵY_�����ۺh���q������(0, 2) = 1 Then
                        �ثe��(16) = 1
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 2
                        '=============�q�����P���q1
                        Do
                            If Val(pagecardnum(�ثe��(16), 5)) = 2 And Val(pagecardnum(�ثe��(16), 6)) = 1 Then
                                atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                                �ثe��(17) = 12
                                FormMainMode.tr�q���P_½�P.Enabled = True
                                Exit Sub
                            End If
                            �ثe��(16) = �ثe��(16) + 1
                        Loop Until �ثe��(16) > 106
                        If �ثe��(16) > 106 Then
'                            FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�0�i�C"
'                            �԰��t����.�۰ʱ��b����
'                            atkingckai(108, 2) = 0
'                            ����ʧ@_�ޯ��ʵ���
                            If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
                                GoTo �ĪG�������_��P�ܤ���
                            Else
                                �ثe��(22) = 33
                                FormMainMode.���ݮɶ�.Enabled = True
                            End If
                        End If
                    End If
                    If �ثe��(16) <= 106 And atking_AI_��ܵY_�����ۺh���q������(0, 2) = 2 Then
                        '==============�q�����P���q2
                        Do
                            If Val(pagecardnum(�ثe��(16), 5)) = 2 And Val(pagecardnum(�ثe��(16), 6)) = 1 Then
                                atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                                �ثe��(17) = 12
                                FormMainMode.tr�q���P_½�P.Enabled = True
                                Exit Sub
                            End If
                            �ثe��(16) = �ثe��(16) + 1
                        Loop Until �ثe��(16) > 106
                        If �ثe��(16) > 106 Then
'                            FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�0�i�C"
'                            �԰��t����.�۰ʱ��b����
'                            atkingckai(108, 2) = 0
'                            ����ʧ@_�ޯ��ʵ���
                            If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
                                GoTo �ĪG�������_��P�ܤ���
                            Else
                                �ثe��(22) = 33
                                FormMainMode.���ݮɶ�.Enabled = True
                            End If
                        End If
                    End If
            End If
        '=====================================================
        Case 5
            FormMainMode.tr�q���P_��P.Enabled = True
        '=====================================================
        Case 6
            If atking_AI_��ܵY_�����ۺh���q������(0, 1) = 11 Then
                    Do
                        If atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 1 Then
                            atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 0
                            atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                            �ثe��(21) = 10
                            FormMainMode.tr�ϥΪ�_��P.Enabled = True
                            Exit Sub
                        End If
                        �ثe��(20) = �ثe��(20) + 1
                    Loop Until �ثe��(20) > 106
                    '=========�q���P�ƧP�_�ο��
�ĪG11_���ܹq���P�_:
                    �ثe��(16) = 1
                    If Val(FormMainMode.pagecomglead) > 8 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 2
                        Erase cardp
                        Erase cardpn
                        For i = 1 To 106
                            atking_AI_��ܵY_�����ۺh���q������(i, 1) = 0
                        Next
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_��ܵY_�����ۺh���q������(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pagecomglead) - 8 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 1 Then
                                �ثe��(17) = 12
                                atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 0
                                atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                                FormMainMode.tr�q���P_½�P.Enabled = True
                                Exit Sub
                            End If
                            �ثe��(16) = �ثe��(16) + 1
                        Loop Until �ثe��(16) > 106
                    ElseIf Val(FormMainMode.pagecomglead) < 8 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 4
                        atking_AI_��ܵY_�����ۺh���q������(0, 3) = 8
                        atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                        If Val(FormMainMode.pageul) < 8 - Val(FormMainMode.pagecomglead) Then
                            �԰��t����.����ʧ@_�~�P
                        End If
                        �ثe��(15) = 33
                        FormMainMode.tr�P��_��P_�q��.Enabled = True
                    Else
'                        FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�8�i�C"
'                        �԰��t����.�۰ʱ��b����
'                        atkingckai(108, 2) = 0
'                        ����ʧ@_�ޯ��ʵ���
                        If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
                            GoTo �ĪG�������_��P�ܤ���
                        Else
                            �ثe��(22) = 33
                            FormMainMode.���ݮɶ�.Enabled = True
                        End If
                    End If
            End If
        '=====================================================
        Case 7
            If atking_AI_��ܵY_�����ۺh���q������(0, 1) = 11 Then
                 Do
                     If atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 1 Then
                         �ثe��(17) = 12
                         atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 0
                         atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                         FormMainMode.tr�q���P_½�P.Enabled = True
                         Exit Sub
                     End If
                     �ثe��(16) = �ثe��(16) + 1
                 Loop Until �ثe��(16) > 106
                 If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
'                        FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�8�i�C"
'                        �԰��t����.�۰ʱ��b����
'                        atkingckai(108, 2) = 0
'                        ����ʧ@_�ޯ��ʵ���
                        GoTo �ĪG�������_��P�ܤ���
                 Else
                        �ثe��(22) = 33
                        FormMainMode.���ݮɶ�.Enabled = True
                 End If
             End If
        '=====================================================
        Case 8
            If atking_AI_��ܵY_�����ۺh���q������(0, 1) = 11 Then
                Select Case atking_AI_��ܵY_�����ۺh���q������(0, 2)
                    Case 3
                        If Val(FormMainMode.pageusglead) < atking_AI_��ܵY_�����ۺh���q������(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           �ثe��(15) = 33
                           atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                           FormMainMode.tr�P��_��P_�ϥΪ�.Enabled = True
                        End If
                        If Val(FormMainMode.pageusglead) >= atking_AI_��ܵY_�����ۺh���q������(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                           GoTo �ĪG11_���ܹq���P�_
                        End If
                    Case 4
                        If Val(FormMainMode.pagecomglead) < atking_AI_��ܵY_�����ۺh���q������(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           �ثe��(15) = 33
                           atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                           FormMainMode.tr�P��_��P_�q��.Enabled = True
                        End If
                        If Val(FormMainMode.pagecomglead) >= atking_AI_��ܵY_�����ۺh���q������(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
'                           FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�8�i�C"
'                           �԰��t����.�۰ʱ��b����
'                            atkingckai(108, 2) = 0
'                            ����ʧ@_�ޯ��ʵ���
                            If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
                                GoTo �ĪG�������_��P�ܤ���
                            Else
                                �ثe��(22) = 33
                                FormMainMode.���ݮɶ�.Enabled = True
                            End If
                        End If
                 End Select
            End If
        '=====================================================
        Case 9
            If atking_AI_��ܵY_�����ۺh���q������(0, 1) = 12 Then
                    Do
                        If atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 1 Then
                            atking_AI_��ܵY_�����ۺh���q������(�ثe��(20), 1) = 0
                            atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                            �ثe��(21) = 10
                            FormMainMode.tr�ϥΪ�_��P.Enabled = True
                            Exit Sub
                        End If
                        �ثe��(20) = �ثe��(20) + 1
                    Loop Until �ثe��(20) > 106
                    '=========�q���P�ƧP�_�ο��
�ĪG12_���ܹq���P�_:
                    �ثe��(16) = 1
                    If Val(FormMainMode.pagecomglead) > 15 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 2
                        Erase cardp
                        Erase cardpn
                        For i = 1 To 106
                            atking_AI_��ܵY_�����ۺh���q������(i, 1) = 0
                        Next
                        Do
                           Randomize
                           i = Int(Rnd() * 106) + 1
                           If cardp(i) = False Then
                                cardp(i) = True
                                cardpn(1) = cardpn(1) + 1
                                If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 1 Then
                                    atking_AI_��ܵY_�����ۺh���q������(i, 1) = 1
                                    cardpn(2) = cardpn(2) + 1
                                End If
                           End If
                           If Val(FormMainMode.pagecomglead) - 15 = cardpn(2) Then
                               Exit Do
                           End If
                        Loop While cardpn(1) < 106
                        Do
                            If atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 1 Then
                                �ثe��(17) = 12
                                atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 0
                                atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                                FormMainMode.tr�q���P_½�P.Enabled = True
                                Exit Sub
                            End If
                            �ثe��(16) = �ثe��(16) + 1
                        Loop Until �ثe��(16) > 106
                    ElseIf Val(FormMainMode.pagecomglead) < 15 Then
                        atking_AI_��ܵY_�����ۺh���q������(0, 2) = 4
                        atking_AI_��ܵY_�����ۺh���q������(0, 3) = 15
                        atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                        If Val(FormMainMode.pageul) < 15 - Val(FormMainMode.pagecomglead) Then
                            �԰��t����.����ʧ@_�~�P
                        End If
                        �ثe��(15) = 33
                        FormMainMode.tr�P��_��P_�q��.Enabled = True
                    Else
'                        FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�15�i�C"
'                        �԰��t����.�۰ʱ��b����
'                        atkingckai(108, 2) = 0
'                        ����ʧ@_�ޯ��ʵ���
                        If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
                            GoTo �ĪG�������_��P�ܤ���
                        Else
                            �ثe��(22) = 33
                            FormMainMode.���ݮɶ�.Enabled = True
                        End If
                    End If
            End If
        '=====================================================
        Case 10
            If atking_AI_��ܵY_�����ۺh���q������(0, 1) = 12 Then
                Do
                    If atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 1 Then
                        �ثe��(17) = 12
                        atking_AI_��ܵY_�����ۺh���q������(�ثe��(16), 1) = 0
                        atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                        FormMainMode.tr�q���P_½�P.Enabled = True
                        Exit Sub
                    End If
                    �ثe��(16) = �ثe��(16) + 1
                Loop Until �ثe��(16) > 106
'                FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�15�i�C"
'                �԰��t����.�۰ʱ��b����
'                atkingckai(108, 2) = 0
'                ����ʧ@_�ޯ��ʵ���
                If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
                    GoTo �ĪG�������_��P�ܤ���
                Else
                    �ثe��(22) = 33
                    FormMainMode.���ݮɶ�.Enabled = True
                End If
            End If
        '=====================================================
       Case 11
           If atking_AI_��ܵY_�����ۺh���q������(0, 1) = 12 Then
                Select Case atking_AI_��ܵY_�����ۺh���q������(0, 2)
                    Case 3
                        If Val(FormMainMode.pageusglead) < atking_AI_��ܵY_�����ۺh���q������(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           �ثe��(15) = 33
                           atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                           FormMainMode.tr�P��_��P_�ϥΪ�.Enabled = True
                        End If
                        If Val(FormMainMode.pageusglead) >= atking_AI_��ܵY_�����ۺh���q������(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
                           GoTo �ĪG12_���ܹq���P�_
                        End If
                    Case 4
                        If Val(FormMainMode.pagecomglead) < atking_AI_��ܵY_�����ۺh���q������(0, 3) And Val(FormMainMode.pageul) > 0 Then
                           �ثe��(15) = 33
                           atking_AI_��ܵY_�����ۺh���q������(0, 4) = atking_AI_��ܵY_�����ۺh���q������(0, 4) + 1
                           FormMainMode.tr�P��_��P_�q��.Enabled = True
                        End If
                        If Val(FormMainMode.pagecomglead) >= atking_AI_��ܵY_�����ۺh���q������(0, 3) Or Val(FormMainMode.pageul) <= 0 Then
'                           FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�15�i�C"
'                           �԰��t����.�۰ʱ��b����
'                           atkingckai(108, 2) = 0
'                           ����ʧ@_�ޯ��ʵ���
                           If atking_AI_��ܵY_�����ۺh���q������(0, 4) >= 2 Then
                                GoTo �ĪG�������_��P�ܤ���
                            Else
                                �ثe��(22) = 33
                                FormMainMode.���ݮɶ�.Enabled = True
                            End If
                        End If
                 End Select
            End If
        Case 12
�ĪG�������_��P�ܤ���:
            '==============�����ޯ���(��P�ܤ���)
            Select Case atking_AI_��ܵY_�����ۺh���q������(0, 1)
                 Case 10
                        FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�0�i�C"
                        �԰��t����.�۰ʱ��b����
                        atkingckai(108, 2) = 0
                        ����ʧ@_�ޯ��ʵ���
                 Case 11
                        FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�8�i�C"
                        �԰��t����.�۰ʱ��b����
                        atkingckai(108, 2) = 0
                        ����ʧ@_�ޯ��ʵ���
                 Case 12
                        FormMainMode.messageus.AddItem "�����ۺh�o��! �ۤv�P��誺��P�ܬ�15�i�C"
                        �԰��t����.�۰ʱ��b����
                        atkingckai(108, 2) = 0
                        ����ʧ@_�ޯ��ʵ���
            End Select
   End Select
End If
End Sub
Sub ����_�o�����c()
Dim tn As Integer
If FormMainMode.comaiatk(1).Caption = "�o�����c" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(109, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(109, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 4) >= 2 And atkingckai(109, 2) = 0 Then
                   atkingckai(109, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 2 And atkingckai(109, 2) = 1 Then
                   atkingckai(109, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7455
                   atkingno(i, 6) = 9075
                   atkingno(i, 7) = 109
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
              '=====================================�����U�@�i�ƥ�d
                tn = Val(FormMainMode.turni) + 1
                If tn <= 18 Then
                    If tn <= �ƥ�d�O���Ȯɼ�(0, 1) Or Formsetting.persontgrecom.Value = 0 Then
                        If pageeventnum(2, tn, 1) <> "" Then
                            ay = Split(�@��t����.�ƥ�d��Ʈw(pageeventnum(2, tn, 1), 3), "=")
                            pagecardnum(88 + tn, 1) = ay(0)
                            pagecardnum(88 + tn, 2) = ay(1)
                            pagecardnum(88 + tn, 3) = ay(2)
                            pagecardnum(88 + tn, 4) = ay(3)
                            pagecardnum(88 + tn, 5) = 2
                            pagecardnum(88 + tn, 6) = 1
                            pagecardnum(88 + tn, 8) = pageeventnum(2, tn, 2)
                            FormMainMode.card(88 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(2, tn, 2) & "-1.bmp")
                            pagecardnum(88 + tn, 11) = 0
                            pageonin(88 + tn) = 1
                        End If
                    End If
                End If
             '=====================================
             If Val(FormMainMode.turni) < 18 And (tn <= �ƥ�d�O���Ȯɼ�(0, 1) Or Formsetting.persontgrecom.Value = 0) Then
                �ثe��(16) = 88 + Val(FormMainMode.turni) + 1
                atking_AI_����_�o�����c������ = 1
                �ثe��(15) = 34
                FormMainMode.tr�P��_�^�P_�q��.Enabled = True
            Else
                atkingckai(109, 2) = 0
            End If
        Case 4
            If Val(FormMainMode.turni) + atking_AI_����_�o�����c������ < 18 And atking_AI_����_�o�����c������ < 2 And _
               (tn <= �ƥ�d�O���Ȯɼ�(0, 1) Or Formsetting.persontgrecom.Value = 0) Then
                '=====================================�����U�@�i�ƥ�d
                tn = Val(FormMainMode.turni) + 2
                If tn <= 18 Then
                        If tn <= �ƥ�d�O���Ȯɼ�(0, 1) Or Formsetting.persontgrecom.Value = 0 Then
                            If pageeventnum(2, tn, 1) <> "" Then
                                ay = Split(�@��t����.�ƥ�d��Ʈw(pageeventnum(2, tn, 1), 3), "=")
                                pagecardnum(88 + tn, 1) = ay(0)
                                pagecardnum(88 + tn, 2) = ay(1)
                                pagecardnum(88 + tn, 3) = ay(2)
                                pagecardnum(88 + tn, 4) = ay(3)
                                pagecardnum(88 + tn, 5) = 2
                                pagecardnum(88 + tn, 6) = 1
                                pagecardnum(88 + tn, 8) = pageeventnum(2, tn, 2)
                                FormMainMode.card(88 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(2, tn, 2) & "-1.bmp")
                                pagecardnum(88 + tn, 11) = 0
                                pageonin(88 + tn) = 1
                            End If
                        End If
                End If
                atking_AI_����_�o�����c������ = atking_AI_����_�o�����c������ + 1
                �ثe��(16) = 88 + Val(FormMainMode.turni) + 2
                �ثe��(15) = 34
                FormMainMode.tr�P��_�^�P_�q��.Enabled = True
            Else
                FormMainMode.turni = Val(FormMainMode.turni) + atking_AI_����_�o�����c������
                turn = Val(FormMainMode.turni)
                atking_AI_����_�o�����c������ = 0
                atkingckai(109, 2) = 0
            End If
   End Select
End If
End Sub
Sub ����_�]���ɤ�()
Dim tn(1 To 3) As Boolean
If FormMainMode.comaiatk(4).Caption = "�]���ɤ�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(110, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
   Select Case atkingckai(110, 1)
        Case 1
             If pageqlead(2) >= 3 And atkingckai(110, 2) = 0 Then
               atkingckai(110, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf pageqlead(2) < 3 And atkingckai(110, 2) = 1 Then
               atkingckai(110, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -840
                   atkingno(i, 4) = -600
                   atkingno(i, 5) = 6870
                   atkingno(i, 6) = 10365
                   atkingno(i, 7) = 110
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
       Case 3
            atkingckai(110, 2) = 0
            '======================
            For i = 1 To 3
                If VBEPerson(2, ����ݾ��H��������(2, i), 1, 2, 1) = "R" Then
                     tn(i) = True
                Else
                     tn(i) = False
                End If
                 If tn(i) = True Then
                     Select Case Val(VBEPerson(2, ����ݾ��H��������(2, i), 1, 2, 2))
                         Case Is <= 2
                              �԰��t����.�^�_����_�q�� 1, i
                         Case Is > 2, Is <= 4
                              �԰��t����.�^�_����_�q�� 2, i
                         Case 5
                              �԰��t����.�^�_����_�q�� 3, i
                     End Select
                 End If
            Next
            '=============================
   End Select
End If
End Sub
Sub ������_��M�_���p()
If FormMainMode.comaiatk(4).Caption = "��M�_���p" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(113, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
   Select Case atkingckai(113, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(113, 2) = 0 Then
                   atkingckai(113, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(113, 2) = 1 Then
                   atkingckai(113, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -360
                   atkingno(i, 5) = 8745
                   atkingno(i, 6) = 10200
                   atkingno(i, 7) = 113
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(113, 2) = 0
             '======================
               Do
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 3) = 22 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                          FormMainMode.personusspe(i).person_turn = 1
                          �H�����`���A��Ʈw(1, i, 2) = 1
                          Exit Do
                      End If
                    Next
                    For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                       If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 1, i, 22, app_path & "gif\���`���A\atkingerr.gif", 0, 1
                          ���`���A�ˬd��(22, 1) = 1
                          ���`���A�ˬd��(22, 2) = 1
                          Exit Do
                       End If
                    Next
               Loop
   End Select
End If
End Sub

Sub �L���S_�]����()
If FormMainMode.comaiatk(3).Caption = "�]����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(114, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�L���S" Then
   Select Case atkingckai(114, 1)
      Case 1
            If movecp < 3 Then
                If atkingpagetot(2, 2) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(114, 2) = 0 Then
                   atkingckai(114, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 2) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(114, 2) = 1 Then
                   atkingckai(114, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
            End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�L���S\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7575
                   atkingno(i, 6) = 9660
                   atkingno(i, 7) = 114
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(114, 2) = 0
             '========================
             �԰��t����.�^�_����_�q�� 1, 1
             '========================
             For i = 18 To (turn + 3) Step -1
                  pageeventnum(2, i, 1) = pageeventnum(2, i - 2, 1)
                  pageeventnum(2, i, 2) = pageeventnum(2, i - 2, 2)
             Next
             For i = (turn + 1) To (turn + 2)
                  pageeventnum(2, i, 1) = "HP�^�_3"
                  pageeventnum(2, i, 2) = �@��t����.�ƥ�d��Ʈw("HP�^�_3", 2)
             Next
   End Select
End If
End Sub
Sub ������S_����ۼv()
Dim m As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "����ۼv" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(116, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "������S" Then
   Select Case atkingckai(116, 1)
        Case 1
            If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 1 And atkingpagetot(2, 3) = 0 And atkingckai(116, 2) = 0 Then
'            If atkingpagetot(2, 3) >= 1 And atkingckai(116, 2) = 0 Then
               atkingckai(116, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            ElseIf (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 1 Or atkingpagetot(2, 3) > 0) And atkingckai(116, 2) = 1 Then
'            ElseIf atkingpagetot(2, 3) < 1 And atkingckai(116, 2) = 1 Then
               atkingckai(116, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
            End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\������S\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6465
                   atkingno(i, 6) = 10455
                   atkingno(i, 7) = 116
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
                atkingckai(116, 2) = 0
                Select Case movecp
                    Case 1
                       Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 29 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 3
                                  �H�����`���A��Ʈw(1, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 29, app_path & "gif\���`���A\����.gif", 0, 3
                                  ���`���A�ˬd��(29, 1) = 1
                                  ���`���A�ˬd��(29, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 2
                        Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 20 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 3
                                  �H�����`���A��Ʈw(1, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 20, app_path & "gif\���`���A\damage.gif", 0, 3
                                  ���`���A�ˬd��(20, 1) = 1
                                  ���`���A�ˬd��(20, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                    Case 3
                        Do
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                              If �H�����`���A��Ʈw(1, i, 3) = 27 And �H�����`���A��Ʈw(1, i, 2) > 0 Then
                                  FormMainMode.personusspe(i).person_turn = 3
                                  �H�����`���A��Ʈw(1, i, 2) = 3
                                  Exit Do
                              End If
                            Next
                            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                               If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 1, i, 27, app_path & "gif\���`���A\�g�Ԥh.gif", 0, 3
                                  ���`���A�ˬd��(27, 1) = 1
                                  ���`���A�ˬd��(27, 2) = 1
                                  Exit Do
                               End If
                            Next
                        Loop
                End Select
   End Select
End If
End Sub
Sub ����P��_SSS()
Dim rrr(1 To 3) As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "S.S.S" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(117, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "����P��" Then
   Select Case atkingckai(117, 1)
        Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                    If pagecardnum(i, 1) = a4a And pagecardnum(i, 2) = 1 Then
                       rrr(1) = rrr(1) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And pagecardnum(i, 2) = 2 Then
                       rrr(2) = rrr(2) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And pagecardnum(i, 2) = 3 Then
                       rrr(3) = rrr(3) + 1
                    End If
                End If
             Next
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1 And rrr(3) >= 1) And atkingckai(117, 2) = 0 Then
'             If pageqlead(2) >= 1 And atkingckai(117, 2) = 0 Then
                atkingckai(117, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or rrr(3) < 1) And atkingckai(117, 2) = 1 Then
'             ElseIf pageqlead(2) < 1 And atkingckai(117, 2) = 1 Then
                atkingckai(117, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
              End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\����P��\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6135
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 117
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
                atkingckai(117, 2) = 0
                Do
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                          If �H�����`���A��Ʈw(2, j, 3) = 32 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                              FormMainMode.personcomspe(j).person_turn = 3
                              �H�����`���A��Ʈw(2, j, 2) = 3
                              Exit Do
                          End If
                     Next
                    For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, j, 32, app_path & "gif\���`���A\�V�P.gif", 0, 3
                          ���`���A�ˬd��(32, 1) = 1
                          ���`���A�ˬd��(32, 2) = 1
                          Exit Do
                      End If
                    Next
                Loop
   End Select
End If
End Sub
Sub �h�g�H_�W�Ťk�D��()
If FormMainMode.comaiatk(3).Caption = "�W�Ťk�D��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(118, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�h�g�H" Then
   Select Case atkingckai(118, 1)
      Case 1
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 5) >= 3 And atkingpagetot(2, 4) >= 2 And atkingckai(118, 2) = 0 Then
               atkingckai(118, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 5) < 3 Or atkingpagetot(2, 4) < 2) And atkingckai(118, 2) = 1 Then
               atkingckai(118, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�h�g�H\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -480
                   atkingno(i, 5) = 5970
                   atkingno(i, 6) = 10365
                   atkingno(i, 7) = 118
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            atkingckai(118, 2) = 0
            '==================
            Do
                For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                      If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                          FormMainMode.personcomspe(j).person_num = 6
                          FormMainMode.personcomspe(j).person_turn = 5
                          �H�����`���A��Ʈw(2, j, 1) = 6
                          �H�����`���A��Ʈw(2, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                   If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                      �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 6, 5
                      ���`���A�ˬd��(1, 1) = 1
                      ���`���A�ˬd��(1, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
            Do
                For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                      If �H�����`���A��Ʈw(2, j, 3) = 2 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                          FormMainMode.personcomspe(j).person_num = 4
                          FormMainMode.personcomspe(j).person_turn = 5
                          �H�����`���A��Ʈw(2, j, 1) = 4
                          �H�����`���A��Ʈw(2, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                   If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                      �԰��t����.�H�����`���A��]�w_��] 2, j, 2, app_path & "gif\���`���A\defup.gif", 4, 5
                      ���`���A�ˬd��(2, 1) = 1
                      ���`���A�ˬd��(2, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
            Do
                For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                      If �H�����`���A��Ʈw(2, j, 3) = 3 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                          FormMainMode.personcomspe(j).person_num = 1
                          FormMainMode.personcomspe(j).person_turn = 5
                          �H�����`���A��Ʈw(2, j, 1) = 1
                          �H�����`���A��Ʈw(2, j, 2) = 5
                          Exit Do
                      End If
                 Next
                For j = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                   If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                      �԰��t����.�H�����`���A��]�w_��] 2, j, 3, app_path & "gif\���`���A\movup.gif", 1, 5
                      ���`���A�ˬd��(3, 1) = 1
                      ���`���A�ˬd��(3, 2) = 1
                      Exit Do
                  End If
                Next
            Loop
   End Select
End If
End Sub
Sub �Ǧh_�]�G���u()
Dim m As Integer
If FormMainMode.comaiatk(1).Caption = "�]�G���u" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(119, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Ǧh" Then
   Select Case atkingckai(119, 1)
      Case 1
            If atkingpagetot(2, 4) >= 1 And atkingckai(119, 2) = 0 Then
               atkingckai(119, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 1 And atkingckai(119, 2) = 1 Then
               atkingckai(119, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǧh\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9255
                   atkingno(i, 7) = 119
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(119, 2) = 0
             Do
                Randomize
                m = Int(Rnd() * 106) + 1
                If Val(pagecardnum(m, 6)) = 1 And Val(pagecardnum(m, 5)) = 1 Then
                     �ثe��(20) = m
                     �ثe��(21) = 1
                     FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
                     Exit Do
                End If
            Loop
   End Select
End If
End Sub
Sub �Ǧh_�]�G����()
Dim m, n As Integer
Dim aw(1 To 2) As Integer
If FormMainMode.comaiatk(2).Caption = "�]�G����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(120, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Ǧh" Then
   Select Case atkingckai(120, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(120, 2) = 0 Then
               atkingckai(120, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(120, 2) = 1 Then
               atkingckai(120, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǧh\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7260
                   atkingno(i, 6) = 8925
                   atkingno(i, 7) = 120
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             �ޯ�ʵe��ܶ��q�� = 11
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.�������q_���q2.Enabled = False
        Case 3
                ���q���A�� = 1
                For m = 1 To 106
                    If Val(pagecardnum(m, 6)) = 2 And Val(pagecardnum(m, 5)) = 1 Then
                        Randomize
                        n = Int(Rnd() * 6) + 1
                        If n Mod 2 = 0 Then
                            FormMainMode.cqen_Click (m)
                        End If
                    End If
                Next
              atkingckai(120, 1) = 4
              FormMainMode.trgoi1_Timer
              FormMainMode.trgoi2_Timer
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(120, 2) = 0
   End Select
End If
End Sub
Sub �Ǧh_�]�G����()
Dim m As Integer
If FormMainMode.comaiatk(3).Caption = "�]�G����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(121, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Ǧh" Then
   Select Case atkingckai(121, 1)
      Case 1
            If atkingpagetot(2, 4) >= 4 And atkingckai(121, 2) = 0 Then
               atkingckai(121, 2) = 1
            End If
            If atkingpagetot(2, 4) < 4 And atkingckai(121, 2) = 1 Then
               atkingckai(121, 2) = 0
             End If
      Case 2
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_�Ǧh_�]�G����O����(i) = 1
                   atking_AI_�Ǧh_�]�G����O����(107) = atking_AI_�Ǧh_�]�G����O����(107) + 1
               End If
            Next
            atking_AI_�Ǧh_�]�G����O����(108) = 1
      Case 3
            atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 4
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�Ǧh\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7425
                   atkingno(i, 6) = 9570
                   atkingno(i, 7) = 121
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
            Do Until atking_AI_�Ǧh_�]�G����O����(108) > 106
                If atking_AI_�Ǧh_�]�G����O����(atking_AI_�Ǧh_�]�G����O����(108)) = 1 Then
                    �ثe��(16) = atking_AI_�Ǧh_�]�G����O����(108)
                    �ثe��(15) = 35
                    FormMainMode.tr�P��_�^�P_�q��.Enabled = True
                    atking_AI_�Ǧh_�]�G����O����(�ثe��(16)) = 0
                    Exit Do
                End If
                atking_AI_�Ǧh_�]�G����O����(108) = atking_AI_�Ǧh_�]�G����O����(108) + 1
            Loop
            If atking_AI_�Ǧh_�]�G����O����(108) >= 106 Then
                If atking_AI_�Ǧh_�]�G����O����(107) < 2 Then
                    atking_AI_�Ǧh_�]�G����O����(107) = atking_AI_�Ǧh_�]�G����O����(107) + 1
                    �ثe��(22) = 34
                    FormMainMode.���ݮɶ�.Enabled = True
                ElseIf atking_AI_�Ǧh_�]�G����O����(107) >= 2 Then
                    atkingckai(121, 2) = 0
                    Erase atking_AI_�Ǧh_�]�G����O����
                    �԰��t����.����ʧ@_�ޯ��ʵ���
                End If
            End If
   End Select
End If
End Sub

Sub ���Y�F_����()
Dim bloodnum As Integer
If FormMainMode.comaiatk(1).Caption = "����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(122, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���Y�F" Then
   Select Case atkingckai(122, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(122, 2) = 0 Then
                atkingckai(122, 2) = 1
             End If
             If atkingpagetot(2, 4) < 2 And atkingckai(122, 2) = 1 Then
                atkingckai(122, 2) = 0
              End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���Y�F\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7035
                   atkingno(i, 6) = 9510
                   atkingno(i, 7) = 122
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_AI_���Y�F_����_��P������
                   Select Case liveus(����H����ԤH��(1, 2))
                       Case Is = liveusmax(����H����ԤH��(1, 2))
                           atking_AI_���Y�F_����_��P������(2) = 4
                           atkingno(i, 11) = 1
                       Case Else
                           atking_AI_���Y�F_����_��P������(2) = 2
                           atkingno(i, 11) = 0
                    End Select
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_AI_���Y�F_����_��P������(2) And atking_AI_���Y�F_����_��P������(1) = 0 Then
               �԰��t����.����ʧ@_�~�P
            End If
            atking_AI_���Y�F_����_��P������(1) = atking_AI_���Y�F_����_��P������(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_���Y�F_����_��P������(1) > atking_AI_���Y�F_����_��P������(2)
                    �ثe��(15) = 36
                    FormMainMode.tr�P��_��P_�q��.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_���Y�F_����_��P������(1) > atking_AI_���Y�F_����_��P������(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_AI_���Y�F_����_��P������(2)) <= 2 Then
                   atkingckai(122, 2) = 0
               Else
                   �ثe��(24) = 35
                   FormMainMode.���ݮɶ�_2.Enabled = True
               End If
            End If
        Case 5
            atkingckai(122, 2) = 0
            �԰��t����.����ʧ@_�ޯ��ʵ���
   End Select
End If
End Sub
Sub ���Y�F_��������()
If FormMainMode.comaiatk(2).Caption = "��������" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(123, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���Y�F" Then
   Select Case atkingckai(123, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(2, 2) >= 2 And atkingpagetot(2, 4) >= 2 And atkingckai(123, 2) = 0 Then
               atkingckai(123, 2) = 1
               �������m��l�`��(2) = �������m��l�`��(2) + 4
            ElseIf (atkingpagetot(2, 2) < 2 Or atkingpagetot(2, 4) < 2) And atkingckai(123, 2) = 1 Then
               atkingckai(123, 2) = 0
               �������m��l�`��(2) = �������m��l�`��(2) - 4
            End If
          End If
        Case 2
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_���Y�F_��������������A��(i) = True
               End If
            Next
            �ثe��(30) = 1
        Case 3
            atkingtrn(2) = Val(atkingtrn(2)) + 1
        Case 4
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���Y�F\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7290
                   atkingno(i, 6) = 9120
                   atkingno(i, 7) = 123
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
        Case 5
            Do
                If atking_AI_���Y�F_��������������A��(�ثe��(30)) = True Then
                    �ثe��(16) = �ثe��(30)
                    �ثe��(15) = 37
                    FormMainMode.tr�P��_�^�P_�q��.Enabled = True
                    atking_AI_���Y�F_��������������A��(�ثe��(16)) = False
                    Exit Do
                End If
                �ثe��(30) = �ثe��(30) + 1
            Loop Until �ثe��(30) >= 106
            If �ثe��(30) >= 106 Then
                If �ثe��(30) < 2 Then
                    �ثe��(30) = �ثe��(30) + 1
                    �ثe��(22) = 35
                    FormMainMode.���ݮɶ�.Enabled = True
                ElseIf �ثe��(30) >= 2 Then
                    atkingckai(123, 2) = 0
                    Erase atking_AI_���Y�F_��������������A��
                    �԰��t����.����ʧ@_�ޯ��ʵ���
                End If
            End If
   End Select
End If
End Sub
Sub ���Y�F_���a�B��()
Dim wtr As Integer, wert(1 To 3) As Boolean, wery As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "���a�B��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(124, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���Y�F" Then
   Select Case atkingckai(124, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 5) >= 4 And atkingpagetot(2, 4) >= 1 And atkingckai(124, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 4
                   atkingckai(124, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 5) < 4 Or atkingpagetot(2, 4) < 1) And atkingckai(124, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 4
                   atkingckai(124, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���Y�F\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7320
                   atkingno(i, 6) = 10170
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If Val(�Y����ˮ`��) > 0 Then
                 Do
                        wtr = Int(Rnd() * 3) + 1
                        If wert(wtr) = False Then
                            wert(wtr) = True
                            wery = wery + 1
                            If liveus(����ݾ��H��������(1, wtr)) > 0 Then
                                �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� 2, wtr
                                 Exit Do
                            End If
                        End If
                 Loop Until wery > 3
             End If
             atkingckai(124, 2) = 0
   End Select
End If
End Sub
Sub ���Y�F_����B()
If FormMainMode.comaiatk(4).Caption = "����B" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(125, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���Y�F" Then
   Select Case atkingckai(125, 1)
      Case 1
           If movecp = 1 Then
                If atkingpagetot(2, 1) >= 1 And atkingpagetot(2, 5) >= 1 _
                   And atkingpagetot(2, 4) >= 1 And atkingpagetot(2, 3) >= 1 And atkingckai(125, 2) = 0 Then
                        �������m��l�`��(2) = �������m��l�`��(2) + 10
                        atkingckai(125, 2) = 1
                        atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 1) < 1 Or atkingpagetot(2, 5) < 1 _
                    Or atkingpagetot(2, 4) < 1 Or atkingpagetot(2, 3) < 1) And atkingckai(125, 2) = 1 Then
                        �������m��l�`��(2) = �������m��l�`��(2) - 10
                        atkingckai(125, 2) = 0
                        atkingtrn(2) = Val(atkingtrn(2)) - 1
                        If atking_AI_���Y�F_����B_�����O�[��������(2) = True Then
                              �������m��l�`��(2) = �������m��l�`��(2) - 15
                              atking_AI_���Y�F_����B_�����O�[��������(2) = False
                        End If
                        If atking_AI_���Y�F_����B_�����O�[��������(1) = True Then
                             �������m��l�`��(2) = �������m��l�`��(2) - 10
                             atking_AI_���Y�F_����B_�����O�[��������(1) = False
                        End If
                End If
                 '=====================
                 If atkingckai(125, 2) = 1 Then
                     If pageqlead(2) >= 10 And atking_AI_���Y�F_����B_�����O�[��������(1) = False Then
                         �������m��l�`��(2) = �������m��l�`��(2) + 10
                         atking_AI_���Y�F_����B_�����O�[��������(1) = True
                     ElseIf pageqlead(2) < 10 And atking_AI_���Y�F_����B_�����O�[��������(1) = True Then
                         �������m��l�`��(2) = �������m��l�`��(2) - 10
                         atking_AI_���Y�F_����B_�����O�[��������(1) = False
                     End If
                     If pageqlead(2) >= 15 And atking_AI_���Y�F_����B_�����O�[��������(2) = False Then
                         �������m��l�`��(2) = �������m��l�`��(2) + 15
                         atking_AI_���Y�F_����B_�����O�[��������(2) = True
                     ElseIf pageqlead(2) < 15 And atking_AI_���Y�F_����B_�����O�[��������(2) = True Then
                         �������m��l�`��(2) = �������m��l�`��(2) - 15
                         atking_AI_���Y�F_����B_�����O�[��������(2) = False
                     End If
                 End If
          End If
      Case 2
             atkingckai(125, 2) = 0
             Erase atking_AI_���Y�F_����B_�����O�[��������
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���Y�F\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8205
                   atkingno(i, 6) = 10080
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
   End Select
End If
End Sub
Sub ��_�w�_���������q()
Dim rrr As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "�w�-���������q" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(126, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
   Select Case atkingckai(126, 1)
        Case 1
             For i = 1 To 106
                If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 3 And Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                   rrr = rrr + 1
                End If
             Next
          If rrr >= 1 And atkingckai(126, 2) = 0 Then
             atkingckai(126, 2) = 1
             atkingtrn(2) = Val(atkingtrn(2)) + 1
          End If
          If rrr < 1 And atkingckai(126, 2) = 1 Then
             atkingckai(126, 2) = 0
             atkingtrn(2) = Val(atkingtrn(2)) - 1
           End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��\��-�w�-���������q_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -360
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8835
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             atkingckai(126, 2) = 0
             If livecom(����H����ԤH��(2, 2)) <= 0 Then
                 For i = 2 To 3
                     If livecom(����ݾ��H��������(2, i)) > 0 Then
                        Do
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                                  If �H�����`���A��Ʈw(2, j, 3) = 1 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 5
                                      FormMainMode.personcomspe(j).person_turn = 3
                                      �H�����`���A��Ʈw(2, j, 1) = 5
                                      �H�����`���A��Ʈw(2, j, 2) = 3
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                               If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 2, j, 1, app_path & "gif\���`���A\atkup.gif", 5, 3
                                  ���`���A�ˬd��(1, 1) = 1
                                  ���`���A�ˬd��(1, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                        ''===========================================
                        Do
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                                  If �H�����`���A��Ʈw(2, j, 3) = 2 And �H�����`���A��Ʈw(2, j, 2) > 0 Then
                                      FormMainMode.personcomspe(j).person_num = 5
                                      FormMainMode.personcomspe(j).person_turn = 3
                                      �H�����`���A��Ʈw(2, j, 1) = 5
                                      �H�����`���A��Ʈw(2, j, 2) = 3
                                      Exit Do
                                  End If
                             Next
                            For j = 14 * (����ݾ��H��������(2, i) - 1) + 1 To 14 * ����ݾ��H��������(2, i)
                               If �H�����`���A��Ʈw(2, j, 2) = 0 Then
                                  �԰��t����.�H�����`���A��]�w_��] 2, j, 2, app_path & "gif\���`���A\defup.gif", 5, 3
                                  ���`���A�ˬd��(2, 1) = 1
                                  ���`���A�ˬd��(2, 2) = 1
                                  Exit Do
                              End If
                            Next
                        Loop
                     End If
                Next
            End If
   End Select
End If
End Sub
Sub ��_EX_�צ�_�L�ɽ��j���׵�()
Dim num(1 To 2) As Integer '��ܤH���Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "Ex�צ�-�L�ɽ��j���׵�" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(127, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "��" Then
   Select Case atkingckai(127, 1)
        Case 1
          If movecp < 3 Then
            If atkingpagetot(2, 4) >= 6 And atkingckai(127, 2) = 0 Then
               atkingckai(127, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 18
            ElseIf atkingpagetot(2, 4) < 6 And atkingckai(127, 2) = 1 Then
               atkingckai(127, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 18
            End If
          End If
        Case 2
             atking_AI_��_�צ�_�L�ɽ��j���׵������� = 0
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\��\��-EX-�צ�-�L�ɽ��j���׵�_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8655
                   atkingno(i, 6) = 0
                   atkingno(i, 7) = 127
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '=============
             atking_AI_��_�צ�_�L�ɽ��j���׵������� = atkingpagetot(1, 2)
        Case 3
             atkingckai(127, 2) = 0
             If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                 num(1) = 1
                 num(2) = liveus(����H����ԤH��(2, 2))
                 For i = 2 To 3
                    If liveus(����ݾ��H��������(1, i)) > 0 And liveus(����ݾ��H��������(1, i)) < num(2) Then
                        num(1) = i
                        num(2) = liveus(����ݾ��H��������(1, i))
                    End If
                Next
                �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� Val(�Y���淾�q�Ȯ��ܼ�(2)), num(1)
            End If
            '=================
            �԰��t����.�ˮ`����_�ޯઽ��_�ϥΪ� Val(atking_AI_��_�צ�_�L�ɽ��j���׵�������), 1
            �Y���淾�q�Ȯ��ܼ�(2) = 0
            �Y����ˮ`�� = 0
   End Select
End If
End Sub
Sub ù��Y_�����ۼv()
If FormMainMode.comaiatk(1).Caption = "�����ۼv" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(128, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "ù��Y" Then
   Select Case atkingckai(128, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 2 And atkingckai(128, 2) = 0 Then
               atkingckai(128, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 5
            ElseIf (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 2) And atkingckai(128, 2) = 1 Then
               atkingckai(128, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 5
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\ù��Y\ù��Y_�����ۼv_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 128
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_ù��Y_�����ۼv�������A��(i) = True
               End If
            Next
            �ثe��(18) = 1
        Case 4
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) <= 0 Then
                Do
                    If atking_AI_ù��Y_�����ۼv�������A��(�ثe��(18)) = True Then
                        �ثe��(16) = �ثe��(18)
                        �ثe��(15) = 38
                        FormMainMode.tr�P��_�^�P_�q��.Enabled = True
                        atking_AI_ù��Y_�����ۼv�������A��(�ثe��(16)) = False
                        Exit Do
                    End If
                    �ثe��(18) = �ثe��(18) + 1
                Loop Until �ثe��(18) >= 106
            End If
            If �ثe��(18) >= 106 Or Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                atkingckai(128, 1) = 6
                FormMainMode.��l���槹�Ұ�.Enabled = True
            End If
        Case 5
'            tr�P��_�^�P_�ϥΪ�.Enabled = True
'            atking_AI_ù��Y_�����ۼv�������A��(�ثe��(16)) = False
        Case 6
            atkingckai(128, 2) = 0
            Erase atking_AI_ù��Y_�����ۼv�������A��
   End Select
End If
End Sub
Sub ù��Y_EX_�����ۼv()
If FormMainMode.comaiatk(1).Caption = "Ex�����ۼv" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(129, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "ù��Y" Then
   Select Case atkingckai(129, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 2) >= 4 And atkingpagetot(2, 4) >= 2 And atkingckai(129, 2) = 0 Then
               atkingckai(129, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 9
            ElseIf (atkingpagetot(2, 2) < 4 Or atkingpagetot(2, 4) < 2) And atkingckai(129, 2) = 1 Then
               atkingckai(129, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 9
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\ù��Y\ù��Y_Ex-�����ۼv_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = -480
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6765
                   atkingno(i, 6) = 9600
                   atkingno(i, 7) = 129
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 2 Then
                   atking_AI_ù��Y_�����ۼv�������A��(i) = True
               End If
            Next
            �ثe��(18) = 1
        Case 4
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) <= 0 Then
                Do
                    If atking_AI_ù��Y_�����ۼv�������A��(�ثe��(18)) = True Then
                        �ثe��(16) = �ثe��(18)
                        �ثe��(15) = 38
                        FormMainMode.tr�P��_�^�P_�q��.Enabled = True
                        atking_AI_ù��Y_�����ۼv�������A��(�ثe��(16)) = False
                        Exit Do
                    End If
                    �ثe��(18) = �ثe��(18) + 1
                Loop Until �ثe��(18) >= 106
            End If
            If �ثe��(18) >= 106 Or Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                atkingckai(129, 1) = 6
                FormMainMode.��l���槹�Ұ�.Enabled = True
            End If
        Case 5
'            tr�P��_�^�P_�ϥΪ�.Enabled = True
'            atking_AI_ù��Y_�����ۼv�������A��(�ثe��(16)) = False
        Case 6
            atkingckai(129, 2) = 0
            Erase atking_AI_ù��Y_�����ۼv�������A��
   End Select
End If
End Sub
Sub �����g_�f��ԧ����j�T()
Dim bloodnum As Integer
If FormMainMode.comaiatk(1).Caption = "�f��ԧ����j�T" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(130, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����g" Then
   Select Case atkingckai(130, 1)
      Case 1
           If movecp > 1 Then
                If atkingpagetot(2, 4) >= 3 And atkingckai(130, 2) = 0 Then
                   atkingckai(130, 2) = 1
                End If
                If atkingpagetot(2, 4) < 3 And atkingckai(130, 2) = 1 Then
                   atkingckai(130, 2) = 0
                 End If
          End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�����g\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = -120
                   atkingno(i, 5) = 7035
                   atkingno(i, 6) = 9540
                   atkingno(i, 7) = 130
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   '================
                   Erase atking_AI_�����g_�f��ԧ����j�T_��P������
                   atking_AI_�����g_�f��ԧ����j�T_��P������(2) = livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2))
                   If atking_AI_�����g_�f��ԧ����j�T_��P������(2) > 2 Then
                       atkingno(i, 11) = 1
                   Else
                       atkingno(i, 11) = 0
                   End If
                   '================
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
        Case 4
             If Val(FormMainMode.pageul.Caption) < atking_AI_�����g_�f��ԧ����j�T_��P������(2) And atking_AI_�����g_�f��ԧ����j�T_��P������(1) = 0 Then
               �԰��t����.����ʧ@_�~�P
            End If
            atking_AI_�����g_�f��ԧ����j�T_��P������(1) = atking_AI_�����g_�f��ԧ����j�T_��P������(1) + 1
            If Val(FormMainMode.pageul.Caption) > 0 Then
                Do Until atking_AI_�����g_�f��ԧ����j�T_��P������(1) > atking_AI_�����g_�f��ԧ����j�T_��P������(2)
                    �ثe��(15) = 39
                    FormMainMode.tr�P��_��P_�q��.Enabled = True
                    Exit Sub
                Loop
            End If
            If atking_AI_�����g_�f��ԧ����j�T_��P������(1) > atking_AI_�����g_�f��ԧ����j�T_��P������(2) Or Val(FormMainMode.pageul.Caption) <= 0 Then
               If Val(atking_AI_�����g_�f��ԧ����j�T_��P������(2)) <= 2 Then
                   atkingckai(130, 2) = 0
               Else
                   �ثe��(24) = 36
                   FormMainMode.���ݮɶ�_2.Enabled = True
               End If
            End If
        Case 5
            atkingckai(130, 2) = 0
            �԰��t����.����ʧ@_�ޯ��ʵ���
   End Select
End If
End Sub
Sub �J�y_�Ѩ����()
If FormMainMode.comaiatk(1).Caption = "�Ѩ����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(131, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�J�y" Then
   Select Case atkingckai(131, 1)
      Case 1
            If atkingpagetot(2, 4) >= 2 And atkingckai(131, 2) = 0 Then
'            If pageqlead(1) >= 1 And atkingckai(131, 2) = 0 Then
               atkingckai(131, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
            End If
            If atkingpagetot(2, 4) < 2 And atkingckai(131, 2) = 1 Then
'            If pageqlead(1) < 1 And atkingckai(131, 2) = 1 Then
               atkingckai(131, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
             End If
      Case 2
             FormMainMode.atkingnumtot.Caption = 1
             Erase atkingno
             '===========================
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�J�y\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 5655
                   atkingno(i, 6) = 9855
                   atkingno(i, 7) = 131
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             '===========================
             �ޯ�ʵe��ܶ��q�� = 11
             FormMainMode.atkingtrtot.Interval = 600
             FormMainMode.atkingtrtot.Enabled = True
             FormMainMode.�������q_���q2.Enabled = False
        Case 3
                ���q���A�� = 1
                �ثe��(21) = 1
                If pageqlead(1) > 0 Then
                    Do
                        Randomize
                        m = Int(Rnd() * 106) + 1
                        If Val(pagecardnum(m, 6)) = 2 And Val(pagecardnum(m, 5)) = 1 Then
                            atking_AI_�J�y_�Ѩ����_�ܵP������(1) = m
                            turnpageoninatking = 1
                            atkingckai(131, 1) = 5
                            FormMainMode.card_Click (m)
                            Exit Do
                        End If
                    Loop
                    FormMainMode.trgoi1_Timer
                    FormMainMode.trgoi2_Timer
                 Else
                    atkingtrn(2) = Val(atkingtrn(2)) - 1
                    atkingckai(131, 2) = 0
                    turnpageoninatking = 0
                    Erase atking_AI_�J�y_�Ѩ����_�ܵP������
                End If
        Case 4
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             atkingckai(131, 2) = 0
             turnpageoninatking = 0
             Erase atking_AI_�J�y_�Ѩ����_�ܵP������
        Case 5
             �ثe��(21) = 1
             atkingckai(131, 1) = 4
             atking_AI_�J�y_�Ѩ����_�ܵP������(2) = �ثe��(5)
             '=========�N�y�Ы��w�ܹq����P
             �԰��t����.�y�Эp��_�q����P
             �԰��t����.����ʧ@_�ϥΪ̵P_���P_�q�� atking_AI_�J�y_�Ѩ����_�ܵP������(1)
             �ثe��(5) = atking_AI_�J�y_�Ѩ����_�ܵP������(2)
             �ثe��(15) = 0
   End Select
End If
End Sub
Sub �J�y_�k�`�p�e()
Dim rrr(1 To 2) As Integer '�P�P�_�Ȯ��ܼ�
Dim au As Integer
If FormMainMode.comaiatk(2).Caption = "�k�`�p�e" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(132, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�J�y" Then
   Select Case atkingckai(132, 1)
      Case 1
            For i = 1 To 106
                If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                    If pagecardnum(i, 1) = a2a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(1) = rrr(1) + 1
                    End If
                    If pagecardnum(i, 1) = a4a And Val(pagecardnum(i, 2)) = 1 Then
                       rrr(2) = rrr(2) + 1
                    End If
                End If
             Next
             '========================
             If (rrr(1) >= 1 And rrr(2) >= 1) And atkingckai(132, 2) = 0 Then
'             If pageqlead(1) >= 1 And atkingckai(132, 2) = 0 Then
                atkingckai(132, 2) = 1
                atkingtrn(2) = Val(atkingtrn(2)) + 1
             ElseIf (rrr(1) < 1 Or rrr(2) < 1) And atkingckai(132, 2) = 1 Then
'             ElseIf pageqlead(1) < 1 And atkingckai(132, 2) = 1 Then
                atkingckai(132, 2) = 0
                atkingtrn(2) = Val(atkingtrn(2)) - 1
              End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�J�y\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7200
                   atkingno(i, 6) = 9990
                   atkingno(i, 7) = 132
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
               Randomize
               m = Int(Rnd() * 2) + 2
               au = 1
               Do
                    If livecom(����ݾ��H��������(2, m)) > 0 Then
                        �԰��t����.�ˮ`����_�ޯઽ��_�q�� 3, m
                        Exit Do
                    End If
                    If au < 2 Then
                        au = au + 1
                        If m = 2 Then
                            m = 3
                        Else
                            m = 2
                        End If
                    Else
                        �԰��t����.�ˮ`����_�ޯઽ��_�q�� 3, 1
                        Exit Do
                    End If
               Loop
               �Y����ˮ`�� = 0
               �Y���淾�q�Ȯ��ܼ�(2) = 0
               atkingckai(132, 2) = 0
   End Select
End If
End Sub
Sub �J�y_�����g��()
Dim p, i, j As Integer
If FormMainMode.comaiatk(3).Caption = "�����g��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(133, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�J�y" Then
   Select Case atkingckai(133, 1)
      Case 1
         If movecp > 1 Then
            If atkingpagetot(2, 5) >= 2 And atkingpagetot(2, 3) >= 1 And atkingckai(133, 2) = 0 Then
               atkingckai(133, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 4
            End If
            If (atkingpagetot(2, 5) < 2 Or atkingpagetot(2, 3) < 1) And atkingckai(133, 2) = 1 Then
               atkingckai(133, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 4
             End If
         End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�J�y\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6555
                   atkingno(i, 6) = 10110
                   atkingno(i, 7) = 0
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
             atkingckai(133, 1) = 3
        Case 3
             If livecom(����H����ԤH��(2, 2)) = livecommax(����H����ԤH��(2, 2)) Then
                    atking_AI_�J�y_�����g����q������(1) = �Y����ˮ`��
                    �Y���淾�q�Ȯ��ܼ�(2) = 0
                    �Y���淾�q�Ȯ��ܼ�(3) = 0
                    '========================================
                       For p = 1 To Val(FormMainMode.��ܦC1.goi1)
                          Randomize Timer
                          i = Int(Rnd() * 6) + 1
                          If i = 1 Or i = 6 Then �Y���淾�q�Ȯ��ܼ�(2) = Val(�Y���淾�q�Ȯ��ܼ�(2)) + 1
                       Next
                       For p = 1 To Val(FormMainMode.��ܦC1.goi2)
                          Randomize Timer
                          j = Int(Rnd() * 6) + 1
                          If j = 1 Or j = 6 Then �Y���淾�q�Ȯ��ܼ�(3) = Val(�Y���淾�q�Ȯ��ܼ�(3)) + 1
                       Next
                       '=============================
                       �ޯ�ʵe��ܶ��q�� = 1
                       atkingckai(133, 1) = 4
                       FormMainMode.��l���槹�Ұ�.Enabled = False
                       �ثe��(22) = 12
                       FormMainMode.���ݮɶ�.Enabled = True
                Else
                       atkingckai(133, 2) = 0
                       FormMainMode.��l���槹�Ұ�.Enabled = True
                       Erase atking_AI_�J�y_�����g����q������
                End If
          Case 4
                atking_AI_�J�y_�����g����q������(2) = �Y����ˮ`��
                '==========================
                �Y���淾�q�Ȯ��ܼ�(2) = atking_AI_�J�y_�����g����q������(1) + atking_AI_�J�y_�����g����q������(2)
                �Y����ˮ`�� = Val(�Y���淾�q�Ȯ��ܼ�(2))
                atkingckai(133, 2) = 0
                Erase atking_AI_�J�y_�����g����q������
   End Select
End If
End Sub
Sub �J�y_�c�N����()
Dim rrr(1 To 2) As Integer '�P�P�_�Ȯ��ܼ�
Dim au As Integer
If FormMainMode.comaiatk(4).Caption = "�c�N����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(134, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�J�y" Then
   Select Case atkingckai(134, 1)
      Case 1
            If movecp > 1 Then
                For i = 1 To 106
                    If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 Then
                        If pagecardnum(i, 1) = a1a And Val(pagecardnum(i, 2)) = 3 Then
                           rrr(1) = rrr(1) + 1
                        End If
                        If pagecardnum(i, 1) = a5a And Val(pagecardnum(i, 2)) = 3 Then
                           rrr(2) = rrr(2) + 1
                        End If
                    End If
                 Next
                 '========================
                 If (rrr(1) >= 1 And rrr(2) >= 1) And atkingpagetot(2, 4) >= 2 And atkingckai(134, 2) = 0 Then
'                 If pageqlead(2) >= 1 And atkingckai(134, 2) = 0 Then
                    atkingckai(134, 2) = 1
                 ElseIf (rrr(1) < 1 Or rrr(2) < 1 Or atkingpagetot(2, 4) < 2) And atkingckai(134, 2) = 1 Then
'                 ElseIf pageqlead(2) < 1 And atkingckai(134, 2) = 1 Then
                    atkingckai(134, 2) = 0
                  End If
            End If
      Case 2
             atkingtrn(2) = Val(atkingtrn(2)) + 1
      Case 3
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�J�y\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7050
                   atkingno(i, 6) = 10005
                   atkingno(i, 7) = 134
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 1
                   Exit For
                 End If
             Next
             atkingtrn(2) = Val(atkingtrn(2)) - 1
             '=====================
              For i = 1 To 106
                   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
                      atking_AI_�J�y_�c�N����������(i) = 1
                      atking_AI_�J�y_�c�N����������(0) = Val(atking_AI_�J�y_�c�N����������(0)) + 1
                   End If
               Next
        Case 4
               For i = 1 To 106
                   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
                      turnpageoninatking = 1
                      ���q���A�� = 1
                      FormMainMode.card_Click (i)
                      �ثe��(15) = 40
                      FormMainMode.��������ˬd.Enabled = False
                      Exit Sub
                    End If
               Next
               If i = 107 And atking_AI_�J�y_�c�N����������(0) > 0 Then
                    For k = 1 To 106
                         If atking_AI_�J�y_�c�N����������(k) = 1 Then
                             atking_AI_�J�y_�c�N����������(k) = 0
                             turnpageoninatking = 1
                             ���q���A�� = 1
                             FormMainMode.card_Click (k)
                             �ثe��(21) = 11
                             FormMainMode.��������ˬd.Enabled = False
                             Exit Sub
                         End If
                    Next
                End If
         Case 5
               turnpageonin = 0
               For k = 1 To 106
                     If atking_AI_�J�y_�c�N����������(k) = 1 Then
                         atking_AI_�J�y_�c�N����������(k) = 0
                         turnpageoninatking = 1
                         ���q���A�� = 1
                         FormMainMode.card_Click (k)
                         �ثe��(21) = 11
                         FormMainMode.��������ˬd.Enabled = False
                         Exit Sub
                     End If
                Next
                If k = 107 Then
                    atkingckai(134, 2) = 0
                    turnpageoninatking = 0
                    turnpageonin = 0
                    ���q���A�� = 4
                    Erase atking_AI_�J�y_�c�N����������
                    �԰��t����.����ʧ@_�ޯ��ʵ���
                End If
   End Select
End If
End Sub
Sub ���_�@����()
Dim cardnum(1 To 2) As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(1).Caption = "�@����" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(135, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "���" Then
   Select Case atkingckai(135, 1)
        Case 1
           If movecp = 2 Then
                 If atkingpagetot(2, 4) >= 3 And atkingckai(135, 2) = 0 Then
'                 If pageqlead(1) >= 1 And atkingckai(135, 2) = 0 Then
                   atkingckai(135, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                ElseIf atkingpagetot(2, 4) < 3 And atkingckai(135, 2) = 1 Then
'                ElseIf pageqlead(1) < 1 And atkingckai(135, 2) = 1 Then
                   atkingckai(135, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                End If
           End If
        Case 2
              For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\���\���_�@����_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 8925
                   atkingno(i, 6) = 9105
                   atkingno(i, 7) = 135
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            '=====================
            For i = 1 To 106
               If Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
                  If Val(pagecardnum(i, 2)) > cardnum(1) Then
                      cardnum(1) = pagecardnum(i, 2)
                      cardnum(2) = i
                  End If
                  If Val(pagecardnum(i, 4)) > cardnum(1) Then
                      cardnum(1) = pagecardnum(i, 4)
                      cardnum(2) = i
                  End If
               End If
            Next
            �ثe��(20) = cardnum(2)
            FormMainMode.tr�ϥΪ̵P_���P.Enabled = True
            �ثe��(21) = 1
            atkingckai(135, 2) = 0
   End Select
End If
End Sub
Sub �ײ��d_�l���K��()
Dim wert As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(1).Caption = "�l���K��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(136, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�ײ��d" Then
   Select Case atkingckai(136, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 1) >= 2 And atkingpagetot(2, 4) >= 1 And atkingckai(136, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 5
                   atkingckai(136, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                   '==========
                   If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                       �������m��l�`��(2) = �������m��l�`��(2) + 5
                   End If
                   '==========
                End If
                If (atkingpagetot(2, 1) < 2 Or atkingpagetot(2, 4) < 1) And atkingckai(136, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 5
                   atkingckai(136, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                   '==========
                   If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                       �������m��l�`��(2) = �������m��l�`��(2) - 5
                   End If
                   '==========
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�ײ��d\atking1_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 9420
                   atkingno(i, 6) = 8940
                   atkingno(i, 7) = 136
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then wert = 2 Else wert = 1
             '====================
             Do
                  For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                      If �H�����`���A��Ʈw(1, i, 3) > 0 Then
                             �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
                             If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                               '===�~�ӤU�@���A���
                                �԰��t����.���`���A�~��_�ϥΪ�
                                If �H�����`���A��Ʈw(1, i, 3) = 15 Then
                                    �԰��t����.�ˮ`����_�ߧY���`_�ϥΪ� 1  '���a�^�X���k0�ɰ��榺�`�ʧ@
                                End If
                             Else
                                FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
                             End If
                     End If
                  Next
                  '=====================
                  wert = Val(wert) - 1
             Loop Until wert <= 0
        Case 4
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 And �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                Randomize
                wert = Int(Rnd() * 3) + 1
                Do
                   For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                     If �H�����`���A��Ʈw(2, i, 2) >= 3 And �H�����`���A��Ʈw(2, i, 3) = 40 Then
                        Exit Do
                     End If
                     If �H�����`���A��Ʈw(2, i, 3) = 40 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) < 3 Then
                         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + (Val(wert) - 1)
                         If �H�����`���A��Ʈw(2, i, 2) > 3 Then �H�����`���A��Ʈw(2, i, 2) = 3
                         FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
                         Exit Do
                     End If
                   Next
                   For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                      If �H�����`���A��Ʈw(2, i, 2) = 0 And (Val(wert) - 1) > 0 Then
                         �԰��t����.�H�����`���A��]�w_��] 2, i, 40, app_path & "gif\���`���A\�{��.gif", 0, (Val(wert) - 1)
                         ���`���A�ˬd��(40, 1) = 1
                         ���`���A�ˬd��(40, 2) = 1
                         Exit Do
                     End If
                   Next
                   If i = 14 * ����H����ԤH��(2, 2) + 1 And (Val(wert) - 1) = 0 Then Exit Do
                Loop
            End If
            atkingckai(136, 2) = 0
            '===============�W���ޯ�ϥε���
            If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                atkingckai(139, 1) = 6
                AI�ޯ�.�ײ��d_�W�� '(���q6)
            End If
            '===============
   End Select
End If
End Sub
Sub �ײ��d_�������H��()
Dim wert As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(2).Caption = "�������H��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(137, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�ײ��d" Then
   Select Case atkingckai(137, 1)
      Case 1
           If movecp < 3 Then
                If atkingpagetot(2, 2) >= 3 And atkingckai(137, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 3
                   atkingckai(137, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 2) < 3 And atkingckai(137, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 3
                   atkingckai(137, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�ײ��d\atking2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7170
                   atkingno(i, 6) = 10440
                   atkingno(i, 7) = 137
                   atkingno(i, 8) = 0
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then wert = 2 Else wert = 3
                '=============
                If Val(�Y���淾�q�Ȯ��ܼ�(2)) Mod Val(wert) = 0 Then
                    �Y���淾�q�Ȯ��ܼ�(2) = 0
                    �Y����ˮ`�� = �Y���淾�q�Ȯ��ܼ�(2)
                End If
            End If
            '======================================
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) <= 0 And �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                Randomize
                wert = Int(Rnd() * 3) + 1
                Do
                   For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                     If �H�����`���A��Ʈw(2, i, 2) >= 3 And �H�����`���A��Ʈw(2, i, 3) = 40 Then
                        Exit Do
                     End If
                     If �H�����`���A��Ʈw(2, i, 3) = 40 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) < 3 Then
                         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + (Val(wert) - 1)
                         If �H�����`���A��Ʈw(2, i, 2) > 3 Then �H�����`���A��Ʈw(2, i, 2) = 3
                         FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
                         Exit Do
                     End If
                   Next
                   For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                      If �H�����`���A��Ʈw(2, i, 2) = 0 And (Val(wert) - 1) > 0 Then
                         �԰��t����.�H�����`���A��]�w_��] 2, i, 40, app_path & "gif\���`���A\�{��.gif", 0, (Val(wert) - 1)
                         ���`���A�ˬd��(40, 1) = 1
                         ���`���A�ˬd��(40, 2) = 1
                         Exit Do
                     End If
                   Next
                   If i = 14 * ����H����ԤH��(2, 2) + 1 And (Val(wert) - 1) = 0 Then Exit Do
                Loop
            End If
            atkingckai(137, 2) = 0
            '===============�W���ޯ�ϥε���
            If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                atkingckai(139, 1) = 6
                AI�ޯ�.�ײ��d_�W�� '(���q6)
            End If
            '===============
   End Select
End If
End Sub
Sub �ײ��d_���c���w��()
Dim wert As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(3).Caption = "���c���w��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(138, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�ײ��d" Then
   Select Case atkingckai(138, 1)
      Case 1
           If movecp = 3 Then
                If atkingpagetot(2, 2) >= 3 And atkingpagetot(2, 4) >= 1 And atkingckai(138, 2) = 0 Then
                   �������m��l�`��(2) = �������m��l�`��(2) + 6
                   atkingckai(138, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If (atkingpagetot(2, 2) < 3 Or atkingpagetot(2, 4) < 1) And atkingckai(138, 2) = 1 Then
                   �������m��l�`��(2) = �������m��l�`��(2) - 6
                   atkingckai(138, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
          End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�ײ��d\atking3_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6450
                   atkingno(i, 6) = 10215
                   atkingno(i, 7) = 138
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                For k = 1 To 3
                    �԰��t����.�^�_����_�q�� 2, k
                Next
            Else
                �԰��t����.�^�_����_�q�� 2, 1
            End If
            '======================================
            If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                    If �H�����`���A��Ʈw(2, i, 3) = 40 Then
                      �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
                      If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                        '===�~�ӤU�@���A���
                         �԰��t����.���`���A�~��_�ϥΪ�
                         ���`���A�ˬd��(40, 2) = 0
                     Else
                         FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
                         ���`���A�ˬd��(40, 1) = 1
                     End If
                   End If
                Next
            End If
            atkingckai(138, 2) = 0
            '===============�W���ޯ�ϥε���
            If �԰��t����.�S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True Then
                atkingckai(139, 1) = 6
                AI�ޯ�.�ײ��d_�W�� '(���q6)
            End If
            '===============
   End Select
End If
End Sub
Sub �ײ��d_�W��()
Dim wert As Integer '�Ȯ��ܼ�
If FormMainMode.comaiatk(4).Caption = "�W��" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(139, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "�ײ��d" Then
   Select Case atkingckai(139, 1)
      Case 1
                If atkingpagetot(2, 4) >= 1 And atkingckai(139, 2) = 0 Then
                   atkingckai(139, 2) = 1
                   atkingtrn(2) = Val(atkingtrn(2)) + 1
                End If
                If atkingpagetot(2, 4) < 1 And atkingckai(139, 2) = 1 Then
                   atkingckai(139, 2) = 0
                   atkingtrn(2) = Val(atkingtrn(2)) - 1
                 End If
      Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\�ײ��d\atking4_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 7920
                   atkingno(i, 6) = 10005
                   atkingno(i, 7) = 139
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
             Do
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                 If �H�����`���A��Ʈw(2, i, 2) >= 3 And �H�����`���A��Ʈw(2, i, 3) = 40 Then
                    atking_AI_�ײ��d_�W���ثe���q������(3) = 2
                    Exit Do
                 End If
                 If �H�����`���A��Ʈw(2, i, 3) = 40 And �H�����`���A��Ʈw(2, i, 2) > 0 And �H�����`���A��Ʈw(2, i, 2) < 3 Then
                     �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) + 1
                     FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
                     atking_AI_�ײ��d_�W���ثe���q������(3) = 1
                     Exit Do
                 End If
               Next
               For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                  If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                     �԰��t����.�H�����`���A��]�w_��] 2, i, 40, app_path & "gif\���`���A\�{��.gif", 0, 1
                     ���`���A�ˬd��(40, 1) = 1
                     ���`���A�ˬd��(40, 2) = 1
                     atking_AI_�ײ��d_�W���ثe���q������(3) = 1
                     Exit Do
                 End If
               Next
            Loop
            '========================�W��3�ɰ���ʦL
            If atking_AI_�ײ��d_�W���ثe���q������(3) = 2 Then
                Do
                    For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                      If �H�����`���A��Ʈw(2, i, 3) = 23 And �H�����`���A��Ʈw(2, i, 2) > 0 Then
                          FormMainMode.personcomspe(i).person_turn = 1
                          �H�����`���A��Ʈw(2, i, 2) = 1
                          Exit Do
                      End If
                    Next
                    For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                       If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                          �԰��t����.�H�����`���A��]�w_��] 2, i, 23, app_path & "gif\���`���A\atkingerr.gif", 0, 1
                          ���`���A�ˬd��(23, 1) = 1
                          ���`���A�ˬd��(23, 2) = 1
                          Exit Do
                       End If
                    Next
               Loop
            End If
        Case 4
            '========================�W��3�ɧ�2�����q-����
            If atking_AI_�ײ��d_�W���ثe���q������(3) = 2 Then
                If Val(atking_AI_�ײ��d_�W���ثe���q������(4)) = 0 Then
                    atking_AI_�ײ��d_�W���ثe���q������(1) = �������m��l�`��(2)
                    atking_AI_�ײ��d_�W���ثe���q������(2) = �������m��l�`��(2) * 2
                    atking_AI_�ײ��d_�W���ثe���q������(4) = 1
                    �������m��l�`��(2) = �������m��l�`��(2) * 2
                ElseIf Val(atking_AI_�ײ��d_�W���ثe���q������(4)) = 1 Then
                    atking_AI_�ײ��d_�W���ثe���q������(1) = atking_AI_�ײ��d_�W���ثe���q������(1) + (�������m��l�`��(2) - atking_AI_�ײ��d_�W���ثe���q������(2))
                    �������m��l�`��(2) = atking_AI_�ײ��d_�W���ثe���q������(1) * 2
                    atking_AI_�ײ��d_�W���ثe���q������(2) = atking_AI_�ײ��d_�W���ثe���q������(1) * 2
                End If
            End If
        Case 5
            '========================�W��3�ɧ�2�����q-�}�l���q�ɲM�����
            atking_AI_�ײ��d_�W���ثe���q������(1) = 0
            atking_AI_�ײ��d_�W���ثe���q������(2) = 0
            atking_AI_�ײ��d_�W���ثe���q������(4) = 0
        Case 6
            '========================�W���ޯ൲��(���q)
            atkingckai(139, 2) = 0
            Erase atking_AI_�ײ��d_�W���ثe���q������
        Case 7
            '========================�󴫨���ɭ��s���J�ޯ�
            If atking_AI_�ײ��d_�W���ثe���q������(3) > 0 Then
                atking_AI_�ײ��d_�W���ثe���q������(1) = 0
                atking_AI_�ײ��d_�W���ثe���q������(2) = 0
                atking_AI_�ײ��d_�W���ثe���q������(4) = 0
            End If
        Case 8
            '========================�W���ޯ൲��(�^�X�������q)
            atkingckai(139, 2) = 0
            If atking_AI_�ײ��d_�W���ثe���q������(3) = 2 Then
                �԰��t����.����ʧ@_�M���Ҧ����`���A_�q��
            End If
            Erase atking_AI_�ײ��d_�W���ثe���q������
   End Select
End If
End Sub
Sub ù��Y_EX_�V�大�b()
If FormMainMode.comaiatk(2).Caption = "Ex�V�大�b" And (����ʧ@_�ˬd�O�_�����w���`���A(2, 23) = False Or atkingckai(140, 2) = 1) _
   And FormMainMode.compi1(����H����ԤH��(2, 2)) = "ù��Y" Then
   Select Case atkingckai(140, 1)
        Case 1
          If movecp = 1 Then
            If atkingpagetot(2, 1) >= 3 And atkingpagetot(2, 3) >= 2 And atkingckai(140, 2) = 0 Then
               atkingckai(140, 2) = 1
               atkingtrn(2) = Val(atkingtrn(2)) + 1
               �������m��l�`��(2) = �������m��l�`��(2) + 9
            ElseIf (atkingpagetot(2, 1) < 3 Or atkingpagetot(2, 3) < 2) And atkingckai(140, 2) = 1 Then
               atkingckai(140, 2) = 0
               atkingtrn(2) = Val(atkingtrn(2)) - 1
               �������m��l�`��(2) = �������m��l�`��(2) - 9
            End If
          End If
        Case 2
             For i = �H���ޯ�Ʀr���� To 1 Step -1
               If atkingno(i, 1) = "" Then
                   atkingno(i, 1) = app_path & "gif\ù��Y\atkingEX2_2.jpg"
                   atkingno(i, 2) = 2
                   atkingno(i, 3) = 0
                   atkingno(i, 4) = 0
                   atkingno(i, 5) = 6225
                   atkingno(i, 6) = 9615
                   atkingno(i, 7) = 140
                   atkingno(i, 8) = 1
                   atkingno(i, 9) = 0
                   atkingno(i, 10) = 0
                   atkingno(i, 11) = 0
                   Exit For
                 End If
             Next
        Case 3
            �^�_����_�q�� 1, 1
        Case 4
            atkingckai(140, 2) = 0
            If Val(�Y���淾�q�Ȯ��ܼ�(2)) > 0 Then
                �^�_����_�q�� 1, 1
            End If
   End Select
End If
End Sub

