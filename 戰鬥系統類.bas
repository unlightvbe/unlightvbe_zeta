Attribute VB_Name = "�԰��t����"
Public Const a1a As String = "ATK-�C"
Public Const a2a As String = "DEF"
Public Const a3a As String = "MOV"
Public Const a4a As String = "SPE"
Public Const a5a As String = "ATK-�j"
Public Const a6a As String = "DRAW"
Public Const a7a As String = "BRK"
Public Const a8a As String = "HPL"
Public Const a9a As String = "HPW"
Public Const b1b As Integer = 1
Public Const b2b As Integer = 2
Public Const b3b As Integer = 3
Public Const b4b As Integer = 4
Public Const b5b As Integer = 5
Public Const b6b As Integer = 6
Public Const b7b As Integer = 7
Public Const b8b As Integer = 8
Public Const b9b As Integer = 9

Public atkingno(1 To 8, 1 To 11) As String '�ޯ�o�ʱƧǼȮɹϤ����|�x�s�ܼ�(�ޯ�o�ʶ���8~1,1.�Ϥ����|/2.(1)�ϥΪ�/(2)�q����/3.Left/4.Top(�y��)/5.�����e��(Width)/6.��������(Height)/7.�ޯ�s��/8.�ޯ���椤�ɱҰʭ�/9.�ޯ���椤���Ϥ��ˬd��/10.��2�i�Ϥ����|)
Public goicheck(1 To 2) As Integer   '����/���m�Ҧ��[��ƭ��ˬd�X
Public pageonin(1 To 106) As Integer  '�P�i���ϭ��ˬd�X
Public liveus(1 To 3) As Integer, livecom(1 To 3) As Integer, liveusmax(1 To 3) As Integer, livecommax(1 To 3) As Integer
Public turn As Integer, atkus(1 To 3) As Integer, atkcom(1 To 3) As Integer, defus(1 To 3) As Integer, defcom(1 To 3) As Integer, pagecheckus As Integer, pagecheckcom As Integer, pagegive As Integer, goidefus As Integer, movecom As Integer, moveus As Integer, movecp As Integer, chkcomck As Integer, uslevel(1 To 3) As Integer, comlevel(1 To 3) As Integer, liveus41(1 To 3) As Integer, livecom41(1 To 3) As Integer, movecheckcom As Integer, movecheckus As Integer
Public nameus(1 To 3) As String, namecom(1 To 3) As String
Public moveturn As Integer  '���������m�Ҧ������ˬd�X(1.�ϥΪ̥���/2.�q������)
Public atkinghelpxy(1 To 2, 1 To 4, 1 To 2) As Integer '�ޯ໡����y�Ы��w���(1.�ϥΪ̤�/2.�q����,��1~4�ӧޯ�,1.Left/2.Top(�y��))
Public pageusleadmax(0 To 1) As Integer   '�ϥΪ̵P���ǭp�ƪ�(0.��P/1.�X�P)
Public pagecomleadmax(0 To 1) As Integer   '�q���P���ǭp�ƪ�(0.��P/1.�X�P)
Public pageqlead(1 To 2) As Integer   '�X�P�p���ܼ�(1.�ϥΪ�/2.�q��)
Public pageglead(1 To 2) As Integer   '��P�p���ܼ�(1.�ϥΪ�/2.�q��)
Public movedsus As Integer   '�ϥΪ̲��ʶ��q�M�w���ܼ�
Public turnpageonin As Integer  '���q�O�_�i�X�P�ܼ�(�@��)
Public turnpageoninatking As Integer  '���q�O�_�i�X�P�ܼ�(�ޯ�ϥ�)
Public goickus As Integer '�P�Ȥ@���ˬd�X
Public atkingck(1 To 161, 1 To 2) As Integer '�ޯඥ�q�ҰʽX(x.�H���ޯ�s��,1.�ޯ���涥�q/2.�ޯ�Ұ��ˬd��)
Public atkingckai(1 To 140, 1 To 2) As Integer 'AI�ޯඥ�q�ҰʽX(x.�H���ޯ�s��,1.�ޯ���涥�q/2.�ޯ�Ұ��ˬd��)
Public atkingtrn(1 To 4) As Integer '�ޯ�p�ƾ��Ȯ��x�s�ܼ�(1.�ϥΪ�(�{)/2.�q��(�{)/3.�ϥΪ�(�ƥ�)/4.�q��(�ƥ�))
Public akhpnm As Integer  '�ޯ໡���Ȯ��ܼ�
Public turnatk As Integer  '���������m���q�ܼ�(1.�ϥΪ̧����B�q�����m,2.�ϥΪ̨��m�B�q������,3.�o�P�B����)
Public trend�Ȯ��ܼ� As Integer '�������q�p�ƾ��Ȯ��ܼ�
Public HP�ˬd�ܼ� As Boolean 'HP�ˬd���q�O�_�w�ˬd�ܼ�
Public HP�ˬd���q�� As Integer 'HP�ˬd���q�ܼ�(1.���ʶ��q��,2.����/���m���q�e,3.��/���m���q��)
Public �Z�����(1 To 2, 1 To 2, 1 To 2) As Integer  '�Z�����Ȯ��x�s���(1.HP���/2.�P����,1.�ϥΪ�/2.�q��,1.Left���/2.Top���)
Public personminixy(1 To 2, 1 To 3, 1 To 3, 1 To 2) As Integer '�p�H���Ϥ��y�Ы��w���(1.�ϥΪ�/2.�q��,��n��,1.��Z��/2.���Z��/3.���Z��,1.Left/2.Top(�y��))
Public �H�����`���A��Ʈw(1 To 2, 1 To 42, 1 To 3) As Integer '���`���A���(1.�ϥΪ�/2.�q��,��x�Ӳ��`���A,1.���A�ƭ�/2.���A�έp��(�Ѿl�^�X/�֭p)/3.���A�s��)
Public ���`���A�ˬd��(1 To 40, 1 To 2) As Integer '���`���A�ҰʽX(x.���`���A�s��,1.���A���涥�q/2.���A�Ұ��ˬd��)
Public �ޯ�ʵe��ܶ��q�� As Integer '�ޯ�ʵe�p�ƾ����q�X(1.����/���m���q-���q,2.���ʶ��q-���q/3.�o�P���q��B���ʶ��q�e/4.���ʶ��q��/5.�������q��/6.���m���q��/7.�^�X������)
Public �������m��l�`��(1 To 4) As Integer '����/���m�Ҧ���l�ƶq���(1.�ϥΪ�(�`)/2.�q��(�`)/3.�ϥΪ�(��)/4.�q��(��))
Public atkingpagetot(1 To 2, 1 To 5) As Integer  '�C���q�X�P�����μƭȲέp���(1.�ϥΪ�/2.�q��,1.�C/2.��/3.��/4.�S/5.�j)
Public ��ƹs�ˬd��(1 To 2) As Boolean '��e���q��l�ƶq�O�_���s�ˬd��(1.�ϥΪ�/2.�q��)
Public pagecardnum(1 To 106, 1 To 11) As String '���εP���(��x�s��(1~70-���P/71~88-�ϥΪ̨ƥ�P/89~106-�q���ƥ�P),1.��������/2.�����ƭ�/3.�ϭ�����/4.�ϭ��ƭ�/5.(1)�ϥΪ�-(2)�q��/6.(1)��P-(2)�X�P-(3)�õP-(4)�P��/7.�X�P����/8.�Ϥ��s��/9.�ثeLeft(�y��)/10.�ثeTop(�y��)/11.(1)�q����X�P(��)-(2)�q���o�X�P(�~))
Public �P�`���q��(1 To 3) As Integer '�P�֦��`���q��(1.�ϥΪ�/2.�q��/3.�`�p)
Public �P���ʼȮ��ܼ�(1 To 3) As Long '�P���ʭp�ƾ��Ȯ��ܼ�(1.Left���/2.Top���/3.�P�i�s��)
Public �ثe��(1 To 33) As Integer '�`�Ȯ��ܼ�
Public �X�P���ǲέp�Ȯ��ܼ�(1 To 4, 1 To 106, 1 To 2) As Integer '�X�P���ǲέp�`�Ȯɸ��(1.�ϥΪ̥X�P/2.�ϥΪ̤�P/3.�q���X�P/4.�q����P,��x����,1.�ثe�P�X�P����/2.�P�i�s��)
Public �Z�����_���P�Ȯɼ�(1 To 106, 1 To 3) As Integer  '���P�ӧO�Z�����Ȯ��x�s�ܼ�(��x����,1.Left���/2.Top���/3.�P�i�s��)
Public ���q���A�� As Integer '�C���q�}�l�������A�ˬd��(1.�}�l���q(�ϥΪ�)/2.�������q(�ϥΪ�)/3.�}�l���q(�q��)/4.�������q(�q��)/5.�洫����)
Public �p�H���Y�����ʤ�V��(1 To 2) As Integer '�p�H���Y�����ʤ�V���A��(1.�ϥΪ�/2.�q��[1.�V��,2.�V�~])
Public ��q�p�ƾ��ʵe�Ȯ��ܼ�(1 To 2, 1 To 2) As Integer '�}�l��l���q-��q�ʵe�p�ƾ��Ȯ��ܼ�(1.�ϥΪ̦��/2.�q�����,1.�C�����ʶq/2.�O�_�w����)
Public �ɶ��b�C���ܤƬ����Ȯ��ܼ�(1 To 4, 1 To 3) As Integer '�ɶ��b�i���C���ܤƶ��q�����Ȯ��ܼ�(1~3(1)����ܤƶq(1(1).�ɶ��b(���~))/2.�ثe�֭p�q/3.�ثe�C��(R,G,B),4.(1)�ɶ��b(�~)���q��-(1)���ܬ�-(2)���ܶ�/2.�ثe�֭p�q/3.�ثe�C��(R))
Public �H���ޯ�Ʀr���� As Integer '�԰��t�Ϊ��-atkingnumtot.Caption���ܼƪ��
Public �}�l�d�����ʰʵe������(1 To 2, 1 To 4) As Integer   '�}�l�ɨC�i�d�����ʰʵe����������(1.�ϥΪ�/2.�q��,1~3.�d��/4.�ثe�ĴX�i)
Public �洫��������Ȯ��ܼ�(1 To 4) As Integer '�洫������������Ȯɼ�(1.�ϥΪ�/2.�q��/3.�O�_��U����/4.�洫���⧹���涥�q��)
Public pageeventnum(1 To 2, 1 To 18, 1 To 2) As String '�ƥ�d�ƦC�������(1.�ϥΪ�/2.�q��,1~18-�s��,1.�ƥ�d�W��/2.�ƥ�d�ɮצW��)
Public �Y����ˮ`�� As Integer '�԰��t�Ϊ��-fm2.Caption���ܼƪ��
Public �԰��Ҧ��ӱѬ����� As Integer '�԰��t�η�e�ӱѬ����Ȯ��ܼ�(1.�ϥΪ̤�ӧQ/2.�ϥΪ̤�ѥ_/3.����)
Public �q���貾�ʶ��q��ܼ� As Integer '���ʶ��q�q�����ܤ���ʼȮ��ܼ�
Public �q����ƥ�d�O�_�X����ܼ� As Boolean '�q������X�ƥ�d�O�_�X���Ȯɬ���
Public �H���d���I���s��������(1 To 7) As Integer '�H���d���I���ޯ໡���H���s���Ȯ��ܼ�(1.(1).�ϥΪ�/(2).�q��,2.��n��,3.�ثe�ϥΪ̤�ϥΤH���s��/4.�ثe��ܤ��ޯ�s��(�ϥΪ̤�ϥΤH��)/5.�ثe��ܤ��ޯ�s��(��L)/6~7.�ثe��ܤ��ޯ�s��(�洫����)
Public �Y���淾�q�Ȯ��ܼ�(1 To 4) As Integer 'Form6���ȷ��q�Ȯ��ܼ�(1.�@�^�X������P�_(1.�e/2.��),2.��l���(�ϥΪ�)-�Y��ᦳ�Ķˮ`��,3.��l���(�q��)-�Y���ˮ`��H(1.�ϥΪ�/2.�q��),4.(1.�ϥΪ̥���/2.�q������))
Public �H�������ˬd�Ȯ��ܼ�(1 To 3) As Integer '�H�������ˬd�p�ƾ������Ȯ��ܼ�(1.�ثe�p��/2.�ϥΪ̼аO/3.�q���аO)
Public ���εP�U�P����������(0 To 31, 1 To 2) As Integer '�U�������εP�P���������Ȯ��ܼ�(0.(1)�ثe�w�o�P�`�ƶq/(2)�ثe�����P�`�ƶq,1~31.(1)�ثe�w�ϥΤ��P��/(2)�ӵP����ϥΤ��`�ƶq)
Public �d���H����T�ɮ�Ū�����Ѭ����� As String '�d���H����T�ɮ�Ū�����Ѯ��ɮצW�����Ȯ��ܼ�
Sub �H���ޯ���O�}��(ByVal k As Boolean, ByVal n As Integer)
Select Case n
   Case 1
      If k = True Then
         FormMainMode.personatk(1).ForeColor = RGB(255, 255, 0)
         FormMainMode.personatk(1).BackColor = RGB(47, 94, 94)
      Else
         FormMainMode.personatk(1).ForeColor = RGB(192, 192, 192)
         FormMainMode.personatk(1).BackColor = RGB(0, 0, 0)
      End If
   Case 2
      If k = True Then
         FormMainMode.personatk(2).ForeColor = RGB(255, 255, 0)
         FormMainMode.personatk(2).BackColor = RGB(47, 94, 94)
      Else
         FormMainMode.personatk(2).ForeColor = RGB(192, 192, 192)
         FormMainMode.personatk(2).BackColor = RGB(0, 0, 0)
      End If
   Case 3
      If k = True Then
         FormMainMode.personatk(3).ForeColor = RGB(255, 255, 0)
         FormMainMode.personatk(3).BackColor = RGB(47, 94, 94)
      Else
         FormMainMode.personatk(3).ForeColor = RGB(192, 192, 192)
         FormMainMode.personatk(3).BackColor = RGB(0, 0, 0)
      End If
   Case 4
      If k = True Then
         FormMainMode.personatk(4).ForeColor = RGB(255, 255, 0)
         FormMainMode.personatk(4).BackColor = RGB(47, 94, 94)
      Else
         FormMainMode.personatk(4).ForeColor = RGB(192, 192, 192)
         FormMainMode.personatk(4).BackColor = RGB(0, 0, 0)
      End If
End Select

End Sub
Sub �H�����`���A��]�w_��](ByVal ���� As Integer, ByVal �ĴX�� As Integer, ByVal ���`���A�s�� As Integer, ByVal ph As String, ByVal num1 As Integer, ByVal num2 As Integer)
If Formsetting.chknewdefferent.Value = 1 Then
    ph = ����ʧ@_���|�ϥηs�����`���A�Ϯ�(ph)
End If
'===================================
Select Case ����
    Case 1
        FormMainMode.personusspe(�ĴX��).���`���A�Ϥ� = ph
        FormMainMode.personusspe(�ĴX��).person_num = num1
        FormMainMode.personusspe(�ĴX��).person_turn = num2
        �H�����`���A��Ʈw(1, �ĴX��, 1) = num1
        �H�����`���A��Ʈw(1, �ĴX��, 2) = num2
        �H�����`���A��Ʈw(1, �ĴX��, 3) = ���`���A�s��
        FormMainMode.personusspe(�ĴX��).Visible = True
    Case 2
        FormMainMode.personcomspe(�ĴX��).���`���A�Ϥ� = ph
        FormMainMode.personcomspe(�ĴX��).person_num = num1
        FormMainMode.personcomspe(�ĴX��).person_turn = num2
        �H�����`���A��Ʈw(2, �ĴX��, 1) = num1
        �H�����`���A��Ʈw(2, �ĴX��, 2) = num2
        �H�����`���A��Ʈw(2, �ĴX��, 3) = ���`���A�s��
        FormMainMode.personcomspe(�ĴX��).Visible = True
End Select

End Sub
Function ����ʧ@_���|�ϥηs�����`���A�Ϯ�(ByVal ph As String) As String
For i = 1 To Len(ph)
    If Mid(ph, i, 1) = "." Then
        ph = Mid(ph, 1, i - 1) & "new" & Right(ph, 4)
        Exit For
    End If
Next
����ʧ@_���|�ϥηs�����`���A�Ϯ� = ph
End Function
Sub �۰ʱ��b����()
FormMainMode.messageus.ListIndex = FormMainMode.messageus.ListCount - 1
End Sub
Sub �ˮ`����_�ޯઽ��_�ϥΪ�(ByVal tot As Integer, ByVal num As Integer)
'===============================
���`���A�ˬd��(35, 1) = 1
���`���A.���@_�ϥΪ� num, tot '(���q1)
'===============================
If atking_��_�u�@�Ҧ����A�Ұʭ� = False Then
    Select Case num
       Case 1
          If tot > 0 And liveus(����H����ԤH��(1, 2)) > 0 Then
              If tot >= liveus(����H����ԤH��(1, 2)) Then
                 FormMainMode.messageus.AddItem "�z����F" & liveus(����H����ԤH��(1, 2)) & "�I�ˮ`�C"
                 �԰��t����.�۰ʱ��b����
                 FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption = 0
                 FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption = 0
                 liveus(����H����ԤH��(1, 2)) = 0
                 FormMainMode.bloodnumus1.Caption = 0
                 FormMainMode.bloodlineout1.Width = 0
                 �P�`���q��(1) = �P�`���q��(1) + 1
              Else
                 FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption = Val(FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption) - tot
                 FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption = Val(FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption) - tot
                 liveus(����H����ԤH��(1, 2)) = liveus(����H����ԤH��(1, 2)) - tot
                 FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
                 FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (�Z�����(1, 1, 1) * tot)
                 FormMainMode.messageus.AddItem "�z����F" & tot & "�I�ˮ`�C"
                 �԰��t����.�۰ʱ��b����
              End If
              �԰��t����.����ˮ`����
           End If
       Case Is > 1
           If tot > 0 And liveus(����ݾ��H��������(1, num)) > 0 Then
              If tot >= liveus(����ݾ��H��������(1, num)) Then
                 liveus(����ݾ��H��������(1, num)) = 0
                 If FormMainMode.uspi1(����ݾ��H��������(1, num)).Caption = "" Then
                     FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption = -liveusmax(����ݾ��H��������(1, num))
                     FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption = -liveusmax(����ݾ��H��������(1, num))
                 Else
                     FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption = 0
                     FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption = 0
                 End If
                 �P�`���q��(1) = �P�`���q��(1) + 1
              Else
                 FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption = Val(FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption) - tot
                 liveus(����ݾ��H��������(1, num)) = liveus(����ݾ��H��������(1, num)) - tot
                 FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption = Val(FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption) - tot
              End If
           End If
    End Select
End If
End Sub
Sub ����ˮ`����()
Select Case movecp
    Case 1
        FormMainMode.wmpse2.Controls.play
        �@��t����.�ˬd���ּ��� 2
    Case Is >= 2
        FormMainMode.wmpse8.Controls.play
        �@��t����.�ˬd���ּ��� 8
End Select
End Sub
Sub ����ʧ@_�ޯ��ʵ���()
FormMainMode.atkingnumtot.Caption = Val(FormMainMode.atkingnumtot) - 1
FormMainMode.atkingtrtot.Interval = 20
FormMainMode.atkingtrtot.Enabled = True
End Sub
Sub �^�_����_�ϥΪ�(ByVal tot As Integer, ByVal num As Integer)
Select Case num
   Case 1
         If liveus(����H����ԤH��(1, 2)) > 0 And tot > 0 Then
               If liveusmax(����H����ԤH��(1, 2)) - liveus(����H����ԤH��(1, 2)) >= tot Then
                    FormMainMode.messageus.AddItem "�A��HP��_�F" & tot & "�I�C"
                    FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + �Z�����(1, 1, 1) * tot
                    liveus(����H����ԤH��(1, 2)) = Val(liveus(����H����ԤH��(1, 2))) + tot
                    FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption = liveus(����H����ԤH��(1, 2))
                    FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption = liveus(����H����ԤH��(1, 2))
                    FormMainMode.bloodnumus1.Caption = liveus(����H����ԤH��(1, 2))
                    �԰��t����.�۰ʱ��b����
              ElseIf liveusmax(����H����ԤH��(1, 2)) - liveus(����H����ԤH��(1, 2)) < tot Then
                    If liveusmax(����H����ԤH��(1, 2)) - liveus(����H����ԤH��(1, 2)) > 0 Then
                       FormMainMode.messageus.AddItem "�A��HP��_�F" & Val(liveusmax(����H����ԤH��(1, 2))) - Val(liveus(����H����ԤH��(1, 2))) & "�I�C"
                       FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + �Z�����(1, 1, 1) * (Val(liveusmax(����H����ԤH��(1, 2))) - Val(liveus(����H����ԤH��(1, 2))))
                       liveus(����H����ԤH��(1, 2)) = Val(liveusmax(����H����ԤH��(1, 2)))
                       FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption = liveus(����H����ԤH��(1, 2))
                       FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption = liveus(����H����ԤH��(1, 2))
                       FormMainMode.bloodnumus1.Caption = liveus(����H����ԤH��(1, 2))
                       �԰��t����.�۰ʱ��b����
                    End If
              End If
        End If
   Case Is > 1
        If liveus(����ݾ��H��������(1, num)) > 0 And tot > 0 Then
               If liveusmax(����ݾ��H��������(1, num)) - liveus(����ݾ��H��������(1, num)) >= tot Then
                    liveus(����ݾ��H��������(1, num)) = Val(liveus(����ݾ��H��������(1, num))) + tot
                    FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption = liveus(����ݾ��H��������(1, num))
                    FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption = liveus(����ݾ��H��������(1, num))
              ElseIf liveusmax(����ݾ��H��������(1, num)) - liveus(����ݾ��H��������(1, num)) < tot Then
                    If liveusmax(����ݾ��H��������(1, num)) - liveus(����ݾ��H��������(1, num)) > 0 Then
                       liveus(����ݾ��H��������(1, num)) = Val(liveusmax(����ݾ��H��������(1, num)))
                       FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption = liveus(����ݾ��H��������(1, num))
                       FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption = liveus(����ݾ��H��������(1, num))
                    End If
              End If
        End If
End Select
End Sub
Sub �^�_����_�q��(ByVal tot As Integer, ByVal num As Integer)
Select Case num
   Case 1
         If livecom(����H����ԤH��(2, 2)) > 0 And tot > 0 Then
               If livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2)) >= tot Then
                    FormMainMode.messageus.AddItem "��誺HP��_�F" & tot & "�I�C"
                    FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - �Z�����(1, 2, 1) * tot
                    livecom(����H����ԤH��(2, 2)) = Val(livecom(����H����ԤH��(2, 2))) + tot
                    FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = livecom(����H����ԤH��(2, 2))
                    FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = livecom(����H����ԤH��(2, 2))
                    FormMainMode.bloodnumcom1.Caption = livecom(����H����ԤH��(2, 2))
                    �԰��t����.�۰ʱ��b����
              ElseIf livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2)) < tot Then
                    If livecommax(����H����ԤH��(2, 2)) - livecom(����H����ԤH��(2, 2)) > 0 Then
                       FormMainMode.messageus.AddItem "��誺HP��_�F" & Val(livecommax(����H����ԤH��(2, 2))) - Val(livecom(����H����ԤH��(2, 2))) & "�I�C"
                       FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - �Z�����(1, 2, 1) * (Val(livecommax(����H����ԤH��(2, 2))) - Val(livecom(����H����ԤH��(2, 2))))
                       livecom(����H����ԤH��(2, 2)) = Val(livecommax(����H����ԤH��(2, 2)))
                       FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = livecom(����H����ԤH��(2, 2))
                       FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = livecom(����H����ԤH��(2, 2))
                       FormMainMode.bloodnumcom1.Caption = livecom(����H����ԤH��(2, 2))
                       �԰��t����.�۰ʱ��b����
                    End If
              End If
        End If
   Case Is > 1
        If livecom(����ݾ��H��������(2, num)) > 0 And tot > 0 Then
               If livecommax(����ݾ��H��������(2, num)) - livecom(����ݾ��H��������(2, num)) >= tot Then
                    livecom(����ݾ��H��������(2, num)) = Val(livecom(����ݾ��H��������(2, num))) + tot
                    FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = Val(FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption) + tot
                    FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = Val(FormMainMode.compi4(����ݾ��H��������(2, num)).Caption) + tot
              ElseIf livecommax(����ݾ��H��������(2, num)) - livecom(����ݾ��H��������(2, num)) < tot Then
                       livecom(����ݾ��H��������(2, num)) = Val(livecommax(����ݾ��H��������(2, num)))
                       If FormMainMode.compi1(����ݾ��H��������(2, num)).Caption = "" Then
                            FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = 0
                            FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = 0
                       Else
                            FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = livecom(����ݾ��H��������(2, num))
                            FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = livecom(����ݾ��H��������(2, num))
                       End If
              End If
        End If
End Select
End Sub
Function �ˮ`����_�ϥΪ�(ByVal tot As Integer)
'===============================
���`���A�ˬd��(35, 1) = 1
���`���A.���@_�ϥΪ� 1, tot '(���q1)
'===============================
If tot > 0 And liveus(����H����ԤH��(1, 2)) > 0 Then
      If tot >= liveus(����H����ԤH��(1, 2)) Then
         FormMainMode.messageus.AddItem "�z����F" & liveus(����H����ԤH��(1, 2)) & "�I�ˮ`�C"
         �԰��t����.�۰ʱ��b����
         FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption = 0
         FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption = 0
         liveus(����H����ԤH��(1, 2)) = 0
         FormMainMode.bloodnumus1.Caption = 0
         FormMainMode.bloodlineout1.Width = 0
         �P�`���q��(1) = �P�`���q��(1) + 1
      Else
         FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption = Val(FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption) - tot
         FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption = Val(FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption) - tot
         liveus(����H����ԤH��(1, 2)) = liveus(����H����ԤH��(1, 2)) - tot
         FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
         FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (�Z�����(1, 1, 1) * tot)
         FormMainMode.messageus.AddItem "�z����F" & tot & "�I�ˮ`�C"
         �԰��t����.�۰ʱ��b����
      End If
�԰��t����.����ˮ`����
End If
End Function
Sub �ˮ`����_�ޯઽ��_�q��(ByVal tot As Integer, ByVal num As Integer)
'===============================
���`���A�ˬd��(36, 1) = 1
���`���A.���@_�q�� num, tot '(���q1)
'===============================
If atking_AI_��_�u�@�Ҧ����A�Ұʭ� = False Then
    Select Case num
        Case 1
           If tot > 0 And livecom(����H����ԤH��(2, 2)) > 0 Then
                    If tot >= livecom(����H����ԤH��(2, 2)) Then
                       FormMainMode.messageus.AddItem "������F" & livecom(����H����ԤH��(2, 2)) & "�I�ˮ`�C"
                       �԰��t����.�۰ʱ��b����
                       FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = 0
                       FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = 0
                       FormMainMode.bloodnumcom1.Caption = 0
                       livecom(����H����ԤH��(2, 2)) = 0
                       FormMainMode.bloodlineout2.Left = 11580
                       �P�`���q��(2) = �P�`���q��(2) + 1
                    Else
                       FormMainMode.messageus.AddItem "������F" & Val(tot) & "�I�ˮ`�C"
                       �԰��t����.�۰ʱ��b����
                       FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = Val(FormMainMode.compi4(����H����ԤH��(2, 2)).Caption) - tot
                       FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = Val(FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption) - tot
                       FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
                       livecom(����H����ԤH��(2, 2)) = livecom(����H����ԤH��(2, 2)) - tot
                       FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (�Z�����(1, 2, 1) * tot)
                    End If
            �԰��t����.����ˮ`����
            End If
        Case Is > 1
           If tot > 0 And livecom(����ݾ��H��������(2, num)) > 0 Then
                    If tot >= livecom(����ݾ��H��������(2, num)) Then
                       If FormMainMode.compi1(����ݾ��H��������(2, num)).Caption = "" Then
                           FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = -livecommax(����ݾ��H��������(2, num))
                           FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = -livecommax(����ݾ��H��������(2, num))
                       Else
                           FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = 0
                           FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = 0
                       End If
                       livecom(����ݾ��H��������(2, num)) = 0
                       �P�`���q��(2) = �P�`���q��(2) + 1
                    Else
                       FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = Val(FormMainMode.compi4(����ݾ��H��������(2, num)).Caption) - tot
                       FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = Val(FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption) - tot
                       livecom(����ݾ��H��������(2, num)) = livecom(����ݾ��H��������(2, num)) - tot
                    End If
            End If
    End Select
End If
End Sub
Function �ˮ`����_�q��(ByVal tot As Integer)
'===============================
���`���A�ˬd��(36, 1) = 1
���`���A.���@_�q�� 1, tot '(���q1)
'===============================
If tot > 0 And livecom(����H����ԤH��(2, 2)) > 0 Then
        If tot >= livecom(����H����ԤH��(2, 2)) Then
           FormMainMode.messageus.AddItem "������F" & livecom(����H����ԤH��(2, 2)) & "�I�ˮ`�C"
           �԰��t����.�۰ʱ��b����
           FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = 0
           FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = 0
           FormMainMode.bloodnumcom1.Caption = 0
           livecom(����H����ԤH��(2, 2)) = 0
           FormMainMode.bloodlineout2.Left = 11580
           �P�`���q��(2) = �P�`���q��(2) + 1
        Else
           FormMainMode.messageus.AddItem "������F" & Val(tot) & "�I�ˮ`�C"
           �԰��t����.�۰ʱ��b����
           FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = Val(FormMainMode.compi4(����H����ԤH��(2, 2)).Caption) - tot
           FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = Val(FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption) - tot
           FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
           livecom(����H����ԤH��(2, 2)) = livecom(����H����ԤH��(2, 2)) - tot
           FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (�Z�����(1, 2, 1) * tot)
        End If
�԰��t����.����ˮ`����
End If
End Function
Sub ����ʧ@_�ϥΪ�_��P(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) - 1
    �ثe��(5) = pagecardnum(n, 7)
    pagecardnum(n, 6) = 3
'    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
'    �P���ǼW�[_��P_�ϥΪ� n
    �ثe��(15) = 4
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
End Sub
Sub ����ʧ@_�P��_�^�P_�ϥΪ�(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
'    �ثe��(9) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 1
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.���εP�^�_���� n
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�ϥΪ� n
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
End Sub
Sub ����ʧ@_�q���P_���P_�ϥΪ�(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    �ثe��(9) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 1
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�ϥΪ̤�P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�ϥΪ� n
    �ثe��(15) = 2
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
End Sub
Sub ����ʧ@_�ϥΪ̵P_���P_�q��(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pageusglead = Val(FormMainMode.pageusglead) - 1
    �ثe��(5) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 2
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�q�� n
    �ثe��(15) = 20
    �԰��t����.���εP�ܭI��
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
End Sub
Sub ����ʧ@_�P��_�^�P_�q��(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
'    �ثe��(5) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 2
    pagecardnum(n, 6) = 1
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �P���ǼW�[_��P_�q�� n
    �԰��t����.���εP�ܭI��
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
End Sub
Sub ����ʧ@_½�P(ByVal n As Integer)
    FormMainMode.card(n).Width = 810
    FormMainMode.card(n).Height = 1260
    FormMainMode.card(n).Picture = LoadPicture(app_path & "card\" & pagecardnum(n, 8) & "-" & pageonin(n) & ".bmp")
    FormMainMode.card(n).Visible = True
    FormMainMode.wmpse4.Controls.stop
    FormMainMode.wmpse4.Controls.play
    �@��t����.�ˬd���ּ��� 4
End Sub
Sub �y�Эp��_�q���X�P()
Dim xy As Long  '�Ȯ��ܼ�(���PLeft)
If pageqlead(2) = 1 Then
    �P���ʼȮ��ܼ�(1) = 5260
    �P���ʼȮ��ܼ�(2) = 1120
ElseIf pageqlead(2) > 1 Then
    xy = (pageqlead(2) - 1) * 460
    �P���ʼȮ��ܼ�(1) = (Val(5260) - xy) + ((pageqlead(2) - 1) * Val(960))
    �P���ʼȮ��ܼ�(2) = 1120
End If

End Sub
Sub �y�Эp��_�q����P()
�P���ʼȮ��ܼ�(1) = 10560 - 240 * (Val(FormMainMode.pagecomglead) - 1) '�p��Left�y��
�P���ʼȮ��ܼ�(2) = -600 '���wTop�y��
End Sub
Sub �y�Эp��_�ϥΪ̥X�P()
Dim xy As Long   '�Ȯ��ܼ�(���PLeft)
If pageqlead(1) = 1 Then
    �P���ʼȮ��ܼ�(1) = 5260
    �P���ʼȮ��ܼ�(2) = 4840
ElseIf pageqlead(1) > 1 Then
    xy = (pageqlead(1) - 1) * 460
    �P���ʼȮ��ܼ�(1) = (Val(5260) - xy) + ((pageqlead(1) - 1) * Val(960))
    �P���ʼȮ��ܼ�(2) = 4840
End If

End Sub
Sub �y�Эp��_�ϥΪ̤�P()
If Val(FormMainMode.pageusglead) <= 9 Then
    �P���ʼȮ��ܼ�(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 1) '�p��Left�y��
Else
   �P���ʼȮ��ܼ�(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 10)
End If

If Val(FormMainMode.pageusglead) <= 9 Then
   �P���ʼȮ��ܼ�(2) = 6700 '���wTop�y��
Else
   �P���ʼȮ��ܼ�(2) = 7980 '���wTop�y��
End If
End Sub
Sub �P���ǼW�[_�X�P_�q��(ByRef m As Integer)
pagecardnum(m, 7) = pagecomleadmax(1) + 1
pagecomleadmax(1) = pagecomleadmax(1) + 1
End Sub
Sub �P���ǼW�[_��P_�q��(ByRef m As Integer)
pagecardnum(m, 7) = pagecomleadmax(0) + 1
pagecomleadmax(0) = pagecomleadmax(0) + 1
End Sub
Sub �P���ǼW�[_��P_�ϥΪ�(ByVal m As Integer)
pagecardnum(m, 7) = pageusleadmax(0) + 1
pageusleadmax(0) = pageusleadmax(0) + 1
End Sub
Sub �P���ǼW�[_�X�P_�ϥΪ�(ByRef m As Integer)
pagecardnum(m, 7) = pageusleadmax(1) + 1
pageusleadmax(1) = pageusleadmax(1) + 1
End Sub
Sub ����ʧ@_�q��_��P(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) - 1
    �ثe��(9) = pagecardnum(n, 7)
    pagecardnum(n, 6) = 3
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '���w�ثeLeft(�y��)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �ثe��(15) = 5
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
End Sub
Sub ����ʧ@_�~�P_��()
For g = 1 To 57
     If pagecardnum(g, 6) = 3 Then
         pagegive = Val(pagegive) - 1
         pagecardnum(g, 6) = 4
     End If
Next
FormMainMode.pageul = 57 - Val(pagegive)
End Sub
Sub ����ʧ@_�~�P()
For g = 1 To 57
     If pagecardnum(g, 6) = 3 Then
         ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) - 1
         pagecardnum(g, 6) = 4
         Select Case pagecardnum(g, 8)
            Case "021"  '==��1�j1��
                 ���εP�U�P����������(1, 1) = Val(���εP�U�P����������(1, 1)) - 1
            Case "019"  '==��1�j2��
                 ���εP�U�P����������(2, 1) = Val(���εP�U�P����������(2, 1)) - 1
            Case "017"  '==��1�j3��
                 ���εP�U�P����������(3, 1) = Val(���εP�U�P����������(3, 1)) - 1
            Case "025"  '==��1��1��
                 ���εP�U�P����������(4, 1) = Val(���εP�U�P����������(4, 1)) - 1
            Case "024"  '==��1��2��
                 ���εP�U�P����������(5, 1) = Val(���εP�U�P����������(5, 1)) - 1
            Case "023"  '==��1��3��
                 ���εP�U�P����������(6, 1) = Val(���εP�U�P����������(6, 1)) - 1
            Case "026"  '==��2�S3��
                 ���εP�U�P����������(7, 1) = Val(���εP�U�P����������(7, 1)) - 1
            Case "027"  '==��3��3��
                 ���εP�U�P����������(8, 1) = Val(���εP�U�P����������(8, 1)) - 1
            Case "001"  '==�C6�C6��
                 ���εP�U�P����������(9, 1) = Val(���εP�U�P����������(9, 1)) - 1
            Case "011"  '==�C1�j1��
                 ���εP�U�P����������(10, 1) = Val(���εP�U�P����������(10, 1)) - 1
            Case "007"  '==�C2�j1��
                 ���εP�U�P����������(11, 1) = Val(���εP�U�P����������(11, 1)) - 1
            Case "006"  '==�C2�j2��
                 ���εP�U�P����������(12, 1) = Val(���εP�U�P����������(12, 1)) - 1
            Case "004"  '==�C3�j3��
                 ���εP�U�P����������(13, 1) = Val(���εP�U�P����������(13, 1)) - 1
            Case "028"  '==�C5�j5��
                 ���εP�U�P����������(14, 1) = Val(���εP�U�P����������(14, 1)) - 1
            Case "012"  '==�C1��1��
                 ���εP�U�P����������(15, 1) = Val(���εP�U�P����������(15, 1)) - 1
            Case "009"  '==�C2��1��
                 ���εP�U�P����������(16, 1) = Val(���εP�U�P����������(16, 1)) - 1
            Case "008"  '==�C2��2��
                 ���εP�U�P����������(17, 1) = Val(���εP�U�P����������(17, 1)) - 1
            Case "005"  '==�C3��3��
                 ���εP�U�P����������(18, 1) = Val(���εP�U�P����������(18, 1)) - 1
            Case "013"  '==�C1�S1��
                 ���εP�U�P����������(19, 1) = Val(���εP�U�P����������(19, 1)) - 1
            Case "010"  '==�C2�S1��
                 ���εP�U�P����������(20, 1) = Val(���εP�U�P����������(20, 1)) - 1
            Case "003"  '==�C4�S1��
                 ���εP�U�P����������(21, 1) = Val(���εP�U�P����������(21, 1)) - 1
            Case "002"  '==�C5�S2��
                 ���εP�U�P����������(22, 1) = Val(���εP�U�P����������(22, 1)) - 1
            Case "015"  '==�j4�j4��
                 ���εP�U�P����������(23, 1) = Val(���εP�U�P����������(23, 1)) - 1
            Case "020"  '==�j2�S1��
                 ���εP�U�P����������(24, 1) = Val(���εP�U�P����������(24, 1)) - 1
            Case "018"  '==�j3�S2��
                 ���εP�U�P����������(25, 1) = Val(���εP�U�P����������(25, 1)) - 1
            Case "016"  '==�j4�S1��
                 ���εP�U�P����������(26, 1) = Val(���εP�U�P����������(26, 1)) - 1
            Case "014"  '==�j5�S2��
                 ���εP�U�P����������(27, 1) = Val(���εP�U�P����������(27, 1)) - 1
            Case "022"  '==��5��5��
                 ���εP�U�P����������(28, 1) = Val(���εP�U�P����������(28, 1)) - 1
            Case "029"  '==��3�S5��
                 ���εP�U�P����������(29, 1) = Val(���εP�U�P����������(29, 1)) - 1
         End Select
     End If
Next
FormMainMode.pageul = Val(���εP�U�P����������(0, 2)) - Val(���εP�U�P����������(0, 1))
End Sub
Sub ����ʧ@_�M���Ҧ����`���A_�q��()
For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
   If �H�����`���A��Ʈw(2, i, 2) > 0 Then
      ���`���A�ˬd��(�H�����`���A��Ʈw(2, i, 3), 2) = 0
      �H�����`���A��Ʈw(2, i, 2) = 0
   End If
Next
�԰��t����.���`���A�~��_�q��
End Sub
Sub ����ʧ@_�M���Ҧ����`���A_�ϥΪ�()
For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
   If �H�����`���A��Ʈw(1, i, 2) > 0 Then
      ���`���A�ˬd��(�H�����`���A��Ʈw(1, i, 3), 2) = 0
      �H�����`���A��Ʈw(1, i, 2) = 0
   End If
Next
�԰��t����.���`���A�~��_�ϥΪ�
End Sub
Sub ����ʧ@_�Z���ܧ�(ByVal m As Integer)
Dim anw(1 To 2) As Integer
Dim anh(1 To 2) As Integer
anw(1) = Val(FormMainMode.personusminijpg.�p�H���Ϥ�width) / 2
anw(2) = Val(FormMainMode.personcomminijpg.�p�H���Ϥ�width) / 2
anh(1) = Val(FormMainMode.personusminijpg.�p�H���Ϥ�height)
anh(2) = Val(FormMainMode.personcomminijpg.�p�H���Ϥ�height)
Select Case m
  Case 1
    FormMainMode.movejpg.�p�H���Ϥ� = app_path & "\gif\short.png"
    FormMainMode.movejpg.Left = 4440
    FormMainMode.movejpg.Top = 2520
'    formmainmode.personusminijpg.Left = personminixy(1, ����H����ԤH��(1, 2), 1, 1)
'    formmainmode.personcomminijpg.Left = personminixy(2, ����H����ԤH��(2, 2), 1, 1)
    FormMainMode.personusminijpg.Left = 4320 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 7080 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
  Case 2
    FormMainMode.movejpg.�p�H���Ϥ� = app_path & "\gif\middle.png"
    FormMainMode.movejpg.Left = 2880
    FormMainMode.movejpg.Top = 2000
'    formmainmode.personusminijpg.Left = personminixy(1, ����H����ԤH��(1, 2), 2, 1)
'    formmainmode.personcomminijpg.Left = personminixy(2, ����H����ԤH��(2, 2), 2, 1)
    FormMainMode.personusminijpg.Left = 2640 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 8680 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
  Case 3
    FormMainMode.movejpg.�p�H���Ϥ� = app_path & "\gif\long.png"
    FormMainMode.movejpg.Left = 1080
    FormMainMode.movejpg.Top = 2360
'    formmainmode.personusminijpg.Left = personminixy(1, ����H����ԤH��(1, 2), 3, 1)
'    formmainmode.personcomminijpg.Left = personminixy(2, ����H����ԤH��(2, 2), 3, 1)
    FormMainMode.personusminijpg.Left = 1040 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 10320 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
End Select
'============�H�U�O���`���A�ˬd�αҰ�
���`���A�ˬd��(33, 1) = 1
���`���A.�G��_�ϥΪ� m  '(���q1)
'=====
���`���A�ˬd��(34, 1) = 1
���`���A.�G��_�q�� m  '(���q1)
'============
movecp = m
End Sub
Sub �p��P���ʶZ�����()
If �P���ʼȮ��ܼ�(1) >= pagecardnum(�P���ʼȮ��ܼ�(3), 9) Then
   �Z�����(2, 1, 1) = (�P���ʼȮ��ܼ�(1) - pagecardnum(�P���ʼȮ��ܼ�(3), 9)) \ 12
Else
   �Z�����(2, 1, 1) = -((pagecardnum(�P���ʼȮ��ܼ�(3), 9) - �P���ʼȮ��ܼ�(1)) \ 12)
End If

If �P���ʼȮ��ܼ�(2) >= pagecardnum(�P���ʼȮ��ܼ�(3), 10) Then
   �Z�����(2, 1, 2) = (�P���ʼȮ��ܼ�(2) - pagecardnum(�P���ʼȮ��ܼ�(3), 10)) \ 12
Else
   �Z�����(2, 1, 2) = -((pagecardnum(�P���ʼȮ��ܼ�(3), 10) - �P���ʼȮ��ܼ�(2)) \ 12)
End If
End Sub
Sub ���`���A�~��_�ϥΪ�()
For k = 1 To 3
    For i = 14 * (����ݾ��H��������(1, k) - 1) + 1 To (14 * ����ݾ��H��������(1, k)) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
             If �H�����`���A��Ʈw(1, i + 1, 2) > 0 Then
                  FormMainMode.personusspe(i).���`���A�Ϥ� = FormMainMode.personusspe(i + 1).���`���A�Ϥ�
                  FormMainMode.personusspe(i).person_num = FormMainMode.personusspe(i + 1).person_num
                  FormMainMode.personusspe(i).person_turn = FormMainMode.personusspe(i + 1).person_turn
                  �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i + 1, 2)
                  �H�����`���A��Ʈw(1, i, 3) = �H�����`���A��Ʈw(1, i + 1, 3)
                  �H�����`���A��Ʈw(1, i, 1) = �H�����`���A��Ʈw(1, i + 1, 1)
                  For j = 1 To 3
                     �H�����`���A��Ʈw(1, i + 1, j) = 0
                  Next
                  FormMainMode.personusspe(i + 1).Visible = False
                  FormMainMode.personusspe(i).Visible = True
             Else
                  For j = 1 To 3
                     �H�����`���A��Ʈw(1, i, j) = 0
                  Next
                  FormMainMode.personusspe(i).Visible = False
             End If
        End If
    Next
Next
End Sub
Sub ���`���A�~��_�q��()
For k = 1 To 3
    For i = 14 * (����ݾ��H��������(2, k) - 1) + 1 To (14 * ����ݾ��H��������(2, k)) - 1
          If �H�����`���A��Ʈw(2, i, 2) = 0 Then
              If �H�����`���A��Ʈw(2, i + 1, 2) > 0 Then
                  FormMainMode.personcomspe(i).���`���A�Ϥ� = FormMainMode.personcomspe(i + 1).���`���A�Ϥ�
                  FormMainMode.personcomspe(i).person_num = FormMainMode.personcomspe(i + 1).person_num
                  FormMainMode.personcomspe(i).person_turn = FormMainMode.personcomspe(i + 1).person_turn
                  �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i + 1, 2)
                  �H�����`���A��Ʈw(2, i, 3) = �H�����`���A��Ʈw(2, i + 1, 3)
                  �H�����`���A��Ʈw(2, i, 1) = �H�����`���A��Ʈw(2, i + 1, 1)
                  For j = 1 To 3
                     �H�����`���A��Ʈw(2, i + 1, j) = 0
                  Next
                  FormMainMode.personcomspe(i + 1).Visible = False
                  FormMainMode.personcomspe(i).Visible = True
              Else
                  For j = 1 To 3
                     �H�����`���A��Ʈw(2, i, j) = 0
                  Next
                  FormMainMode.personcomspe(i).Visible = False
              End If
          End If
    Next
Next
End Sub
Sub �S��_�v��L_�������A_�ϥΪ�()
Select Case atking_�v��L_�����Ҧ����A��(1)
   Case 1
            If atking_�v��L_�����Ҧ����A��(5) = 0 Then
                atking_�v��L_�����Ҧ����A��(3) = �������m��l�`��(1)
                atking_�v��L_�����Ҧ����A��(4) = �������m��l�`��(1) * 2
                atking_�v��L_�����Ҧ����A��(5) = 1
                �������m��l�`��(1) = �������m��l�`��(1) * 2
            ElseIf atking_�v��L_�����Ҧ����A��(5) = 1 Then
                atking_�v��L_�����Ҧ����A��(3) = atking_�v��L_�����Ҧ����A��(3) + (�������m��l�`��(1) - atking_�v��L_�����Ҧ����A��(4))
                �������m��l�`��(1) = atking_�v��L_�����Ҧ����A��(3) * 2
                atking_�v��L_�����Ҧ����A��(4) = atking_�v��L_�����Ҧ����A��(3) * 2
            End If
    Case 2
           atking_�v��L_�����Ҧ����A��(3) = 0
           atking_�v��L_�����Ҧ����A��(4) = 0
           atking_�v��L_�����Ҧ����A��(5) = 0
    Case 3
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.�p�H���Ϥ� = app_path & "gif\�v��L\�@��\Staciamini1.png"
            FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = app_path & "gif\�v��L\�@��\Staciaminidown1.png"
            FormMainMode.personusminijpg.�p�H���v�lLeft = 10
            FormMainMode.personusminijpg.�p�H���v�ltop�t = -50
            FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\�v��L\�@��\Staciaperson1.png"
            FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = app_path & "gif\�v��L\�@��\Staciaf1.png"
            atking_�v��L_�����Ҧ����A��(2) = 0
            atking_�v��L_�����Ҧ����A��(3) = 0
            atking_�v��L_�����Ҧ����A��(4) = 0
            atking_�v��L_�����Ҧ����A��(5) = 0
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personusminijpg.Visible = True
    Case 4
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.�p�H���Ϥ� = app_path & "gif\�v��L\����\Staciamini1.png"
            FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = app_path & "gif\�v��L\����\Staciaminidown1.png"
            FormMainMode.personusminijpg.�p�H���v�lLeft = -90
            FormMainMode.personusminijpg.�p�H���v�ltop�t = -60
            FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\�v��L\����\Staciaperson1.png"
            FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = app_path & "gif\�v��L\����\Staciaf1.png"
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personusminijpg.Visible = True
End Select
End Sub
Sub �S��_�v��L_�������A_�q��()
Select Case atking_AI_�v��L_�����Ҧ����A��(1)
   Case 1
            If atking_AI_�v��L_�����Ҧ����A��(5) = 0 Then
                atking_AI_�v��L_�����Ҧ����A��(3) = �������m��l�`��(2)
                atking_AI_�v��L_�����Ҧ����A��(4) = �������m��l�`��(2) * 2
                atking_AI_�v��L_�����Ҧ����A��(5) = 1
                �������m��l�`��(2) = �������m��l�`��(2) * 2
            ElseIf atking_AI_�v��L_�����Ҧ����A��(5) = 1 Then
                atking_AI_�v��L_�����Ҧ����A��(3) = atking_AI_�v��L_�����Ҧ����A��(3) + (�������m��l�`��(2) - atking_AI_�v��L_�����Ҧ����A��(4))
                �������m��l�`��(2) = atking_AI_�v��L_�����Ҧ����A��(3) * 2
                atking_AI_�v��L_�����Ҧ����A��(4) = atking_AI_�v��L_�����Ҧ����A��(3) * 2
            End If
    Case 2
           atking_AI_�v��L_�����Ҧ����A��(3) = 0
           atking_AI_�v��L_�����Ҧ����A��(4) = 0
           atking_AI_�v��L_�����Ҧ����A��(5) = 0
    Case 3
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.�p�H���Ϥ� = app_path & "gif\�v��L\�@��\Staciamini2.png"
            FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = app_path & "gif\�v��L\�@��\Staciaminidown2.png"
            FormMainMode.personcomminijpg.�p�H���v�lLeft = 10
            FormMainMode.personcomminijpg.�p�H���v�ltop�t = -50
            FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\�v��L\�@��\Staciaperson2.png"
            FormMainMode.��ܦC1.�q����p�H���Ϥ� = app_path & "gif\�v��L\�@��\Staciaf2.png"
            atking_AI_�v��L_�����Ҧ����A��(2) = 0
            atking_AI_�v��L_�����Ҧ����A��(3) = 0
            atking_AI_�v��L_�����Ҧ����A��(4) = 0
            atking_AI_�v��L_�����Ҧ����A��(5) = 0
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personcomminijpg.Visible = True
    Case 4
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.�p�H���Ϥ� = app_path & "gif\�v��L\����\Staciamini2.png"
            FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = app_path & "gif\�v��L\����\Staciaminidown2.png"
            FormMainMode.personcomminijpg.�p�H���v�lLeft = 90
            FormMainMode.personcomminijpg.�p�H���v�ltop�t = -60
            FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\�v��L\����\Staciaperson2.png"
            FormMainMode.��ܦC1.�q����p�H���Ϥ� = app_path & "gif\�v��L\����\Staciaf2.png"
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personcomminijpg.Visible = True
End Select
End Sub

Sub �S��_������_�������A_�ϥΪ�()
Select Case atking_������_�����Ҧ����A��(1)
   Case 1
            �������m��l�`��(1) = 10
    Case 2
           �������m��l�`��(1) = 10
           �԰��t����.�����g�J��ܦC�ƭ� 1, 10
    Case 3
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.�p�H���Ϥ� = app_path & "gif\������\�@��\Nenemmini1.png"
            FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = app_path & "gif\������\�@��\Nenemminidown1.png"
            FormMainMode.personusminijpg.�p�H���v�lLeft = 10
            FormMainMode.personusminijpg.�p�H���v�ltop�t = -20
            FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\������\�@��\Nenemperson1.png"
            FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = app_path & "gif\������\�@��\Nenemf1.png"
            atking_������_�����Ҧ����A��(2) = 0
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personusminijpg.Visible = True
    Case 4
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.�p�H���Ϥ� = app_path & "gif\������\����\Nenemmini1.png"
            FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = app_path & "gif\������\����\Nenemminidown1.png"
            FormMainMode.personusminijpg.�p�H���v�lLeft = 20
            FormMainMode.personusminijpg.�p�H���v�ltop�t = -90
            FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\������\����\Nenemperson1.png"
            FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = app_path & "gif\������\����\Nenemf1.png"
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personusminijpg.Visible = True
End Select
End Sub
Sub �S��_������_�������A_�q��()
Select Case atking_AI_������_�����Ҧ����A��(1)
   Case 1
            �������m��l�`��(2) = 10
    Case 2
           �������m��l�`��(2) = 10
           �԰��t����.�����g�J��ܦC�ƭ� 2, 10
    Case 3
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.�p�H���Ϥ� = app_path & "gif\������\�@��\Nenemmini2.png"
            FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = app_path & "gif\������\�@��\Nenemminidown2.png"
            FormMainMode.personcomminijpg.�p�H���v�lLeft = 10
            FormMainMode.personcomminijpg.�p�H���v�ltop�t = -20
            FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\������\�@��\Nenemperson2.png"
            FormMainMode.��ܦC1.�q����p�H���Ϥ� = app_path & "gif\������\�@��\Nenemf2.png"
            atking_AI_������_�����Ҧ����A��(2) = 0
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personcomminijpg.Visible = True
    Case 4
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.�p�H���Ϥ� = app_path & "gif\������\����\Nenemmini2.png"
            FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = app_path & "gif\������\����\Nenemminidown2.png"
            FormMainMode.personcomminijpg.�p�H���v�lLeft = 20
            FormMainMode.personcomminijpg.�p�H���v�ltop�t = -90
            FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\������\����\Nenemperson2.png"
            FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = app_path & "gif\������\����\Nenemf2.png"
            �԰��t����.����ʧ@_�Z���ܧ� movecp
            FormMainMode.personcomminijpg.Visible = True
End Select
End Sub

Sub �S��_����_�@���ø��_�ϥΪ�()
Dim m As Integer
Randomize
m = Int(Rnd() * 3) + 1
Select Case m
    Case 1
       FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\����\Blauperson1-1.png"
    Case 2
       FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\����\Blauperson1-2.png"
    Case 3
       FormDice.jpgus.�j�H���Ϥ� = app_path & "gif\����\Blauperson1-3.png"
End Select
End Sub
Sub �S��_����_�@���ø��_�q��()
Dim m As Integer
Randomize
m = Int(Rnd() * 3) + 1
Select Case m
    Case 1
       FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\����\Blauperson2-1.png"
    Case 2
       FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\����\Blauperson2-2.png"
    Case 3
       FormDice.jpgcom.�j�H���Ϥ� = app_path & "gif\����\Blauperson2-3.png"
End Select
End Sub
Function �S��_�ײ��d_�ˬd�W���O�_�Ұ�_�ϥΪ�() As Boolean
If atkingck(49, 2) = 1 And atking_�ײ��d_�W���ثe���q������(3) > 0 Then
    �S��_�ײ��d_�ˬd�W���O�_�Ұ�_�ϥΪ� = True
Else
    �S��_�ײ��d_�ˬd�W���O�_�Ұ�_�ϥΪ� = False
End If
End Function
Function �S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q��() As Boolean
If atkingckai(139, 2) = 1 And atking_AI_�ײ��d_�W���ثe���q������(3) > 0 Then
    �S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = True
Else
    �S��_�ײ��d_�ˬd�W���O�_�Ұ�_�q�� = False
End If
End Function
Sub comatk_AI_����_���j�¤�_�C(ByVal i As Integer)
            If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
               If pagecardnum(i, 1) = a1a Then
                  pagecardnum(i, 11) = 1
              ElseIf pagecardnum(i, 3) = a1a Then
                  cspce = pagecardnum(i, 1)
                  cspme = pagecardnum(i, 2)
                  pagecardnum(i, 1) = pagecardnum(i, 3)
                  pagecardnum(i, 2) = pagecardnum(i, 4)
                  pagecardnum(i, 3) = cspce
                  pagecardnum(i, 4) = cspme
                  If pageonin(i) = 2 Then
                     pageonin(i) = 1
                  Else
                     pageonin(i) = 2
                  End If
                  pagecardnum(i, 11) = 1
               End If
            End If

End Sub
Sub comatk_AI_����_���b�B_��(j As Integer)
If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 Then
     If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) = 1 Then
       pagecardnum(j, 11) = 1
     ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) = 1 Then
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
       pagecardnum(j, 11) = 1
     End If
  End If
End Sub
Sub comatk_AI_�Ǧh_�]�G����_��(j As Integer)
If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 Then
     If pagecardnum(j, 1) = a3a And Val(pagecardnum(j, 2)) >= 1 Then
       pagecardnum(j, 11) = 1
     ElseIf pagecardnum(j, 3) = a3a And Val(pagecardnum(j, 4)) >= 1 Then
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
       pagecardnum(j, 11) = 1
     End If
  End If
End Sub
Sub comatk_AI_����_�۱��ɦV_�S(ByVal a As Integer)
            If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 5)) <> 1 Then
               If pagecardnum(a, 1) = a4a Then
                  pagecardnum(a, 11) = 1
              ElseIf pagecardnum(a, 3) = a4a Then
                  cspce = pagecardnum(a, 1)
                  cspme = pagecardnum(a, 2)
                  pagecardnum(a, 1) = pagecardnum(a, 3)
                  pagecardnum(a, 2) = pagecardnum(a, 4)
                  pagecardnum(a, 3) = cspce
                  pagecardnum(a, 4) = cspme
                  If pageonin(a) = 2 Then
                     pageonin(a) = 1
                  Else
                     pageonin(a) = 2
                  End If
                  pagecardnum(a, 11) = 1
               End If
            End If

End Sub
Sub comatk_AI_����_�h�g�H_�����_�S(ByVal a As Integer)
            If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
               If pagecardnum(a, 1) = a4a Then
                  pagecardnum(a, 11) = 1
              ElseIf pagecardnum(a, 3) = a4a Then
                  cspce = pagecardnum(a, 1)
                  cspme = pagecardnum(a, 2)
                  pagecardnum(a, 1) = pagecardnum(a, 3)
                  pagecardnum(a, 2) = pagecardnum(a, 4)
                  pagecardnum(a, 3) = cspce
                  pagecardnum(a, 4) = cspme
                  If pageonin(a) = 2 Then
                     pageonin(a) = 1
                  Else
                     pageonin(a) = 2
                  End If
                  pagecardnum(a, 11) = 1
               End If
            End If

End Sub
Sub comatk_AI_��_�צ�_�L�ɽ��j���׵�_�S(ByVal a As Integer)
            If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 Then
               If pagecardnum(a, 1) = a4a Then
                  pagecardnum(a, 11) = 1
              ElseIf pagecardnum(a, 3) = a4a Then
                  cspce = pagecardnum(a, 1)
                  cspme = pagecardnum(a, 2)
                  pagecardnum(a, 1) = pagecardnum(a, 3)
                  pagecardnum(a, 2) = pagecardnum(a, 4)
                  pagecardnum(a, 3) = cspce
                  pagecardnum(a, 4) = cspme
                  If pageonin(a) = 2 Then
                     pageonin(a) = 1
                  Else
                     pageonin(a) = 2
                  End If
                  pagecardnum(a, 11) = 1
               End If
            End If

End Sub
Sub �����g�J��ܦC�ƭ�(ByVal n As Integer, ByVal num As Integer)
If num < 0 Then num = 0
Select Case n
    Case 1
        FormMainMode.��ܦC1.goi1 = num
    Case 2
        FormMainMode.��ܦC1.goi2 = num
End Select
End Sub
Sub �p�H���Y�����槹�P�__�ϥΪ�()
If turnatk = 1 Or turnatk = 2 Then
   turnpageonin = 1
'   ���q���A�� = 1
End If
If turnatk = 3 Then
    FormMainMode.trtimeline.Enabled = True
End If
End Sub
Sub �p�H���Y�����槹�P�__�q��()
If turnatk = 1 Or turnatk = 2 Or turnatk = 3 Then
   ���q���A�� = 3
   FormMainMode.�q���X�P.Enabled = True
End If
End Sub
Sub ���εP�ܭI��()
FormMainMode.card(�P���ʼȮ��ܼ�(3)).Width = 720
FormMainMode.card(�P���ʼȮ��ܼ�(3)).Height = 990
FormMainMode.card(�P���ʼȮ��ܼ�(3)).Picture = LoadPicture(app_path & "card\cardback.bmp")
End Sub
Sub ���εP�^�_����(ByVal num As Integer)
FormMainMode.card(num).Width = 810
FormMainMode.card(num).Height = 1260
FormMainMode.card(num).Picture = LoadPicture(app_path & "card\" & pagecardnum(num, 8) & "-" & pageonin(num) & ".bmp")
End Sub
Sub �X�P���ǭp��_�ϥΪ�_��P()
Dim pagegustot As Integer '�Ȯ��ܼ�

For i = 1 To 106
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(2, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
    pagegustot = Val(pagegustot) + 1
    �X�P���ǲέp�Ȯ��ܼ�(2, pagegustot, 1) = Val(pagecardnum(i, 7))
    �X�P���ǲέp�Ȯ��ܼ�(2, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(2, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(2, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(2, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(2, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(2, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(2, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(2, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(2, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(2, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(2, o, 2) = h
   End If
  Next
Next
'MsgBox 123
End Sub
Sub �X�P���ǭp��_�ϥΪ�_�X�P()
Dim pagegustot As Integer '�Ȯ��ܼ�

For i = 1 To 106
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(1, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
    pagegustot = Val(pagegustot) + 1
    �X�P���ǲέp�Ȯ��ܼ�(1, pagegustot, 1) = Val(pagecardnum(i, 7))
    �X�P���ǲέp�Ȯ��ܼ�(1, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(1, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(1, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(1, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(1, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(1, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(1, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(1, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(1, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(1, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(1, o, 2) = h
   End If
  Next
Next

End Sub
Sub �X�P���ǭp��_�q��_��P()
Dim pagegustot As Integer '�Ȯ��ܼ�

For i = 1 To 106
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(4, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
       pagegustot = Val(pagegustot) + 1
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 1) = Val(pagecardnum(i, 7))
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 2) = i
   ElseIf Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 1 Then
       pagegustot = Val(pagegustot) + 1
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 1) = Val(pagecardnum(i, 7))
       �X�P���ǲέp�Ȯ��ܼ�(4, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(4, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(4, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(4, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(4, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(4, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(4, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(4, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(4, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(4, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(4, o, 2) = h
   End If
  Next
Next
End Sub
Sub �X�P���ǭp��_�q��_�X�P()
Dim pagegustot As Integer '�Ȯ��ܼ�

For i = 1 To 106
   For j = 1 To 2
      �X�P���ǲέp�Ȯ��ܼ�(3, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 2 Then
       pagegustot = Val(pagegustot) + 1
       �X�P���ǲέp�Ȯ��ܼ�(3, pagegustot, 1) = Val(pagecardnum(i, 7))
       �X�P���ǲέp�Ȯ��ܼ�(3, pagegustot, 2) = i
    End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If �X�P���ǲέp�Ȯ��ܼ�(3, o, 1) > �X�P���ǲέp�Ȯ��ܼ�(3, i, 1) Then
    g = �X�P���ǲέp�Ȯ��ܼ�(3, i, 1)
    h = �X�P���ǲέp�Ȯ��ܼ�(3, i, 2)
    �X�P���ǲέp�Ȯ��ܼ�(3, i, 1) = �X�P���ǲέp�Ȯ��ܼ�(3, o, 1)
    �X�P���ǲέp�Ȯ��ܼ�(3, i, 2) = �X�P���ǲέp�Ȯ��ܼ�(3, o, 2)
    �X�P���ǲέp�Ȯ��ܼ�(3, o, 1) = g
    �X�P���ǲέp�Ȯ��ܼ�(3, o, 2) = h
   End If
  Next
Next
End Sub
Sub ���P�p��Z�����_�ϥΪ�()
For i = 1 To 106
    �Z�����_���P�Ȯɼ�(i, 1) = 0
    �Z�����_���P�Ȯɼ�(i, 2) = 0
Next

�԰��t����.�X�P���ǭp��_�ϥΪ�_�X�P
For i = 1 To pageqlead(1)
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = �X�P���ǲέp�Ȯ��ܼ�(1, i, 2)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2), 9) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2)).Left  '���w�ثeLeft(�y��)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2), 10) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(1, i, 2)).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �Z�����_���P�Ȯɼ�(i, 1) = �Z�����(2, 1, 1)
    �Z�����_���P�Ȯɼ�(i, 2) = �Z�����(2, 1, 2)
    �Z�����_���P�Ȯɼ�(i, 3) = �X�P���ǲέp�Ȯ��ܼ�(1, i, 2)
Next
End Sub
Sub ���P�p��Z�����_�q��()
For i = 1 To 106
    �Z�����_���P�Ȯɼ�(i, 1) = 0
    �Z�����_���P�Ȯɼ�(i, 2) = 0
Next

�԰��t����.�X�P���ǭp��_�q��_�X�P
For i = 1 To pageqlead(2)
    �P���ʼȮ��ܼ�(1) = 240
    �P���ʼȮ��ܼ�(2) = 960
    �P���ʼȮ��ܼ�(3) = �X�P���ǲέp�Ȯ��ܼ�(3, i, 2)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2), 9) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2)).Left  '���w�ثeLeft(�y��)
    pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2), 10) = FormMainMode.card(�X�P���ǲέp�Ȯ��ܼ�(3, i, 2)).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �Z�����_���P�Ȯɼ�(i, 1) = �Z�����(2, 1, 1)
    �Z�����_���P�Ȯɼ�(i, 2) = �Z�����(2, 1, 2)
    �Z�����_���P�Ȯɼ�(i, 3) = �X�P���ǲέp�Ȯ��ܼ�(3, i, 2)
Next
End Sub
Sub �ޯ�Ұʼƶq�ˬd()
FormMainMode.atkingnumtot.Caption = Val(atkingtrn(1)) + Val(atkingtrn(2))
Erase atkingno
End Sub
Sub �ޯ໡�����J_�ϥΪ�(ByVal n As Integer)
Dim ahmt As String
FormMainMode.atkinghelpt1.Caption = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 2)
FormMainMode.atkinghelpt2.Caption = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 3)
FormMainMode.atkinghelpt3.Caption = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 4)
ahmt = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 5)
For i = 1 To Len(ahmt)
    If Mid(ahmt, i, 1) = "&" Then
        Mid(ahmt, i, 1) = Chr(10)
    End If
Next
FormMainMode.atkinghelpt4.Caption = ahmt
If VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 6) <> "" Then
    FormMainMode.atkinghelpt3.FontSize = Val(VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 6))
Else
    FormMainMode.atkinghelpt3.FontSize = 10
End If
If VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 7) <> "" Then
    FormMainMode.atkinghelpt4.FontSize = Val(VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 7))
Else
    FormMainMode.atkinghelpt4.FontSize = 10
End If
End Sub
Sub �ޯ໡�����J_�q��(ByVal n As Integer)
Dim ahmt As String
FormMainMode.atkinghelpt1.Caption = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 2)
FormMainMode.atkinghelpt2.Caption = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 3)
FormMainMode.atkinghelpt3.Caption = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 4)
ahmt = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 5)
For i = 1 To Len(ahmt)
    If Mid(ahmt, i, 1) = "&" Then
        Mid(ahmt, i, 1) = Chr(10)
    End If
Next
FormMainMode.atkinghelpt4.Caption = ahmt

If VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 6) <> "" Then
    FormMainMode.atkinghelpt3.FontSize = Val(VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 6))
Else
    FormMainMode.atkinghelpt3.FontSize = 10
End If
If VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 7) <> "" Then
    FormMainMode.atkinghelpt4.FontSize = Val(VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 7))
Else
    FormMainMode.atkinghelpt4.FontSize = 10
End If
End Sub
Sub ���q�R���ո`�]�w()
If Formsetting.cksemute.Value = 1 Then
   FormMainMode.wmpse1.settings.mute = True
   FormMainMode.wmpse2.settings.mute = True
   FormMainMode.wmpse3.settings.mute = True
   FormMainMode.wmpse4.settings.mute = True
   FormMainMode.wmpse5.settings.mute = True
   FormMainMode.wmpse6.settings.mute = True
   FormMainMode.wmpse7.settings.mute = True
   FormMainMode.wmpse8.settings.mute = True
   FormMainMode.wmpse9.settings.mute = True
Else
   FormMainMode.wmpse1.settings.mute = False
   FormMainMode.wmpse2.settings.mute = False
   FormMainMode.wmpse3.settings.mute = False
   FormMainMode.wmpse4.settings.mute = False
   FormMainMode.wmpse5.settings.mute = False
   FormMainMode.wmpse6.settings.mute = False
   FormMainMode.wmpse7.settings.mute = False
   FormMainMode.wmpse8.settings.mute = False
   FormMainMode.wmpse9.settings.mute = False
End If
End Sub
Sub �ɶ��b_���]()
FormMainMode.timelineout1.X1 = 0
FormMainMode.timelineout2.X2 = 11310
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 1) = 23
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 2) = 77
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(1, 3) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 1) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 2) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(2, 3) = 0
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 1) = 111
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 2) = 251
�ɶ��b�C���ܤƬ����Ȯ��ܼ�(3, 3) = 50
FormMainMode.timelineout1.BorderColor = RGB(111, 251, 50)
FormMainMode.timelineout2.BorderColor = RGB(111, 251, 50)
End Sub
Sub �ɶ��b_����()
FormMainMode.trtimeline.Enabled = False
FormMainMode.timelinein1.BorderColor = RGB(0, 0, 0)
FormMainMode.timelinein2.BorderColor = RGB(0, 0, 0)
End Sub
Sub �ɶ��b_����()
FormMainMode.timeup.Visible = False
FormMainMode.timelinein1.Visible = False
FormMainMode.timelinein2.Visible = False
FormMainMode.timelineout1.Visible = False
FormMainMode.timelineout2.Visible = False
End Sub
Sub �ɶ��b_���()
FormMainMode.timeup.Visible = True
FormMainMode.timelinein1.Visible = True
FormMainMode.timelinein2.Visible = True
FormMainMode.timelineout1.Visible = True
FormMainMode.timelineout2.Visible = True
End Sub
Sub ���q����P�_()
If Val(�Y���淾�q�Ȯ��ܼ�(4)) = 1 Then
   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
    Case 1
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
'           cn3.Visible = True
           �Y���淾�q�Ȯ��ܼ�(1) = 2
           �ثe��(22) = 14
           FormMainMode.���ݮɶ�.Enabled = True
       Else
'           cn4.Visible = True
           �ثe��(22) = 15
           FormMainMode.���ݮɶ�.Enabled = True
       End If
    Case 2
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
'          cn4.Visible = True
          �ثe��(22) = 15
          FormMainMode.���ݮɶ�.Enabled = True
       Else
'          cn2.Visible = True
          �Y���淾�q�Ȯ��ܼ�(1) = 2
          �ثe��(22) = 13
          FormMainMode.���ݮɶ�.Enabled = True
       End If
    End Select
Else
   Select Case Val(�Y���淾�q�Ȯ��ܼ�(1))
    Case 1
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
'          cn4.Visible = True
          �ثe��(22) = 15
          FormMainMode.���ݮɶ�.Enabled = True
       Else
'          cn2.Visible = True
          �Y���淾�q�Ȯ��ܼ�(1) = 2
          �ثe��(22) = 13
          FormMainMode.���ݮɶ�.Enabled = True
       End If
    Case 2
       If �Y���淾�q�Ȯ��ܼ�(4) = 1 Then
'           cn3.Visible = True
           �Y���淾�q�Ȯ��ܼ�(1) = 2
           �ثe��(22) = 14
           FormMainMode.���ݮɶ�.Enabled = True
       Else
'           cn4.Visible = True
           �ثe��(22) = 15
           FormMainMode.���ݮɶ�.Enabled = True
       End If
    End Select
  End If
End Sub
Sub �q���P_�������P(ByVal Index As Integer)
If pagecardnum(Index, 6) = 1 And pagecardnum(Index, 5) = 2 Then
   pagecardnum(Index, 6) = 2
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + Val(pagecardnum(Index, 2))
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
         �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + Val(pagecardnum(Index, 2))
   End If
   '===================
    �ثe��(9) = pagecardnum(Index, 7)
    pagecardnum(Index, 7) = Val(pagecomleadmax(1)) + 1
    pagecomleadmax(1) = Val(pagecomleadmax(1)) + 1
    pageqlead(2) = Val(pageqlead(2)) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) + 1
    pagecardnum(Index, 11) = 2
   '===================�H�U�O�X�P���
    �ثe��(7) = 0
    �԰��t����.�X�P���ǭp��_�q��_�X�P
    FormMainMode.�q���X�P_�X�P���_�a��.Enabled = True
    '============�H�U�O�ޯ��ˬd�αҰ�
    atkingckai(1, 1) = 2
    If turnatk = 2 Then
       AI�ޯ�.����_�۱��ɦV Index '(���q2)
       AI�ޯ�.������_�r�֩�� Index '(���q1)
    End If
    If turnatk = 2 And atkingckai(26, 2) = 1 Then
        atkingckai(26, 1) = 2
        AI�ޯ�.��̬d�w_���t���C Index '(���q2)
        atkingckai(26, 1) = 1
    End If
    If turnatk = 2 And atkingckai(98, 2) = 1 Then
        atkingckai(98, 1) = 2
        AI�ޯ�.�S�{��_���M�C�{ Index  '(���q2)
        atkingckai(98, 1) = 1
    End If
   '=============�H�U�O�P����(�X�P)(�q��)
    �԰��t����.�y�Эp��_�q���X�P
    �P���ʼȮ��ܼ�(3) = Index
    pagecardnum(Index, 9) = FormMainMode.card(Index).Left  '���w�ثeLeft(�y��)
    pagecardnum(Index, 10) = FormMainMode.card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �ثe��(15) = 0
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
   '================�H�U�O��P���
   �ثe��(8) = 0
   �ثe��(17) = 1
   '===================�H�U�O�ƥ�d�ˬd�αҰ�
   If pagecardnum(Index, 1) = a6a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.���|_�q�� Index, pagecardnum(Index, 2)
   End If
   If turnatk = 1 Or turnatk = 2 Then
        If pagecardnum(Index, 1) = a7a Then
            �ƥ�d�O���Ȯɼ�(2, 3) = 1
            �ƥ�d.�A�G�N_�q�� Index, pagecardnum(Index, 2)
        End If
   End If
   If pagecardnum(Index, 1) = a8a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.HP�^�__�q�� Index, pagecardnum(Index, 2)
   End If
   If pagecardnum(Index, 1) = a9a Then
       �ƥ�d�O���Ȯɼ�(2, 3) = 1
       �ƥ�d.�t��_�q�� Index, pagecardnum(Index, 2)
   End If
   '===================
End If

End Sub
Sub �q���P_�������P_�~(ByVal Index As Integer)
If pagecardnum(Index, 6) = 2 And pagecardnum(Index, 5) = 2 Then
   pagecardnum(Index, 6) = 1
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - Val(pagecardnum(Index, 2))
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 2))
         �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - Val(pagecardnum(Index, 2))
   End If
   '================
   �ثe��(9) = pagecardnum(Index, 7)
    pagecardnum(Index, 7) = Val(pagecomleadmax(0)) + 1
    pagecomleadmax(0) = Val(pagecomleadmax(0)) + 1
    pageqlead(2) = Val(pageqlead(2)) - 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
    pagecardnum(Index, 11) = 0
   '============�H�U�O�ޯ��ˬd�αҰ�
    atkingckai(1, 1) = 2
    If turnatk = 2 Then
       AI�ޯ�.����_�۱��ɦV Index '(���q2)
       AI�ޯ�.������_�r�֩�� Index '(���q1)
    End If
    If turnatk = 2 And atkingckai(26, 2) = 1 Then
        atkingckai(26, 1) = 2
        AI�ޯ�.��̬d�w_���t���C Index '(���q2)
        atkingckai(26, 1) = 1
    End If
    If turnatk = 2 And atkingckai(98, 2) = 1 Then
        atkingckai(98, 1) = 2
        AI�ޯ�.�S�{��_���M�C�{ Index  '(���q2)
        atkingckai(98, 1) = 1
    End If
   '=============�H�U�O�P����(�^�P)(�q��)
    �԰��t����.�y�Эp��_�q����P
    �P���ʼȮ��ܼ�(3) = Index
    pagecardnum(Index, 9) = FormMainMode.card(Index).Left  '���w�ثeLeft(�y��)
    pagecardnum(Index, 10) = FormMainMode.card(Index).Top  '���w�ثeTop(�y��)
    �԰��t����.�p��P���ʶZ�����
    �԰��t����.���εP�ܭI��
    �ثe��(15) = 0
    FormMainMode.�P����.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    �@��t����.�ˬd���ּ��� 1
   '================�H�U�O�X�P���
   �ثe��(7) = 0
   �԰��t����.�X�P���ǭp��_�q��_�X�P
   FormMainMode.�q���X�P_�X�P���_�a�k.Enabled = True
   '=====================�H�U�O�ޯ��ˬd�αҰ�(�J�y-�Ѩ����)
   If turnatk = 2 And atkingck(157, 2) = 1 And atkingck(157, 1) = 5 Then
        �ޯ�.�J�y_�Ѩ���� '(���q5)
   End If
    '====================
End If
End Sub
Sub �q���P_������P_�~(ByVal Index As Integer)
uspce = pagecardnum(Index, 1)
uspme = pagecardnum(Index, 2)
pagecardnum(Index, 1) = pagecardnum(Index, 3)
pagecardnum(Index, 2) = pagecardnum(Index, 4)
pagecardnum(Index, 3) = uspce
pagecardnum(Index, 4) = uspme
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
�@��t����.�ˬd���ּ��� 3
If pageonin(Index) = 1 Then
   pageonin(Index) = 2
   FormMainMode.card(Index).Picture = LoadPicture(app_path & "card\" & pagecardnum(Index, 8) & "-" & pageonin(Index) & ".bmp")
Else
   pageonin(Index) = 1
   FormMainMode.card(Index).Picture = LoadPicture(app_path & "card\" & pagecardnum(Index, 8) & "-" & pageonin(Index) & ".bmp")
End If
'goickus = 0

   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + pagecardnum(Index, 2)
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + pagecardnum(Index, 2)
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
          �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + pagecardnum(Index, 2)
      If turnatk = 1 Then
         �������m��l�`��(2) = �������m��l�`��(2) + Val(pagecardnum(Index, 2))
         �������m��l�`��(4) = �������m��l�`��(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + pagecardnum(Index, 2)
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + pagecardnum(Index, 2)
   End If
'======================================
   If pagecardnum(Index, 3) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - pagecardnum(Index, 4)
      If turnatk = 2 And movecp = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - pagecardnum(Index, 4)
      If turnatk = 2 And movecp > 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - pagecardnum(Index, 4)
      If turnatk = 1 Then
          �������m��l�`��(2) = �������m��l�`��(2) - Val(pagecardnum(Index, 4))
          �������m��l�`��(4) = �������m��l�`��(4) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - pagecardnum(Index, 4)
   End If
   If pagecardnum(Index, 3) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - pagecardnum(Index, 4)
   End If
    '============�H�U�O�ޯ��ˬd�αҰ�
    If turnatk = 2 Then
        atkingckai(26, 1) = 3
        AI�ޯ�.��̬d�w_���t���C Index '(���q3)
        atkingckai(1, 1) = 3
        AI�ޯ�.����_�۱��ɦV Index  '(���q3)
        atkingckai(111, 1) = 2
        AI�ޯ�.������_�r�֩�� Index '(���q2)
    End If
    If turnatk = 2 Then
        atkingckai(98, 1) = 3
        AI�ޯ�.�S�{��_���M�C�{ Index '(���q3)
    End If
    '=================
    atkingckai(1, 1) = 1
    atkingckai(111, 1) = 1
    Call FormMainMode.pagecomqlead_Change
End Sub
Sub ��ƹs����P�_()
FormDice.outprocess
End Sub
Sub ����HP�ˬd()
Dim inp As Integer 'RND�Ȯ��ܼ�
Dim person(1 To 2) As Integer
Erase �H�������ˬd�Ȯ��ܼ�
If livecom(����H����ԤH��(2, 2)) <= 0 Then
   �H�������ˬd�Ȯ��ܼ�(3) = 1
   If livecom(����ݾ��H��������(2, 2)) > 0 Then
'       �H���洫_�q��_���w�洫 2
       person(2) = 2
       �洫��������Ȯ��ܼ�(2) = 1
'       �P�`���q��(2) = �P�`���q��(2) + 1
   ElseIf livecom(����ݾ��H��������(2, 3)) > 0 Then
'       �H���洫_�q��_���w�洫 3
       �洫��������Ȯ��ܼ�(2) = 1
       person(2) = 2
'       �P�`���q��(2) = �P�`���q��(2) + 1
   Else
       person(2) = 1
   End If
End If
If Val(FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption) <= 0 Then
   �H�������ˬd�Ȯ��ܼ�(2) = 1
   If Val(FormMainMode.usbi1(����ݾ��H��������(1, 2)).Caption) > 0 Or Val(FormMainMode.usbi1(����ݾ��H��������(1, 3)).Caption) > 0 Then
'       ����ʧ@_�洫�H������_��l
       person(1) = 2
       �洫��������Ȯ��ܼ�(1) = 1
'       �P�`���q��(1) = �P�`���q��(1) + 1
   Else
       person(1) = 1
   End If
End If

If person(1) = 2 Or person(2) = 2 Then
   �ثe��(22) = 21
   FormMainMode.�H�������ˬd.Enabled = True
   Exit Sub
ElseIf person(1) = 0 And person(2) = 1 Then
   �԰��Ҧ��ӱѬ����� = 1
   �ثe��(22) = 36
   FormMainMode.�H�������ˬd.Enabled = True
ElseIf person(1) = 1 And person(2) = 0 Then
   �ثe��(22) = 36
   �԰��Ҧ��ӱѬ����� = 2
   FormMainMode.�H�������ˬd.Enabled = True
ElseIf person(1) = 1 And person(2) = 1 Then
   Randomize
   inp = Int(Rnd() * 2) + 1
   Select Case inp
       Case 1
           �԰��Ҧ��ӱѬ����� = 1
           �ثe��(22) = 36
           FormMainMode.�H�������ˬd.Enabled = True
       Case 2
           �԰��Ҧ��ӱѬ����� = 2
           �ثe��(22) = 36
           FormMainMode.�H�������ˬd.Enabled = True
    End Select
End If

If FormMainMode.�H�������ˬd.Enabled = False Then
  Select Case HP�ˬd���q��
     Case 1
       '----------�H�U�����q�~����]���ʶ��q3�^
        �ثe��(22) = 4
        FormMainMode.���ݮɶ�.Enabled = True
     Case 2
'         atkingnumtot = 0
          �ثe��(22) = 11
          FormMainMode.���ݮɶ�.Enabled = True
     Case 3
        �԰��t����.���q����P�_
     Case 4
        FormMainMode.NextTurn_���q2.Enabled = True
'     Case 5
'        �ثe��(26) = 1
'        formmainmode.��l���槹�Ұ�.Enabled = True
  End Select
End If
End Sub
Function ����HP�ˬd_�����^�X�ˬd() As Boolean
Dim num(1 To 2) As Integer '��ܤH���Ȯ��ܼ�
If turn >= Val(Formsetting.ckendturnnum.Text) And Formsetting.ckendturn.Value = 1 Then
        ����HP�ˬd_�����^�X�ˬd = True
        '==============
        For i = 1 To 3
            If liveus(����ݾ��H��������(1, i)) > 0 Then
                num(1) = Val(num(1)) + Val(liveus(����ݾ��H��������(1, i)))
            End If
            If livecom(����ݾ��H��������(2, i)) > 0 Then
                num(2) = Val(num(2)) + Val(livecom(����ݾ��H��������(2, i)))
            End If
         Next
        '==============
        If num(1) > num(2) Then
           �԰��Ҧ��ӱѬ����� = 1
           FormMainMode.trend.Enabled = True
        ElseIf num(1) < num(2) Then
           �԰��Ҧ��ӱѬ����� = 2
           FormMainMode.trend.Enabled = True
        ElseIf num(1) = num(2) Then
            '�L����ѥ_
            �԰��Ҧ��ӱѬ����� = 2
            FormMainMode.trend.Enabled = True
        End If
Else
     ����HP�ˬd_�����^�X�ˬd = False
End If
End Function

Sub checkpage()

For i = 1 To �ثe��(11)
  If �ثe��(10) = 1 Then
   FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
   pageqlead(1) = Val(pageqlead(1)) - 1
  ElseIf �ثe��(10) = 2 Then
   FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
   pageqlead(2) = Val(pageqlead(2)) - 1
  End If
Next
End Sub
Sub chkcom()
If goicheck(2) = 0 Then
  If atkingpagetot(2, 1) > 0 And movecp = 1 Then
    �������m��l�`��(2) = �������m��l�`��(2) + atkcom(����H����ԤH��(2, 2))
    �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
    goicheck(2) = 1
  ElseIf atkingpagetot(2, 5) > 0 And movecp > 1 Then
    �������m��l�`��(2) = �������m��l�`��(2) + atkcom(����H����ԤH��(2, 2))
    �������m��l�`��(4) = �������m��l�`��(4) + atkcom(����H����ԤH��(2, 2))
    goicheck(2) = 1
  End If
  If goicheck(2) = 1 Then
    '=========�H�U�O�ޯ��ˬd�εo��
        ���`���A�ˬd��(1, 1) = 1
        ���`���A.ATK�[_�q�� '(���q1)
        '=======
        ���`���A�ˬd��(26, 1) = 1
        ���`���A.�t��_�q�� '(���q1)
        '=======
        ���`���A�ˬd��(4, 1) = 1
        ���`���A.ATK��_�q�� '(���q1)
        '=======
        ���`���A�ˬd��(25, 1) = 1
        ���`���A.��O�C�U_�q�� '(���q1)
     '==============
  End If
End If
End Sub
Sub chkdef()
If goidefus = 0 Then
 �������m��l�`��(1) = �������m��l�`��(1) + defus(����H����ԤH��(1, 2))
 �������m��l�`��(3) = �������m��l�`��(3) + defus(����H����ԤH��(1, 2))
 FormMainMode.��ܦC1.goi1 = Val(FormMainMode.��ܦC1.goi1) + defus(����H����ԤH��(1, 2))
 goidefus = 1
   '=========�H�U�O�ޯ��ˬd�εo��
'   If ���`���A�ˬd��(8, 2) = 1 Then
      ���`���A�ˬd��(8, 1) = 1
      ���`���A.DEF�[_�ϥΪ� '(���q1)
'   End If
'   If ���`���A�ˬd��(11, 2) = 1 Then
      ���`���A�ˬd��(11, 1) = 1
      ���`���A.DEF��_�ϥΪ� '(���q1)
'   End If
   ���`���A�ˬd��(13, 1) = 1
   ���`���A.�t��_�ϥΪ� '(���q1)
   '====
   ���`���A�ˬd��(24, 1) = 1
   ���`���A.��O�C�U_�ϥΪ� '(���q1)
   '====
   ���`���A�ˬd��(39, 1) = 1
   ���`���A.�{��_�ϥΪ� '(���q1)
   '==============
End If
End Sub
Sub chkdefcom()
If chkcomck = 0 Then
 �������m��l�`��(2) = �������m��l�`��(2) + defcom(����H����ԤH��(2, 2))
 �������m��l�`��(4) = �������m��l�`��(4) + defcom(����H����ԤH��(2, 2))
 FormMainMode.��ܦC1.goi2 = Val(FormMainMode.��ܦC1.goi2) + defcom(����H����ԤH��(2, 2))
 chkcomck = 1
    '=========�H�U�O�ޯ��ˬd�εo��
'   If ���`���A�ˬd��(8, 2) = 1 Then
      ���`���A�ˬd��(2, 1) = 1
      ���`���A.DEF�[_�q��  '(���q1)
'   End If
'   If ���`���A�ˬd��(11, 2) = 1 Then
      ���`���A�ˬd��(5, 1) = 1
      ���`���A.DEF��_�q�� '(���q1)
'   End If
   ���`���A�ˬd��(26, 1) = 2
   ���`���A.�t��_�q�� '(���q2)
   '===
   ���`���A�ˬd��(25, 1) = 1
   ���`���A.��O�C�U_�q�� '(���q1)
   '==============
End If
End Sub
Sub chkus1()
If goicheck(1) = 0 Then
 If atkingpagetot(1, 1) > 0 Then
   �������m��l�`��(1) = �������m��l�`��(1) + atkus(����H����ԤH��(1, 2))
   �������m��l�`��(3) = �������m��l�`��(3) + atkus(����H����ԤH��(1, 2))
   goicheck(1) = 1
   '=========�H�U�O�ޯ��ˬd�εo��
'   If ���`���A�ˬd��(13, 2) = 1 Then
      ���`���A�ˬd��(13, 1) = 1
      ���`���A.�t��_�ϥΪ� '(���q1)
'   End If
'   If ���`���A�ˬd��(24, 2) = 1 Then
      ���`���A�ˬd��(24, 1) = 1
      ���`���A.��O�C�U_�ϥΪ� '(���q1)
'   End If
'   If ���`���A�ˬd��(7, 2) = 1 Then
      ���`���A�ˬd��(7, 1) = 1
      ���`���A.ATK�[_�ϥΪ� '(���q1)
'   End If
'   If ���`���A�ˬd��(10, 2) = 1 Then
      ���`���A�ˬd��(10, 1) = 1
      ���`���A.ATK��_�ϥΪ� '(���q1)
'   End If
    '====
    ���`���A�ˬd��(39, 1) = 1
    ���`���A.�{��_�ϥΪ�  '(���q1)
   '==============
  End If
End If
End Sub
Sub chkus2()
If goicheck(1) = 0 Then
  If atkingpagetot(1, 5) > 0 Then
   �������m��l�`��(1) = �������m��l�`��(1) + atkus(����H����ԤH��(1, 2))
   �������m��l�`��(3) = �������m��l�`��(3) + atkus(����H����ԤH��(1, 2))
   goicheck(1) = 1
   '=========�H�U�O�ޯ��ˬd�εo��
'   If ���`���A�ˬd��(13, 2) = 1 Then
      ���`���A�ˬd��(13, 1) = 1
      ���`���A.�t��_�ϥΪ� '(���q1)
'   End If
'   If ���`���A�ˬd��(24, 2) = 1 Then
      ���`���A�ˬd��(24, 1) = 1
      ���`���A.��O�C�U_�ϥΪ� '(���q1)
'   End If
'   If ���`���A�ˬd��(7, 2) = 1 Then
      ���`���A�ˬd��(7, 1) = 1
      ���`���A.ATK�[_�ϥΪ� '(���q1)
'   End If
'   If ���`���A�ˬd��(10, 2) = 1 Then
      ���`���A�ˬd��(10, 1) = 1
      ���`���A.ATK��_�ϥΪ� '(���q1)
'   End If
    '====
    ���`���A�ˬd��(39, 1) = 1
    ���`���A.�{��_�ϥΪ�  '(���q1)
   '==============
  End If
End If
End Sub
Sub cleanatkingpagetot()
For i = 1 To 2
     For j = 1 To 5
        atkingpagetot(i, j) = 0
     Next
Next
End Sub
Sub comatk1()

For a = 1 To 106
  If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 Then
     If pagecardnum(a, 1) = a1a Then
       pagecardnum(a, 11) = 1
     ElseIf pagecardnum(a, 3) = a1a Then
       cspce = pagecardnum(a, 1)
       cspme = pagecardnum(a, 2)
       pagecardnum(a, 1) = pagecardnum(a, 3)
       pagecardnum(a, 2) = pagecardnum(a, 4)
       pagecardnum(a, 3) = cspce
       pagecardnum(a, 4) = cspme
       If pageonin(a) = 2 Then
          pageonin(a) = 1
       Else
          pageonin(a) = 2
       End If
       pagecardnum(a, 11) = 1
     End If
  End If
Next
End Sub
Sub comatk2()

For j = 1 To 106
  If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
     If pagecardnum(j, 1) = a5a Then
       pagecardnum(j, 11) = 1
     ElseIf pagecardnum(j, 3) = a5a Then
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
       pagecardnum(j, 11) = 1
     End If
  End If
Next
End Sub
Sub comatk_���z��AI�޾ɵ{��_�W�X�P�i��(ByVal turn As Integer, ByVal movecpre As Integer, ByVal choose As Integer)
Dim werstr As String, werbo As Boolean
If movecpre = 1 And turn = 1 Then
   werstr = a1a
ElseIf movecpre > 1 And turn = 1 Then
   werstr = a5a
ElseIf turn = 2 Then
   werstr = a2a
End If
'=================================
For a = 1 To 106
    werbo = False
    For k = 1 To 10
        If a = cardAInumOvertenrecord(k) Then
            werbo = True
        End If
    Next
    If Val(pagecardnum(a, 6)) = 1 And Val(pagecardnum(a, 5)) = 2 And Val(pagecardnum(a, 11)) <> 1 And werbo = False Then
            If pagecardnum(a, 1) = werstr Then
              pagecardnum(a, 11) = 1
            ElseIf pagecardnum(a, 3) = werstr Then
              cspce = pagecardnum(a, 1)
              cspme = pagecardnum(a, 2)
              pagecardnum(a, 1) = pagecardnum(a, 3)
              pagecardnum(a, 2) = pagecardnum(a, 4)
              pagecardnum(a, 3) = cspce
              pagecardnum(a, 4) = cspme
              If pageonin(a) = 2 Then
                 pageonin(a) = 1
              Else
                 pageonin(a) = 2
              End If
              pagecardnum(a, 11) = 1
            End If
            If choose = 1 And pagecardnum(a, 11) = 0 Then
                pagecardnum(a, 11) = 1
            End If
    End If
Next
End Sub
Sub getpage_��(ByVal k As Integer, m As Integer)
Select Case m
            Case 1
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b6b
               pagecardnum(m, 4) = b6b
               pagecardnum(m, 3) = a1a
               'checkus(m) = 1
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\001-1.bmp")
               pagecardnum(m, 8) = "001"
               pageonin(m) = 1
               pagecardnum(m, 5) = k
               pagecardnum(m, 6) = 1
               'getpageus(k) = 1
             'End If
            Case 2
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b5b
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\002-1.bmp")
               pagecardnum(m, 8) = "002"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 3
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b5b
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 1) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\002-2.bmp")
               pagecardnum(m, 8) = "002"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 4
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b4b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\003-1.bmp")
               pagecardnum(m, 8) = "003"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 5
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b4b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\003-2.bmp")
               pagecardnum(m, 8) = "003"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 6
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 4) = b3b
               pagecardnum(m, 3) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\004-1.bmp")
               pagecardnum(m, 8) = "004"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 7
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b3b
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 1) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\004-2.bmp")
               pagecardnum(m, 8) = "004"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 8
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 4) = b3b
               pagecardnum(m, 3) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\005-1.bmp")
               pagecardnum(m, 8) = "005"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 9
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b3b
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 1) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\005-2.bmp")
               pagecardnum(m, 8) = "005"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 10
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 3) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\006-1.bmp")
               pagecardnum(m, 8) = "006"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 11
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 1) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\006-2.bmp")
               pagecardnum(m, 8) = "006"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 12
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\007-1.bmp")
               pagecardnum(m, 8) = "007"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 13
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\007-2.bmp")
               pagecardnum(m, 8) = "007"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 14
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 3) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\008-1.bmp")
               pagecardnum(m, 8) = "008"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 15
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 1) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\008-2.bmp")
               pagecardnum(m, 8) = "008"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 16
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\009-1.bmp")
               pagecardnum(m, 8) = "009"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 17
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\009-2.bmp")
               pagecardnum(m, 8) = "009"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 18
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\010-1.bmp")
               pagecardnum(m, 8) = "010"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 19
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\011-1.bmp")
               pagecardnum(m, 8) = "011"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 20
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\011-1.bmp")
               pagecardnum(m, 8) = "011"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 21
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\011-2.bmp")
               pagecardnum(m, 8) = "011"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 22
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\011-2.bmp")
               pagecardnum(m, 8) = "011"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 23
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\012-1.bmp")
               pagecardnum(m, 8) = "012"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 24
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\012-1.bmp")
               pagecardnum(m, 8) = "012"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 25
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\012-2.bmp")
               pagecardnum(m, 8) = "012"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 26
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a1a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\012-2.bmp")
               pagecardnum(m, 8) = "012"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 27
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a1a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\013-1.bmp")
               pagecardnum(m, 8) = "013"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 28
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b5b
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\014-1.bmp")
               pagecardnum(m, 8) = "014"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 29
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a5a
               pagecardnum(m, 4) = b5b
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 1) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\014-2.bmp")
               pagecardnum(m, 8) = "014"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 30
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b4b
               pagecardnum(m, 4) = b4b
               pagecardnum(m, 3) = a5a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\015-1.bmp")
               pagecardnum(m, 8) = "015"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 31
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b4b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\016-1.bmp")
               pagecardnum(m, 8) = "016"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 32
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a5a
               pagecardnum(m, 4) = b4b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\016-2.bmp")
               pagecardnum(m, 8) = "016"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 33
              'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\017-1.bmp")
               pagecardnum(m, 8) = "017"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 34
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a5a
               pagecardnum(m, 4) = b3b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\017-2.bmp")
               pagecardnum(m, 8) = "017"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 35
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\018-1.bmp")
               pagecardnum(m, 8) = "018"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 36
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\019-1.bmp")
               pagecardnum(m, 8) = "019"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 37
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a5a
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\019-2.bmp")
               pagecardnum(m, 8) = "019"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 38
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a4a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\020-1.bmp")
               pagecardnum(m, 8) = "020"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 39
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\021-1.bmp")
               pagecardnum(m, 8) = "021"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 40
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\021-1.bmp")
               pagecardnum(m, 8) = "021"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 41
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a5a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\021-1.bmp")
               pagecardnum(m, 8) = "021"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 42
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a5a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\021-2.bmp")
               pagecardnum(m, 8) = "021"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 43
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a5a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\021-2.bmp")
               pagecardnum(m, 8) = "021"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 44
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a5a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\021-2.bmp")
               pagecardnum(m, 8) = "021"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 45
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a2a
               pagecardnum(m, 2) = b5b
               pagecardnum(m, 4) = b5b
               pagecardnum(m, 3) = a2a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\022-1.bmp")
               pagecardnum(m, 8) = "022"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 46
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a2a
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\023-1.bmp")
               pagecardnum(m, 8) = "023"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 47
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a2a
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\024-1.bmp")
               pagecardnum(m, 8) = "024"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 48
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a2a
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\024-2.bmp")
               pagecardnum(m, 8) = "024"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 49
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a2a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\025-1.bmp")
               pagecardnum(m, 8) = "025"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 50
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a2a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\025-1.bmp")
               pagecardnum(m, 8) = "025"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 51
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a2a
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\025-1.bmp")
               pagecardnum(m, 8) = "025"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 52
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a2a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\025-2.bmp")
               pagecardnum(m, 8) = "025"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 53
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a2a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\025-2.bmp")
               pagecardnum(m, 8) = "025"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 54
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a2a
               pagecardnum(m, 4) = b1b
               pagecardnum(m, 2) = b1b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\025-2.bmp")
               pagecardnum(m, 8) = "025"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 55
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a4a
               pagecardnum(m, 4) = b3b
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\026-2.bmp")
               pagecardnum(m, 8) = "026"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 56
             'If checkus(m) = 0 Then
               pagecardnum(m, 3) = a4a
               pagecardnum(m, 4) = b3b
               pagecardnum(m, 2) = b2b
               pagecardnum(m, 1) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\026-2.bmp")
               pagecardnum(m, 8) = "026"
               pageonin(m) = 2
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
            Case 57
             'If checkus(m) = 0 Then
               pagecardnum(m, 1) = a4a
               pagecardnum(m, 2) = b3b
               pagecardnum(m, 4) = b2b
               pagecardnum(m, 3) = a3a
               'checkus(m) = 1
               pagecardnum(m, 5) = k
               pagegive = Val(pagegive) + 1
               FormMainMode.card(m).Picture = LoadPicture(app_path & "card\026-1.bmp")
               pagecardnum(m, 8) = "026"
               pageonin(m) = 1
               pagecardnum(m, 6) = 1
                'getpageus(k) = 1
             'End If
End Select
        Select Case k
                      Case 1 '�ϥΪ�
                          pagecardnum(m, 11) = 0
                          FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                          FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
                          �԰��t����.�y�Эp��_�ϥΪ̤�P
                          �P���ʼȮ��ܼ�(3) = m
                          pagecardnum(m, 9) = 240 '���w�ثeLeft(�y��)
                          pagecardnum(m, 10) = 960 '���w�ثeTop(�y��)
                          FormMainMode.card(m).Left = 240
                          FormMainMode.card(m).Top = 960
                          �԰��t����.�p��P���ʶZ�����
                          �԰��t����.���εP�^�_���� (�P���ʼȮ��ܼ�(3))
                          FormMainMode.card(m).Visible = True
                          �԰��t����.�P���ǼW�[_��P_�ϥΪ� m
                          FormMainMode.�P����.Enabled = True
                          FormMainMode.wmpse1.Controls.stop
                          FormMainMode.wmpse1.Controls.play
                          �@��t����.�ˬd���ּ��� 1
                      Case 2 '�q��
                          pagecardnum(m, 11) = 0
                          FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                          FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
                          �԰��t����.�y�Эp��_�q����P
                          �P���ʼȮ��ܼ�(3) = m
                          pagecardnum(m, 9) = 240 '���w�ثeLeft(�y��)
                          pagecardnum(m, 10) = 960 '���w�ثeTop(�y��)
                          FormMainMode.card(m).Left = 240
                          FormMainMode.card(m).Top = 960
                          �԰��t����.�p��P���ʶZ�����
                          �԰��t����.���εP�ܭI��
                          FormMainMode.card(m).Visible = True
                          �԰��t����.�P���ǼW�[_��P_�q�� m
                          FormMainMode.�P����.Enabled = True
                          FormMainMode.wmpse1.Controls.stop
                          FormMainMode.wmpse1.Controls.play
                          �@��t����.�ˬd���ּ��� 1
        End Select
End Sub
Sub moveatkin()
Do
    For j = 71 To 106
      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
         If pagecardnum(j, 1) = a3a And pagecardnum(j, 3) = a3a Then '���ʳ歱�ƥ�d�u��
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            �ثe��(25) = �ثe��(25) + Val(pagecardnum(j, 2))
         End If
         If �ثe��(25) >= 2 Then Exit Do
      End If
    Next
    For j = 1 To 106
      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
         If pagecardnum(j, 1) = a3a Then
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            �ثe��(25) = �ثe��(25) + 1
         ElseIf pagecardnum(j, 3) = a3a Then
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
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            �ثe��(25) = �ثe��(25) + Val(pagecardnum(j, 2))
         End If
         If �ثe��(25) >= 2 Then Exit Do
      End If
    Next
    Exit Do
Loop
'movecheckcom = movecom
End Sub
Sub movetnus()
FormMainMode.messageus.AddItem "�A���D���v�C"
'formmainmode.messageus.AddItem "�{�b���Z��" & movecp & "�C"
�԰��t����.�۰ʱ��b����
FormMainMode.move3.Picture = LoadPicture(app_path & "gif\atk1.gif")
FormMainMode.move4.Picture = LoadPicture(app_path & "gif\def1.gif")
FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\atk2.gif")
FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\def2.gif")
moveturn = 1
'cn2.Visible = True
FormMainMode.cnmove2.Visible = False
�Y���淾�q�Ȯ��ܼ�(1) = 1
End Sub
Sub movetncom()
FormMainMode.messageus.AddItem "��観�D���v�C"
'formmainmode.messageus.AddItem "�{�b���Z��" & movecp & "�C"
�԰��t����.�۰ʱ��b����
FormMainMode.move3.Picture = LoadPicture(app_path & "gif\def1.gif")
FormMainMode.move4.Picture = LoadPicture(app_path & "gif\atk1.gif")
FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\def2.gif")
FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\atk2.gif")
moveturn = 2
'cn3.Visible = True
FormMainMode.cnmove2.Visible = False
�Y���淾�q�Ȯ��ܼ�(1) = 1
End Sub
Sub �H���洫_�ϥΪ�_���w�洫(ByVal num As Integer)
Dim ae As Integer
ae = ����H����ԤH��(1, 2)
����H����ԤH��(1, 2) = ����ݾ��H��������(1, num)
����ݾ��H��������(1, 1) = ����H����ԤH��(1, 2)
����ݾ��H��������(1, num) = ae
FormMainMode.uspiin(����ݾ��H��������(1, num)).Left = 2520 * (num - 1)
FormMainMode.uspiin(����ݾ��H��������(1, num)).Visible = True
FormMainMode.cardus(����ݾ��H��������(1, num)).Visible = False

FormMainMode.uspiin(����H����ԤH��(1, 2)).Left = 0
FormMainMode.uspiin(����H����ԤH��(1, 2)).Visible = False
FormMainMode.cardus(����H����ԤH��(1, 2)).Left = 0
FormMainMode.cardus(����H����ԤH��(1, 2)).Top = 6240
FormMainMode.cardus(����H����ԤH��(1, 2)).ZOrder
FormMainMode.cardus(����H����ԤH��(1, 2)).Visible = True
For n = 1 To 4
    If VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 1) = "" Then
       FormMainMode.personatk(n).Caption = ""
       FormMainMode.personatk(n).Visible = False
    Else
       FormMainMode.personatk(n).Caption = VBEPerson(1, ����H����ԤH��(1, 2), 3, n, 1)
       If VBEPerson(1, ����H����ԤH��(1, 2), 2, 3, 5) = 1 Then
           FormMainMode.personatk(n).FontSize = 12
       Else
           FormMainMode.personatk(n).FontSize = VBEPerson(1, ����H����ԤH��(1, 2), 2, 3, n)
       End If
       FormMainMode.personatk(n).Visible = True
    End If
Next
FormMainMode.personusminijpg.Visible = False
FormMainMode.personusminijpg.�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 1)
FormMainMode.personusminijpg.�p�H���v�l�Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 2)
FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 4)
FormMainMode.personusminijpg.�p�H���v�lLeft = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 5))
FormMainMode.personusminijpg.�p�H���v�ltop�t = Val(VBEPerson(1, ����H����ԤH��(1, 2), 2, 1, 6))
FormDice.jpgus.�j�H���Ϥ� = VBEPerson(1, ����H����ԤH��(1, 2), 1, 5, 3)
FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�left = -(FormMainMode.��ܦC1.�ϥΪ̤�p�H���Ϥ�width)
FormMainMode.personusminijpg.Visible = True
'--------------------------�p��s�Z�����(HP���)
�Z�����(1, 1, 1) = 5295 \ liveusmax(����H����ԤH��(1, 2))
FormMainMode.bloodlineout1.Width = (�Z�����(1, 1, 1) * liveus(����H����ԤH��(1, 2)))
FormMainMode.bloodnumus1.Caption = liveus(����H����ԤH��(1, 2))
FormMainMode.bloodnumus2.Caption = liveusmax(����H����ԤH��(1, 2))
'========================
����ʧ@_�Z���ܧ� movecp
'========================�H�U�O�ޯ��ˬd�αҰ�
If FormMainMode.uspi1(����H����ԤH��(1, 2)).Caption = "�v��L" Then
    If atking_�v��L_�����Ҧ����A��(2) = 1 Then
       atking_�v��L_�����Ҧ����A��(1) = 4
       �԰��t����.�S��_�v��L_�������A_�ϥΪ� '(���q4)
    End If
End If
If FormMainMode.uspi1(����H����ԤH��(1, 2)).Caption = "������" Then
    If atking_������_�����Ҧ����A��(2) = 1 Then
       atking_������_�����Ҧ����A��(1) = 4
       �԰��t����.�S��_������_�������A_�ϥΪ� '(���q4)
    End If
End If
'=============================
For i = 1 To 4
    �԰��t����.�H���ޯ���O�}�� False, i
Next
'=============================
If FormMainMode.uspi1(����H����ԤH��(1, 2)).Caption = "�ײ��d" And atking_�ײ��d_�W���ثe���q������(3) > 0 Then
    atkingck(49, 2) = 1
    atkingck(49, 1) = 7
    �ޯ�.�ײ��d_�W��  '(���q7)
End If
'==========
End Sub

Sub �H���洫_�q��_���w�洫(ByVal num As Integer)
Dim ae As Integer
ae = ����H����ԤH��(2, 2)
����H����ԤH��(2, 2) = ����ݾ��H��������(2, num)
����ݾ��H��������(2, num) = ae
����ݾ��H��������(2, 1) = ����H����ԤH��(2, 2)
FormMainMode.compiin(����ݾ��H��������(2, num)).Left = 2520 * (num - 1)
FormMainMode.compiin(����H����ԤH��(2, 2)).Left = 0
For n = 1 To 4
    If VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 1) = "" Then
       FormMainMode.comaiatk(n).Caption = ""
       FormMainMode.comaiatk(n).Visible = False
    Else
       FormMainMode.comaiatk(n).Caption = VBEPerson(2, ����H����ԤH��(2, 2), 3, n, 1)
       If VBEPerson(2, ����H����ԤH��(2, 2), 2, 3, 5) = 1 Then
           FormMainMode.comaiatk(n).FontSize = 12
       Else
           FormMainMode.comaiatk(n).FontSize = VBEPerson(2, ����H����ԤH��(2, 2), 2, 3, n)
       End If
       FormMainMode.comaiatk(n).Visible = True
    End If
Next
FormMainMode.personcomminijpg.Visible = False
'====================
FormMainMode.personcomminijpg.�p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 1)
FormMainMode.personcomminijpg.�p�H���v�l�Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 2)
FormMainMode.��ܦC1.�q����p�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 4)
FormMainMode.personcomminijpg.�p�H���v�lLeft = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 5)
FormMainMode.personcomminijpg.�p�H���v�ltop�t = VBEPerson(2, ����H����ԤH��(2, 2), 2, 1, 6)
FormMainMode.cardcom(����H����ԤH��(2, 2)).Picture = LoadPicture(VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 5))
FormDice.jpgcom.�j�H���Ϥ� = VBEPerson(2, ����H����ԤH��(2, 2), 1, 5, 3)
FormMainMode.��ܦC1.�q����p�H���Ϥ�left = FormMainMode.ScaleWidth
FormMainMode.personcomminijpg.Left = personminixy(2, ����H����ԤH��(2, 2), movecp, 1)
FormMainMode.personcomminijpg.Top = personminixy(2, ����H����ԤH��(2, 2), movecp, 2)
FormMainMode.personcomminijpg.Visible = True
FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = livecom(����H����ԤH��(2, 2))
FormMainMode.cardcompi2(����H����ԤH��(2, 2)).Caption = atkcom(����H����ԤH��(2, 2))
FormMainMode.cardcompi3(����H����ԤH��(2, 2)).Caption = defcom(����H����ԤH��(2, 2))
FormMainMode.compi1(����H����ԤH��(2, 2)).Caption = namecom(����H����ԤH��(2, 2))
FormMainMode.compi2(����H����ԤH��(2, 2)).Caption = comlevel(����H����ԤH��(2, 2))
FormMainMode.compiatk(����H����ԤH��(2, 2)).Caption = atkcom(����H����ԤH��(2, 2))
FormMainMode.compidef(����H����ԤH��(2, 2)).Caption = defcom(����H����ԤH��(2, 2))
FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = livecom(����H����ԤH��(2, 2))
FormMainMode.compi5(����H����ԤH��(2, 2)).Caption = livecommax(����H����ԤH��(2, 2))
'--------------------------�p��s�Z�����(HP���)
�Z�����(1, 2, 1) = (11340 - 6060) \ livecommax(����H����ԤH��(2, 2))
FormMainMode.bloodlineout2.Left = 11340 - (�Z�����(1, 2, 1) * livecom(����H����ԤH��(2, 2)))
FormMainMode.bloodnumcom1.Caption = livecom(����H����ԤH��(2, 2))
FormMainMode.bloodnumcom2.Caption = livecommax(����H����ԤH��(2, 2))
'==============================
����ʧ@_�Z���ܧ� movecp
'=============================
If FormMainMode.compi1(����H����ԤH��(2, 2)).Caption = "�ײ��d" And atking_AI_�ײ��d_�W���ثe���q������(3) > 0 Then
    atkingckai(139, 2) = 1
    atkingckai(139, 1) = 7
    AI�ޯ�.�ײ��d_�W��  '(���q7)
End If
'==========
End Sub
Sub ����ʧ@_�洫�H������_�ϥΪ�_��l()
Dim i As Integer
Dim ne As Integer
For i = 2 To 3
   Formchangeperson.card(i - 1).Picture = FormMainMode.cardus(����ݾ��H��������(1, i)).Picture
   Formchangeperson.cardhp(i - 1).Caption = FormMainMode.usbi1(����ݾ��H��������(1, i)).Caption
   Formchangeperson.cardatk(i - 1).Caption = FormMainMode.usbi2(����ݾ��H��������(1, i)).Caption
   Formchangeperson.carddef(i - 1).Caption = FormMainMode.usbi3(����ݾ��H��������(1, i)).Caption
Next
ne = 1
For k = 2 To 3
    For j = 14 * (����ݾ��H��������(1, k) - 1) + 1 To 14 * ����ݾ��H��������(1, k)
'        For i = 14 * (k - 2) + 1 To 14 * (k - 1)
            If �H�����`���A��Ʈw(1, j, 2) > 0 Then
                Formchangeperson.personusspe(ne).person_turn = FormMainMode.personusspe(j).person_turn
                Formchangeperson.personusspe(ne).person_num = FormMainMode.personusspe(j).person_num
                Formchangeperson.personusspe(ne).���`���A�Ϥ� = FormMainMode.personusspe(j).���`���A�Ϥ�
                Formchangeperson.personusspe(ne).Visible = True
            Else
                Formchangeperson.personusspe(ne).Visible = False
            End If
            ne = ne + 1
'        Next
    Next
Next
�洫��������Ȯ��ܼ�(1) = 0
For k = 1 To 2
     Formchangeperson.PEAFcardback(k).Visible = False
Next
If Formsetting.chkusenewaipersonauto.Value = 1 Then
    Formchangeperson.�ϥΪ̤贼�z��AI_�۰ʱ����H.Enabled = True
End If
Formchangeperson.Left = FormMainMode.Left + 2430
Formchangeperson.Top = FormMainMode.Top + 1655
Formchangeperson.Show 1
End Sub
Sub ����ʧ@_�洫�H������_�q��_��l()
Select Case �洫��������Ȯ��ܼ�(2)
    Case 1
       �洫��������Ȯ��ܼ�(2) = 0
       �ثe��(22) = 18
       FormMainMode.���ݮɶ�.Enabled = True
    Case 0
       �ثe��(22) = 19
       FormMainMode.���ݮɶ�.Enabled = True
End Select

End Sub
Sub ����ʧ@_�洫�H������_�q��_�洫()
If livecom(����ݾ��H��������(2, 2)) > 0 Then
       �H���洫_�q��_���w�洫 2
ElseIf livecom(����ݾ��H��������(2, 3)) > 0 Then
       �H���洫_�q��_���w�洫 3
End If
����ʧ@_�洫�H������_��������
End Sub
Sub ����ʧ@_�洫�H������_��l()
If (�洫��������Ȯ��ܼ�(1) = 1 Or �洫��������Ȯ��ܼ�(2) = 1) And �洫��������Ȯ��ܼ�(3) = 0 Then
    turnatk = 6
    ���q���A�� = 5
    �԰��t����.�ɶ��b_���]
    FormMainMode.��ܦC1.��ܦC�Ϥ� = App.Path & "\gif\linechange.png"
    FormMainMode.��ܦC1.Visible = True
    FormMainMode.��ܦC1.goi1��� = False
    FormMainMode.��ܦC1.goi2��� = False
    �԰��t����.�ɶ��b_���
    FormMainMode.trtimeline.Enabled = True
    �p�H���Y�����ʤ�V��(1) = 2
    �p�H���Y�����ʤ�V��(2) = 2
    FormMainMode.�p�H���Y������_�ϥΪ�.Enabled = True
    FormMainMode.�p�H���Y������_�q��.Enabled = True
    �洫��������Ȯ��ܼ�(3) = 1
    FormMainMode.��ܦC1.���ʶ��q��ܭ� = 0
    FormMainMode.��ܦC1.���ʶ��q����� = False
End If
If �洫��������Ȯ��ܼ�(1) = 1 Then
    ����ʧ@_�洫�H������_�ϥΪ�_��l
ElseIf �洫��������Ȯ��ܼ�(2) = 1 Then
    ����ʧ@_�洫�H������_�q��_��l
End If
End Sub
Sub ����ʧ@_���ʶ��q��ܰ���()
'===========�洫������
If �洫��������Ȯ��ܼ�(1) = 1 Or �洫��������Ȯ��ܼ�(2) = 1 Then
    ����ʧ@_�洫�H������_��l
Else
    �洫��������Ȯ��ܼ�(3) = 0
    �ثe��(22) = 17
    FormMainMode.���ݮɶ�.Enabled = True
End If
End Sub
Sub ����ʧ@_�H�����`�洫���q��ܰ���()
If �洫��������Ȯ��ܼ�(1) = 1 Or �洫��������Ȯ��ܼ�(2) = 1 Then
    ����ʧ@_�洫�H������_��l
Else
    �洫��������Ȯ��ܼ�(3) = 0
    �ثe��(22) = 20
    FormMainMode.���ݮɶ�.Enabled = True
End If
End Sub
Sub ����ʧ@_�洫�H������_��������()
   Formchangeperson.Hide
   �԰��t����.�ɶ��b_����
   Select Case �洫��������Ȯ��ܼ�(4)
      Case 1
         ����ʧ@_���ʶ��q��ܰ���
      Case 2
         ����ʧ@_�H�����`�洫���q��ܰ���
    End Select
End Sub
Sub �ƥ�d�B�z_���w_�ϥΪ̤�()
Dim kp(1 To 18)  As Integer '�ƥ�d�аO�Ȯɼ�
Dim m, km As Integer
If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Then
    Do
        Randomize
        m = Int(Rnd() * 18) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(1, km, 1) = Formsetting.personus(m).Text
            pageeventnum(1, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personus(m).Text, 2)
        End If
    Loop Until km >= 18
ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then
    Do
        Randomize
        m = Int(Rnd() * 6) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(1, km, 1) = Formsetting.personus(m).Text
            pageeventnum(1, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personus(m).Text, 2)
        End If
    Loop Until km >= 6
    For i = 7 To 12
        pageeventnum(1, i, 1) = Formsetting.personus(i).Text
        pageeventnum(1, i, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).Text, 2)
    Next
End If
End Sub
Sub �ƥ�d�B�z_���w_�q����()
Dim kp(1 To 18)  As Integer '�ƥ�d�аO�Ȯɼ�
Dim m, km As Integer
If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Then
    Do
        Randomize
        m = Int(Rnd() * 18) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(2, km, 1) = Formsetting.personcom(m).Text
            pageeventnum(2, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(m).Text, 2)
        End If
    Loop Until km >= 18
ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then
    Do
        Randomize
        m = Int(Rnd() * 6) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(2, km, 1) = Formsetting.personcom(m).Text
            pageeventnum(2, km, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(m).Text, 2)
        End If
    Loop Until km >= 6
    For i = 7 To 11
        pageeventnum(2, i, 1) = Formsetting.personcom(i).Text
        pageeventnum(2, i, 2) = �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).Text, 2)
    Next
End If
End Sub
Sub �ƥ�d�B�z_��l_�ϥΪ̤�()
Dim ck As Boolean
Dim m As Integer
If Formsetting.persontgruonus(1).Value = True Then '=====(�L)
    For i = 1 To 18
       Randomize
       m = Int(Rnd() * 3) + 1
       Select Case m
          Case 1
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "�C1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
          Case 2
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "�j1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
          Case 3
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "��1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
       End Select
    Next
ElseIf Formsetting.persontgruonus(2).Value = True Then '=====�ۭq
   If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgreus.Value = 0 Then
        For i = 1 To 18
'            If Formsetting.personus(i).Text = "(�L)" Then
            If �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�C1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�j1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "��1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
         For i = 1 To 18
            If Formsetting.personus(i).Text = "(�L)" Or i >= 7 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�C1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�j1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "��1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    End If
ElseIf Formsetting.persontgruonus(3).Value = True Then '===============��̤ܳj��
    If Formsetting.persontgreus.Value = 1 Then  '===��u�W�h
         For i = 1 To 18
             Select Case Formsetting.persontgus(i).Caption
                 Case 0
                      Randomize
                      m = Int(Rnd() * 8) + 1
                      Select Case m
                          Case 1
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C3/�j1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j3/�C1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 3
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 4
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 5
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 6
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 7
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j3/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 8
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�S2" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                     End Select
                 Case 1
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C5/�j3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C5/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�C8" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 2
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j5/�C3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j5/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�j8" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 3
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��5/��1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��7" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "HP�^�_3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 4
                      Randomize
                      m = Int(Rnd() * 2) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��3/�S3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "��5" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 5
                    For j = 0 To Formsetting.personus(i).ListCount - 1
                        If Formsetting.personus(i).List(j) = "���|5" Then
                            Formsetting.personus(i).ListIndex = j
                        End If
                     Next
                 Case 6
                    For j = 0 To Formsetting.personus(i).ListCount - 1
                        If Formsetting.personus(i).List(j) = "�A�G�N5" Then
                            Formsetting.personus(i).ListIndex = j
                        End If
                     Next
                 Case 7
                      Randomize
                      m = Int(Rnd() * 2) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�S3/��3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "�S5" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
             End Select
         Next
         If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�C1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "�j1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "��1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else  '================================����u�W�h
        For i = 1 To 18
            Do
               Randomize
               m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
               '==============================
                    Select Case Formsetting.personus(i).List(m)
                        Case "�C8"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�j8"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��7"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "HP�^�_3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "���|5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�A�G�N5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�S5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�C5/�j3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�j5/�C3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��5/��1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�j5/��1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�C5/��1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "��3/�S3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "�S3/��3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                    End Select
            Loop
        Next
    End If
ElseIf Formsetting.persontgruonus(4).Value = True Then '=====�H��
    If Formsetting.persontgreus.Value = 1 Then '===��u�W�h
        For i = 1 To 18
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
                If �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).List(m), 1) = Formsetting.persontgus(i).Caption Or _
                   �@��t����.�ƥ�d��Ʈw(Formsetting.personus(i).List(m), 1) = 0 Then
                   Formsetting.personus(i).ListIndex = m
                   Exit Do
                End If
             Loop
         Next
        If �ƥ�d�O���Ȯɼ�(0, 1) = 12 Then
            For i = 7 To 18
                   Randomize
                   m = Int(Rnd() * 3) + 1
                   Select Case m
                      Case 1
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "�C1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                      Case 2
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "�j1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                      Case 3
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "��1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                   End Select
            Next
        End If
    Else '=============================����u�W�h
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
            Formsetting.personus(i).ListIndex = m
         Next
    End If
End If
End Sub
Sub �ƥ�d�B�z_��l_�q����()
Dim m As Integer
Dim ay() As String
If Formsetting.persontgruoncom(1).Value = True Then '=====(�L)
    For i = 1 To 18
       Randomize
       m = Int(Rnd() * 3) + 1
       Select Case m
          Case 1
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "�C1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
          Case 2
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "�j1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
          Case 3
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "��1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
       End Select
    Next
ElseIf Formsetting.persontgruoncom(2).Value = True Then '=====�ۭq
   If �ƥ�d�O���Ȯɼ�(0, 1) = 18 Or Formsetting.persontgrecom.Value = 0 Then
        For i = 1 To 18
'            If Formsetting.personcom(i).Text = "(�L)" Then
            If �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    ElseIf �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
         For i = 1 To 18
            If Formsetting.personcom(i).Text = "(�L)" Or i >= 7 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    End If
ElseIf Formsetting.persontgruoncom(3).Value = True Then '=====��̤ܳj��
    If Formsetting.persontgrecom.Value = 1 Then  '===��u�W�h
         For i = 1 To 18
             Select Case Formsetting.persontgcom(i).Caption
                 Case 0
                      Randomize
                      m = Int(Rnd() * 8) + 1
                      Select Case m
                          Case 1
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C3/�j1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j3/�C1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 3
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 4
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 5
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 6
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 7
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j3/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 8
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�S2" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                     End Select
                 Case 1
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C5/�j3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C5/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�C8" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 2
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j5/�C3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j5/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "�j8" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 3
                      Randomize
                      m = Int(Rnd() * 3) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��5/��1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��7" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "HP�^�_3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 4
                      Randomize
                      m = Int(Rnd() * 2) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��3/�S3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "��5" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 5
                    For j = 0 To Formsetting.personcom(i).ListCount - 1
                        If Formsetting.personcom(i).List(j) = "���|5" Then
                            Formsetting.personcom(i).ListIndex = j
                        End If
                     Next
                 Case 6
                    For j = 0 To Formsetting.personcom(i).ListCount - 1
                        If Formsetting.personcom(i).List(j) = "�A�G�N5" Then
                            Formsetting.personcom(i).ListIndex = j
                        End If
                     Next
                 Case 7
                        For j = 0 To Formsetting.personcom(i).ListCount - 1
                           If Formsetting.personcom(i).List(j) = "�S3/��3" Then
                               Formsetting.personcom(i).ListIndex = j
                           End If
                        Next
             End Select
         Next
         If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else  '================================����u�W�h
        For i = 1 To 18
            Do
               Randomize
               m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
               '==============================
                    Select Case Formsetting.personcom(i).List(m)
                        Case "�C8"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�j8"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��7"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "HP�^�_3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "���|5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�A�G�N5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�C5/�j3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�j5/�C3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��5/��1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�j5/��1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�C5/��1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "��3/�S3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "�S3/��3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                    End Select
            Loop
        Next
    End If
ElseIf Formsetting.persontgruoncom(4).Value = True Then '=====�H��
    If Formsetting.persontgrecom.Value = 1 Then '===��u�W�h
        For i = 1 To 18
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                If �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = Formsetting.persontgcom(i).Caption Or _
                   �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = 0 Then
                   Formsetting.personcom(i).ListIndex = m
                   Exit Do
                End If
             Loop
         Next
         If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else '=============================����u�W�h
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
            Formsetting.personcom(i).ListIndex = m
         Next
    End If
ElseIf Formsetting.persontgruoncom(5).Value = True Then '=====�H��(���t�S)
    If Formsetting.persontgrecom.Value = 1 Then '===��u�W�h
        For i = 1 To 18
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                If �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = Formsetting.persontgcom(i).Caption Or _
                   �@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 1) = 0 Then
                   ay = Split(�@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 3), "=")
                   If ay(0) = a4a And ay(2) = a4a Then
                   Else
                        Formsetting.personcom(i).ListIndex = m
                        Exit Do
                   End If
                End If
             Loop
         Next
         If �ƥ�d�O���Ȯɼ�(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�C1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "�j1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "��1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else '=============================����u�W�h
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
            ay = Split(�@��t����.�ƥ�d��Ʈw(Formsetting.personcom(i).List(m), 3), "=")
            If ay(0) = a4a And ay(2) = a4a Then
                 i = i - 1
            Else
                 Formsetting.personcom(i).ListIndex = m
            End If
         Next
    End If
End If
End Sub
Sub �ƥ�d�B�z_����_�ϥΪ̤�()
Dim tn As Integer
Dim ay() As String
tn = Val(FormMainMode.turni.Caption)
If tn <= 18 Then
    If tn <= �ƥ�d�O���Ȯɼ�(0, 1) Or Formsetting.persontgreus.Value = 0 Then
        If pageeventnum(1, tn, 1) <> "" Then
            ay = Split(�@��t����.�ƥ�d��Ʈw(pageeventnum(1, tn, 1), 3), "=")
            pagecardnum(70 + tn, 1) = ay(0)
            pagecardnum(70 + tn, 2) = ay(1)
            pagecardnum(70 + tn, 3) = ay(2)
            pagecardnum(70 + tn, 4) = ay(3)
            pagecardnum(70 + tn, 5) = 1
            pagecardnum(70 + tn, 6) = 1
            pagecardnum(70 + tn, 8) = pageeventnum(1, tn, 2)
            pagecardnum(70 + tn, 11) = 0
            FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
            FormMainMode.card(70 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(1, tn, 2) & "-1.bmp")
            pageonin(70 + tn) = 1
            �԰��t����.�y�Эp��_�ϥΪ̤�P
            �P���ʼȮ��ܼ�(3) = 70 + tn
            �԰��t����.�P���ǼW�[_��P_�ϥΪ� 70 + tn
            pagecardnum(70 + tn, 9) = �P���ʼȮ��ܼ�(1) '���w�ثeLeft(�y��)
            pagecardnum(70 + tn, 10) = �P���ʼȮ��ܼ�(2) '���w�ثeTop(�y��)
            FormMainMode.card(70 + tn).Left = �P���ʼȮ��ܼ�(1)
            FormMainMode.card(70 + tn).Top = �P���ʼȮ��ܼ�(2)
            FormMainMode.card(70 + tn).ZOrder
            FormMainMode.card(70 + tn).Visible = True
        End If
    End If
End If
End Sub
Sub �ƥ�d�B�z_����_�q����()
Dim tn As Integer
Dim ay() As String
tn = Val(FormMainMode.turni.Caption)
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
            pagecardnum(88 + tn, 11) = 0
            FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
            FormMainMode.card(88 + tn).Picture = LoadPicture(app_path & "card\" & pageeventnum(2, tn, 2) & "-1.bmp")
            pageonin(88 + tn) = 1
            �԰��t����.�y�Эp��_�q����P
            �P���ʼȮ��ܼ�(3) = 88 + tn
            �԰��t����.���εP�ܭI��
            �԰��t����.�P���ǼW�[_��P_�q�� 88 + tn
            pagecardnum(88 + tn, 9) = �P���ʼȮ��ܼ�(1) '���w�ثeLeft(�y��)
            pagecardnum(88 + tn, 10) = �P���ʼȮ��ܼ�(2) '���w�ثeTop(�y��)
            FormMainMode.card(88 + tn).Left = �P���ʼȮ��ܼ�(1)
            FormMainMode.card(88 + tn).Top = �P���ʼȮ��ܼ�(2)
            FormMainMode.card(88 + tn).ZOrder
            FormMainMode.card(88 + tn).Visible = True
            For i = 1 To 3
                FormMainMode.compiin(i).ZOrder
            Next
        End If
    End If
End If
End Sub
Sub �ƥ�d�B�z_�p��i��()
If ����H����ԤH��(1, 1) > 1 Or ����H����ԤH��(2, 1) > 1 Then
    �ƥ�d�O���Ȯɼ�(0, 1) = 18
Else
    �ƥ�d�O���Ȯɼ�(0, 1) = 12
End If
End Sub
Function ����ʧ@_�ˬd�O�_�����w���`���A(ByVal uscom As Integer, ByVal num As Integer) As Boolean
����ʧ@_�ˬd�O�_�����w���`���A = False
Select Case uscom
   Case 1
        For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
           If �H�����`���A��Ʈw(1, i, 3) = num Then
               ����ʧ@_�ˬd�O�_�����w���`���A = True
           End If
        Next
   Case 2
        For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
            If �H�����`���A��Ʈw(2, i, 3) = num Then
                ����ʧ@_�ˬd�O�_�����w���`���A = True
            End If
        Next
End Select
End Function
Sub ����ʧ@_���m���q�����ɧޯ�Ұ�()
atkingtrn(1) = 0
atkingtrn(2) = 0
'=================�H�U�O�ޯ��ˬd�αҰ�(�^�X�������q1)
If turnatk = 2 And atkingck(64, 2) = 1 Then
   atkingck(64, 1) = 3
   �ޯ�.����_Jackpot  '(���q3)
End If
If turnatk = 2 And atkingck(146, 2) = 1 Then
   atkingck(146, 1) = 3
   �ޯ�.�Ǧh_�]�G����  '(���q3)
End If
If turnatk = 2 And atkingck(100, 2) = 1 Then
   atkingck(100, 1) = 2
   �ޯ�.�S�{��_�t�v���l  '(���q2)
End If
If turnatk = 2 And atkingck(111, 2) = 1 Then
   atkingck(111, 1) = 3
   �ޯ�.���Y�F_��������  '(���q3)
End If
'=================
�ޯ�ʵe��ܶ��q�� = 9
�԰��t����.�ޯ�Ұʼƶq�ˬd
'===================
If turnatk = 2 And atkingck(64, 2) = 1 Then
   atkingck(64, 1) = 4
   �ޯ�.����_Jackpot  '(���q4)
End If
If turnatk = 2 And atkingck(146, 2) = 1 Then
   atkingck(146, 1) = 4
   �ޯ�.�Ǧh_�]�G����  '(���q4)
End If
If turnatk = 2 And atkingck(100, 2) = 1 Then
   atkingck(100, 1) = 3
   �ޯ�.�S�{��_�t�v���l  '(���q3)
End If
If turnatk = 2 And atkingck(111, 2) = 1 Then
   atkingck(111, 1) = 4
   �ޯ�.���Y�F_��������  '(���q4)
End If
'================
FormMainMode.atkingtrtot.Interval = 600
FormMainMode.atkingtrtot.Enabled = True
End Sub
Sub ����ʧ@_�������q�����ɧޯ�Ұ�()
atkingtrn(1) = 0
atkingtrn(2) = 0
'=================�H�U�O�ޯ��ˬd�αҰ�(�^�X�������q1)
If turnatk = 1 And atkingckai(31, 2) = 1 Then
   atkingckai(31, 1) = 3
   AI�ޯ�.����_Jackpot  '(���q3)
End If
If turnatk = 1 And atkingckai(97, 2) = 1 Then
   atkingckai(97, 1) = 2
   AI�ޯ�.�S�{��_�t�v���l  '(���q2)
End If
If turnatk = 1 And atkingckai(121, 2) = 1 Then
   atkingckai(121, 1) = 3
   AI�ޯ�.�Ǧh_�]�G����  '(���q3)
End If
If turnatk = 1 And atkingckai(123, 2) = 1 Then
   atkingckai(123, 1) = 3
   AI�ޯ�.���Y�F_��������  '(���q3)
End If
'=================
�ޯ�ʵe��ܶ��q�� = 9
�԰��t����.�ޯ�Ұʼƶq�ˬd
'===================
If turnatk = 1 And atkingckai(31, 2) = 1 Then
   atkingckai(31, 1) = 4
   AI�ޯ�.����_Jackpot  '(���q4)
End If
If turnatk = 1 And atkingckai(97, 2) = 1 Then
   atkingckai(97, 1) = 3
   AI�ޯ�.�S�{��_�t�v���l  '(���q3)
End If
If turnatk = 1 And atkingckai(121, 2) = 1 Then
   atkingckai(121, 1) = 4
   AI�ޯ�.�Ǧh_�]�G����  '(���q4)
End If
If turnatk = 1 And atkingckai(123, 2) = 1 Then
   atkingckai(123, 1) = 4
   AI�ޯ�.���Y�F_��������  '(���q4)
End If
'=================
FormMainMode.atkingtrtot.Interval = 600
FormMainMode.atkingtrtot.Enabled = True
End Sub
Sub �ޯ໡�����J_�H���d���I��_�ϥΪ�(ByVal n As Integer)
Dim strw() As String
If ����H����ԤH��(1, 2) = n Then
    For i = 5 To 8
        FormMainMode.PEAFpersoncardback_text(i) = VBEPerson(1, n, 3, i - 4, 1)
        '========
        FormMainMode.PEAFpersoncardback_turn(i).�������O = 3
        FormMainMode.PEAFpersoncardback_turn(i).�Ϥ� = app_path & "gif\�d���I��\CBturn.png"
        FormMainMode.PEAFpersoncardback_turn(i).���ؽs�� = Val(VBEPerson(1, n, 3, i - 4, 8))
        '============================
        Select Case i - 4
            Case 1
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range1(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range1(k + 3).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range1(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range1(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range1(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range1(k + 3).���ؽs�� = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range2(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range2(k + 3).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range2(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range2(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range2(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range2(k + 3).���ؽs�� = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range3(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range3(k + 3).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range3(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range3(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range3(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range3(k + 3).���ؽs�� = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range4(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range4(k + 3).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range4(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range4(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k + 3).�������O = 2
                             FormMainMode.PEAFpersoncardback_range4(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range4(k + 3).���ؽs�� = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, n, 3, i - 4, 10), "&")
        Select Case i - 4
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num1(k + 6).Visible = False
                    Next
              Case 2
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num2(k + 6).Visible = False
                    Next
              Case 3
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num3(k + 6).Visible = False
                    Next
              Case 4
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num4(k + 6).Visible = False
                    Next
        End Select
    Next
    FormMainMode.PEAFpersoncardback_main(2).Caption = ""
'===================================================================
Else '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'===================================================================
    For i = 1 To 4
        FormMainMode.PEAFpersoncardback_text(i) = VBEPerson(1, n, 3, i, 1)
        '========
        FormMainMode.PEAFpersoncardback_turn(i).�������O = 3
        FormMainMode.PEAFpersoncardback_turn(i).�Ϥ� = app_path & "gif\�d���I��\CBturn.png"
        FormMainMode.PEAFpersoncardback_turn(i).���ؽs�� = Val(VBEPerson(1, n, 3, i, 8))
        '============================
        Select Case i
            Case 1
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range1(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range1(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range2(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range2(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range3(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range3(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range4(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k).�������O = 2
                             FormMainMode.PEAFpersoncardback_range4(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, n, 3, i, 10), "&")
        Select Case i
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num1(k + 1).Visible = False
                    Next
              Case 2
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num2(k + 1).Visible = False
                    Next
              Case 3
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num3(k + 1).Visible = False
                    Next
              Case 4
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).�������O = 1
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).Visible = True
                            Else
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            FormMainMode.PEAFpersoncardback_num4(k + 1).Visible = False
                    Next
        End Select
    Next
    FormMainMode.PEAFpersoncardback_main(1).Caption = ""
End If
End Sub
Sub �ޯ໡�����J_�H���d���I��_�q��(ByVal n As Integer)
Dim strw() As String
For i = 1 To 4
    FormMainMode.PEAFpersoncardback_text(i) = VBEPerson(2, n, 3, i, 1)
    '========
    FormMainMode.PEAFpersoncardback_turn(i).�������O = 3
    FormMainMode.PEAFpersoncardback_turn(i).�Ϥ� = app_path & "gif\�d���I��\CBturn.png"
    FormMainMode.PEAFpersoncardback_turn(i).���ؽs�� = Val(VBEPerson(2, n, 3, i, 8))
    '============================
    Select Case i
        Case 1
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range1(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range1(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range1(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range1(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range1(k).���ؽs�� = 2
                    Next
              End If
        Case 2
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range2(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range2(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range2(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range2(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range2(k).���ؽs�� = 2
                    Next
              End If
        Case 3
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range3(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range3(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range3(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range3(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range3(k).���ؽs�� = 2
                    Next
              End If
        Case 4
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range4(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range4(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range4(k).�������O = 2
                         FormMainMode.PEAFpersoncardback_range4(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range4(k).���ؽs�� = 2
                    Next
              End If
    End Select
    '=========================================
    strw = Split(VBEPerson(2, n, 3, i, 10), "&")
    Select Case i
          Case 1
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               FormMainMode.PEAFpersoncardback_num1(k + 1).�������O = 1
                               FormMainMode.PEAFpersoncardback_num1(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num1(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                               FormMainMode.PEAFpersoncardback_num1(k + 1).Visible = True
                        Else
                               FormMainMode.PEAFpersoncardback_num1(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        FormMainMode.PEAFpersoncardback_num1(k + 1).Visible = False
                Next
          Case 2
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               FormMainMode.PEAFpersoncardback_num2(k + 1).�������O = 1
                               FormMainMode.PEAFpersoncardback_num2(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num2(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                               FormMainMode.PEAFpersoncardback_num2(k + 1).Visible = True
                        Else
                               FormMainMode.PEAFpersoncardback_num2(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        FormMainMode.PEAFpersoncardback_num2(k + 1).Visible = False
                Next
          Case 3
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               FormMainMode.PEAFpersoncardback_num3(k + 1).�������O = 1
                               FormMainMode.PEAFpersoncardback_num3(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num3(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                               FormMainMode.PEAFpersoncardback_num3(k + 1).Visible = True
                        Else
                               FormMainMode.PEAFpersoncardback_num3(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        FormMainMode.PEAFpersoncardback_num3(k + 1).Visible = False
                Next
          Case 4
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               FormMainMode.PEAFpersoncardback_num4(k + 1).�������O = 1
                               FormMainMode.PEAFpersoncardback_num4(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num4(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                               FormMainMode.PEAFpersoncardback_num4(k + 1).Visible = True
                        Else
                               FormMainMode.PEAFpersoncardback_num4(k + 1).Visible = False
                        End If
                Next
                For k = UBound(strw) + 1 To 4
                        FormMainMode.PEAFpersoncardback_num4(k + 1).Visible = False
                Next
    End Select
Next
FormMainMode.PEAFpersoncardback_main(1).Caption = ""
End Sub

Sub ����ʧ@_�H���d���I���Ѱ��G��(ByVal n As Integer)
Select Case n
      Case 1
            For k = 1 To 4
                 FormMainMode.PEAFcardbackBR(k).Opacity = 0
            Next
      Case 2
            For k = 1 To 4
                 FormMainMode.PEAFcardbackBR(k + 4).Opacity = 0
            Next
End Select
End Sub
Sub �ޯ໡�����J_�H���d���I��_�洫����(ByVal n As Integer)
Dim strw() As String
If n = 2 Then
    For i = 5 To 8
        Formchangeperson.PEAFpersoncardback_text(i) = VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 1)
        '========
        Formchangeperson.PEAFpersoncardback_turn(i).�������O = 3
        Formchangeperson.PEAFpersoncardback_turn(i).�Ϥ� = app_path & "gif\�d���I��\CBturn.png"
        Formchangeperson.PEAFpersoncardback_turn(i).���ؽs�� = Val(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 8))
        '============================
        Select Case i - 4
            Case 1
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range1(k + 3).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range1(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range1(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).���ؽs�� = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range2(k + 3).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range2(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range2(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).���ؽs�� = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range3(k + 3).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range3(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range3(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).���ؽs�� = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range4(k + 3).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range4(k + 3).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range4(k + 3).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).���ؽs�� = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i - 4, 10), "&")
        Select Case i - 4
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num1(k + 6).Visible = False
                    Next
              Case 2
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num2(k + 6).Visible = False
                    Next
              Case 3
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num3(k + 6).Visible = False
                    Next
              Case 4
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num4(k + 6).Visible = False
                    Next
        End Select
    Next
    Formchangeperson.PEAFpersoncardback_main(2).Caption = ""
    �H���d���I���s��������(7) = 0
'===================================================================
Else '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'===================================================================
    For i = 1 To 4
        Formchangeperson.PEAFpersoncardback_text(i) = VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 1)
        '========
        Formchangeperson.PEAFpersoncardback_turn(i).�������O = 3
        Formchangeperson.PEAFpersoncardback_turn(i).�Ϥ� = app_path & "gif\�d���I��\CBturn.png"
        Formchangeperson.PEAFpersoncardback_turn(i).���ؽs�� = Val(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 8))
        '============================
        Select Case i
            Case 1
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range1(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range1(k).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range1(k).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range1(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range1(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range1(k).���ؽs�� = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range2(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range2(k).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range2(k).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range2(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range2(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range2(k).���ؽs�� = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range3(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range3(k).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range3(k).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range3(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range3(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range3(k).���ؽs�� = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range4(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             If Mid(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range4(k).���ؽs�� = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range4(k).���ؽs�� = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range4(k).���ؽs�� = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k).�������O = 2
                             Formchangeperson.PEAFpersoncardback_range4(k).�Ϥ� = app_path & "gif\�d���I��\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range4(k).���ؽs�� = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, ����ݾ��H��������(1, n + 1), 3, i, 10), "&")
        Select Case i
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num1(k + 1).Visible = False
                    Next
              Case 2
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num2(k + 1).Visible = False
                    Next
              Case 3
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num3(k + 1).Visible = False
                    Next
              Case 4
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).�������O = 1
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).�Ϥ� = app_path & "gif\�d���I��\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).���ؽs�� = Val(Mid(strw(k), 2, 1))
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).Visible = True
                            Else
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).Visible = False
                            End If
                    Next
                    For k = UBound(strw) + 1 To 4
                            Formchangeperson.PEAFpersoncardback_num4(k + 1).Visible = False
                    Next
        End Select
    Next
    Formchangeperson.PEAFpersoncardback_main(1).Caption = ""
    �H���d���I���s��������(6) = 0
End If
End Sub
Sub getpage(ByVal k As Integer, m As Integer)
Dim qwp As Integer, n As Integer, uspce As String, uspme As String, yne As Boolean
If Val(���εP�U�P����������(0, 1)) < Val(���εP�U�P����������(0, 2)) Then
    yne = False
    Do
            Randomize
            qwp = Int(Rnd() * 29) + 1
            Select Case qwp
                    Case 1  '==��1�j1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\021-2.bmp")
                            pagecardnum(m, 8) = "021"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 2  '==��1�j2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\019-2.bmp")
                            pagecardnum(m, 8) = "019"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 3  '==��1�j3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\017-2.bmp")
                            pagecardnum(m, 8) = "017"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 4  '==��1��1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\025-2.bmp")
                            pagecardnum(m, 8) = "025"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 5  '==��1��2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\024-2.bmp")
                            pagecardnum(m, 8) = "024"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 6  '==��1��3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\023-2.bmp")
                            pagecardnum(m, 8) = "023"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 7  '==��2�S3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\026-2.bmp")
                            pagecardnum(m, 8) = "026"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 8  '==��3��3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a3a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a3a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\027-2.bmp")
                            pagecardnum(m, 8) = "027"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 9  '==�C6�C6��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b6b
                            pagecardnum(m, 3) = a1a
                            pagecardnum(m, 4) = b6b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\001-2.bmp")
                            pagecardnum(m, 8) = "001"
                            pageonin(m) = 2
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 10  '==�C1�j1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\011-1.bmp")
                            pagecardnum(m, 8) = "011"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 11  '==�C2�j1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\007-1.bmp")
                            pagecardnum(m, 8) = "007"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 12  '==�C2�j2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\006-1.bmp")
                            pagecardnum(m, 8) = "006"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 13  '==�C3�j3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\004-1.bmp")
                            pagecardnum(m, 8) = "004"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 14  '==�C5�j5��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b5b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\028-1.bmp")
                            pagecardnum(m, 8) = "028"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 15  '==�C1��1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\012-1.bmp")
                            pagecardnum(m, 8) = "012"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 16  '==�C2��1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\009-1.bmp")
                            pagecardnum(m, 8) = "009"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 17  '==�C2��2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\008-1.bmp")
                            pagecardnum(m, 8) = "008"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 18  '==�C3��3��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b3b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\005-1.bmp")
                            pagecardnum(m, 8) = "005"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 19  '==�C1�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b1b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\013-1.bmp")
                            pagecardnum(m, 8) = "013"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 20  '==�C2�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\010-1.bmp")
                            pagecardnum(m, 8) = "010"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 21  '==�C4�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b4b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\003-1.bmp")
                            pagecardnum(m, 8) = "003"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 22  '==�C5�S2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a1a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\002-1.bmp")
                            pagecardnum(m, 8) = "002"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 23  '==�j4�j4��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b4b
                            pagecardnum(m, 3) = a5a
                            pagecardnum(m, 4) = b4b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\015-1.bmp")
                            pagecardnum(m, 8) = "015"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 24  '==�j2�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b2b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\020-1.bmp")
                            pagecardnum(m, 8) = "020"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 25  '==�j3�S2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\018-1.bmp")
                            pagecardnum(m, 8) = "018"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 26  '==�j4�S1��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b4b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b1b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\016-1.bmp")
                            pagecardnum(m, 8) = "016"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 27  '==�j5�S2��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a5a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b2b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\014-1.bmp")
                            pagecardnum(m, 8) = "014"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 28  '==��5��5��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a2a
                            pagecardnum(m, 2) = b5b
                            pagecardnum(m, 3) = a2a
                            pagecardnum(m, 4) = b5b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\022-1.bmp")
                            pagecardnum(m, 8) = "022"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
                    Case 29  '==��3�S5��
                         If Val(���εP�U�P����������(qwp, 1)) < Val(���εP�U�P����������(qwp, 2)) Then
                            ���εP�U�P����������(qwp, 1) = Val(���εP�U�P����������(qwp, 1)) + 1
                            ���εP�U�P����������(0, 1) = Val(���εP�U�P����������(0, 1)) + 1
                            pagecardnum(m, 1) = a2a
                            pagecardnum(m, 2) = b3b
                            pagecardnum(m, 3) = a4a
                            pagecardnum(m, 4) = b5b
                            pagecardnum(m, 5) = k
                            FormMainMode.card(m).Picture = LoadPicture(app_path & "card\029-1.bmp")
                            pagecardnum(m, 8) = "029"
                            pageonin(m) = 1
                            pagecardnum(m, 6) = 1
                            yne = True
                         End If
             End Select
     Loop Until yne = True
     '==================================�H����P
     Randomize
     n = Int(Rnd() * 2) + 1
     If n = 2 Then
        uspce = pagecardnum(m, 1)
        uspme = pagecardnum(m, 2)
        pagecardnum(m, 1) = pagecardnum(m, 3)
        pagecardnum(m, 2) = pagecardnum(m, 4)
        pagecardnum(m, 3) = uspce
        pagecardnum(m, 4) = uspme
        If pageonin(m) = 1 Then
           pageonin(m) = 2
           FormMainMode.card(m).Picture = LoadPicture(app_path & "card\" & pagecardnum(m, 8) & "-" & pageonin(m) & ".bmp")
        Else
           pageonin(m) = 1
           FormMainMode.card(m).Picture = LoadPicture(app_path & "card\" & pagecardnum(m, 8) & "-" & pageonin(m) & ".bmp")
        End If
     End If
     '==============================================
     Select Case k
            Case 1 '�ϥΪ�
                pagecardnum(m, 11) = 0
                FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
                �԰��t����.�y�Эp��_�ϥΪ̤�P
                �P���ʼȮ��ܼ�(3) = m
                pagecardnum(m, 9) = 240 '���w�ثeLeft(�y��)
                pagecardnum(m, 10) = 960 '���w�ثeTop(�y��)
                FormMainMode.card(m).Left = 240
                FormMainMode.card(m).Top = 960
                �԰��t����.�p��P���ʶZ�����
                �԰��t����.���εP�^�_���� (�P���ʼȮ��ܼ�(3))
                FormMainMode.card(m).Visible = True
                �԰��t����.�P���ǼW�[_��P_�ϥΪ� m
                FormMainMode.�P����.Enabled = True
                FormMainMode.wmpse1.Controls.stop
                FormMainMode.wmpse1.Controls.play
                �@��t����.�ˬd���ּ��� 1
            Case 2 '�q��
                pagecardnum(m, 11) = 0
                FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
                �԰��t����.�y�Эp��_�q����P
                �P���ʼȮ��ܼ�(3) = m
                pagecardnum(m, 9) = 240 '���w�ثeLeft(�y��)
                pagecardnum(m, 10) = 960 '���w�ثeTop(�y��)
                FormMainMode.card(m).Left = 240
                FormMainMode.card(m).Top = 960
                �԰��t����.�p��P���ʶZ�����
                �԰��t����.���εP�ܭI��
                FormMainMode.card(m).Visible = True
                �԰��t����.�P���ǼW�[_��P_�q�� m
                FormMainMode.�P����.Enabled = True
                FormMainMode.wmpse1.Controls.stop
                FormMainMode.wmpse1.Controls.play
                �@��t����.�ˬd���ּ��� 1
        End Select
End If
End Sub
Sub ���εP�a�ϵP�����t�m(ByVal name As String)
Select Case name
     Case "�ܤB���|�櫰��"
           ���εP�U�P����������(0, 2) = 57
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "���b�˪L"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 0
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 0
           ���εP�U�P����������(20, 2) = 0
           ���εP�U�P����������(21, 2) = 1
           ���εP�U�P����������(22, 2) = 1
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "�U������"
           ���εP�U�P����������(0, 2) = 55
           ���εP�U�P����������(1, 2) = 2
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "�B�ʴ�`(�s)"
           ���εP�U�P����������(0, 2) = 53
           ���εP�U�P����������(1, 2) = 4
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 2
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�H��Ӧa"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 4
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 0
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "���Y����"
           ���εP�U�P����������(0, 2) = 54
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 0
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 0
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
    Case "���ɯ"
           ���εP�U�P����������(0, 2) = 52
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 2
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 0
           ���εP�U�P����������(20, 2) = 0
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 0
           ���εP�U�P����������(29, 2) = 1
    Case "ÿ�e�઺���"
           ���εP�U�P����������(0, 2) = 49
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 1
           ���εP�U�P����������(3, 2) = 1
           ���εP�U�P����������(4, 2) = 3
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 1
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 2
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 1
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�]��ù�e�����J"
           ���εP�U�P����������(0, 2) = 42
           ���εP�U�P����������(1, 2) = 0
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 2
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 0
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 0
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�ƨg�s��"
           ���εP�U�P����������(0, 2) = 47
           ���εP�U�P����������(1, 2) = 2
           ���εP�U�P����������(2, 2) = 0
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 2
           ���εP�U�P����������(5, 2) = 0
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�]�k�s��"
           ���εP�U�P����������(0, 2) = 52
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 3
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 3
           ���εP�U�P����������(11, 2) = 1
           ���εP�U�P����������(12, 2) = 1
           ���εP�U�P����������(13, 2) = 0
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "�Q�i�����´�"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 1
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 1
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 2
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 1
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 0
           ���εP�U�P����������(21, 2) = 1
           ���εP�U�P����������(22, 2) = 1
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 0
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 1
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 1
    Case "���]�������۰}"
           ���εP�U�P����������(0, 2) = 50
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 2
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 0
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 1
           ���εP�U�P����������(22, 2) = 1
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 1
           ���εP�U�P����������(27, 2) = 1
           ���εP�U�P����������(28, 2) = 0
           ���εP�U�P����������(29, 2) = 0
    Case Else
           ���εP�U�P����������(0, 2) = 57
           ���εP�U�P����������(1, 2) = 6
           ���εP�U�P����������(2, 2) = 2
           ���εP�U�P����������(3, 2) = 2
           ���εP�U�P����������(4, 2) = 6
           ���εP�U�P����������(5, 2) = 2
           ���εP�U�P����������(6, 2) = 1
           ���εP�U�P����������(7, 2) = 3
           ���εP�U�P����������(8, 2) = 0
           ���εP�U�P����������(9, 2) = 1
           ���εP�U�P����������(10, 2) = 4
           ���εP�U�P����������(11, 2) = 2
           ���εP�U�P����������(12, 2) = 2
           ���εP�U�P����������(13, 2) = 2
           ���εP�U�P����������(14, 2) = 0
           ���εP�U�P����������(15, 2) = 4
           ���εP�U�P����������(16, 2) = 2
           ���εP�U�P����������(17, 2) = 2
           ���εP�U�P����������(18, 2) = 2
           ���εP�U�P����������(19, 2) = 1
           ���εP�U�P����������(20, 2) = 1
           ���εP�U�P����������(21, 2) = 2
           ���εP�U�P����������(22, 2) = 2
           ���εP�U�P����������(23, 2) = 1
           ���εP�U�P����������(24, 2) = 1
           ���εP�U�P����������(25, 2) = 1
           ���εP�U�P����������(26, 2) = 2
           ���εP�U�P����������(27, 2) = 2
           ���εP�U�P����������(28, 2) = 1
           ���εP�U�P����������(29, 2) = 0
End Select
End Sub
Sub ���εP���ϥ��ˬd()
For i = Val(���εP�U�P����������(0, 2)) + 1 To 70
     pagecardnum(i, 6) = 5
Next
End Sub
Sub �ˮ`����_�ߧY���`_�ϥΪ�(ByVal num As Integer)
Select Case num
   Case 1
        FormMainMode.messageus.AddItem "�z����F" & liveus(����H����ԤH��(1, 2)) & "�I�ˮ`�C"
        �԰��t����.�۰ʱ��b����
        FormMainMode.usbi1(����H����ԤH��(1, 2)).Caption = 0
        FormMainMode.uspi4(����H����ԤH��(1, 2)).Caption = 0
        liveus(����H����ԤH��(1, 2)) = 0
        FormMainMode.bloodnumus1.Caption = 0
        FormMainMode.bloodlineout1.Width = 0
        �P�`���q��(1) = �P�`���q��(1) + 1
        �԰��t����.����ˮ`����
   Case Is > 1
        liveus(����ݾ��H��������(1, num)) = 0
        If FormMainMode.uspi1(����ݾ��H��������(1, num)).Caption = "" Then
            FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption = -liveusmax(����ݾ��H��������(1, num))
            FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption = -liveusmax(����ݾ��H��������(1, num))
        Else
            FormMainMode.usbi1(����ݾ��H��������(1, num)).Caption = 0
            FormMainMode.uspi4(����ݾ��H��������(1, num)).Caption = 0
        End If
        �P�`���q��(1) = �P�`���q��(1) + 1
End Select
End Sub
Sub �ˮ`����_�ߧY���`_�q��(ByVal num As Integer)
Select Case num
    Case 1
        FormMainMode.messageus.AddItem "������F" & livecom(����H����ԤH��(2, 2)) & "�I�ˮ`�C"
        �԰��t����.�۰ʱ��b����
        FormMainMode.compi4(����H����ԤH��(2, 2)).Caption = 0
        FormMainMode.cardcompi1(����H����ԤH��(2, 2)).Caption = 0
        FormMainMode.bloodnumcom1.Caption = 0
        livecom(����H����ԤH��(2, 2)) = 0
        FormMainMode.bloodlineout2.Left = 11580
        �P�`���q��(2) = �P�`���q��(2) + 1
        �԰��t����.����ˮ`����
    Case Is > 1
        If FormMainMode.compi1(����ݾ��H��������(2, num)).Caption = "" Then
            FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = -livecommax(����ݾ��H��������(2, num))
            FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = -livecommax(����ݾ��H��������(2, num))
        Else
            FormMainMode.compi4(����ݾ��H��������(2, num)).Caption = 0
            FormMainMode.cardcompi1(����ݾ��H��������(2, num)).Caption = 0
        End If
        livecom(����ݾ��H��������(2, num)) = 0
        �P�`���q��(2) = �P�`���q��(2) + 1
End Select
End Sub

