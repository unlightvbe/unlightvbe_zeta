Attribute VB_Name = "���z��AI�t����"
Option Explicit
Public cardcountAInum() As String  '���εP�p��Ȯɰ򥻸��(��x�i,1.��������/2.�����ƭ�/3.�ϭ�����/4.�ϭ��ƭ�/5.�P�s��)
Public cardcountAInumMOV() As String  '���εP�p��Ȯɰ򥻸��-���ʶ��q��-�쥻(��x�i,1.��������/2.�����ƭ�/3.�ϭ�����/4.�ϭ��ƭ�/5.�P�s��)
Dim cardAIn() As Integer '�ƦC�զX�p��Ȯ��ܼ�
Dim cardAInumans As String '�ƦC�զX�p��Ȯ��ܼ�
Public cardAInumnm() As String '�ƦC�զX�p��̲׼ƭ�
Public cardAInumFinal() As Integer '�ƦC�զX�p��̲״����
Public cardAInumFinal2() As Integer '�ƦC�զX�p��̲״����-�ƦC��
Public cardAInumcase(1 To 5, 1 To 2) As Integer '���εP�p��έp���(1.ATK-�C/2.DEF/3.MOV/4.SPE/5.ATK-�j,1.�զX�U�̧C�ƭ�/2.�զX�U�̰��ƭ�)
Public cardAInumcaseperson() As Integer '���εP�p��έp�Ȯɸ��-�ӧO�զX
Public cardAInumuscom As Integer '��P�֦��̵P�ưO���Ȯ��ܼ�
Public cardAInumcasepersonTER() As Integer '���εP�p��έp�Ȯɸ��-�ӧO�զX-�ӧO�d���ƭȭp�Ʋέp
Public cardAInumselect1 As Integer  '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰������
Public cardAInumselect4 As Integer  '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰��ӧO�[�`�����
Public cardAInumselect2 As String '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰�����ȤU�s����-��l
Public cardAInumselect3() As String '���εP�p��έp��ǼȮ��ܼ�-�ثe�̰�����ȤU�s����-�}�C
Public cardAInumchoose As Integer '���εP�p��̲׿�ܲզX�s��
Public cardAInumMOVmain(1 To 2, 1 To 15) As String 'AI-���ʶ��q��-�զX�Ȯɬ���
Public cardAInumMOVnm() As String 'AI-���ʶ��q��-���V��-�p��ƦC�զX��Ȯɬ���
Public cardAInumMOVnmtot() As String 'AI-���ʶ��q��-���V��-�`�@�ƦC�զX�������ƼȮɬ���
Public cardAInumMOVFinal(1 To 3) As String 'AI-���ʶ��q��-���V��-�̲׵��G������(1.�̲ױƦC�զX��/2.�̲ױƦC�զX�s��/3.�̲׿�w�ؼжZ��[1.��/2.��])
Public �O�_���ʶ��q����p�P�_�{�� As Boolean 'AI-���ʶ��q��-�O�_�����p�P�_�{�ǼаO��
Public cardAInumOvertenrecord(1 To 10) As Integer 'AI�޾ɵ{��-�W�X�P�i��-�P�����Ȯ��ܼ�(1~10.�P�s��)
Public personatkingtfr(1 To 5) As Integer '�p��ӧO�ޯ�-�O�_��Ex��(1~4.(1)��/(2)�L,5.�O�_���ʦL)
Sub ���z��AI�t�έp��_�@���q_��l(ByVal pagenumber As Integer)
Erase cardcountAInum
Erase cardAInumnm
Erase cardAInumcase
Erase cardAInumselect3
cardAInumans = ""
cardAInumselect1 = 0
cardAInumselect4 = 0
cardAInumselect2 = ""
cardAInumchoose = 0
cardAInumuscom = pagenumber
ReDim cardcountAInum(1 To cardAInumuscom, 1 To 5) As String
ReDim cardAInumcaseperson(1 To 2 ^ cardAInumuscom, 1 To 2, 1 To 15) As Integer
ReDim cardAInumcasepersonTER(1 To 2 ^ cardAInumuscom, 1 To 5, 1 To 10) As Integer
ReDim cardAInumFinal(1 To 2 ^ cardAInumuscom, 1 To 4) As Integer
ReDim cardAInumFinal2(1 To 2 ^ cardAInumuscom, 1 To 4) As Integer
'=========�p�⥿�ϭ��ƦC�զX�ƭ�
���z��AI�t����.�ƦC�զX�p�� pagenumber
End Sub
Sub ���z��AI�t�έp��_�@���q_���o�P�����(ByVal �O�_�@�� As Boolean, ByVal uscom As Integer)
Dim i As Integer
If �O�_�@�� = True Then
        '=========�^���ثe�P�����
        Select Case uscom
            Case 1
                �԰��t����.�X�P���ǭp��_�ϥΪ�_��P
            Case 2
                �԰��t����.�X�P���ǭp��_�q��_��P
        End Select
        Dim w As Integer '�Ȯ��ܼ�
        w = 2 * uscom '(2-�ϥΪ̤�P/4-�q����P)
        For i = 1 To pageglead(uscom)
            cardcountAInum(i, 5) = �X�P���ǲέp�Ȯ��ܼ�(w, i, 2)
            cardcountAInum(i, 1) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 1)
            cardcountAInum(i, 2) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 2)
            cardcountAInum(i, 3) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 3)
            cardcountAInum(i, 4) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 4)
        Next
End If
'======================
���z��AI�t����.�ƦC�զX�έp�ƭȭp��_��P�`�p
���z��AI�t����.�ƦC�զX�έp�ƭȭp��_�ӧO�զX
End Sub
Sub ���z��AI�t�έp��_�G���q_�p������_��l(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
Dim wnum As Integer, whnum As Integer, i As Integer, j As Integer '�Ȯ��ܼ�
Select Case turn
    Case 1 '===�������q
         If uscom = 1 Then whnum = atkus(����H����ԤH��(1, 2)) Else whnum = atkcom(����H����ԤH��(2, 2))
         '==========================
         For i = 0 To (2 ^ cardAInumuscom) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If (cardcountAInum(j, 1) = a1a And movecpre = 1) Or (cardcountAInum(j, 1) = a5a And movecpre > 1) Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                          End If
                      Case 1
                          If (cardcountAInum(j, 3) = a1a And movecpre = 1) Or (cardcountAInum(j, 3) = a5a And movecpre > 1) Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
             If Val(wnum) > 0 Then
                 cardAInumFinal(i + 1, 1) = Val(cardAInumFinal(i + 1, 1)) + Val(whnum)
             End If
         Next
    Case 2  '===���m���q
         For i = 0 To (2 ^ cardAInumuscom) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If cardcountAInum(j, 1) = a2a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                          End If
                      Case 1
                          If cardcountAInum(j, 3) = a2a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
         Next
    Case 3  '===���ʶ��q
         For i = 0 To (2 ^ cardAInumuscom) - 1
             wnum = 0
             For j = 1 To cardAInumuscom
                 Select Case Mid(cardAInumnm(i), j, 1)
                      Case 0
                          If cardcountAInum(j, 1) = a3a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 2))
                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 2))
                          End If
                      Case 1
                          If cardcountAInum(j, 3) = a3a Then
                              wnum = Val(wnum) + Val(cardcountAInum(j, 4))
                              cardAInumcaseperson(i + 1, 2, j) = Val(cardcountAInum(j, 4))
                          End If
                 End Select
             Next
             cardAInumFinal(i + 1, 1) = Val(wnum)
             cardAInumFinal(i + 1, 2) = i + 1
         Next
End Select
End Sub
Sub ���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ�(ByVal name As String, ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
���z��AI�t����.�ˬd�H���ޯ�O�_��EX�� uscom, name
If personatkingtfr(5) = 1 Then
   Exit Sub '���ʦL���A�ɵL�k�o�ʧޯ�
End If
Select Case name
     Case "��B�����S"
           ���z��AI�H����.��B�����S turn, movecpre, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "����"
           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "���"
           ���z��AI�H����.��� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�j�|�˺��h"
           ���z��AI�H����.�j�|�˺��h turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "���["
           ���z��AI�H����.���[ turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�v��L"
           ���z��AI�H����.�v��L turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "C.C."
           ���z��AI�H����.CC turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "��ܵY"
           ���z��AI�H����.��ܵY turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "����"
           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "����"
           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "������"
           ���z��AI�H����.������ turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "��̬d�w"
           ���z��AI�H����.��̬d�w turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "������"
           ���z��AI�H����.������ turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�Q��"
           ���z��AI�H����.�Q�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�L���S"
           ���z��AI�H����.�L���S turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "���纸"
           ���z��AI�H����.���纸 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "������S"
           ���z��AI�H����.������S turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�w�ǥ���"
           ���z��AI�H����.�w�ǥ��� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "����P��"
           ���z��AI�H����.����P�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�h�g�H"
           ���z��AI�H����.�h�g�H turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�Ǧh"
           ���z��AI�H����.�Ǧh turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "���_�i���h"
           ���z��AI�H����.���_�i���h turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�S�{��"
           ���z��AI�H����.�S�{�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "����"
           ���z��AI�H����.���� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "���Y�F"
           ���z��AI�H����.���Y�F turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "��"
           ���z��AI�H����.�� turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "ù��Y"
           ���z��AI�H����.ù��Y turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�����g"
           ���z��AI�H����.�����g turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�J�y"
           ���z��AI�H����.�J�y turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�����i"
           ���z��AI�H����.�����i turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "�ײ��d"
           ���z��AI�H����.�ײ��d turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
End Select

End Sub
Sub �ƦC�զX�p��(ByVal qnum As Integer)
Dim i As Integer
'===========
ReDim cardAIn(1 To Val(qnum))
Erase cardAInumnm
cardAInumans = ""
Dim s As Integer
For i = 1 To qnum   '���]�϶��ƭ�
    cardAIn(i) = 0
Next
s = 1
'================
Do
    For i = qnum To 1 Step -1
        cardAInumans = cardAInumans & cardAIn(i)
    Next
    '================
    cardAIn(1) = cardAIn(1) + 1
    ���z��AI�t����.�ƦC�զX�p��_�϶��i�� qnum '�@[qnum]���
    '================
    s = s + 1
    cardAInumans = cardAInumans & "="
Loop Until s > (2 ^ qnum)
cardAInumnm = Split(cardAInumans, "=")
'Dim h As Integer
'h = 1
'For i = 0 To (2 ^ qnum) - 1
'    Print nm(i)
'    h = h + 1
'    If h > 50 Then
'        Cls
'        h = 1
'    End If
'Next

End Sub
Sub �ƦC�զX�p��_�϶��i��(ByVal num As Integer)
Dim i As Integer
For i = 1 To num - 1
    If cardAIn(i) = 2 Then
        cardAIn(i + 1) = cardAIn(i + 1) + 1
        cardAIn(i) = 0
    End If
Next

End Sub
Sub �ƦC�զX�έp�ƭȭp��_��P�`�p()
Dim we As Integer, i As Integer, j As Integer '�Ȯ��ܼ�
For i = 1 To cardAInumuscom
    For j = 1 To 2
        we = 2 * j
        Select Case cardcountAInum(i, j)
             Case a1a
                  If cardcountAInum(i, we) < cardAInumcase(1, 1) Or cardAInumcase(1, 1) = 0 Then
                      cardAInumcase(1, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(1, 2) Or cardAInumcase(1, 2) = 0 Then
                      cardAInumcase(1, 2) = cardcountAInum(i, we)
                  End If
             Case a2a
                  If cardcountAInum(i, we) < cardAInumcase(2, 1) Or cardAInumcase(2, 1) = 0 Then
                      cardAInumcase(2, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(2, 2) Or cardAInumcase(2, 2) = 0 Then
                      cardAInumcase(2, 2) = cardcountAInum(i, we)
                  End If
             Case a3a
                  If cardcountAInum(i, we) < cardAInumcase(3, 1) Or cardAInumcase(3, 1) = 0 Then
                      cardAInumcase(3, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(3, 2) Or cardAInumcase(3, 2) = 0 Then
                      cardAInumcase(3, 2) = cardcountAInum(i, we)
                  End If
             Case a4a
                  If cardcountAInum(i, we) < cardAInumcase(4, 1) Or cardAInumcase(4, 1) = 0 Then
                      cardAInumcase(4, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(4, 2) Or cardAInumcase(4, 2) = 0 Then
                      cardAInumcase(4, 2) = cardcountAInum(i, we)
                  End If
             Case a5a
                  If cardcountAInum(i, we) < cardAInumcase(5, 1) Or cardAInumcase(5, 1) = 0 Then
                      cardAInumcase(5, 1) = cardcountAInum(i, we)
                  End If
                  If cardcountAInum(i, we) > cardAInumcase(5, 2) Or cardAInumcase(5, 2) = 0 Then
                      cardAInumcase(5, 2) = cardcountAInum(i, we)
                  End If
        End Select
    Next
Next
End Sub
Sub �ƦC�զX�έp�ƭȭp��_�ӧO�զX()
Dim we As Integer, i As Integer, j As Integer '�Ȯ��ܼ�
For i = 1 To 2 ^ cardAInumuscom
    For j = 1 To cardAInumuscom
        Select Case Mid(cardAInumnm(i - 1), j, 1)
            Case 0
                 we = 2
                  Select Case cardcountAInum(j, 1)
                     Case a1a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 1) Or cardAInumcaseperson(i, 1, 1) = 0 Then
                              cardAInumcaseperson(i, 1, 1) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 2) Or cardAInumcaseperson(i, 1, 2) = 0 Then
                              cardAInumcaseperson(i, 1, 2) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 11) = cardAInumcaseperson(i, 1, 11) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 1, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 1, cardcountAInum(j, we))) + 1
                     Case a2a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 3) Or cardAInumcaseperson(i, 1, 3) = 0 Then
                              cardAInumcaseperson(i, 1, 3) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 4) Or cardAInumcaseperson(i, 1, 4) = 0 Then
                              cardAInumcaseperson(i, 1, 4) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 12) = cardAInumcaseperson(i, 1, 12) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 2, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 2, cardcountAInum(j, we))) + 1
                     Case a3a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 5) Or cardAInumcaseperson(i, 1, 5) = 0 Then
                              cardAInumcaseperson(i, 1, 5) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 6) Or cardAInumcaseperson(i, 1, 6) = 0 Then
                              cardAInumcaseperson(i, 1, 6) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 13) = cardAInumcaseperson(i, 1, 13) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 3, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 3, cardcountAInum(j, we))) + 1
                     Case a4a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 7) Or cardAInumcaseperson(i, 1, 7) = 0 Then
                              cardAInumcaseperson(i, 1, 7) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 8) Or cardAInumcaseperson(i, 1, 8) = 0 Then
                              cardAInumcaseperson(i, 1, 8) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 14) = cardAInumcaseperson(i, 1, 14) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 4, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 4, cardcountAInum(j, we))) + 1
                     Case a5a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 9) Or cardAInumcaseperson(i, 1, 9) = 0 Then
                              cardAInumcaseperson(i, 1, 9) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 10) Or cardAInumcaseperson(i, 1, 10) = 0 Then
                              cardAInumcaseperson(i, 1, 10) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 15) = cardAInumcaseperson(i, 1, 15) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 5, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 5, cardcountAInum(j, we))) + 1
                End Select
            Case 1
                 we = 4
                  Select Case cardcountAInum(j, 3)
                     Case a1a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 1) Or cardAInumcaseperson(i, 1, 1) = 0 Then
                              cardAInumcaseperson(i, 1, 1) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 2) Or cardAInumcaseperson(i, 1, 2) = 0 Then
                              cardAInumcaseperson(i, 1, 2) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 11) = cardAInumcaseperson(i, 1, 11) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 1, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 1, cardcountAInum(j, we))) + 1
                     Case a2a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 3) Or cardAInumcaseperson(i, 1, 3) = 0 Then
                              cardAInumcaseperson(i, 1, 3) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 4) Or cardAInumcaseperson(i, 1, 4) = 0 Then
                              cardAInumcaseperson(i, 1, 4) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 12) = cardAInumcaseperson(i, 1, 12) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 2, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 2, cardcountAInum(j, we))) + 1
                     Case a3a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 5) Or cardAInumcaseperson(i, 1, 5) = 0 Then
                              cardAInumcaseperson(i, 1, 5) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 6) Or cardAInumcaseperson(i, 1, 6) = 0 Then
                              cardAInumcaseperson(i, 1, 6) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 13) = cardAInumcaseperson(i, 1, 13) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 3, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 3, cardcountAInum(j, we))) + 1
                     Case a4a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 7) Or cardAInumcaseperson(i, 1, 7) = 0 Then
                              cardAInumcaseperson(i, 1, 7) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 8) Or cardAInumcaseperson(i, 1, 8) = 0 Then
                              cardAInumcaseperson(i, 1, 8) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 14) = cardAInumcaseperson(i, 1, 14) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 4, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 4, cardcountAInum(j, we))) + 1
                     Case a5a
                          If cardcountAInum(j, we) < cardAInumcaseperson(i, 1, 9) Or cardAInumcaseperson(i, 1, 9) = 0 Then
                              cardAInumcaseperson(i, 1, 9) = cardcountAInum(j, we)
                          End If
                          If cardcountAInum(j, we) > cardAInumcaseperson(i, 1, 10) Or cardAInumcaseperson(i, 1, 10) = 0 Then
                              cardAInumcaseperson(i, 1, 10) = cardcountAInum(j, we)
                          End If
                          cardAInumcaseperson(i, 1, 15) = cardAInumcaseperson(i, 1, 15) + cardcountAInum(j, we)
                          cardAInumcasepersonTER(i, 5, cardcountAInum(j, we)) = Val(cardAInumcasepersonTER(i, 5, cardcountAInum(j, we))) + 1
                End Select
        End Select
    Next
Next
End Sub
Sub ���z��AI�t�έp��_�T���q_�έp�ƦC()
Dim i As Integer, j As Integer, k As Integer
'=================�ƻs���e
For k = 1 To 2 ^ cardAInumuscom
    cardAInumFinal2(k, 1) = cardAInumFinal(k, 1)
    cardAInumFinal2(k, 2) = cardAInumFinal(k, 2)
Next
'=================
Dim wer As Integer, wes As Integer
For i = 2 ^ cardAInumuscom To 1 Step -1
    For j = 1 To i - 1
        If Val(cardAInumFinal2(j, 1)) < Val(cardAInumFinal2(j + 1, 1)) Then
            wer = cardAInumFinal2(j + 1, 1)
            wes = cardAInumFinal2(j + 1, 2)
            cardAInumFinal2(j + 1, 1) = cardAInumFinal2(j, 1)
            cardAInumFinal2(j + 1, 2) = cardAInumFinal2(j, 2)
            cardAInumFinal2(j, 1) = wer
            cardAInumFinal2(j, 2) = wes
        End If
    Next
Next
End Sub
Sub ���z��AI�t�έp��_�|���q_���_1_��l()
Dim i As Integer
For i = 1 To 2 ^ cardAInumuscom
    If Val(cardAInumFinal2(i, 1)) > Val(cardAInumselect1) Then
        cardAInumselect1 = cardAInumFinal2(i, 1)
    End If
Next
'====================
If cardAInumselect1 < 0 Then cardAInumselect1 = 0 '�h���`����Ȭ��t�Ƥ��զX
'====================
For i = 1 To 2 ^ cardAInumuscom
    If cardAInumFinal2(i, 1) = cardAInumselect1 Then
        cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
    End If
Next
'====================
If cardAInumselect2 = "" Then  '�S������զX�ŦX����
    cardAInumselect2 = "-10=-10"
End If
End Sub
Sub ���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1()
Dim i As Integer, j As Integer
cardAInumselect3 = Split(cardAInumselect2, "=")
If UBound(cardAInumselect3) > 1 Then
    For i = 1 To 2 ^ cardAInumuscom
        For j = 1 To cardAInumuscom
             If cardAInumcaseperson(Val(cardAInumFinal2(i, 2)), 2, j) < 0 Then
                 cardAInumFinal2(i, 3) = 1
             End If
             cardAInumFinal2(i, 4) = Val(cardAInumFinal2(i, 4)) + Val(cardAInumcaseperson(cardAInumFinal2(i, 2), 2, j))
        Next
    Next
    '===============
    Erase cardAInumselect3
    cardAInumselect2 = ""
    '======
    For i = 1 To 2 ^ cardAInumuscom
        If cardAInumFinal2(i, 1) = cardAInumselect1 And cardAInumFinal2(i, 3) = 0 Then
            cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
        End If
    Next
    cardAInumselect3 = Split(cardAInumselect2, "=")
End If
End Sub
Sub ���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2()
Dim i As Integer
If UBound(cardAInumselect3) > 1 Then
    Dim wer As Integer
    For i = 1 To 2 ^ cardAInumuscom
         If Val(cardAInumFinal2(i, 4)) > Val(wer) And cardAInumFinal2(i, 1) = cardAInumselect1 Then
             wer = cardAInumFinal2(i, 4)
         End If
    Next
    '===============
    Erase cardAInumselect3
    cardAInumselect2 = ""
    cardAInumselect4 = wer
    '======
    For i = 1 To 2 ^ cardAInumuscom
        If cardAInumFinal2(i, 4) = wer And cardAInumFinal2(i, 1) = cardAInumselect1 And cardAInumFinal2(i, 3) = 0 Then
            cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
        End If
    Next
    cardAInumselect3 = Split(cardAInumselect2, "=")
End If
End Sub
Sub ���z��AI�t�έp��_�|���q_���_3_��ܲզX()
If UBound(cardAInumselect3) > 1 Then
    Dim wtr As Integer '�Ȯ��ܼ�
    wtr = Int(Rnd() * UBound(cardAInumselect3)) + 1
    cardAInumchoose = cardAInumselect3(wtr)
Else
    cardAInumchoose = cardAInumselect3(1)
End If
End Sub
Sub ���z��AI�t�έp��_�̫ᶥ�q_����P(ByVal choose As Integer, ByVal uscom As Integer)
Dim wer As Integer, i As Integer, cspce As String, cspme As String '�Ȯ��ܼ�
If choose = 1 Then
    wer = 0
Else
    wer = 1
End If
'=================
Dim pu As Integer '�Ȯ��ܼ�
'=====
If cardAInumchoose = -10 Then  '==�S������զX�ŦX�X�P����
    Exit Sub
End If
'=======================�p�զX�ŦX�X�P���󪺸�
Select Case uscom
     Case 1 '==�ϥΪ̤�
            For i = 1 To cardAInumuscom
                    pu = cardcountAInum(i, 5)
                    If Mid(cardAInumnm(cardAInumchoose - 1), i, 1) = 1 And cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 4
                    ElseIf cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 3
                    End If
            Next
     Case 2 '==�q����
            For i = 1 To cardAInumuscom
                    pu = cardcountAInum(i, 5)
                    If Mid(cardAInumnm(cardAInumchoose - 1), i, 1) = 1 And cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        cspce = pagecardnum(pu, 1)
                        cspme = pagecardnum(pu, 2)
                        pagecardnum(pu, 1) = pagecardnum(pu, 3)
                        pagecardnum(pu, 2) = pagecardnum(pu, 4)
                        pagecardnum(pu, 3) = cspce
                        pagecardnum(pu, 4) = cspme
                        If pageonin(pu) = 2 Then
                           pageonin(pu) = 1
                        Else
                           pageonin(pu) = 2
                        End If
                    End If
                    If cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 1
                    End If
            Next
End Select
End Sub
Sub ���z��AI�t�έp��_�ȮɶץX(ByVal uscom As Integer)
Dim i As Integer, k As Integer
If Formsetting.checktest.Value = 1 Then
'    Open App.Path & "\test\out1.txt" For Output As #1
    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & �԰��t����.turn & "turn_" & �԰��t����.turnatk & "_" & uscom & "_1.txt" For Output As #1
    For i = 1 To 2 ^ cardAInumuscom
        Print #1, cardAInumnm(Val(cardAInumFinal2(i, 2)) - 1) & "=" & cardAInumFinal2(i, 1) & "/" & cardAInumFinal2(i, 4) & "#" & cardAInumFinal2(i, 2) & "@";
        For k = 1 To cardAInumuscom
            Print #1, cardAInumcaseperson(Val(cardAInumFinal2(i, 2)), 2, k) & "=";
        Next
        Print #1,
    Next
    Close
    'MsgBox "�w�ץX����1"
End If
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_����1(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer)
���z��AI�t����.���z��AI�t�έp��_�@���q_��l uscom
���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turn, movecpre, uscom
���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turn, movecpre, uscom
���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
���z��AI�t����.���z��AI�t�έp��_�ȮɶץX uscom
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_���(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer)
If Val(pageglead(uscom)) > 10 Then
    ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_�W�X�P�i�� uscom, turn, name, movecpre, choose
ElseIf Val(pageglead(uscom)) > 0 And Val(pageglead(uscom)) <= 10 Then
    ���z��AI�t����.���z��AI�t�έp��_�@���q_��l pageglead(uscom)
    ���z��AI�t����.���z��AI�t�έp��_�@���q_���o�P����� True, uscom
    ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turn, movecpre, uscom
    ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turn, movecpre, uscom
    ���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_1_��l
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2
    ���z��AI�t����.���z��AI�t�έp��_�ȮɶץX uscom
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_3_��ܲզX
    If turn = 3 And cardAInumchoose > 0 Then
        ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_���ʶ��q�� uscom, turn, name, movecpre, choose, pageglead(uscom)
    Else
        ���z��AI�t����.���z��AI�t�έp��_�̫ᶥ�q_����P choose, uscom
    End If
End If
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_���ʶ��q��(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer, ByVal pagenumber As Integer)
If Val(pagenumber) > 0 Then
    Select Case ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P���(uscom)
        Case True
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�@���q_�ǳƶi����
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�G���q_�i����p�ƦC�զX��p�� pagenumber, uscom
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�T���q_�i����p����ȭp�� uscom, name, choose, movecpre, pagenumber
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�|���q_�έp���p����ȤΧP�_ uscom
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�����q_����P choose, uscom, pagenumber
        Case False
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_�_�w��_�@���q_���]�����_�ӧO
            ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_�_�w��_�G���q_��ܦ�� uscom
            ���z��AI�t����.���z��AI�t�έp��_�̫ᶥ�q_����P choose, uscom
    End Select
End If
End Sub
Function ���z��AI�t��_�ثe�i���椧�H���P�_(ByVal name As String) As Boolean
If Formsetting.chkusenewai.Value = 0 Then
    ���z��AI�t��_�ثe�i���椧�H���P�_ = False
    Exit Function
End If
Select Case name
    Case "��B�����S"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�j�|�˺��h"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���["
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�v��L"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "C.C."
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "��ܵY"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "������"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "��̬d�w"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "������"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�Q��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�L���S"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���纸"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "������S"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�w�ǥ���"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����P��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�h�g�H"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�Ǧh"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���_�i���h"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�S�{��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "����"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "���Y�F"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "��"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "ù��Y"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�����g"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�J�y"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�����i"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case "�ײ��d"
            ���z��AI�t��_�ثe�i���椧�H���P�_ = True
    Case Else
            ���z��AI�t��_�ثe�i���椧�H���P�_ = False
End Select
End Function
Function ���h��(ByVal num As Integer) As Single
Dim w As Double, i As Integer
w = 1
If num <> 0 Then
    For i = 1 To Val(num)
        w = Val(w) * Val(i)
    Next
Else
    w = 1
End If
���h�� = w
End Function
Function ���h��_��C(ByVal c1 As Integer, ByVal c2 As Integer) As Single

���h��_��C = (���z��AI�t����.���h��(c1) / ���z��AI�t����.���h��(Val(c1) - Val(c2))) / ���z��AI�t����.���h��(c2)

End Function
Sub ���z��AI�t�έp��_���ʶ��q��_���o�p�⤧�ƦC�զX(ByVal n1 As Integer, ByVal n2 As Integer)
Dim wtstr As String, wtall As Integer, wtpnum() As String, wtn As Integer, i As Integer, j As Integer
'===================
���z��AI�t����.�ƦC�զX�p�� n1
wtall = ���z��AI�t����.���h��_��C(n1, n2)
ReDim cardAInumMOVnm(1 To wtall) As String
'====================
For i = 1 To 2 ^ n1
    wtn = 0
    For j = 1 To n1
        If Val(Mid(cardAInumnm(i - 1), j, 1)) = 1 Then
            wtn = wtn + 1
        End If
    Next
    If wtn = n2 Then '==��n2�i�X�P���զX
        wtstr = wtstr & "=" & i
    End If
Next
wtpnum = Split(wtstr, "=")
'If UBound(wtpnum) = wtall Then
'    MsgBox wtstr
'    For i = 1 To UBound(wtpnum)
'        Debug.Print wtpnum(i) & "=" & cardAInumnm(wtpnum(i) - 1)
'    Next
'Else
'    MsgBox "����"
'End If
For i = 1 To UBound(wtpnum)
    cardAInumMOVnm(i) = cardAInumnm(wtpnum(i) - 1)
Next
End Sub
Function ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P���(ByVal uscom As Integer) As Boolean
Erase cardAInumMOVmain
Erase cardAInumMOVnm
Erase cardAInumMOVnmtot
Dim wtmovnum As Integer, i As Integer '�Ȯ��ܼ�
If cardAInumchoose = -10 Then
    ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P��� = False
    Exit Function
End If
'============�����ثe�զX
cardAInumMOVmain(1, 1) = cardAInumselect1
cardAInumMOVmain(1, 2) = cardAInumselect4
cardAInumMOVmain(1, 3) = cardAInumnm(cardAInumchoose - 1)
cardAInumMOVmain(1, 4) = cardAInumcaseperson(cardAInumchoose, 1, 13)
cardAInumMOVmain(1, 5) = cardAInumchoose
For i = 1 To cardAInumuscom
    cardAInumMOVmain(2, i) = cardAInumcaseperson(cardAInumchoose, 2, i)
Next
'==============�p�⦳�Ĳ��ʼ�
wtmovnum = cardAInumMOVmain(1, 4)
For i = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
    If (�H�����`���A��Ʈw(uscom, i, 3) = 6 And uscom = 2) Or (�H�����`���A��Ʈw(uscom, i, 3) = 12 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) - Val(�H�����`���A��Ʈw(uscom, i, 1))
    End If
    If (�H�����`���A��Ʈw(uscom, i, 3) = 3 And uscom = 2) Or (�H�����`���A��Ʈw(uscom, i, 3) = 9 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) + Val(�H�����`���A��Ʈw(uscom, i, 1))
    End If
    If (�H�����`���A��Ʈw(uscom, i, 3) = 16 And uscom = 1) Or (�H�����`���A��Ʈw(uscom, i, 3) = 17 And uscom = 2) Then
        wtmovnum = -100
    End If
Next
'=====================
If wtmovnum >= 2 Then
    ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P��� = True
Else
    ���z��AI�t�έp��_���ʶ��q��_�P�_�X�P��� = False
End If
End Function
Sub ���z��AI�t�έp��_���ʶ��q��_�_�w��_�@���q_���]�����_�ӧO()
Dim i As Integer
For i = 1 To cardAInumuscom
    If cardAInumcaseperson(cardAInumchoose, 2, i) < 10 Then
        cardAInumcaseperson(cardAInumchoose, 2, i) = 0
        cardAInumMOVmain(2, i) = 0
    End If
Next
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�@���q_�ǳƶi����()
Dim wercnum As Integer, werct As String, werpnum As Integer, k As Integer, q As Integer
ReDim cardcountAInumMOV(1 To cardAInumuscom, 1 To 5) As String
�O�_���ʶ��q����p�P�_�{�� = True
For k = 1 To cardAInumuscom
    Select Case Mid(cardAInumMOVmain(1, 3), k, 1)
         Case 0
              If cardcountAInum(k, 1) = a3a And cardAInumMOVmain(2, k) < 10 Then
                  wercnum = Val(wercnum) + 1
                  werct = werct & "=" & k
'              ElseIf cardAInumMOVmain(2, k) >= 10 Then
'                 werpnum = Val(werpnum) + 1
              End If
         Case 1
              If cardcountAInum(k, 3) = a3a And cardAInumMOVmain(2, k) < 10 Then
                  wercnum = Val(wercnum) + 1
                  werct = werct & "=" & k
'              ElseIf cardAInumMOVmain(2, k) >= 10 Then
'                 werpnum = Val(werpnum) + 1
              End If
    End Select
    For q = 1 To 5
         cardcountAInumMOV(k, q) = cardcountAInum(k, q)
    Next
Next
'===============
'If Val(werpnum) >= 1 Then werpnum = 1
'===============
ReDim cardAInumMOVnmtot(0 To (2 ^ wercnum), 1 To 8) As String
cardAInumMOVnmtot(0, 1) = werct
cardAInumMOVnmtot(0, 2) = 1
cardAInumMOVnmtot(0, 3) = wercnum
cardAInumMOVnmtot(0, 4) = 1
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�G���q_�i����p�ƦC�զX��p��(ByVal pagenumber As Integer, ByVal uscom As Integer)
Dim weru As Integer, wernum As Integer, werqr As String
Dim werstru As String
Dim werpstr() As String
Dim wermovnm As Integer, wermovynm As Integer, i As Integer, k As Integer
'============�i����p�����ʵP�ƦC�զX�p��
For i = 1 To Val(cardAInumMOVnmtot(0, 3))
       ���z��AI�t�έp��_���ʶ��q��_���o�p�⤧�ƦC�զX Val(cardAInumMOVnmtot(0, 3)), i
       weru = 1
       wernum = ���h��_��C(Val(cardAInumMOVnmtot(0, 3)), i)
        For k = Val(cardAInumMOVnmtot(0, 2)) To (Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)) - 1
             cardAInumMOVnmtot(k, 1) = cardAInumMOVnm(weru)
             weru = Val(weru) + 1
        Next
        cardAInumMOVnmtot(0, 2) = Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)
Next
'=====================�i��Ѿl���ʵP���ƦC�զX���X
'werpstr = Split(cardAInumMOVnmtot(1, 1), "=")
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
    weru = 0
    werstru = ""
    wermovnm = 0
    wermovynm = 0
    For k = 1 To pagenumber
        Select Case Mid(cardAInumMOVmain(1, 3), k, 1)
              Case 0
                    If cardcountAInum(k, 1) = a3a And i <= (2 ^ Val(cardAInumMOVnmtot(0, 3)) - 1) And cardAInumMOVmain(2, k) < 10 Then
                        weru = Val(weru) + 1
                        If Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 1 Then
                            werstru = werstru & "1"
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
                            wermovynm = Val(wermovynm) + 1
                        ElseIf Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 0 Then
'                            werstru = werstru & "1"
'                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
'                            wermovynm = Val(wermovynm) + 1
'                        Else
                            werstru = werstru & "n"
                        End If
                    ElseIf cardAInumMOVmain(2, k) >= 10 Then
                        werstru = werstru & "1"
                        If cardcountAInum(k, 1) = a3a Then
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 2))
                            wermovynm = Val(wermovynm) + 1
                        End If
                    Else
                        werstru = werstru & "n"
                    End If
              Case 1
                    If cardcountAInum(k, 3) = a3a And i <= (2 ^ Val(cardAInumMOVnmtot(0, 3)) - 1) And cardAInumMOVmain(2, k) < 10 Then
                        weru = Val(weru) + 1
                        If Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 1 Then
                            werstru = werstru & "1"
                            wermovynm = Val(wermovynm) + 1
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
                        ElseIf Mid(cardAInumMOVnmtot(i, 1), weru, 1) = 0 Then
'                            werstru = werstru & "1"
'                            wermovynm = Val(wermovynm) + 1
'                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
'                        Else
                            werstru = werstru & "n"
                        End If
                    ElseIf cardAInumMOVmain(2, k) >= 10 Then
                        werstru = werstru & "1"
                        If cardcountAInum(k, 3) = a3a Then
                            wermovnm = Val(wermovnm) + Val(cardcountAInum(k, 4))
                            wermovynm = Val(wermovynm) + 1
                        End If
                    Else
                        werstru = werstru & "n"
                    End If
        End Select
    Next
    cardAInumMOVnmtot(i, 2) = werstru
    cardAInumMOVnmtot(i, 6) = wermovnm
    cardAInumMOVnmtot(i, 7) = wermovynm
Next
'=========================���եζץX
If Formsetting.checktest.Value = 1 Then
'    Open App.Path & "\test\out2.txt" For Output As #1
    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & �԰��t����.turn & "turn_" & �԰��t����.turnatk & "_" & uscom & "_2.txt" For Output As #1
    For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
        Print #1, cardAInumMOVnmtot(i, 2)
    Next
    Print #1, cardAInumMOVmain(1, 5) & "=" & cardAInumMOVmain(1, 3)
    Close
    'MsgBox "�w�ץX����2"
End If
'==============================
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�T���q_�i����p����ȭp��(ByVal uscom As Integer, ByVal name As String, ByVal choose As Integer, ByVal movecpre As Integer, ByVal pagenumber As Integer)
Dim weru As Integer, wertp As Integer, movecpren As Integer, turnm As Integer, werucount As Boolean, i As Integer, k As Integer, q As Integer, wp As Integer, wds As Integer
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
    For k = 1 To 2
         '===========�N����ಾ�ܫݹB����
         weru = 0
         For wp = 1 To pagenumber
              If Mid(cardAInumMOVnmtot(i, 2), wp, 1) = "n" Then
                  weru = Val(weru) + 1
              End If
         Next
         If Val(weru) > 0 Then
                 ���z��AI�t����.���z��AI�t�έp��_�@���q_��l weru
                 wertp = 0
                 '=======
                 For q = 1 To pagenumber
                     If Mid(cardAInumMOVnmtot(i, 2), q, 1) = "n" Then
                           wertp = Val(wertp) + 1
                           For wds = 1 To 5
                                 cardcountAInum(wertp, wds) = cardcountAInumMOV(q, wds)
                           Next
                    End If
                Next
                '========================
                If k = 1 Then movecpren = 1 Else movecpren = 3
                If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) And werucount = True Then
                    turnm = 2
                    movecpren = movecpre
                Else
                    turnm = 1
                End If
                '========================
                ���z��AI�t����.���z��AI�t�έp��_�@���q_���o�P����� False, uscom
                ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turnm, movecpren, uscom
                ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turnm, movecpren, uscom
                ���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_1_��l
                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1
                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2
                ���z��AI�t����.���z��AI�t�έp��_�|���q_���_3_��ܲզX
        Else
                cardAInumselect1 = 0
        End If
        '=======================�N���s���p�����x�s
        If k = 1 And werucount = False Then
           movecpren = 3
        ElseIf k = 2 And werucount = False Then
           movecpren = 4
        ElseIf werucount = True Then
           movecpren = 8
        End If
        '=========
        cardAInumMOVnmtot(i, movecpren) = cardAInumselect1
        '=========
        If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) And k = 2 And werucount = False Then
           werucount = True
           k = 0
        ElseIf werucount = True Then
           k = 2
        End If
        '==========================
    Next
Next
'=========================���եζץX
If Formsetting.checktest.Value = 1 Then
'    Open App.Path & "\test\out3.txt" For Output As #1
    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & �԰��t����.turn & "turn_" & �԰��t����.turnatk & "_" & uscom & "_3.txt" For Output As #1
    For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
        Print #1, i & "=" & cardAInumMOVnmtot(i, 2) & "=";
        For k = 3 To 4
              Print #1, cardAInumMOVnmtot(i, k) & "#";
        Next
        If i = 2 ^ Val(cardAInumMOVnmtot(0, 3)) Then
            Print #1, cardAInumMOVnmtot(i, 8);
        End If
        Print #1,
    Next
    
    Close
    'MsgBox "�w�ץX����3"
End If
'==============================
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�|���q_�έp���p����ȤΧP�_(ByVal uscom As Integer)
Dim atk1max As Integer, atk2max As Integer, defmax As Integer, chemax As Integer, chestr As String
Dim wtmovnum As Integer, i As Integer
'==================�z��O�_�ŦX���ʶq
For i = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
    If (�H�����`���A��Ʈw(uscom, i, 3) = 6 And uscom = 2) Or (�H�����`���A��Ʈw(uscom, i, 3) = 12 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) - Val(�H�����`���A��Ʈw(uscom, i, 1))
    End If
    If (�H�����`���A��Ʈw(uscom, i, 3) = 3 And uscom = 2) Or (�H�����`���A��Ʈw(uscom, i, 3) = 9 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) + Val(�H�����`���A��Ʈw(uscom, i, 1))
    End If
    If (�H�����`���A��Ʈw(uscom, i, 3) = 16 And uscom = 1) Or (�H�����`���A��Ʈw(uscom, i, 3) = 17 And uscom = 2) Then
        wtmovnum = -100
    End If
Next
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, 6)) + Val(wtmovnum) < 2 Then
         cardAInumMOVnmtot(i, 5) = "x"
     Else
         cardAInumMOVnmtot(i, 5) = "y"
     End If
Next
'===================
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, 3)) > Val(atk1max) And cardAInumMOVnmtot(i, 5) = "y" Then
         atk1max = cardAInumMOVnmtot(i, 3)
     End If
     If Val(cardAInumMOVnmtot(i, 4)) > Val(atk2max) And cardAInumMOVnmtot(i, 5) = "y" Then
         atk2max = cardAInumMOVnmtot(i, 4)
     End If
Next
defmax = cardAInumMOVnmtot(2 ^ Val(cardAInumMOVnmtot(0, 3)), 8)
'==================
If Val(atk1max) >= Val(atk2max) And Val(atk1max) >= Val(defmax) Then
    chemax = 1
ElseIf Val(atk1max) <= Val(atk2max) And Val(atk2max) >= Val(defmax) Then
    chemax = 2
ElseIf Val(defmax) >= Val(atk1max) And Val(defmax) >= Val(atk2max) Then
    chemax = 3
Else
    chemax = 3
End If
'==================
Select Case chemax
     Case 1
           cardAInumMOVFinal(3) = 1
           cardAInumMOVFinal(2) = atk1max
           ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�T�{���_��̲ܳײզX 1, atk1max
     Case 2
           cardAInumMOVFinal(3) = 2
           cardAInumMOVFinal(2) = atk2max
           ���z��AI�t����.���z��AI�t�έp��_���ʶ��q��_���V��_�T�{���_��̲ܳײզX 2, atk2max
     Case 3
           cardAInumMOVFinal(1) = cardAInumMOVnmtot(2 ^ Val(cardAInumMOVnmtot(0, 3)), 2)
           cardAInumMOVFinal(3) = 3
           cardAInumMOVFinal(2) = defmax
End Select
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�T�{���_��̲ܳײզX(ByVal movche As Integer, ByVal atkmax As Integer)
Dim werstr As String, werg() As String, werg2() As String, werg3() As String
Dim werpagenum As Integer, werpgnumstr As String
Dim wermovmaxnum As Integer, wermvaxstr As String
Dim werrndnum As Integer, werche As Integer, i As Integer, k As Integer
'==========================
If movche = 1 Then werche = 3 Else werche = 4
'==========================
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
     If Val(cardAInumMOVnmtot(i, werche)) = Val(atkmax) Then
         werstr = werstr & "=" & i
     End If
Next
werg = Split(werstr, "=")
If UBound(werg) > 1 Then
        '====================================
        werpagenum = 0 '==�ت����̤j���X�P��
        For k = 1 To UBound(werg)
            If cardAInumMOVnmtot(werg(k), 7) > werpagenum Then
                werpagenum = cardAInumMOVnmtot(werg(k), 7)
            End If
        Next
        For k = 1 To UBound(werg)
            If cardAInumMOVnmtot(werg(k), 7) = werpagenum Then
                werpgnumstr = werpgnumstr & "=" & werg(k)
            End If
        Next
        werg2 = Split(werpgnumstr, "=")
        If UBound(werg2) > 1 Then
                '====================================
                wermovmaxnum = 0 '==�ت����̤j�����ʼ�
                For k = 1 To UBound(werg2)
                    If Val(cardAInumMOVnmtot(werg(k), 6)) > Val(wermovmaxnum) Then
                        wermovmaxnum = cardAInumMOVnmtot(werg(k), 6)
                    End If
                Next
                For k = 1 To UBound(werg2)
                    If Val(cardAInumMOVnmtot(werg(k), 6)) = wermovmaxnum Then
                        wermvaxstr = wermvaxstr & "=" & werg2(k)
                    End If
                Next
                werg3 = Split(wermvaxstr, "=")
                If UBound(werg3) > 1 Then
                     Randomize
                     werrndnum = Int(Rnd() * UBound(werg3)) + 1
                     cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg3(werrndnum), 2)
                Else
                     cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg3(1), 2)
                End If
                '==========================================
        Else
                cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg2(1), 2)
        End If
        '====================================
Else
        cardAInumMOVFinal(1) = cardAInumMOVnmtot(werg(1), 2)
End If
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_���V��_�����q_����P(ByVal choose As Integer, ByVal uscom As Integer, ByVal pagenumber As Integer)
Dim wer As Integer '�Ȯ��ܼ�
If choose = 1 Then
    wer = 0
Else
    wer = 1
End If
'=================
Dim pu As Integer, i As Integer, cspce As String, cspme As String '�Ȯ��ܼ�
'=======================�p�զX�ŦX�X�P���󪺸�
Select Case uscom
     Case 1 '==�ϥΪ̤�
            For i = 1 To pagenumber
                    pu = cardcountAInumMOV(i, 5)
                    If Mid(cardAInumMOVFinal(1), i, 1) = 1 Then
                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 And Val(cardAInumMOVmain(2, i)) >= wer Then
                                pagecardnum(pu, 11) = 4
                            ElseIf Val(cardAInumMOVmain(2, i)) >= wer Then
                                pagecardnum(pu, 11) = 3
                            End If
                    End If
            Next
            '===================��ܦ��
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      �ثe��(33) = 3
                 Case 2
                      �ثe��(33) = 1
                 Case 3
                      �ثe��(33) = 2
            End Select
     Case 2 '==�q����
            For i = 1 To pagenumber
                    pu = cardcountAInumMOV(i, 5)
                    If Mid(cardAInumMOVFinal(1), i, 1) = 1 Then
                            If Mid(cardAInumMOVmain(1, 3), i, 1) = 1 And Val(cardAInumMOVmain(2, i)) >= wer Then
                                cspce = pagecardnum(pu, 1)
                                cspme = pagecardnum(pu, 2)
                                pagecardnum(pu, 1) = pagecardnum(pu, 3)
                                pagecardnum(pu, 2) = pagecardnum(pu, 4)
                                pagecardnum(pu, 3) = cspce
                                pagecardnum(pu, 4) = cspme
                                If pageonin(pu) = 2 Then
                                   pageonin(pu) = 1
                                Else
                                   pageonin(pu) = 2
                                End If
                            End If
                            If Val(cardAInumMOVmain(2, i)) >= wer Then
                                pagecardnum(pu, 11) = 1
                            End If
                    End If
            Next
            '===================��ܦ��
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      �q���貾�ʶ��q��ܼ� = 3
                 Case 2
                      �q���貾�ʶ��q��ܼ� = 1
                 Case 3
                      �q���貾�ʶ��q��ܼ� = 2
            End Select
End Select

�O�_���ʶ��q����p�P�_�{�� = False
End Sub
Sub ���z��AI�t�έp��_�޾ɵ{��_�W�X�P�i��(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer)
Dim i As Integer, w As Integer
If Val(pageglead(uscom)) > 10 Then
    Erase cardAInumOvertenrecord
    ���z��AI�t����.���z��AI�t�έp��_�@���q_��l 10
    '=========�^���ثe�P�����(�e10�i)
        Select Case uscom
            Case 1
                �԰��t����.�X�P���ǭp��_�ϥΪ�_��P
            Case 2
                �԰��t����.�X�P���ǭp��_�q��_��P
        End Select
        w = 2 * uscom '(2-�ϥΪ̤�P/4-�q����P)
        For i = 1 To 10
            cardcountAInum(i, 5) = �X�P���ǲέp�Ȯ��ܼ�(w, i, 2)
            cardAInumOvertenrecord(i) = �X�P���ǲέp�Ȯ��ܼ�(w, i, 2)
            cardcountAInum(i, 1) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 1)
            cardcountAInum(i, 2) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 2)
            cardcountAInum(i, 3) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 3)
            cardcountAInum(i, 4) = pagecardnum(�X�P���ǲέp�Ȯ��ܼ�(w, i, 2), 4)
        Next
     '========================
    ���z��AI�t����.���z��AI�t�έp��_�@���q_���o�P����� False, uscom
    ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_��l turn, movecpre, uscom
    ���z��AI�t����.���z��AI�t�έp��_�G���q_�p������_�ӧO�ޯ� name, turn, movecpre, uscom
    ���z��AI�t����.���z��AI�t�έp��_�T���q_�έp�ƦC
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_1_��l
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__1
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_2_�W�B��ǧP�__2
    ���z��AI�t����.���z��AI�t�έp��_�ȮɶץX uscom
    ���z��AI�t����.���z��AI�t�έp��_�|���q_���_3_��ܲզX
    If turn = 3 And cardAInumchoose > 0 Then
        ���z��AI�t����.���z��AI�t�έp��_�޾ɵ{��_���ʶ��q�� uscom, turn, name, movecpre, choose, 10
    Else
        ���z��AI�t����.���z��AI�t�έp��_�̫ᶥ�q_����P choose, uscom
    End If
    '==========================
    If turn <> 3 Then
        �԰��t����.comatk_���z��AI�޾ɵ{��_�W�X�P�i�� turn, movecpre, choose
    End If
End If
End Sub
Sub �ˬd�H���ޯ�O�_��EX��(ByVal uscom As Integer, ByVal name As String)
Erase personatkingtfr
Dim i As Integer, k As Integer
For i = 1 To 3
     If VBEPerson(uscom, i, 1, 1, 1) = name Then
         For k = 1 To 4
               If Mid(VBEPerson(uscom, i, 3, k, 1), 1, 2) = "Ex" Then
                   personatkingtfr(k) = 1
               Else
                   personatkingtfr(k) = 0
               End If
          Next
          For k = 14 * (i - 1) + 1 To 14 * i
                If (�H�����`���A��Ʈw(uscom, k, 3) = 22 And uscom = 1) Or _
                    (�H�����`���A��Ʈw(uscom, k, 3) = 23 And uscom = 2) Then
                    personatkingtfr(5) = 1
                End If
          Next
     End If
Next
End Sub
Sub ���z��AI�t��_�ϥΪ̥X�P���q�P�_����()
Dim i As Integer
For i = 1 To 106
    If Val(pagecardnum(i, 11)) = 4 And Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
        FormMainMode.cgen_Click (i)
        pagecardnum(i, 11) = 3
    End If
Next
End Sub
Sub ���z��AI�t�έp��_���ʶ��q��_�_�w��_�G���q_��ܦ��(ByVal uscom As Integer)
Select Case uscom
    Case 1
        �ثe��(33) = 2
    Case 2
         �q���貾�ʶ��q��ܼ� = 2
End Select
End Sub
