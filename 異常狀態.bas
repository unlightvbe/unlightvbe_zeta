Attribute VB_Name = "���`���A"
Option Explicit
Public ���`���A_�V�P������(1 To 4) As Integer '���`���A-�V�P-��q�����Ȯ��ܼ�(1.�����ƭ�(��l)/2.�����ƭ�(�ܧ��)/3.�ƭȬ����O�_�Ұ�/4.�������m�Ҧ����q��)
Public ���`���A_AI_�V�P������(1 To 4) As Integer '���`���A-AI-�V�P-��q�����Ȯ��ܼ�(1.�����ƭ�(��l)/2.�����ƭ�(�ܧ��)/3.�ƭȬ����O�_�Ұ�/4.�������m�Ҧ����q��)
Sub ATK�[_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(7, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 7 Then
        atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "+" & �H�����`���A��Ʈw(1, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 7 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(7, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(7, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub ATK�[_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(1, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 1 Then
        atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "+" & �H�����`���A��Ʈw(2, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 1 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(1, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(1, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub ATK��_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(10, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 10 Then
        atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "-" & �H�����`���A��Ʈw(1, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 10 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(10, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(10, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub ATK��_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(4, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 4 Then
        atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "-" & �H�����`���A��Ʈw(2, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 4 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(4, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(4, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub DEF�[_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(8, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 8 Then
        atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "+" & �H�����`���A��Ʈw(1, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 8 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(8, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(8, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub DEF��_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(11, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 11 Then
        atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "+" & �H�����`���A��Ʈw(1, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 11 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(11, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(11, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub DEF�[_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(2, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 2 Then
        atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "+" & �H�����`���A��Ʈw(2, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 2 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(2, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(2, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub DEF��_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(5, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 5 Then
        atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "-" & �H�����`���A��Ʈw(2, i, 1) & "="
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 5 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(5, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(5, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub MOV�[_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(9, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 9 Then
            moveus = moveus + �H�����`���A��Ʈw(1, i, 1)
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 9 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(9, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(9, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub MOV��_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(12, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 12 Then
           If moveus > 0 Then
               moveus = moveus - �H�����`���A��Ʈw(1, i, 1)
               If moveus < 0 Then moveus = 0
           End If
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 12 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(12, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(12, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub MOV�[_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(3, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 3 Then
           movecom = movecom + �H�����`���A��Ʈw(2, i, 1)
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 3 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(3, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(3, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub MOV��_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(6, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 6 Then
           If movecom > 0 Then
               movecom = movecom - �H�����`���A��Ʈw(2, i, 1)
               If movecom < 0 Then
                   movecom = 0
                   movecheckcom = 0
                End If
           End If
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 6 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(6, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(6, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub ����_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(14, 1)
    Case 1
        For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
          If �H�����`���A��Ʈw(1, i, 3) = 14 Then
             �Y���淾�q�Ȯ��ܼ�(2) = 0
             �Y����ˮ`�� = 0
          End If
        Next
    Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 14 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(14, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub ����_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(18, 1)
    Case 1
        For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
          If �H�����`���A��Ʈw(2, i, 3) = 18 Then
             �Y���淾�q�Ȯ��ܼ�(2) = 0
             �Y����ˮ`�� = 0
          End If
        Next
    Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 18 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(18, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub ���r_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(20, 1)
    Case 1
        For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
          If �H�����`���A��Ʈw(1, i, 3) = 20 Then
            �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
            �ˮ`����_�ϥΪ� (1)
            If �H�����`���A��Ʈw(1, i, 2) = 0 Then
              '===�~�ӤU�@���A���
               �԰��t����.���`���A�~��_�ϥΪ�
               ���`���A�ˬd��(21, 2) = 0
           Else
               FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
           End If
         End If
        Next
End Select
End Sub
Sub ���r_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(21, 1)
    Case 1
        For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
          If �H�����`���A��Ʈw(2, i, 3) = 21 Then
            �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
            �ˮ`����_�q�� (1)
            If �H�����`���A��Ʈw(2, i, 2) = 0 Then
              '===�~�ӤU�@���A���
               �԰��t����.���`���A�~��_�q��
               ���`���A�ˬd��(21, 2) = 0
           Else
               FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
           End If
         End If
        Next
    Case 2
        movecom = 0
        movecheckcom = 0
End Select
End Sub
Sub ���a_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(15, 1)
    Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 15 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
            �԰��t����.�ˮ`����_�ߧY���`_�ϥΪ� 1
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(15, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub ���a_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(19, 1)
    Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 19 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
            �԰��t����.�ˮ`����_�ߧY���`_�q�� 1
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(19, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub �ʦL_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(22, 1)
    Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 22 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(22, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub �ʦL_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(23, 1)
    Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 23 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(23, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub ��O�C�U_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(24, 1)
  Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 24 Then
        atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "-" & �H�����`���A��Ʈw(1, i, 2) * 1 & "="
        Exit For
       End If
     Next
End Select
End Sub
Sub ��O�C�U_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(25, 1)
  Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 25 Then
        atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "-" & �H�����`���A��Ʈw(2, i, 2) * 1 & "="
        Exit For
       End If
     Next
End Select
End Sub
Sub �·�_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(16, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 16 Then
        moveus = 0
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 16 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(16, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(16, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub �·�_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(17, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 17 Then
        movecom = 0
        movecheckcom = 0
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 17 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(17, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(17, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub �t��_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(13, 1)
  Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 13 Then
        atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "+" & �H�����`���A��Ʈw(1, i, 2) * 1 & "="
        Exit For
       End If
     Next
End Select
End Sub
Sub �t��_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(26, 1)
  Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 26 Then
        atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "+" & �H�����`���A��Ʈw(2, i, 2) * 1 & "="
        Exit For
       End If
     Next
End Select
End Sub
Sub ����_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(29, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 29 Then
            �Y����ˮ`�� = �Y����ˮ`�� \ 2
            �Y���淾�q�Ȯ��ܼ�(2) = �Y����ˮ`��
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 29 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(29, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(29, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub �g�Ԥh_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(27, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 27 Then
            �Y����ˮ`�� = �Y����ˮ`�� * 2
            �Y���淾�q�Ȯ��ܼ�(2) = �Y����ˮ`��
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 27 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(27, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
            ���`���A�ˬd��(27, 1) = 1
        End If
      End If
     Next
End Select
End Sub

Sub �g�Ԥh_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(28, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 28 Then
            �Y����ˮ`�� = �Y����ˮ`�� * 2
            �Y���淾�q�Ȯ��ܼ�(2) = �Y����ˮ`��
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 28 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(28, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(28, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub ����_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(30, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 30 Then
            �Y����ˮ`�� = �Y����ˮ`�� \ 2
            �Y���淾�q�Ȯ��ܼ�(2) = �Y����ˮ`��
       End If
     Next
   Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 30 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(30, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
            ���`���A�ˬd��(30, 1) = 1
        End If
      End If
     Next
End Select
End Sub
Sub �V�P_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(31, 1)
   Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 31 Then
            If Val(���`���A_�V�P������(3)) = 0 Then
                ���`���A_�V�P������(1) = �������m��l�`��(1)
                ���`���A_�V�P������(2) = �������m��l�`��(1) * 2
                ���`���A_�V�P������(3) = 1
                �������m��l�`��(1) = �������m��l�`��(1) * 2
            ElseIf Val(���`���A_�V�P������(3)) = 1 Then
                ���`���A_�V�P������(1) = ���`���A_�V�P������(1) + (�������m��l�`��(1) - ���`���A_�V�P������(2))
                �������m��l�`��(1) = ���`���A_�V�P������(1) * 2
                ���`���A_�V�P������(2) = ���`���A_�V�P������(1) * 2
            End If
       End If
     Next
   Case 2
        For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
          If �H�����`���A��Ʈw(1, i, 3) = 31 Then
            �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
            If �H�����`���A��Ʈw(1, i, 2) = 0 Then
              '===�~�ӤU�@���A���
               �԰��t����.���`���A�~��_�ϥΪ�
               ���`���A�ˬd��(31, 2) = 0
           Else
               FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
               ���`���A�ˬd��(31, 1) = 1
           End If
         End If
        Next
   Case 3
        Erase ���`���A_�V�P������
End Select
End Sub
Sub �V�P_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(32, 1)
   Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 32 Then
            If Val(���`���A_AI_�V�P������(3)) = 0 Then
                ���`���A_AI_�V�P������(1) = �������m��l�`��(2)
                ���`���A_AI_�V�P������(2) = �������m��l�`��(2) * 2
                ���`���A_AI_�V�P������(3) = 1
                �������m��l�`��(2) = �������m��l�`��(2) * 2
            ElseIf Val(���`���A_AI_�V�P������(3)) = 1 Then
                ���`���A_AI_�V�P������(1) = ���`���A_AI_�V�P������(1) + (�������m��l�`��(2) - ���`���A_AI_�V�P������(2))
                �������m��l�`��(2) = ���`���A_AI_�V�P������(1) * 2
                ���`���A_AI_�V�P������(2) = ���`���A_AI_�V�P������(1) * 2
            End If
       End If
     Next
   Case 2
        For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
          If �H�����`���A��Ʈw(2, i, 3) = 32 Then
            �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
            If �H�����`���A��Ʈw(2, i, 2) = 0 Then
              '===�~�ӤU�@���A���
               �԰��t����.���`���A�~��_�q��
               ���`���A�ˬd��(32, 2) = 0
           Else
               FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
               ���`���A�ˬd��(32, 1) = 1
           End If
         End If
        Next
   Case 3
        Erase ���`���A_AI_�V�P������
End Select
End Sub

Sub �G��_�ϥΪ�(ByVal moveend As Integer)
Dim dge As Integer, i As Integer
Select Case ���`���A�ˬd��(33, 1)
    Case 1
        If movecp > 0 Then
            For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
              If �H�����`���A��Ʈw(1, i, 3) = 33 Then
                 dge = Abs(moveend - movecp)
                 �ˮ`����_�ޯઽ��_�ϥΪ� dge, 1
              End If
            Next
        End If
    Case 2
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
       If �H�����`���A��Ʈw(1, i, 3) = 33 Then
         �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
         If �H�����`���A��Ʈw(1, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�ϥΪ�
            ���`���A�ˬd��(33, 2) = 0
        Else
            FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub �G��_�q��(ByVal moveend As Integer)
Dim dge As Integer, i As Integer
Select Case ���`���A�ˬd��(34, 1)
    Case 1
        If movecp > 0 Then
            For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
              If �H�����`���A��Ʈw(2, i, 3) = 34 Then
                 dge = Abs(moveend - movecp)
                 �ˮ`����_�ޯઽ��_�q�� dge, 1
              End If
            Next
        End If
    Case 2
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
       If �H�����`���A��Ʈw(2, i, 3) = 34 Then
         �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
         If �H�����`���A��Ʈw(2, i, 2) = 0 Then
           '===�~�ӤU�@���A���
            �԰��t����.���`���A�~��_�q��
            ���`���A�ˬd��(34, 2) = 0
        Else
            FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
        End If
      End If
     Next
End Select
End Sub
Sub ���@_�ϥΪ�(ByVal num As Integer, ByRef tot As Integer)
Dim i As Integer
Select Case ���`���A�ˬd��(35, 1)
  Case 1
     For i = 14 * (����ݾ��H��������(1, num) - 1) + 1 To 14 * ����ݾ��H��������(1, num)
       If �H�����`���A��Ʈw(1, i, 3) = 35 Then
            If tot > 0 Then
                tot = 0
                '================
                If num = 1 Then
                    FormMainMode.messageus.AddItem "���@�ĪG�o��!    �����쪺�ˮ`�L�Ĥ�"
                Else
                    FormMainMode.messageus.AddItem "���@�ĪG�o��!    �ݾ����������쪺�ˮ`�L�Ĥ�"
                End If
                �԰��t����.�۰ʱ��b����
                '================
                �H�����`���A��Ʈw(1, i, 2) = 0
                 If �H�����`���A��Ʈw(1, i, 2) = 0 Then
                   '===�~�ӤU�@���A���
                    �԰��t����.���`���A�~��_�ϥΪ�
                    ���`���A�ˬd��(35, 2) = 0
                End If
            End If
            Exit For
       End If
     Next
End Select
End Sub
Sub ���@_�q��(ByVal num As Integer, ByRef tot As Integer)
Dim i As Integer
Select Case ���`���A�ˬd��(36, 1)
  Case 1
     For i = 14 * (����ݾ��H��������(2, num) - 1) + 1 To 14 * ����ݾ��H��������(2, num)
       If �H�����`���A��Ʈw(2, i, 3) = 36 Then
            If tot > 0 Then
                tot = 0
                '================
                If num = 1 Then
                    FormMainMode.messageus.AddItem "���@�ĪG�o��!    �������쪺�ˮ`�L�Ĥ�"
                Else
                    FormMainMode.messageus.AddItem "���@�ĪG�o��!    �����ݾ��������쪺�ˮ`�L�Ĥ�"
                End If
                �԰��t����.�۰ʱ��b����
                '================
                �H�����`���A��Ʈw(2, i, 2) = 0
                 If �H�����`���A��Ʈw(2, i, 2) = 0 Then
                   '===�~�ӤU�@���A���
                    �԰��t����.���`���A�~��_�q��
                    ���`���A�ˬd��(36, 2) = 0
                End If
            End If
            Exit For
       End If
     Next
End Select
End Sub
Sub �A��_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(37, 1)
    Case 1
        For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
          If �H�����`���A��Ʈw(1, i, 3) = 37 Then
            �H�����`���A��Ʈw(1, i, 2) = �H�����`���A��Ʈw(1, i, 2) - 1
            �԰��t����.�^�_����_�ϥΪ� 1, 1
            If �H�����`���A��Ʈw(1, i, 2) = 0 Then
              '===�~�ӤU�@���A���
               �԰��t����.���`���A�~��_�ϥΪ�
               ���`���A�ˬd��(37, 2) = 0
           Else
               FormMainMode.personusspe(i).person_turn = �H�����`���A��Ʈw(1, i, 2)
           End If
         End If
        Next
End Select
End Sub
Sub �A��_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(38, 1)
    Case 1
        For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
          If �H�����`���A��Ʈw(2, i, 3) = 38 Then
            �H�����`���A��Ʈw(2, i, 2) = �H�����`���A��Ʈw(2, i, 2) - 1
            �԰��t����.�^�_����_�q�� 1, 1
            If �H�����`���A��Ʈw(2, i, 2) = 0 Then
              '===�~�ӤU�@���A���
               �԰��t����.���`���A�~��_�q��
               ���`���A�ˬd��(38, 2) = 0
           Else
               FormMainMode.personcomspe(i).person_turn = �H�����`���A��Ʈw(2, i, 2)
           End If
         End If
        Next
End Select
End Sub
Sub �{��_�ϥΪ�()
Dim i As Integer
Select Case ���`���A�ˬd��(39, 1)
  Case 1
     For i = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
        If �H�����`���A��Ʈw(1, i, 3) = 39 Then
             If �H�����`���A��Ʈw(1, i, 2) < 3 Then
                atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "+" & �H�����`���A��Ʈw(1, i, 2) * 1 & "="
                Exit For
             ElseIf �H�����`���A��Ʈw(1, i, 2) >= 3 Then
                atkingckdice(1, 1, 3) = atkingckdice(1, 1, 3) & "+" & 5 & "="
                Exit For
             End If
        End If
     Next
End Select
End Sub
Sub �{��_�q��()
Dim i As Integer
Select Case ���`���A�ˬd��(40, 1)
  Case 1
     For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
        If �H�����`���A��Ʈw(2, i, 3) = 40 Then
             If �H�����`���A��Ʈw(2, i, 2) < 3 Then
                atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "+" & �H�����`���A��Ʈw(2, i, 2) * 1 & "="
                Exit For
             ElseIf �H�����`���A��Ʈw(2, i, 2) >= 3 Then
                atkingckdice(2, 2, 3) = atkingckdice(2, 2, 3) & "+" & 5 & "="
                Exit For
             End If
        End If
     Next
End Select
End Sub
