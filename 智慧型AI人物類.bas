Attribute VB_Name = "���z��AI�H����"
Public �L���S_���q�B�z�O����(1 To 3) As Integer '���z��AI-�L���S-�Բ��P�_������(1.��e���q���/2.�ؼе������^�X��/3.���֪��z�ѬO�_�o��)
Sub ��B�����S(ByVal turn As Integer, ByVal movecpre As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer '�Ȯ��ܼ�
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================��K�g��
                 If Pn1 = 0 Then
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 10) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 10) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================�p��
                 If Pn2 = 0 Then
                         If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 14) >= 2 Then
                                     cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                     '=====================
                                     sq = 1
                                     wnm = 0
                                     Do '==���q�S�d�d�����C�B�}�l
                                        If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                              For k = 1 To cardAInumuscom
                                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                        Case 0
                                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                                 wnm = wnm + sq
                                                             End If
                                                        Case 1
                                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                                 wnm = wnm + sq
                                                             End If
                                                   End Select
                                                   If wnm >= 2 Then Exit Do
                                              Next
                                        End If
                                        sq = sq + 1
                                    Loop Until sq > 10
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
               '===========���L
               If Pn3 = 0 Then
                    If movecpre = 1 Then
                            If cardAInumcaseperson(i, 1, 14) >= 2 And cardAInumcaseperson(i, 1, 12) >= 2 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '=====================
                                  sq = 1
                                  wnm = 0
                                  Do '==���q�S�d�d�����C�B�}�l
                                     If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                              wnm = wnm + sq
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                              wnm = wnm + sq
                                                          End If
                                                End Select
                                                If wnm >= 2 Then Exit Do
                                           Next
                                     End If
                                     sq = sq + 1
                                 Loop Until sq > 10
                            End If
                     End If
               End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
                '=================����
                If Pn4 = 0 Then
                        Dim werp As Integer
                        werp = 0
                        For k = 1 To cardAInumuscom
                              If cardAInumcaseperson(i, 2, k) > 0 Then
                                  werp = Val(werp) + 1
                              End If
                        Next
                        If Val(werp) >= 3 Then
        '                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                            werp = 0
                            For k = 1 To cardAInumuscom
                                If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                    cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                    werp = Val(werp) + 1
                                End If
                            Next
                            If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                        ElseIf Val(werp) < 3 Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 3 Then
                                werp = 0
                                '==============1.�w��w�w�w�X�P�������@�[��
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.��ƭȬ�1���P�@�[��
                                If werp < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.��ƭȬ�2���P�@�[��
                                If Val(werp) < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                            End If
                        End If
                End If
         Next
End Select

End Sub
Sub ����(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, livewer As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
'=============================
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�۱��ɦV
                 If Pn1 = 0 Then
                        If cardAInumcaseperson(i, 1, 14) >= 1 Then
                            If cardAInumcaseperson(i, 1, 14) < livewer Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + Val(cardAInumcaseperson(i, 1, 14)) * 5
                            '======================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 2)) * 5
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 4)) * 5
                                                   End If
                                         End Select
                                    Next
                            ElseIf cardAInumcaseperson(i, 1, 14) >= livewer Then '==�զX�U�S�d�ƶW�L�ۨ���q��
                                    cardAInumFinal(i, 1) = -10000
                            End If
                        End If
                 End If
                 '================���b�B
                 If Pn4 = 0 Then
                         If movecpre = 3 Then
                               If cardAInumcasepersonTER(i, 3, 1) >= 1 Then
                                     '=====================
                                    For k = 1 To cardAInumuscom
                                         Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0 '===����p���ƥ�d
                                                       If cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a3a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 2
                                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 2
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a3a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 2
                                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 2
                                                       End If
                                             End Select
                                    Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
                '===========�����
                If Pn2 = 0 Then
                        If cardAInumcasepersonTER(i, 4, 3) >= 1 And _
                        ((����ʧ@_�ˬd�O�_�����w���`���A(uscom, 14) = False And uscom = 1) Or _
                        (����ʧ@_�ˬd�O�_�����w���`���A(uscom, 18) = False And uscom = 2)) Then
                              If (uscom = 2 And �������m��l�`��(1) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And �O�_���ʶ��q����p�P�_�{�� = False) Or _
                                 (uscom = 1 And �������m��l�`��(2) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And �O�_���ʶ��q����p�P�_�{�� = False) Or _
                                 (�O�_���ʶ��q����p�P�_�{�� = True And Val(livewer) <= 3) Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10000
                              '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                         End Select
                                    Next
                              End If
                        End If
                End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
                '=====================���j�¤�
                If Pn3 = 0 Then
                        If movecpre < 3 Then
                            Dim werp As Integer
                            werp = 0
                            If cardAInumcaseperson(i, 1, 11) >= 3 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '=====================
                                    For p = Val(cardAInumcaseperson(i, 1, 2)) To 1 Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 2))
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 4))
                                                       End If
                                             End Select
                                        Next
                                    Next
                            End If
                        End If
                End If
        Next
End Select
End Sub
Sub ���(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, livewer As Integer, livewermax As Integer '�Ȯ��ܼ�
If uscom = 1 Then
    livewer = liveus(����H����ԤH��(1, 2))
    livewermax = liveusmax(����H����ԤH��(1, 2))
Else
    livewer = livecom(����H����ԤH��(2, 2))
    livewermax = livecommax(����H����ԤH��(2, 2))
End If
'================
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�Q�T����
                 If Pn4 = 0 Then
                        If movecpre < 3 Then
                               If cardAInumcasepersonTER(i, 1, 3) >= 1 And cardAInumcasepersonTER(i, 5, 3) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If (cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 3) Or (cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 3) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 30
                                                  Else
                                                      cardAInumcaseperson(i, 2, k) = 0
                                                  End If
                                             Case 1
                                                  If (cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 3) Or (cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 3) Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 30
                                                  Else
                                                      cardAInumcaseperson(i, 2, k) = 0
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
           Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
                '====================��Ө���
                If Pn2 = 0 Then
                        If cardAInumcaseperson(i, 1, 13) >= 1 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                              '=====================
                              sq = 1
                              wnm = 0
                              Do '==���q���d�d�����C�B�}�l
                                 If cardAInumcasepersonTER(i, 3, sq) >= 1 Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + (11 - sq)
                                                          wnm = wnm + sq
                                                          If cardcountAInum(k, 3) <> a4a Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          End If
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + (11 - sq)
                                                          wnm = wnm + sq
                                                          If cardcountAInum(k, 1) <> a4a Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          End If
                                                      End If
                                            End Select
                                            If wnm >= 1 Then Exit Do
                                       Next
                                 End If
                                 sq = sq + 1
                             Loop Until sq > 10
                        End If
                End If
                '=====================�E���F��
                If Pn3 = 0 Then
                        If movecpre > 1 Then
                            If cardAInumcaseperson(i, 1, 12) >= 5 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  If livewer = livewermax Then
                                      cardAInumFinal(i, 1) = cardAInumFinal(i, 1) - ((cardAInumcaseperson(i, 1, 14) - cardAInumcaseperson(i, 1, 7)) * 2)
                                  End If
                                  '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 5 * Val(cardcountAInum(k, 2))
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 5 * Val(cardcountAInum(k, 4))
                                                   End If
                                         End Select
                                    Next
                            End If
                        End If
                End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
                '=====================�@����
                If Pn1 = 0 Then
                        If movecpre = 2 Then
                            Dim werp As Integer
                            werp = 0
                            If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '=====================
                                    For p = Val(cardAInumcaseperson(i, 1, 8)) To 1 Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 2))
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           werp = Val(werp) + Val(cardcountAInum(k, 4))
                                                       End If
                                             End Select
                                        Next
                                    Next
                            End If
                        End If
                End If
        Next
End Select

End Sub
Sub �j�|�˺��h(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�r��
                 If Pn1 = 0 Then
                        If movecpre = 1 Then
                               If cardAInumcasepersonTER(i, 1, 1) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
                '===========�大����
                If Pn3 = 0 Then
                        If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                              '=====================
                              sq = 1
                              wnm = 0
                              Do '==���q�S�d�d�����C�B�}�l
                                 If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                          wnm = wnm + sq
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + sq
                                                          wnm = wnm + sq
                                                      End If
                                            End Select
                                            If wnm >= 2 Then Exit Do
                                       Next
                                 End If
                                 sq = sq + 1
                             Loop Until sq > 10
                        End If
                End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
                '=====================�����[��
                If Pn2 = 0 Then
                        werp = 0
                        If cardAInumcaseperson(i, 1, 14) >= 2 And cardAInumcaseperson(i, 1, 13) = 0 Then
                              For k = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
                                    If �H�����`���A��Ʈw(2, k, 3) = 17 And usocm = 2 Then
                                        werp = 1
                                    ElseIf �H�����`���A��Ʈw(1, k, 3) = 16 And usocm = 1 Then
                                        werp = 1
                                    End If
                                    If �H�����`���A��Ʈw(2, k, 3) = 6 And usocm = 2 Then
                                        werp = 1
                                    ElseIf �H�����`���A��Ʈw(1, k, 3) = 12 And usocm = 1 Then
                                        werp = 1
                                    End If
                              Next
                              If werp = 1 Then
                                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                    werp = 0
                                    '=====================
                                      For p = Val(cardAInumcaseperson(i, 1, 8)) To 1 Step -1
                                          For k = 1 To cardAInumuscom
                                              Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                    Case 0
                                                         If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 2 Then
                                                             cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                             werp = Val(werp) + Val(cardcountAInum(k, 2))
                                                         End If
                                                    Case 1
                                                         If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 2 Then
                                                             cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                             werp = Val(werp) + Val(cardcountAInum(k, 4))
                                                         End If
                                               End Select
                                          Next
                                      Next
                                End If
                        End If
                End If
                '=====================�믫�O�l��
                If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcasepersonTER(i, 1, 1) >= 1 And cardAInumcasepersonTER(i, 5, 1) >= 1 And cardAInumcasepersonTER(i, 4, 1) >= 1 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                              '=====================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If (cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 1 And Val(weryu(3)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 1) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 1) = a5a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 1) = a4a Then
                                                       weryu(3) = weryu(3) + 1
                                                   End If
                                               End If
                                          Case 1
                                               If (cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 1 And Val(weryu(3)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 3) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 3) = a5a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 3) = a4a Then
                                                       weryu(3) = weryu(3) + 1
                                                   End If
                                               End If
                                     End Select
                                Next
                        End If
                End If
        Next
End Select

End Sub
Sub ���[(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, weryu(1 To 3) As Integer, livewer As Integer, livewermax As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
If uscom = 1 Then livewermax = liveusmax(����H����ԤH��(1, 2)) Else livewermax = livecommax(����H����ԤH��(2, 2))
'=============================
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�R�Ĥ��I
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 2 Then
                                   If cardAInumcaseperson(i, 1, 11) >= 6 And movecpre = 2 Then
                                       cardAInumFinal(i, 1) = 0 '==�}�R�Ĥ��I�ɤ����P�O�d���Ų��P�}
                                   Else
                                        For k = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
                                              If �H�����`���A��Ʈw(uscom, k, 3) = 26 And uscom = 2 Then
                                                  werp = �H�����`���A��Ʈw(uscom, k, 2)
                                              ElseIf �H�����`���A��Ʈw(uscom, k, 3) = 13 And uscom = 1 Then
                                                  werp = �H�����`���A��Ʈw(uscom, k, 2)
                                              End If
                                        Next
                                        If werp > 0 And cardAInumFinal(i, 1) > 0 Then
                                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (Val(werp) - 6) * 3
                                                '======================
                                                For k = 1 To cardAInumuscom
                                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                          Case 0
                                                               If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                   weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 2))
                                                               End If
                                                               If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                   weryu(2) = Val(weryu(2)) + Val(cardcountAInum(k, 2))
                                                               End If
                                                          Case 1
                                                               If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                   weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 4))
                                                               End If
                                                               If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                   weryu(2) = Val(weryu(2)) + Val(cardcountAInum(k, 4))
                                                               End If
                                                     End Select
                                                Next
                                        End If
                                    End If
                               End If
                        End If
                 End If
                 '=================�O�d���Ų�
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 6 Then
                                   For k = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
                                         If �H�����`���A��Ʈw(uscom, k, 3) = 26 And uscom = 2 Then
                                             werp = �H�����`���A��Ʈw(uscom, k, 2)
                                         ElseIf �H�����`���A��Ʈw(uscom, k, 3) = 13 And uscom = 1 Then
                                             werp = �H�����`���A��Ʈw(uscom, k, 2)
                                         End If
                                   Next
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (9 - Val(werp)) * 3
                                   If livewer = livewermax And werp = 9 Then
                                       cardAInumFinal(i, 1) = 0
                                   ElseIf cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 2 And movecpre = 2 Then
                                       cardAInumFinal(i, 1) = 0
                                   End If
                                   '======================
                                   If cardAInumFinal(i, 1) > 0 Then
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a1a And weryu(1) < 6 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 2))
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a1a And weryu(1) < 6 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 4))
                                                       End If
                                             End Select
                                        Next
                                    End If
                               End If
                        End If
                 End If
                 '=================�ԷX���T��
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 6 Then
                                   For k = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
                                         If �H�����`���A��Ʈw(uscom, k, 3) = 26 And uscom = 2 Then
                                             werp = �H�����`���A��Ʈw(uscom, k, 2)
                                         ElseIf �H�����`���A��Ʈw(uscom, k, 3) = 13 And uscom = 1 Then
                                             werp = �H�����`���A��Ʈw(uscom, k, 2)
                                         End If
                                   Next
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (Val(werp) - 8) * 3 + (3 - Val(livewer)) * 5
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 2))
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + Val(cardcountAInum(k, 4))
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
                '=====================���㤧��
                If Pn1 = 0 Then
                        werp = 0
                        If cardAInumcaseperson(i, 1, 14) >= 1 Then
                            For k = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
                                  If �H�����`���A��Ʈw(uscom, k, 3) = 26 And uscom = 2 Then
                                      werp = �H�����`���A��Ʈw(uscom, k, 2)
                                  ElseIf �H�����`���A��Ʈw(uscom, k, 3) = 13 And uscom = 1 Then
                                      werp = �H�����`���A��Ʈw(uscom, k, 2)
                                  End If
                            Next
                            If werp < 9 Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                '=====================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   Exit For
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   Exit For
                                               End If
                                     End Select
                                Next
                            End If
                        End If
                End If
        Next
End Select

End Sub
Sub �v��L(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�M�̤���
                 If Pn2 = 0 Then
                        werp = 0
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 6 Then
                                   If atking_�v��L_�����Ҧ����A��(2) = 1 Then
                                       For k = 1 To 3
                                             Select Case uscom
                                                   Case 1
                                                        If liveus(����ݾ��H��������(uscom, 1)) > 0 Then
                                                            werp = Val(werp) + 1
                                                        End If
                                                   Case 2
                                                        If livecom(����ݾ��H��������(uscom, 1)) > 0 Then
                                                            werp = Val(werp) + 1
                                                        End If
                                            End Select
                                        Next
                                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (livewer - werp) * 4
                                   Else
                                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   End If
                                   '======================
                                   werp = 0
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And werp < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      werp = Val(werp) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And werp < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      werp = Val(werp) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================�R�B���K��
                 If Pn4 = 0 Then
                         Erase weryu
                         werp = 0
                         If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 9 Then
                                     For k = 1 To 3
                                          weryu(k) = 999   '�ت����̧CHP�q
                                     Next
                                     Select Case uscom
                                          Case 1
                                               For k = 2 To 3
                                                     If Val(liveus(����ݾ��H��������(1, k))) < Val(weryu(1)) And Val(liveus(����ݾ��H��������(1, k))) > 0 Then
                                                         weryu(1) = liveus(����ݾ��H��������(1, k))
                                                    End If
                                               Next
                                               For k = 1 To 3
                                                     If Val(livecom(����ݾ��H��������(2, k))) < Val(weryu(2)) And Val(livecom(����ݾ��H��������(2, k))) > 0 Then
                                                         weryu(2) = livecom(����ݾ��H��������(2, k))
                                                    End If
                                               Next
                                          Case 2
                                               For k = 2 To 3
                                                     If Val(livecom(����ݾ��H��������(2, k))) < Val(weryu(1)) And Val(livecom(����ݾ��H��������(2, k))) > 0 Then
                                                         weryu(1) = livecom(����ݾ��H��������(2, k))
                                                    End If
                                               Next
                                               For k = 1 To 3
                                                     If Val(liveus(����ݾ��H��������(1, k))) < Val(weryu(2)) And Val(liveus(����ݾ��H��������(1, k))) > 0 Then
                                                         weryu(2) = liveus(����ݾ��H��������(1, k))
                                                    End If
                                               Next
                                     End Select
                                     If Val(weryu(2)) < Val(weryu(1)) Then
                                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                         '=====================
                                            For k = 1 To cardAInumuscom
                                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                      Case 0
                                                           If cardcountAInum(k, 1) = a1a Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           End If
                                                      Case 1
                                                           If cardcountAInum(k, 3) = a1a Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           End If
                                                 End Select
                                            Next
                                     End If
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
               '=================�ɶ��ؤl
               If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 12) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   werp = 0
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
                '=================�������x
                If Pn1 = 0 Then
                        werp = 0
                        For k = 1 To cardAInumuscom
                              If cardAInumcaseperson(i, 2, k) > 0 Then
                                  werp = Val(werp) + 1
                              End If
                        Next
                        If Val(werp) >= 3 Then
        '                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                            werp = 0
                            For k = 1 To cardAInumuscom
                                If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                    cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                    werp = Val(werp) + 1
                                End If
                            Next
                            If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                        ElseIf Val(werp) < 3 Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 3 Then
                                werp = 0
                                '==============1.�w��w�w�w�X�P�������@�[��
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.��ƭȬ�1���P�@�[��
                                If werp < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.��ƭȬ�2���P�@�[��
                                If Val(werp) < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                        End If
                End If
         Next
End Select

End Sub
Sub CC(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�ջȾԾ�
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 15) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================���W�q�Ϥ�N�M
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 14) >= 6 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 6 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================��l����
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================���ߪŶ�
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
            End If
         Next
End Select

End Sub
Sub ��ܵY(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================��������
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 15) >= 2 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   For k = 14 * (����H����ԤH��(uscom, 2) - 1) + 1 To 14 * ����H����ԤH��(uscom, 2)
                                         If �H�����`���A��Ʈw(uscom, k, 3) = 25 And uscom = 2 Then
                                             werp = �H�����`���A��Ʈw(uscom, k, 2)
                                         ElseIf �H�����`���A��Ʈw(uscom, k, 3) = 24 And uscom = 1 Then
                                             werp = �H�����`���A��Ʈw(uscom, k, 2)
                                         End If
                                   Next
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (9 - werp) * 5
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�E�����q
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre > 1 Then
                           If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================��k���Ӫ�
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre < 3 Then
                           If cardAInumcaseperson(i, 1, 14) >= 2 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '=====================
                                  sq = 1
                                  wnm = 0
                                  Do '==���q�S�d�d�����C�B�}�l
                                     If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                End Select
                                                If wnm >= 2 Then Exit Do
                                           Next
                                     End If
                                     sq = sq + 1
                                 Loop Until sq > 10
                           End If
                     End If
                End If
              '=====================�����ۺh
                If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                                If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And _
                                    cardAInumcaseperson(i, 1, 13) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 Then
                                      cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                                      '=====================
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If (cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 1) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 1) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 1) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 1) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 1) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                                  Case 1
                                                       If (cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 3) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 3) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 3) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 3) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 3) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                             End Select
                                        Next
                                End If
                        End If
                End If
         Next
End Select

End Sub
Sub ����(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�ɶ��z�u
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                 End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�ɶ��l�y
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre < 3 Then
                           If cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�o�����c
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre > 1 Then
                           If cardAInumcaseperson(i, 1, 14) >= 2 Then
                               '=====�p�U2�^�X���ݽ�Ʀ^�X�~���`����ȥ[��
                               If Val(�԰��t����.turn) + 2 = 3 Or Val(�԰��t����.turn) + 2 = 5 Or Val(�԰��t����.turn) + 2 = 7 Or _
                                  Val(�԰��t����.turn) + 2 = 11 Or Val(�԰��t����.turn) + 2 = 13 Or Val(�԰��t����.turn) + 2 = 17 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               End If
                               '=====================
                                  sq = 1
                                  wnm = 0
                                  Do '==���q�S�d�d�����C�B�}�l
                                     If (cardAInumcasepersonTER(i, 4, sq) >= 2 And sq = 1) Or (cardAInumcasepersonTER(i, 4, sq) >= 1 And sq > 1) Then
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = sq Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              wnm = wnm + sq
                                                          End If
                                                End Select
                                                If wnm >= 2 Then Exit Do
                                           Next
                                     End If
                                     sq = sq + 1
                                 Loop Until sq > 10
                           End If
                     End If
                End If
              '=====================�]���ɤ�
              If Pn4 = 0 Then
                        werp = 0
                        wnm = 0
                        Erase weryu
                        For k = 1 To cardAInumuscom
                              If cardAInumcaseperson(i, 2, k) > 0 Then
                                  werp = Val(werp) + 1
                              End If
                        Next
                        For k = 1 To 3
                            If VBEPerson(uscom, ����ݾ��H��������(uscom, k), 1, 2, 1) = "R" Then
                                 weryu(1) = Val(weryu(1)) + 1
                                 Select Case uscom
                                     Case 1
                                          If (Val(liveus(����ݾ��H��������(1, k))) + 3) <= Val(liveusmax(����ݾ��H��������(1, k))) Then
                                              weryu(2) = Val(weryu(2)) + 1
                                          End If
                                     Case 2
                                          If (Val(livecom(����ݾ��H��������(2, k))) + 3) <= Val(livecommax(����ݾ��H��������(2, k))) Then
                                              weryu(2) = Val(weryu(2)) + 1
                                          End If
                                End Select
                            End If
                        Next
                        If Val(werp) >= 3 Then
        '                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                            werp = 0
                            For k = 1 To cardAInumuscom
                                If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                    cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                    werp = Val(werp) + 1
                                End If
                            Next
                            If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                        ElseIf Val(werp) < 3 And Val(weryu(1)) >= 1 Then
                            If Val(weryu(2)) = 0 Then
                                wnm = 2
                            ElseIf Val(weryu(2)) > 0 Then
                                wnm = 3
                            End If
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= wnm And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= wnm And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 3 Then
                                werp = 0
                                '==============1.�w��w�w�w�X�P�������@�[��
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.��ƭȬ�1���P�@�[��
                                If werp < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.��ƭȬ�2���P�@�[��
                                If Val(werp) < 3 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============4.��ƭȬ�3���P�@�[��
                                If Val(werp) < 3 And Val(weryu(2)) > 0 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                        End If
                End If
         Next
End Select

End Sub
Sub ����(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=====================Lowball
                If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And _
                            cardAInumcaseperson(i, 1, 13) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 Then
                              cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                              '=====================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If (cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a3a And Val(weryu(3)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a4a And Val(weryu(4)) < 1) Or _
                                                   (cardcountAInum(k, 1) = a5a And Val(weryu(5)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 1) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 1) = a2a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 1) = a3a Then
                                                       weryu(3) = weryu(3) + 1
                                                   ElseIf cardcountAInum(k, 1) = a4a Then
                                                       weryu(4) = weryu(4) + 1
                                                   ElseIf cardcountAInum(k, 1) = a5a Then
                                                       weryu(5) = weryu(5) + 1
                                                   End If
                                               End If
                                          Case 1
                                               If (cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a3a And Val(weryu(3)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a4a And Val(weryu(4)) < 1) Or _
                                                   (cardcountAInum(k, 3) = a5a And Val(weryu(5)) < 1) Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   If cardcountAInum(k, 3) = a1a Then
                                                       weryu(1) = weryu(1) + 1
                                                   ElseIf cardcountAInum(k, 3) = a2a Then
                                                       weryu(2) = weryu(2) + 1
                                                   ElseIf cardcountAInum(k, 3) = a3a Then
                                                       weryu(3) = weryu(3) + 1
                                                   ElseIf cardcountAInum(k, 3) = a4a Then
                                                       weryu(4) = weryu(4) + 1
                                                   ElseIf cardcountAInum(k, 3) = a5a Then
                                                       weryu(5) = weryu(5) + 1
                                                   End If
                                               End If
                                     End Select
                                Next
                        End If
                 End If
                 '=================Gamble
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If Len(cardAInumnm(i - 1)) >= 3 Then
                                   For k = 1 To cardAInumuscom
                                         If cardAInumcaseperson(i, 2, k) > 0 Then
                                             werp = Val(werp) + 1
                                         End If
                                   Next
                                   If Val(werp) >= 3 Then
                                       werp = 0
                                       For k = 1 To cardAInumuscom
                                           If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 3 Then
                                               cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                               werp = Val(werp) + 1
                                           End If
                                       Next
                                       If Val(werp) >= 3 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   Else
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a5a) _
                                                           And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a5a) _
                                                           And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                      End If
                                            End Select
                                       Next
                                       If Val(werp) >= 3 Then
                                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                               '======================
                                               werp = 0
                                               For k = 1 To cardAInumuscom
                                                     If cardAInumcaseperson(i, 2, k) > 0 Then
                                                         werp = Val(werp) + 1
                                                     End If
                                               Next
                                               For k = 1 To cardAInumuscom
                                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                         Case 0
                                                              If cardAInumcaseperson(i, 2, k) = 0 And Val(werp) < 3 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                         Case 1
                                                              If cardAInumcaseperson(i, 2, k) = 0 And Val(werp) < 3 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                    End Select
                                               Next
                                       End If
                                   End If
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================High hand
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
             End If
              '=================Jackpot
             If Pn2 = 0 Then
                     werp = 0
                     Erase weryu
                     If movecpre = 2 Then
                            If cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                '======================
                                werp = 0
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                              If Val(cardcountAInum(k, 2)) < 6 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                              End If
                                          Case 1
                                              If Val(cardcountAInum(k, 4)) < 6 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                              End If
                                     End Select
                                Next
                            End If
                    End If
            End If
        Next
    Case 3 '==���ʶ��q��
        
End Select

End Sub
Sub ������(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewer41 As Integer, weryu(1 To 5) As Integer '�Ȯ��ܼ�
If uscom = 1 Then
    livewer = liveus(����H����ԤH��(1, 2))
    livewer41 = liveus41(����H����ԤH��(1, 2))
Else
    livewer = livecom(����H����ԤH��(2, 2))
    livewer41 = livecom41(����H����ԤH��(1, 2))
End If
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
              Erase weryu
              weryu(1) = 999 '�ت����̧CHP��
            '=================�r�֩��
            If Pn3 = 0 Then
                    If cardAInumcaseperson(i, 1, 14) >= 2 Then
                        Select Case uscom
                             Case 1
                                   For k = 1 To 3
                                         If Val(liveus(����ݾ��H��������(1, k))) < Val(weryu(1)) And Val(liveus(����ݾ��H��������(1, k))) > 0 Then
                                             weryu(1) = liveus(����ݾ��H��������(1, k))
                                             weryu(2) = liveus41(����ݾ��H��������(1, k))
                                             weryu(3) = k
                                        End If
                                   Next
                             Case 2
                                   For k = 1 To 3
                                         If Val(livecom(����ݾ��H��������(2, k))) < Val(weryu(1)) And Val(livecom(����ݾ��H��������(2, k))) > 0 Then
                                             weryu(1) = livecom(����ݾ��H��������(2, k))
                                             weryu(2) = livecom41(����ݾ��H��������(2, k))
                                             weryu(3) = k
                                        End If
                                   Next
                        End Select
                        If cardAInumcaseperson(i, 1, 14) < Val(weryu(1)) Or _
                           (Val(weryu(1)) < Val(weryu(2)) And cardAInumcaseperson(i, 1, 14) >= 10 And weryu(3) <> 1) Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + Val(cardAInumcaseperson(i, 1, 14)) * 5
                        '======================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 2)) * 5
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + Val(cardcountAInum(k, 4)) * 5
                                               End If
                                     End Select
                                Next
                        ElseIf cardAInumcaseperson(i, 1, 14) >= Val(weryu(1)) Then '==�զX�U�S�d�ƶW�L�ۨ�/�ݾ������̧C��q��
                                cardAInumFinal(i, 1) = -10000
                        End If
                    End If
            End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�ŬX�`�g
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    weryu(1) = 999 '�ت����̧CHP��
                     If movecpre < 3 Then
                           If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 12) >= 2 Then
                               Select Case uscom
                                    Case 1
                                          For k = 2 To 3
                                                If Val(liveus(����ݾ��H��������(1, k))) < Val(weryu(1)) And Val(liveus(����ݾ��H��������(1, k))) > 0 Then
                                                    weryu(1) = liveus(����ݾ��H��������(1, k))
                                                    weryu(2) = k
                                               End If
                                          Next
                                    Case 2
                                          For k = 2 To 3
                                                If Val(livecom(����ݾ��H��������(2, k))) < Val(weryu(1)) And Val(livecom(����ݾ��H��������(2, k))) > 0 Then
                                                    weryu(1) = livecom(����ݾ��H��������(2, k))
                                                    weryu(2) = k
                                               End If
                                          Next
                               End Select
                               If Val(weryu(2)) <> 0 And livewer < Val(weryu(1)) Then
                               Else   '====���F�ݾ�������q����ۨ���q�~
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                       '======================
                                       werp = 0
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                      If cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                                      If cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               End If
                           End If
                     End If
             End If
              '=================���K�W��
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre > 1 Then
                           If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 Then
                               Select Case uscom
                                    Case 1
                                           weryu(1) = liveus(����ݾ��H��������(1, 2))
                                           weryu(2) = liveus(����ݾ��H��������(1, 3))
                                           weryu(3) = liveus41(����ݾ��H��������(1, 2))
                                           weryu(4) = liveus41(����ݾ��H��������(1, 3))
                                    Case 2
                                           weryu(1) = livecom(����ݾ��H��������(2, 2))
                                           weryu(2) = livecom(����ݾ��H��������(2, 3))
                                           weryu(3) = livecom41(����ݾ��H��������(2, 2))
                                           weryu(4) = livecom41(����ݾ��H��������(2, 3))
                               End Select
                               If (Val(weryu(1)) < Val(weryu(3)) And Val(weryu(1)) > 0) And _
                                  (Val(weryu(2)) < Val(weryu(4)) And Val(weryu(2)) > 0) And _
                                   Val(livewer) < Val(livewer41) Then
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                                       '======================
                                       werp = 0
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                      If cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                      End If
                                                      If cardcountAInum(k, 1) = a4a And Val(weryu(3)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                                      If cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                      End If
                                                      If cardcountAInum(k, 3) = a4a And Val(weryu(3)) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               Else
                                       cardAInumFinal(i, 1) = -100
                               End If
                           End If
                     End If
            End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================��������
             If Pn1 = 0 Then
                     werp = 0
                     Erase weryu
                     For k = 1 To cardAInumuscom
                          If cardAInumcaseperson(i, 2, k) > 0 Then
                              werp = Val(werp) + 1
                          End If
                     Next
                     Select Case uscom
                         Case 1
                                weryu(1) = liveus(����ݾ��H��������(1, 2))
                                weryu(2) = liveus(����ݾ��H��������(1, 3))
                         Case 2
                                weryu(1) = livecom(����ݾ��H��������(2, 2))
                                weryu(2) = livecom(����ݾ��H��������(2, 3))
                     End Select
                      If Val(werp) >= 2 And Val(weryu(1)) <> 1 And Val(weryu(2)) <> 1 Then
                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                          werp = 0
                          '======================
                            For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                         Case 1
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                    End Select
                               Next
                      ElseIf Val(werp) < 2 And Val(weryu(1)) <> 1 And Val(weryu(2)) <> 1 Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 2 Then
                                werp = 0
                                '==============1.�w��w�w�w�X�P�������@�[��
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.��ƭȬ�1���P�@�[��
                                If werp < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.��ƭȬ�2���P�@�[��
                                If Val(werp) < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 3 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 2 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                      Else
                            cardAInumFinal(i, 1) = -10000
                      End If
                End If
         Next
End Select

End Sub
Sub ��̬d�w(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewermax As Integer, weryu(1 To 3) As Integer  '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
If uscom = 1 Then livewermax = liveusmax(����H����ԤH��(1, 2)) Else livewermax = livecommax(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�s�g
                 If Pn1 = 0 Then
                        If movecpre > 1 Then
                               If cardAInumcasepersonTER(i, 5, 1) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================���t���C
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 11) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + Int(Val(cardAInumcaseperson(i, 1, 11)) / 2 + 0.9)
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) > 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) > 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                                   '=============
                                   If cardAInumcasepersonTER(i, 1, 1) Mod 2 <> 0 Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      End If
                                            End Select
                                       Next
                                   ElseIf cardAInumcasepersonTER(i, 1, 1) > 0 Then
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 3) <> a2a _
                                                          And cardcountAInum(k, 2) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + 1
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 1) <> a2a _
                                                          And cardcountAInum(k, 4) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + 1
                                                      End If
                                            End Select
                                       Next
                                       If Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                           For k = 1 To cardAInumuscom
                                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                         Case 0
                                                              If cardcountAInum(k, 1) = a1a _
                                                                  And cardcountAInum(k, 2) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                         Case 1
                                                              If cardcountAInum(k, 3) = a1a _
                                                                  And cardcountAInum(k, 4) = 1 And Val(werp) < Val(cardAInumcasepersonTER(i, 1, 1)) - 1 Then
                                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                  werp = Val(werp) + 1
                                                              End If
                                                    End Select
                                               Next
                                        End If
                                   End If
                               End If
                        End If
                End If
                 '=================����@��
                 If Pn3 = 0 Then
                        werp = 0
                        wnm = 0
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                   wnm = (livewermax - livewer) * 2
                                   If wnm > 16 Then wnm = 16
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + wnm
                                   '======================
                                   For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              werp = Val(werp) + p
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              werp = Val(werp) + p
                                                          End If
                                                End Select
                                           Next
                                     Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================���}����
            If Pn4 = 0 Then
                    If cardAInumcasepersonTER(i, 2, 1) >= 2 Then
                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                        '======================
                        For k = 1 To cardAInumuscom
                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                  Case 0
                                       If cardcountAInum(k, 1) = a2a And cardcountAInum(k, 2) = 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       End If
                                  Case 1
                                       If cardcountAInum(k, 3) = a2a And cardcountAInum(k, 4) = 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       End If
                             End Select
                        Next
                    End If
            End If
        Next
    Case 3 '==���ʶ��q��
        
End Select

End Sub

Sub ������(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�Q���{��
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================�{�q�ۭ���
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 2 Then
                               If cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a3a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a3a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================�ۼv�C�R
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcasepersonTER(i, 1, 1) >= 1 And cardAInumcasepersonTER(i, 1, 2) >= 1 And cardAInumcasepersonTER(i, 1, 3) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================��M�_���p
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre = 1 Then
                           If cardAInumcaseperson(i, 1, 14) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                                 For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + p
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          werp = Val(werp) + p
                                                      End If
                                            End Select
                                       Next
                                 Next
                           End If
                    End If
            End If
         Next
End Select

End Sub
Sub �Q��(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                '=================�T�v����
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                         If cardAInumcaseperson(i, 1, 14) >= 1 Then
                             cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                             '======================
                               For k = 1 To cardAInumuscom
                                      Select Case Mid(cardAInumnm(i - 1), k, 1)
                                            Case 0
                                                 If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                            Case 1
                                                 If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                       End Select
                                  Next
                         End If
                 End If
                 '=================�r��
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 14) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================�I��
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   Select Case uscom
                                        Case 1
                                              For k = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
                                                     If �H�����`���A��Ʈw(2, k, 3) = 17 Then
                                                         weryu(1) = �H�����`���A��Ʈw(2, k, 2)
                                                     End If
                                               Next
                                        Case 2
                                               For k = 14 * (����H����ԤH��(1, 2) - 1) + 1 To 14 * ����H����ԤH��(1, 2)
                                                     If �H�����`���A��Ʈw(1, k, 3) = 16 Then
                                                         weryu(1) = �H�����`���A��Ʈw(1, k, 2)
                                                     End If
                                               Next
                                   End Select
                                   If (weryu(1) >= 2 And �O�_���ʶ��q����p�P�_�{�� = True) Or (weryu(1) >= 1 And �O�_���ʶ��q����p�P�_�{�� = False) Then
                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                           '======================
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                          End If
                                                          If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                          End If
                                                          If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                          End If
                                                End Select
                                           Next
                                   End If
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�������T��
            If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��

End Select

End Sub
Sub �L���S(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewer41 As Integer, livewermax As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then
    livewer = liveus(����H����ԤH��(1, 2))
    livewer41 = liveus41(����H����ԤH��(1, 2))
    livewermax = liveusmax(����H����ԤH��(1, 2))
Else
    livewer = livecom(����H����ԤH��(2, 2))
    livewer41 = livecom41(����H����ԤH��(1, 2))
    livewermax = livecommax(����H����ԤH��(2, 2))
End If
If �L���S_���q�B�z�O����(2) = Val(�԰��t����.turn) Then
    �L���S_���q�B�z�O����(1) = 0
End If
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�V����
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   If �L���S_���q�B�z�O����(1) = 0 Then
                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                           '======================
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                          End If
                                                          If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                          End If
                                                          If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                          End If
                                                End Select
                                           Next
                                   End If
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�j�t��
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    werp = Val(werp) + p
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    werp = Val(werp) + p
                                                End If
                                      End Select
                                 Next
                           Next
                     End If
            End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�]����
              If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   If �L���S_���q�B�z�O����(1) <> 3 Then
                                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                           If livewer <= livewer41 Then
                                               �L���S_���q�B�z�O����(1) = 2
                                               �L���S_���q�B�z�O����(2) = �԰��t����.turn + 2
                                           End If
                                           '======================
                                           For k = 1 To cardAInumuscom
                                               Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                     Case 0
                                                          If cardcountAInum(k, 1) = a2a And weryu(1) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                          End If
                                                          If cardcountAInum(k, 1) = a3a And weryu(2) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                          End If
                                                     Case 1
                                                          If cardcountAInum(k, 3) = a2a And weryu(1) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                          End If
                                                          If cardcountAInum(k, 3) = a3a And weryu(2) < 1 Then
                                                              cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                              weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                          End If
                                                End Select
                                           Next
                                     End If
                               End If
                        End If
                 End If
                 '=================���֪��z��
                 If Pn4 = 0 Then
                         werp = 0
                         Erase weryu
                         If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                If �L���S_���q�B�z�O����(1) = 0 And �L���S_���q�B�z�O����(3) = 0 Then
                                        Select Case uscom
                                             Case 1
                                                  For k = 2 To 3
                                                       If liveus(����ݾ��H��������(1, k)) <= 0 Then
                                                           werp = Val(werp) + 1
                                                       End If
                                                  Next
                                             Case 2
                                                  For k = 2 To 3
                                                       If livecom(����ݾ��H��������(2, k)) <= 0 Then
                                                           werp = Val(werp) + 1
                                                       End If
                                                  Next
                                        End Select
                                        If werp = 2 And livewer + 1 >= livewermax Then
                                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                                �L���S_���q�B�z�O����(1) = 3
                                                �L���S_���q�B�z�O����(2) = �԰��t����.turn + 2
                                                �L���S_���q�B�z�O����(3) = 1
                                                '======================
                                                For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                                      For k = 1 To cardAInumuscom
                                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                                Case 0
                                                                     If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(werp) < 3 Then
                                                                         cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                         werp = Val(werp) + p
                                                                     End If
                                                                Case 1
                                                                     If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(werp) < 3 Then
                                                                         cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                                         werp = Val(werp) + p
                                                                     End If
                                                           End Select
                                                      Next
                                                Next
                                        End If
                                End If
                        End If
                End If
         Next
End Select

End Sub
Sub ���纸(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================Rud-913
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a3a And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a3a And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================Chr-799
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And weryu(2) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================Wil-846
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================Von-541
            If Pn2 = 0 Then
                werp = 0
                Erase weryu
                If cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 Then
                    cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                    '======================
                    For k = 1 To cardAInumuscom
                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                              Case 0
                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                   End If
                                   If cardcountAInum(k, 1) = a2a And weryu(2) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                   End If
                              Case 1
                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                   End If
                                   If cardcountAInum(k, 3) = a2a And weryu(2) < 1 Then
                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                   End If
                         End Select
                    Next
                End If
            End If
        Next
    Case 3 '==���ʶ��q��
        
End Select

End Sub
Sub ������S(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewermax As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then
    livewer = liveus(����H����ԤH��(1, 2))
    livewermax = liveusmax(����H����ԤH��(1, 2))
Else
    livewer = livecom(����H����ԤH��(2, 2))
    livewermax = livecommax(����H����ԤH��(2, 2))
End If
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================���
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================�a���y���~
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (Val(cardAInumcaseperson(i, 1, 11)) + Val(cardAInumcaseperson(i, 1, 15)))
                               weryu(2) = ((Val(cardAInumcaseperson(i, 1, 11)) + Val(cardAInumcaseperson(i, 1, 15))) \ 5) * 5
                               '======================
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a1a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                              If cardcountAInum(k, 1) = a5a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a1a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                              If cardcountAInum(k, 3) = a5a And weryu(1) < weryu(2) Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                 End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�鱫
            If Pn2 = 0 Then
                 werp = 0
                 Erase weryu
                 If movecpre = 1 Then
                        If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                            cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            '======================
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 5)) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                           End If
                                      Case 1
                                           If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 5)) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                           End If
                                 End Select
                            Next
                        End If
                 End If
            End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================����ۼv
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 13) = 0 Then
                           cardAInumFinal(i, 1) = 1
                           '======================
                             For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 1)) And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                               If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 9)) And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 1)) And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                               If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 9)) And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                             Next
                     End If
              End If
         Next
End Select

End Sub
Sub �w�ǥ���(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�`�W
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 14) >= 3 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10 + Int(Val(cardAInumcaseperson(i, 1, 14)) / 2 + 0.9)
                               '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                           End If
                 End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�ƨg����
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    If movecpre = 1 Then
                           If cardAInumcaseperson(i, 1, 14) >= 1 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) And Val(weryu(1)) < 1 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                           End If
                    End If
             End If
              '=================�·t�x��
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 3)) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 5)) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 3)) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 5)) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
              End If
              '=================�F�z���������¼�
              If Pn1 = 0 Then
                    If movecpre = 3 Then
                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                    End If
              End If
        Next
    Case 3 '==���ʶ��q��
        
End Select

End Sub
Sub ����P��(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================C.T.L
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
                 '=================B.P.A
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================L.A.R
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
              End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=====================S.S.S
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                    If cardAInumcasepersonTER(i, 4, 1) >= 1 And cardAInumcasepersonTER(i, 4, 2) >= 1 And cardAInumcasepersonTER(i, 4, 3) >= 1 Then
                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                          '=====================
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 1 And Val(weryu(1)) < 1) Or _
                                               (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 2 And Val(weryu(2)) < 1) Or _
                                               (cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 And Val(weryu(3)) < 1) Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(cardcountAInum(k, 2)) = weryu(cardcountAInum(k, 2)) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 1 And Val(weryu(1)) < 1) Or _
                                               (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 2 And Val(weryu(2)) < 1) Or _
                                               (cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 And Val(weryu(3)) < 1) Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(cardcountAInum(k, 4)) = weryu(cardcountAInum(k, 4)) + 1
                                           End If
                                 End Select
                            Next
                    End If
             End If
         Next
End Select

End Sub
Sub �h�g�H(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�ݭh�ɦV
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                           If cardAInumcaseperson(i, 1, 14) >= 2 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               Next
                           End If
                End If
             '=================�ߦ���
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If movecpre = 1 Then
                           If cardAInumcaseperson(i, 1, 11) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                               '======================
                               werp = 0
                                 For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                         For k = 1 To cardAInumuscom
                                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                   Case 0
                                                        If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                        End If
                                                   Case 1
                                                        If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                        End If
                                              End Select
                                         Next
                                 Next
                           End If
                     End If
             End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
                '===========�����
                If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcasepersonTER(i, 4, 3) >= 1 Then
                              If (uscom = 2 And �������m��l�`��(1) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And �O�_���ʶ��q����p�P�_�{�� = False) Or _
                                 (uscom = 1 And �������m��l�`��(2) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And �O�_���ʶ��q����p�P�_�{�� = False) Or _
                                 (�O�_���ʶ��q����p�P�_�{�� = True And Val(livewer) <= 3) Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10000
                              '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 100
                                                       Exit For
                                                   End If
                                         End Select
                                    Next
                              End If
                        End If
                End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�W�Ťk�D��
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 15) >= 3 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                         Select Case uscom
                              Case 1
                                   If ����ʧ@_�ˬd�O�_�����w���`���A(1, 7) = True Then
                                       werp = 1
                                   End If
                                   If ����ʧ@_�ˬd�O�_�����w���`���A(1, 8) = True Then
                                       werp = 1
                                   End If
                                   If ����ʧ@_�ˬd�O�_�����w���`���A(1, 9) = True Then
                                       werp = 1
                                   End If
                              Case 2
                                   If ����ʧ@_�ˬd�O�_�����w���`���A(2, 1) = True Then
                                       werp = 1
                                   End If
                                   If ����ʧ@_�ˬd�O�_�����w���`���A(2, 2) = True Then
                                       werp = 1
                                   End If
                                   If ����ʧ@_�ˬd�O�_�����w���`���A(2, 3) = True Then
                                       werp = 1
                                   End If
                         End Select
                         If werp = 0 Then
                               cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                               '======================
                                 For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                   End If
                                                   If cardcountAInum(k, 1) = a5a And weryu(2) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                   End If
                                                   If cardcountAInum(k, 1) = a4a And weryu(3) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                   End If
                                                   If cardcountAInum(k, 3) = a5a And weryu(2) < 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                   End If
                                                   If cardcountAInum(k, 3) = a4a And weryu(3) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                                   End If
                                         End Select
                                    Next
                           End If
                     End If
                End If
         Next
End Select

End Sub
Sub �Ǧh(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�]�G����
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcaseperson(i, 1, 13) >= 1 Then
                            cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            '======================
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                           End If
                                      Case 1
                                           If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                           End If
                                 End Select
                            Next
                        End If
                 End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�]�G����
              If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
              End If
              '=================�]�G����
              If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 4 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 4 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 4 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
              End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�]�G���u
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
             End If
         Next
End Select

End Sub
Sub ���_�i���h(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�P�R�j��
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcasepersonTER(i, 5, 1) >= 1 And cardAInumcasepersonTER(i, 5, 2) >= 1 And cardAInumcasepersonTER(i, 5, 3) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 30
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 2 And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + 1
                                                  End If
                                                  If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 3 And weryu(3) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(3) = Val(weryu(3)) + 1
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================�T�v����
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                         If cardAInumcaseperson(i, 1, 14) >= 1 Then
                             cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                             '======================
                               For k = 1 To cardAInumuscom
                                      Select Case Mid(cardAInumnm(i - 1), k, 1)
                                            Case 0
                                                 If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                            Case 1
                                                 If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     Exit For
                                                 End If
                                       End Select
                                  Next
                         End If
                 End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================���@�g��
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 15) >= 1 And movecpre > 1 Then
                         If ((�������m��l�`��(uscom) >= 30 Or cardAInumcaseperson(i, 1, 9) = 1) And �O�_���ʶ��q����p�P�_�{�� = False) Or _
                             �O�_���ʶ��q����p�P�_�{�� = True Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                '======================
                                If �������m��l�`��(uscom) >= 30 Then
                                    werp = Int((�������m��l�`��(uscom) - 30) / 2 + 0.9)
                                Else
                                    werp = 0
                                End If
                                If werp > 0 Then
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a5a And Val(weryu(1)) < werp Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a5a And Val(weryu(1)) < werp Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                Else
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                End If
                         End If
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�j�a�Y�a
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 3 And movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                               For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                              End If
                                         Case 1
                                              If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 3 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                              End If
                                    End Select
                               Next
                          Next
                     End If
            End If
         Next
End Select

End Sub
Sub �S�{��(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�G�����F
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================���M�C�{
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10 + cardAInumcaseperson(i, 1, 11) * 5
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 2) * 5
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + cardcountAInum(k, 4) * 5
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�a�g���t
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 And _
                         movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================�t�v���l
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 And _
                         movecpre < 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 1) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 1) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        
End Select

End Sub
Sub ����(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�ڤ��]��
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 3 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 3 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================�ڹҷn�x
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 5
                                   If livewer <= 2 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 95
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�K�a�ڦ�
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 15) >= 1 And cardAInumcaseperson(i, 1, 12) >= 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(2)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 9) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(2)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================���Ϥ۹�
             If Pn3 = 0 And movecpre < 3 Then
                     werp = 0
                     Erase weryu
                     For k = 1 To cardAInumuscom
                          If cardAInumcaseperson(i, 2, k) > 0 Then
                              werp = Val(werp) + 1
                          End If
                     Next
                      If Val(werp) >= 2 And Val(livewer) >= 5 Then
                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                          werp = 0
                          '======================
                            For k = 1 To cardAInumuscom
                                   Select Case Mid(cardAInumnm(i - 1), k, 1)
                                         Case 0
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                         Case 1
                                              If cardAInumcaseperson(i, 2, k) > 0 And Val(werp) < 2 Then
                                                  cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                  werp = Val(werp) + 1
                                              End If
                                    End Select
                               Next
                      ElseIf Val(werp) < 2 And (cardAInumcaseperson(i, 1, 13) < 2 Or Val(livewer) >= 5) Then
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                               And cardcountAInum(k, 2) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                      Case 1
                                           If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                               And cardcountAInum(k, 4) <= 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                               werp = Val(werp) + 1
                                           End If
                                 End Select
                            Next
                            If Val(werp) >= 2 Then
                                werp = 0
                                '==============1.�w��w�w�w�X�P�������@�[��
                                For k = 1 To cardAInumuscom
                                      If cardAInumcaseperson(i, 2, k) > 0 Then
                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                          werp = Val(werp) + 1
                                      End If
                                Next
                                '==============2.��ƭȬ�1���P�@�[��
                                If werp < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 1 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 1 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
                                '==============3.��ƭȬ�2���P�@�[��
                                If Val(werp) < 2 Then
                                    For k = 1 To cardAInumuscom
                                          Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If (cardcountAInum(k, 1) = a1a Or cardcountAInum(k, 1) = a2a Or cardcountAInum(k, 1) = a4a Or cardcountAInum(k, 1) = a5a) _
                                                          And cardcountAInum(k, 2) = 2 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                                 Case 1
                                                      If (cardcountAInum(k, 3) = a1a Or cardcountAInum(k, 3) = a2a Or cardcountAInum(k, 3) = a4a Or cardcountAInum(k, 3) = a5a) _
                                                          And cardcountAInum(k, 4) = 2 And werp < 2 And cardAInumcaseperson(i, 2, k) = 0 Then
                                                          werp = Val(werp) + 1
                                                          cardAInumcaseperson(i, 2, k) = Val(cardAInumcaseperson(i, 2, k)) + 10
                                                      End If
                                            End Select
                                    Next
                                End If
        '                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + ((10 * 3) - (werp - 3) * 2)
                                If Val(werp) >= 2 Then cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                            End If
                      Else
                            cardAInumFinal(i, 1) = -10
                      End If
                End If
         Next
End Select

End Sub
Sub ���Y�F(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================���a�B��
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                '=====================����B
                If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                                If cardAInumcaseperson(i, 1, 11) >= 1 And _
                                    cardAInumcaseperson(i, 1, 13) >= 1 And cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 15) >= 1 Then
                                      cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                      If cardAInumuscom >= 10 Then
                                          cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                      End If
                                      '=====================
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If (cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 1) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 1) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 1) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 1) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 1) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 1) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                                  Case 1
                                                       If (cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a3a And Val(weryu(3)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a4a And Val(weryu(4)) < 1) Or _
                                                           (cardcountAInum(k, 3) = a5a And Val(weryu(5)) < 1) Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           If cardcountAInum(k, 3) = a1a Then
                                                               weryu(1) = weryu(1) + 1
                                                           ElseIf cardcountAInum(k, 3) = a2a Then
                                                               weryu(2) = weryu(2) + 1
                                                           ElseIf cardcountAInum(k, 3) = a3a Then
                                                               weryu(3) = weryu(3) + 1
                                                           ElseIf cardcountAInum(k, 3) = a4a Then
                                                               weryu(4) = weryu(4) + 1
                                                           ElseIf cardcountAInum(k, 3) = a5a Then
                                                               weryu(5) = weryu(5) + 1
                                                           End If
                                                       End If
                                             End Select
                                        Next
                                End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================��������
             If Pn2 = 0 And movecpre < 3 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 2 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a4a And Val(weryu(2)) < 2 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================����
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         Select Case uscom
                              Case 1
                                   If livecom(����H����ԤH��(2, 2)) = livecommax(����H����ԤH��(2, 2)) Then
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   End If
                              Case 2
                                   If liveus(����H����ԤH��(1, 2)) = liveusmax(����H����ԤH��(1, 2)) Then
                                       cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   End If
                         End Select
                         '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                       For k = 1 To cardAInumuscom
                                           Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                 Case 0
                                                      If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                      End If
                                                 Case 1
                                                      If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                          cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                          weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                      End If
                                            End Select
                                       Next
                               Next
                     End If
            End If
         Next
End Select

End Sub
Sub ��(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================���ۦ�-�[���⪺�L��
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 4 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 4 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                '=================Ex���ۦ�-�[���⪺�L��
                 If Pn1 = 1 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 5 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                 '=================�צ�-�L�ɽ��j���׵�
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 4 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                    For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                            For k = 1 To cardAInumuscom
                                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                      Case 0
                                                           If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 4 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                           End If
                                                      Case 1
                                                           If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 4 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                           End If
                                                 End Select
                                            Next
                                    Next
                               End If
                        End If
                End If
                '=================Ex�צ�-�L�ɽ��j���׵�
                 If Pn4 = 1 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 14) >= 6 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                    For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                            For k = 1 To cardAInumuscom
                                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                      Case 0
                                                           If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 6 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                           End If
                                                      Case 1
                                                           If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 6 Then
                                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                           End If
                                                 End Select
                                            Next
                                    Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�󫵦�-�[�ʯP���u�@
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================Ex�󫵦�-�[�ʯP���u�@
             If Pn2 = 1 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 3 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 3 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(2)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '===========�w�-���������q/Ex�w�-���������q
                If Pn3 = 0 Or Pn3 = 1 Then
                        If cardAInumcasepersonTER(i, 4, 3) >= 1 Then
                              If (�������m��l�`��(uscom) > (Val(livewer) + cardAInumcaseperson(i, 1, 12)) * 3 And �O�_���ʶ��q����p�P�_�{�� = False) Or _
                                 (�O�_���ʶ��q����p�P�_�{�� = True And Val(livewer) <= 3) Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                              '=====================
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       Exit For
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 3 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       Exit For
                                                   End If
                                         End Select
                                    Next
                              End If
                        End If
                End If
        Next
    Case 3 '==���ʶ��q��
        
End Select

End Sub
Sub ù��Y(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�V�大�b
                 If Pn2 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
                '=================Ex�V�大�b
                 If Pn2 = 1 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 3 And cardAInumcaseperson(i, 1, 13) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For p = Val(cardAInumcaseperson(i, 1, 6)) To Val(cardAInumcaseperson(i, 1, 5)) Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                    Next
                               End If
                        End If
                End If
                 '=================��������¶
                 If Pn4 = 0 Then
                       werp = 0
                       Erase weryu
                       If cardAInumcaseperson(i, 1, 14) >= 2 Then
                           cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                           '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                    For k = 1 To cardAInumuscom
                                        Select Case Mid(cardAInumnm(i - 1), k, 1)
                                              Case 0
                                                   If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                   End If
                                              Case 1
                                                   If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                       cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                       weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                   End If
                                         End Select
                                    Next
                            Next
                       End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�����ۼv
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 2 And movecpre = 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                     End If
             End If
             '=================Ex�����ۼv
             If Pn1 = 1 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 And movecpre = 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And weryu(1) < 2 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Next
                     End If
             End If
             '=================�C�G����L
             If Pn3 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 5 And cardAInumcaseperson(i, 1, 14) >= 1 And movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                            For k = 1 To cardAInumuscom
                                Select Case Mid(cardAInumnm(i - 1), k, 1)
                                      Case 0
                                           If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                           End If
                                      Case 1
                                           If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(1) < 1 Then
                                               cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                               weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                           End If
                                 End Select
                            Next
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        
End Select

End Sub
Sub �����g(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, livewermax As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then
    livewer = liveus(����H����ԤH��(1, 2))
    livewermax = liveusmax(����H����ԤH��(1, 2))
Else
    livewer = livecom(����H����ԤH��(2, 2))
    livewermax = livecommax(����H����ԤH��(2, 2))
End If
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================��������
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 3 Then
                               If cardAInumcaseperson(i, 1, 15) >= 4 And cardAInumcaseperson(i, 1, 14) >= 2 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                         For k = 1 To cardAInumuscom
                                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                   Case 0
                                                        If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                        End If
                                                   Case 1
                                                        If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                        End If
                                              End Select
                                         Next
                                    Next
                               End If
                        End If
                End If
                 '=================�g�����b�P�ݦ大�j
                 If Pn4 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre = 1 Then
                               If cardAInumcaseperson(i, 1, 11) >= 5 And cardAInumcaseperson(i, 1, 15) >= 5 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a5a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a5a And weryu(1) < 5 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================���ɷP��
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 And cardAInumcaseperson(i, 1, 12) >= 1 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                        If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And Val(weryu(1)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                                        If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And Val(weryu(3)) < 1 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�f��ԧ����j�T
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 3 And movecpre > 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + (livewermax - livewer) * 2
                         '======================
                           For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                End If
                                      End Select
                                 Next
                            Next
                     End If
            End If
         Next
End Select

End Sub
Sub �J�y(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 3) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�����g��
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre > 1 Then
                               If cardAInumcaseperson(i, 1, 15) >= 2 And cardAInumcaseperson(i, 1, 13) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a3a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a3a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 5) And weryu(1) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�Ѩ����
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 2 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 2 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                End If
                                      End Select
                                 Next
                            Next
                     End If
             End If
             '=================�k�`�p�e
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    If cardAInumcasepersonTER(i, 2, 1) >= 1 And cardAInumcasepersonTER(i, 4, 1) >= 1 Then
                        Select Case uscom
                            Case 1
                                   weryu(1) = liveus(����ݾ��H��������(1, 2))
                                   weryu(2) = liveus(����ݾ��H��������(1, 3))
                            Case 2
                                   weryu(1) = livecom(����ݾ��H��������(2, 2))
                                   weryu(2) = livecom(����ݾ��H��������(2, 3))
                        End Select
                        If �������m��l�`��(uscom) >= 30 And weryu(1) > 3 And weryu(2) > 3 And �O�_���ʶ��q����p�P�_�{�� = False Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                '======================
                                For k = 1 To cardAInumuscom
                                    Select Case Mid(cardAInumnm(i - 1), k, 1)
                                          Case 0
                                               If cardcountAInum(k, 1) = a2a And cardcountAInum(k, 2) = 1 And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                               End If
                                               If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = 1 And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                               End If
                                          Case 1
                                               If cardcountAInum(k, 3) = a2a And cardcountAInum(k, 4) = 1 And weryu(1) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                               End If
                                               If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = 1 And weryu(2) < 1 Then
                                                   cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                   weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                               End If
                                     End Select
                                Next
                        Else
                                cardAInumFinal(i, 1) = 0
                        End If
                    End If
            End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�c�N����
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                    If cardAInumcasepersonTER(i, 1, 3) >= 1 And cardAInumcasepersonTER(i, 5, 3) >= 1 And cardAInumcaseperson(i, 1, 14) >= 2 And movecpre > 1 Then
                        cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 100
                        '======================
                        For k = 1 To cardAInumuscom
                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                  Case 0
                                       If cardcountAInum(k, 1) = a1a And cardcountAInum(k, 2) = 3 And weryu(1) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(1) = Val(weryu(1)) + 1
                                       End If
                                       If cardcountAInum(k, 1) = a5a And cardcountAInum(k, 2) = 3 And weryu(2) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(2) = Val(weryu(2)) + 1
                                       End If
                                  Case 1
                                       If cardcountAInum(k, 3) = a1a And cardcountAInum(k, 4) = 3 And weryu(1) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(1) = Val(weryu(1)) + 1
                                       End If
                                       If cardcountAInum(k, 3) = a5a And cardcountAInum(k, 4) = 3 And weryu(2) < 1 Then
                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                           weryu(2) = Val(weryu(2)) + 1
                                       End If
                             End Select
                        Next
                        For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                             For k = 1 To cardAInumuscom
                                 Select Case Mid(cardAInumnm(i - 1), k, 1)
                                       Case 0
                                            If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(3)) < 2 Then
                                                cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                weryu(3) = Val(weryu(3)) + cardcountAInum(k, 2)
                                            End If
                                       Case 1
                                            If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(3)) < 2 Then
                                                cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                weryu(3) = Val(weryu(3)) + cardcountAInum(k, 4)
                                            End If
                                  End Select
                             Next
                        Next
                    End If
            End If
         Next
End Select

End Sub
Sub �����i(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�t���¥�
                 If Pn3 = 0 Then
                        werp = 0
                        Erase weryu
                        If cardAInumcaseperson(i, 1, 14) >= 3 Then
                            cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                            '======================
                            For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                 For k = 1 To cardAInumuscom
                                     Select Case Mid(cardAInumnm(i - 1), k, 1)
                                           Case 0
                                                If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                End If
                                           Case 1
                                                If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 3 Then
                                                    cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                    weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                End If
                                      End Select
                                 Next
                            Next
                        End If
                End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================���y����
             If Pn1 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 11) >= 5 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a1a And Val(weryu(1)) < 5 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a1a And Val(weryu(1)) < 5 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================�զʦX
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                    For k = 1 To cardAInumuscom
                          If cardAInumcaseperson(i, 2, k) > 0 Then
                              werp = Val(werp) + 1
                          End If
                    Next
                     If Val(werp) >= 2 And movecpre < 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                     End If
             End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================���٤Ѩ�
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 5 Then
                         Select Case uscom
                                Case 1
                                     For k = 2 To 3
                                            For p = 14 * (����ݾ��H��������(1, k) - 1) + 1 To 14 * ����ݾ��H��������(1, k)
                                               If �H�����`���A��Ʈw(1, p, 3) = 35 Then
                                                   werp = 1
                                               End If
                                            Next
                                     Next
                                Case 2
                                     For k = 2 To 3
                                            For p = 14 * (����ݾ��H��������(2, k) - 1) + 1 To 14 * ����ݾ��H��������(2, k)
                                                If �H�����`���A��Ʈw(2, p, 3) = 36 Then
                                                    werp = 1
                                                End If
                                            Next
                                    Next
                         End Select
                         Select Case uscom
                             Case 1
                                    weryu(1) = liveus(����ݾ��H��������(1, 2))
                                    weryu(2) = liveus(����ݾ��H��������(1, 3))
                                    weryu(3) = liveus41(����ݾ��H��������(1, 2))
                                    weryu(4) = liveus41(����ݾ��H��������(1, 3))
                             Case 2
                                    weryu(1) = livecom(����ݾ��H��������(2, 2))
                                    weryu(2) = livecom(����ݾ��H��������(2, 3))
                                    weryu(3) = livecom41(����ݾ��H��������(2, 2))
                                    weryu(4) = livecom41(����ݾ��H��������(2, 3))
                        End Select
                         If (weryu(1) <= weryu(3) And weryu(1) > 0 And weryu(2) <= weryu(4) And weryu(2) > 0) Or _
                             ((����ʧ@_�ˬd�O�_�����w���`���A(uscom, 37) = False And uscom = 1 And weryu(1) = 0 And weryu(2) = 0) Or _
                             (����ʧ@_�ˬd�O�_�����w���`���A(uscom, 38) = False And uscom = 2 And weryu(1) = 0 And weryu(2) = 0)) Or _
                             (werp = 0 And (weryu(1) > 0 Or weryu(2) > 0)) Then
                                cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 20
                                '======================
                                  For p = Val(cardAInumcaseperson(i, 1, 8)) To Val(cardAInumcaseperson(i, 1, 7)) Step -1
                                        For k = 1 To cardAInumuscom
                                            Select Case Mid(cardAInumnm(i - 1), k, 1)
                                                  Case 0
                                                       If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = p And Val(weryu(1)) < 5 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                       End If
                                                  Case 1
                                                       If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = p And Val(weryu(1)) < 5 Then
                                                           cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                           weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                       End If
                                             End Select
                                        Next
                                   Next
                         End If
                     End If
            End If
         Next
End Select

End Sub
Sub �ײ��d(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer, ByVal Pn1 As Integer, ByVal Pn2 As Integer, ByVal Pn3 As Integer, ByVal Pn4 As Integer)
Dim wnm As Integer, sq As Integer, werp As Integer, livewer As Integer, weryu(1 To 5) As Integer '�Ȯ��ܼ�
If uscom = 1 Then livewer = liveus(����H����ԤH��(1, 2)) Else livewer = livecom(����H����ԤH��(2, 2))
Select Case turn
    Case 1 '==�������q��
          For i = 1 To 2 ^ cardAInumuscom
                 '=================�l���K��
                 If Pn1 = 0 Then
                        werp = 0
                        Erase weryu
                        If movecpre < 3 Then
                               If cardAInumcaseperson(i, 1, 11) >= 2 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                   cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                   '======================
                                   For k = 1 To cardAInumuscom
                                       Select Case Mid(cardAInumnm(i - 1), k, 1)
                                             Case 0
                                                  If cardcountAInum(k, 1) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                  End If
                                                  If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                  End If
                                             Case 1
                                                  If cardcountAInum(k, 3) = a1a And weryu(1) < 2 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                  End If
                                                  If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                      cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                      weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                  End If
                                        End Select
                                   Next
                               End If
                        End If
                 End If
        Next
    Case 2 '==���m���q��
        For i = 1 To 2 ^ cardAInumuscom
            '=================�������H��
             If Pn2 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 12) >= 3 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                         werp = 0
                         For k = 1 To cardAInumuscom
                             Select Case Mid(cardAInumnm(i - 1), k, 1)
                                   Case 0
                                        If cardcountAInum(k, 1) = a2a And Val(weryu(1)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                        End If
                                   Case 1
                                        If cardcountAInum(k, 3) = a2a And Val(weryu(1)) < 3 Then
                                            cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                            weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                        End If
                              End Select
                         Next
                     End If
             End If
             '=================���c���w��
                If Pn3 = 0 Then
                       werp = 0
                       Erase weryu
                       If movecpre = 3 Then
                              If cardAInumcaseperson(i, 1, 12) >= 3 And cardAInumcaseperson(i, 1, 14) >= 1 Then
                                  cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                                  '======================
                                  For k = 1 To cardAInumuscom
                                      Select Case Mid(cardAInumnm(i - 1), k, 1)
                                            Case 0
                                                 If cardcountAInum(k, 1) = a2a And weryu(1) < 3 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(1) = Val(weryu(1)) + cardcountAInum(k, 2)
                                                 End If
                                                 If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(2) = Val(weryu(2)) + cardcountAInum(k, 2)
                                                 End If
                                            Case 1
                                                 If cardcountAInum(k, 3) = a2a And weryu(1) < 3 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(1) = Val(weryu(1)) + cardcountAInum(k, 4)
                                                 End If
                                                 If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = cardAInumcaseperson(i, 1, 7) And weryu(2) < 1 Then
                                                     cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                     weryu(2) = Val(weryu(2)) + cardcountAInum(k, 4)
                                                 End If
                                       End Select
                                  Next
                              End If
                       End If
                End If
        Next
    Case 3 '==���ʶ��q��
        For i = 1 To 2 ^ cardAInumuscom
             '=================�W��
             If Pn4 = 0 Then
                    werp = 0
                    Erase weryu
                     If cardAInumcaseperson(i, 1, 14) >= 1 Then
                         cardAInumFinal(i, 1) = cardAInumFinal(i, 1) + 10
                         '======================
                           For k = 1 To cardAInumuscom
                                  Select Case Mid(cardAInumnm(i - 1), k, 1)
                                        Case 0
                                             If cardcountAInum(k, 1) = a4a And cardcountAInum(k, 2) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                        Case 1
                                             If cardcountAInum(k, 3) = a4a And cardcountAInum(k, 4) = Val(cardAInumcaseperson(i, 1, 7)) Then
                                                 cardAInumcaseperson(i, 2, k) = cardAInumcaseperson(i, 2, k) + 10
                                                 Exit For
                                             End If
                                   End Select
                              Next
                     End If
            End If
         Next
End Select

End Sub

