Attribute VB_Name = "AI�H��"
Option Explicit
Sub ���H���q��(ByVal n As Integer)
Dim ay As Integer, i As Integer, j As Integer, cspce As String, cspme As String
Select Case n
    Case 1
        '================���`���A-MOV��-���Ĳ��ʭȧP�_�B�z
        For i = 14 * (����H����ԤH��(2, 2) - 1) + 1 To 14 * ����H����ԤH��(2, 2)
            If �H�����`���A��Ʈw(2, i, 3) = 6 Then
                ay = ay + �H�����`���A��Ʈw(2, i, 1)
            End If
            If �H�����`���A��Ʈw(2, i, 3) = 17 Then
                ay = 99
                Exit For
            End If
        Next
        If �ثe��(25) <= Val(ay) Then
            For i = 1 To 106
               If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 1 Then
                  pagecardnum(i, 11) = 0
               End If
            Next
        End If
    Case 2
        '================�b�̫��P���q�P�_�B�z
        For j = 1 To 106
           If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
             pagecardnum(j, 11) = 1
             If FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" _
                Or FormMainMode.compi1(����H����ԤH��(2, 2)) = "���H" _
                Or FormMainMode.compi1(����H����ԤH��(2, 2)) = "������" Then
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
             ElseIf FormMainMode.compi1(����H����ԤH��(2, 2)) = "������S" _
                 Or FormMainMode.compi1(����H����ԤH��(2, 2)) = "C.C." _
                 Or FormMainMode.compi1(����H����ԤH��(2, 2)) = "���[" _
                 Or FormMainMode.compi1(����H����ԤH��(2, 2)) = "�S�{��" _
                 Or FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Ǧh" Then
                       If pagecardnum(j, 1) = a4a Then
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
             Else
                  If pagecardnum(j, 3) = a4a Then
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
             If pagecardnum(j, 3) = a3a And pagecardnum(j, 4) = 1 Then '�ಾ�ʵP(����)
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
End Select
End Sub
Sub ��̬d�w(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "��̬d�w" Then
    Select Case n
        Case 1
            '===========�̾ڶZ���X�P
            Select Case movecp
                Case 1
                    For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a Then
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
                Case Is > 1
                      For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a1a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a1a And Val(pagecardnum(j, 2)) >= 2 Then
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
            End Select
        Case 2
        
    End Select
End If
End Sub
Sub ��B�����S(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
Dim aw As Integer
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "��B�����S" Then
    Select Case n
        Case 1
            If movecp = 1 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 2 Then
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
                                 Exit For
                            End If
                        End If
                 Next
            End If
        Case 2
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) = 1 Then
                       aw = Val(aw) + 1
                   End If
            Next
            If aw = 2 Then
                For j = 1 To 106
                       If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                               pagecardnum(j, 11) = 1
                               If pagecardnum(j, 3) = a3a Then
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
                                Exit For
                           End If
                       End If
                Next
            ElseIf aw < 2 Then
                aw = 0
                For j = 1 To 106
                    If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                        If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                            aw = Val(aw) + 1
                        End If
                    End If
                Next
                If Val(aw) >= 3 Then
                    aw = 0
                    For j = 1 To 106
                       If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                           If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                               pagecardnum(j, 11) = 1
                               aw = Val(aw) + 1
                               If pagecardnum(j, 3) = a3a Then
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
                       End If
                       If Val(aw) >= 3 Then Exit For
                    Next
                End If
            End If
    End Select
End If
End Sub
Sub �v��L(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String, i As Integer
Dim aw(1 To 2) As Integer
Dim ae As Integer
Dim num(1 To 2, 1 To 2) As Integer '��ܤH���Ȯ��ܼ�
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "�v��L" Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If movecp = 3 And Val(aw(2)) >= 9 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a1a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a1a Then
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
            End If
        Case 2
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
               aw(1) = 0
           Else
               aw(1) = 1
           End If
            '===============================
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a1a Then
                           aw(2) = Val(aw(2)) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If aw(1) = 1 And Val(aw(2)) >= 9 Then
                For j = 1 To 106
                       If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a3a And pagecardnum(j, 3) <> a1a Then
                                pagecardnum(j, 11) = 1
                             ElseIf pagecardnum(j, 3) = a3a And pagecardnum(j, 1) <> a1a Then
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
            Else
               For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) = 1 Then
                       ae = Val(ae) + 1
                   End If
                Next
                If ae = 2 Then
                    For j = 1 To 106
                           If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                               If Val(pagecardnum(j, 2)) <= 2 And Val(pagecardnum(j, 4)) <= 2 Then
                                   pagecardnum(j, 11) = 1
                                   If pagecardnum(j, 3) = a3a Then
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
                                    Exit For
                               End If
                           End If
                    Next
                End If
            End If
    End Select
End If
End Sub
Sub CC(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
Dim aw As Integer
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "C.C." Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If movecp = 1 And Val(aw) >= 6 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a Then
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
            End If
        Case 2
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                        If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) = 1 And Val(pagecardnum(j, 4)) <= 3 Then
                            pagecardnum(j, 11) = 1
                            Exit For
                        ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) = 1 And Val(pagecardnum(j, 2)) <= 3 Then
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
                             Exit For
                        End If
                   End If
            Next
    End Select
End If
End Sub
Sub ����(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "����" Then
    Select Case n
        Case 1
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 2 Then
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
                                 Exit For
                            End If
                        End If
                 Next
    End Select
End If
End Sub
Sub �Q��(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "�Q��" Then
    Select Case n
        Case 1
            If movecp = 1 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 3 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 3 Then
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
                                 Exit For
                            End If
                        End If
                 Next
            End If
    End Select
End If
End Sub
Sub �L���S(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "�L���S" Then
    Select Case n
        Case 1
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 2 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 2 Then
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
                                 Exit For
                            End If
                        End If
                 Next
    End Select
End If
End Sub
Sub �w�ǥ���(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
Dim aw As Integer
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "�w�ǥ���" Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 Then
                       If pagecardnum(j, 1) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a4a Then
                           aw = Val(aw) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If Val(aw) >= 3 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a Then
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
            End If
    End Select
End If
End Sub
Sub ����P��(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "����P��" Then
    Select Case n
        Case 1
            If movecp = 1 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 3 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 3 Then
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
                                 Exit For
                            End If
                        End If
                 Next
              Else
                 For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a Then
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
                                 Exit For
                            End If
                        End If
                 Next
             End If
    End Select
End If
End Sub
Sub �h�g�H(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
Dim aw As Integer
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "�h�g�H" Then
    Select Case n
        Case 1
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) <> 3 Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) <> 3 Then
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
    End Select
End If
End Sub
Sub ���_�i���h(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "���_�i���h" Then
    Select Case n
        Case 1
            If movecp > 1 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 3 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 3 Then
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
                                 Exit For
                            End If
                        End If
                 Next
              End If
          Case 2
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 1 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 1 Then
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
                                 Exit For
                            End If
                        End If
                 Next
    End Select
End If
End Sub
Sub ������S(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "������S" Then
    Select Case n
        Case 1
             If FormMainMode.comaiatk(1).Caption = "���" And movecp < 3 Then
                For j = 106 To 1 Step -1
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a4a And Val(pagecardnum(j, 2)) >= 1 Then
                                pagecardnum(j, 11) = 1
                                Exit For
                            ElseIf pagecardnum(j, 3) = a4a And Val(pagecardnum(j, 4)) >= 1 Then
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
                                 Exit For
                            End If
                        End If
                 Next
             End If
    End Select
End If
End Sub
Sub �����i(ByVal n As Integer)
Dim j As Integer, cspce As String, cspme As String
Dim aw As Integer
If FormMainMode.compi1(����H����ԤH��(2, 2)) = "�����i" Then
    Select Case n
        Case 1
            For j = 1 To 106
                   If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                       If pagecardnum(j, 1) = a1a Then
                           aw = Val(aw) + Val(pagecardnum(j, 2))
                       ElseIf pagecardnum(j, 3) = a1a Then
                           aw = Val(aw) + Val(pagecardnum(j, 4))
                       End If
                   End If
            Next
            If movecp > 1 And Val(aw) >= 5 Then
                For j = 1 To 106
                        If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
                            If pagecardnum(j, 1) = a1a Then
                                pagecardnum(j, 11) = 1
                            ElseIf pagecardnum(j, 3) = a1a Then
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
            End If
    End Select
End If
End Sub

