Attribute VB_Name = "�H���t����"
Option Explicit
Public totpersonnumber As Integer '�{�b�ثe�B�z�ĴX�H�Ȯɼ�
Public �`�@�H���W�� As String '�ثe�`�@Ū�J�H���W��
Public �`�@�H���ɮצW As String '�ثe�`�@Ū�J�H���ɮצW
Public ���ϥΪ̨ƥ� As Boolean '������O�_���ϥΪ̨ƥ�Ȯɼ�
Public ���q���ƥ� As Boolean '������O�_���q���ƥ�Ȯɼ�
Public VBEPerson(1 To 2, 1 To 3, 1 To 4, 1 To 30, 1 To 10) As String 'VBE�H���Τ@��ưO���ܼ�
Dim app_path As String  '���|�]�w�X
Sub �d���H����TŪ�J_�춥�q(ByVal filename As String)
Dim textlinea As String 'Ū�J���ɤ@��B�z�Ȯ��ܼ�
Dim ���r��() As String
Dim textcheck As Boolean '�P�_����ˬd�X���T���ܼ�
'MsgBox filename
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   ���r�� = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "XzPF5922" Then
'           MsgBox "Ū�J�ɮ׮ɵo�Ϳ��~!"
           �d���H����T�ɮ�Ū�����Ѭ����� = �d���H����T�ɮ�Ū�����Ѭ����� & "=" & filename
           Exit Do
       Else
           textcheck = True
           �[�J�`�@�H���ɮצW�r�� filename
       End If
   End If
   If textlinea = "" Then
       GoTo ���L����r��
   End If
   Select Case ���r��(0)
       Case "MenuList"
           
       Case "MenuName"
           �[�J�`�@�H���W�٦r�� ���r��(1)
           ��s�H���M��_�ϥΪ̤�_��]
           ��s�H���M��_�q����_��]
       Case "EndFirst"
           Exit Do
   End Select
���L����r��:
Loop
Close
End Sub
Sub �d���H����TŪ�J_�G���q_�ϥΪ�(ByVal personname As String, ByVal Index As Integer)
Dim textlinea As String 'Ū�J���ɤ@��B�z�Ȯ��ܼ�
Dim ���r��() As String
Dim textcheck As Boolean '�P�_����ˬd�X���T���ܼ�
Dim filename As String '�ؼФH���ɮצW�Ȯɼ�
Dim at() As String, aw() As String
Dim i As Integer
at = Split(�`�@�H���W��, "=")
aw = Split(�`�@�H���ɮצW, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
FormMainMode.personlevelus(Index).Clear
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   ���r�� = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "XzPF5922" Then
           MsgBox "Ū�J�ɮ׮ɵo�Ϳ��~!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo ���L����r��
   End If
   Select Case ���r��(0)
       Case "MenuList"
           For i = 1 To UBound(���r��)
               FormMainMode.personlevelus(Index).AddItem ���r��(i)
           Next
       Case "EndFirst"
           Exit Do
   End Select
���L����r��:
Loop
Close
End Sub
Sub �d���H����TŪ�J_�G���q_�q��(ByVal personname As String, ByVal Index As Integer)
Dim textlinea As String 'Ū�J���ɤ@��B�z�Ȯ��ܼ�
Dim ���r��() As String
Dim textcheck As Boolean '�P�_����ˬd�X���T���ܼ�
Dim filename As String '�ؼФH���ɮצW�Ȯɼ�
Dim at() As String, aw() As String
Dim i As Integer
at = Split(�`�@�H���W��, "=")
aw = Split(�`�@�H���ɮצW, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
FormMainMode.personlevelcom(Index).Clear
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   ���r�� = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "XzPF5922" Then
           MsgBox "Ū�J�ɮ׮ɵo�Ϳ��~!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo ���L����r��
   End If
   Select Case ���r��(0)
       Case "MenuList"
           For i = 1 To UBound(���r��)
               FormMainMode.personlevelcom(Index).AddItem ���r��(i)
           Next
       Case "EndFirst"
           Exit Do
   End Select
���L����r��:
Loop
Close
End Sub
Sub �d���H����TŪ�J_�T���q_�ϥΪ�(ByVal personname As String, ByVal personlevel As String, ByVal Index As Integer, ByVal uscom As Integer)
Dim textlinea As String 'Ū�J���ɤ@��B�z�Ȯ��ܼ�
Dim ���r��() As String
Dim textcheck As Boolean '�P�_����ˬd�X���T���ܼ�
Dim filename As String '�ؼФH���ɮצW�Ȯɼ�
Dim ���L��T As Boolean '�O�_���L�ثe�Ϭq�Ȯɼ�
Dim at() As String, aw() As String
Dim i As Integer
at = Split(�`�@�H���W��, "=")
aw = Split(�`�@�H���ɮצW, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
'======================
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   ���r�� = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "XzPF5922" Then
           MsgBox "Ū�J�ɮ׮ɵo�Ϳ��~!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo ���L����r��
   End If
   If ���L��T = False Then
        Select Case ���r��(0)
            Case "StartPerson"
               If ���r��(1) <> personlevel Or ���r��(2) <> personname Or ���r��(3) <> uscom Then
                   ���L��T = True
               End If
            Case "cardjpgus"
               VBEPerson(1, Index, 1, 5, 5) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personhp"
               VBEPerson(1, Index, 1, 3, 1) = ���r��(1)
            Case "personatk"
               VBEPerson(1, Index, 1, 3, 2) = ���r��(1)
            Case "persondef"
               VBEPerson(1, Index, 1, 3, 3) = ���r��(1)
            Case "personname"
               VBEPerson(1, Index, 1, 1, 1) = ���r��(1)
            Case "personengname"
               VBEPerson(1, Index, 1, 1, 2) = ���r��(1)
            Case "personpname"
               VBEPerson(1, Index, 1, 1, 3) = ���r��(1)
            Case "personlevel1"
               VBEPerson(1, Index, 1, 2, 1) = ���r��(1)
            Case "personlevel2"
               VBEPerson(1, Index, 1, 2, 2) = ���r��(1)
            Case "cardid"
               VBEPerson(1, Index, 1, 4, 1) = ���r��(1)
            Case "persontg"
               VBEPerson(1, Index, 1, 3, 4) = ���r��(1)
            Case "personbig"
               VBEPerson(1, Index, 1, 5, 3) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personmini"
               VBEPerson(1, Index, 1, 5, 1) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personf"
               VBEPerson(1, Index, 1, 5, 4) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personsmalldown"
               VBEPerson(1, Index, 1, 5, 2) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personfleftall"
               VBEPerson(1, Index, 2, 4, 1) = ���r��(1)
            Case "atkingfontck"
               VBEPerson(1, Index, 2, 3, 5) = ���r��(1)
            Case "bight"
               VBEPerson(1, Index, 2, 2, 1) = ���r��(1)
            Case "bigtop"
               VBEPerson(1, Index, 2, 2, 3) = ���r��(1)
            Case "bigwh"
               VBEPerson(1, Index, 2, 2, 2) = ���r��(1)
            Case "minileft1"
               VBEPerson(1, Index, 2, 1, 1) = ���r��(1)
            Case "minileft2"
               VBEPerson(1, Index, 2, 1, 2) = ���r��(1)
            Case "minileft3"
               VBEPerson(1, Index, 2, 1, 3) = ���r��(1)
            Case "minitop"
               VBEPerson(1, Index, 2, 1, 4) = ���r��(1)
            Case "atkingjpgleftallzero"
               VBEPerson(1, Index, 2, 2, 5) = ���r��(1)
            Case "bigleftall"
               VBEPerson(1, Index, 2, 2, 4) = ���r��(1)
            Case "smalldownleft"
               VBEPerson(1, Index, 2, 1, 5) = ���r��(1)
            Case "smalldowntop"
               VBEPerson(1, Index, 2, 1, 6) = ���r��(1)
            Case "atkingfont1"
               VBEPerson(1, Index, 2, 3, 1) = ���r��(1)
            Case "atkingfont2"
               VBEPerson(1, Index, 2, 3, 2) = ���r��(1)
            Case "atkingfont3"
               VBEPerson(1, Index, 2, 3, 3) = ���r��(1)
            Case "atkingfont4"
               VBEPerson(1, Index, 2, 3, 4) = ���r��(1)
            Case "atkingcfont(1)"
               VBEPerson(1, Index, 3, 1, 6) = ���r��(1)
            Case "atkingcfont(2)"
               VBEPerson(1, Index, 3, 2, 6) = ���r��(1)
            Case "atkingcfont(3)"
               VBEPerson(1, Index, 3, 3, 6) = ���r��(1)
            Case "atkingcfont(4)"
               VBEPerson(1, Index, 3, 4, 6) = ���r��(1)
            Case "atkingdfont(1)"
               VBEPerson(1, Index, 3, 1, 7) = ���r��(1)
            Case "atkingdfont(2)"
               VBEPerson(1, Index, 3, 2, 7) = ���r��(1)
            Case "atkingdfont(3)"
               VBEPerson(1, Index, 3, 3, 7) = ���r��(1)
            Case "atkingdfont(4)"
               VBEPerson(1, Index, 3, 4, 7) = ���r��(1)
            Case "atkingname(1)"
               VBEPerson(1, Index, 3, 1, 1) = ���r��(1)
            Case "atkingname(2)"
               VBEPerson(1, Index, 3, 2, 1) = ���r��(1)
            Case "atkingname(3)"
               VBEPerson(1, Index, 3, 3, 1) = ���r��(1)
            Case "atkingname(4)"
               VBEPerson(1, Index, 3, 4, 1) = ���r��(1)
            Case "atkinga(1)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(1)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(1)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(1)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
            Case "atkinga(2)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(2)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(2)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(2)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
            Case "atkinga(3)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(3)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(3)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(3)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
            Case "atkinga(4)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(4)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(4)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(4)"
               VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(1, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
        End Select
   End If
   If ���r��(0) = "EndPerson" Then
       ���L��T = False
   End If
���L����r��:
Loop
Close
End Sub
Sub �d���H����TŪ�J_�T���q_�q��(ByVal personname As String, ByVal personlevel As String, ByVal Index As Integer, ByVal uscom As Integer)
Dim textlinea As String 'Ū�J���ɤ@��B�z�Ȯ��ܼ�
Dim ���r��() As String
Dim textcheck As Boolean '�P�_����ˬd�X���T���ܼ�
Dim filename As String '�ؼФH���ɮצW�Ȯɼ�
Dim ���L��T As Boolean '�O�_���L�ثe�Ϭq�Ȯɼ�
Dim at() As String, aw() As String
Dim i As Integer
at = Split(�`�@�H���W��, "=")
aw = Split(�`�@�H���ɮצW, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
'======================
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   ���r�� = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "XzPF5922" Then
           MsgBox "Ū�J�ɮ׮ɵo�Ϳ��~!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo ���L����r��
   End If
   If ���L��T = False Then
        Select Case ���r��(0)
            Case "StartPerson"
               If ���r��(1) <> personlevel Or ���r��(2) <> personname Or ���r��(3) <> uscom Then
                   ���L��T = True
               End If
            Case "cardjpgcom"
               VBEPerson(2, Index, 1, 5, 5) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personhp"
               VBEPerson(2, Index, 1, 3, 1) = ���r��(1)
            Case "personatk"
               VBEPerson(2, Index, 1, 3, 2) = ���r��(1)
            Case "persondef"
               VBEPerson(2, Index, 1, 3, 3) = ���r��(1)
            Case "personname"
               VBEPerson(2, Index, 1, 1, 1) = ���r��(1)
            Case "personengname"
               VBEPerson(2, Index, 1, 1, 2) = ���r��(1)
            Case "personpname"
               VBEPerson(2, Index, 1, 1, 3) = ���r��(1)
            Case "personlevel1"
               VBEPerson(2, Index, 1, 2, 1) = ���r��(1)
            Case "personlevel2"
               VBEPerson(2, Index, 1, 2, 2) = ���r��(1)
            Case "cardid"
               VBEPerson(2, Index, 1, 4, 1) = ���r��(1)
            Case "persontg"
               VBEPerson(2, Index, 1, 3, 4) = ���r��(1)
            Case "personcomai"
               VBEPerson(2, Index, 1, 4, 3) = ���r��(1)
            Case "personbig"
               VBEPerson(2, Index, 1, 5, 3) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personmini"
               VBEPerson(2, Index, 1, 5, 1) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personf"
               VBEPerson(2, Index, 1, 5, 4) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personsmalldown"
               VBEPerson(2, Index, 1, 5, 2) = app_path & "gif\" & personname & "\" & ���r��(1)
            Case "personfleftall"
               VBEPerson(2, Index, 2, 4, 1) = ���r��(1)
            Case "atkingfontck"
               VBEPerson(2, Index, 2, 3, 5) = ���r��(1)
            Case "bight"
               VBEPerson(2, Index, 2, 2, 1) = ���r��(1)
            Case "bigtop"
               VBEPerson(2, Index, 2, 2, 3) = ���r��(1)
            Case "bigwh"
               VBEPerson(2, Index, 2, 2, 2) = ���r��(1)
            Case "minileft1"
               VBEPerson(2, Index, 2, 1, 1) = ���r��(1)
            Case "minileft2"
               VBEPerson(2, Index, 2, 1, 2) = ���r��(1)
            Case "minileft3"
               VBEPerson(2, Index, 2, 1, 3) = ���r��(1)
            Case "minitop"
               VBEPerson(2, Index, 2, 1, 4) = ���r��(1)
            Case "atkingjpgleftallzero"
               VBEPerson(2, Index, 2, 2, 5) = ���r��(1)
            Case "bigleftall"
               VBEPerson(2, Index, 2, 2, 4) = ���r��(1)
            Case "smalldownleft"
               VBEPerson(2, Index, 2, 1, 5) = ���r��(1)
            Case "smalldowntop"
               VBEPerson(2, Index, 2, 1, 6) = ���r��(1)
            Case "atkingfont1"
               VBEPerson(2, Index, 2, 3, 1) = ���r��(1)
            Case "atkingfont2"
               VBEPerson(2, Index, 2, 3, 2) = ���r��(1)
            Case "atkingfont3"
               VBEPerson(2, Index, 2, 3, 3) = ���r��(1)
            Case "atkingfont4"
               VBEPerson(2, Index, 2, 3, 4) = ���r��(1)
            Case "atkingcfont(1)"
               VBEPerson(2, Index, 3, 1, 6) = ���r��(1)
            Case "atkingcfont(2)"
               VBEPerson(2, Index, 3, 2, 6) = ���r��(1)
            Case "atkingcfont(3)"
               VBEPerson(2, Index, 3, 3, 6) = ���r��(1)
            Case "atkingcfont(4)"
               VBEPerson(2, Index, 3, 4, 6) = ���r��(1)
            Case "atkingdfont(1)"
               VBEPerson(2, Index, 3, 1, 7) = ���r��(1)
            Case "atkingdfont(2)"
               VBEPerson(2, Index, 3, 2, 7) = ���r��(1)
            Case "atkingdfont(3)"
               VBEPerson(2, Index, 3, 3, 7) = ���r��(1)
            Case "atkingdfont(4)"
               VBEPerson(2, Index, 3, 4, 7) = ���r��(1)
            Case "atkingname(1)"
               VBEPerson(2, Index, 3, 1, 1) = ���r��(1)
            Case "atkingname(2)"
               VBEPerson(2, Index, 3, 2, 1) = ���r��(1)
            Case "atkingname(3)"
               VBEPerson(2, Index, 3, 3, 1) = ���r��(1)
            Case "atkingname(4)"
               VBEPerson(2, Index, 3, 4, 1) = ���r��(1)
            Case "atkinga(1)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(1)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(1)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(1)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
            Case "atkinga(2)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(2)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(2)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(2)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
            Case "atkinga(3)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(3)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(3)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(3)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
            Case "atkinga(4)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 2) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 8) = ���r��(2)
            Case "atkingb(4)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 3) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 9) = ���r��(2)
            Case "atkingc(4)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 4) = ���r��(1)
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 10) = ���r��(2)
            Case "atkingd(4)"
               VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = ""
               For i = 1 To UBound(���r��)
                    VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) = VBEPerson(2, Index, 3, Val(Mid(���r��(0), 9, 1)), 5) & ���r��(i)
               Next
        End Select
   End If
   If ���r��(0) = "EndPerson" Then
       ���L��T = False
   End If
���L����r��:
Loop
Close
End Sub
Sub �d���H����TŪ�J_�|���q_�ϥΪ�(ByVal personname As String, ByVal Index As Integer)
Dim textlinea As String 'Ū�J���ɤ@��B�z�Ȯ��ܼ�
Dim ���r��() As String
Dim textcheck As Boolean '�P�_����ˬd�X���T���ܼ�
Dim filename As String '�ؼФH���ɮצW�Ȯɼ�
Dim ���L��T As Boolean '�O�_���L�ثe�Ϭq�Ȯɼ�
Dim persontalka As Integer '��Ƽg�J�Ȯ��ܼ�
Dim at() As String, aw() As String
Dim i As Integer
at = Split(�`�@�H���W��, "=")
aw = Split(�`�@�H���ɮצW, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
'======================
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   ���r�� = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "XzPF5922" Then
           MsgBox "Ū�J�ɮ׮ɵo�Ϳ��~!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
        GoTo ���L����r��
   End If
   If ���L��T = False Then
        If Left(���r��(0), 4) = "Talk" Then
           If ���r��(1) = "" Then
                GoTo ���L����r��
           End If
        End If
        '=====================
        Select Case ���r��(0)
           Case "StartTalk"
               If ���r��(1) <> personname Then
                   ���L��T = True
               End If
            Case "TalkA1"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA2"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA3"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA4"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA5"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA6"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA7"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA8"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA9"
               persontalka = Right(���r��(0), 1)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA10"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA11"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA12"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA13"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA14"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA15"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA16"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA17"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA18"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA19"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA20"
               persontalka = Right(���r��(0), 2)
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(1, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkB1"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB2"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB3"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB4"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB5"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB6"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB7"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB8"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB9"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB10"
               persontalka = Val(Right(���r��(0), 1)) + 20
               VBEPerson(1, Index, 4, persontalka, 1) = ���r��(1)
        End Select
   End If
   If ���r��(0) = "EndTalk" Then
       ���L��T = False
   End If
���L����r��:
Loop
Close
End Sub
Sub �d���H����TŪ�J_�|���q_�q��(ByVal personname As String, ByVal Index As Integer)
Dim textlinea As String 'Ū�J���ɤ@��B�z�Ȯ��ܼ�
Dim ���r��() As String
Dim textcheck As Boolean '�P�_����ˬd�X���T���ܼ�
Dim filename As String '�ؼФH���ɮצW�Ȯɼ�
Dim ���L��T As Boolean '�O�_���L�ثe�Ϭq�Ȯɼ�
Dim persontalka As Integer '��Ƽg�J�Ȯ��ܼ�
Dim at() As String, aw() As String
Dim i As Integer
at = Split(�`�@�H���W��, "=")
aw = Split(�`�@�H���ɮצW, "=")
For i = 0 To UBound(at)
   If at(i) = personname Then
       filename = aw(i)
   End If
Next
'======================
app_path = App.Path
If Right$(app_path, 1) <> "\" Then app_path = app_path & "\"
'======================
Open filename For Input As #1
Do Until EOF(1)
   Line Input #1, textlinea
   ���r�� = Split(textlinea, "=")
   If textcheck = False Then
       If textlinea <> "XzPF5922" Then
           MsgBox "Ū�J�ɮ׮ɵo�Ϳ��~!"
           Exit Do
       Else
           textcheck = True
       End If
   End If
   If textlinea = "" Then
       GoTo ���L����r��
   End If
   If ���L��T = False Then
        Select Case ���r��(0)
           Case "StartTalk"
               If ���r��(1) <> personname Then
                   ���L��T = True
               End If
            Case "TalkA1"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA2"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA3"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA4"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA5"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA6"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA7"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA8"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA9"
               persontalka = Right(���r��(0), 1)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkA10"
               persontalka = Right(���r��(0), 2)
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
               VBEPerson(2, Index, 4, persontalka, 2) = ���r��(2)
            Case "TalkB1"
               persontalka = Val(Right(���r��(0), 1)) + 10
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB2"
               persontalka = Val(Right(���r��(0), 1)) + 10
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB3"
               persontalka = Val(Right(���r��(0), 1)) + 10
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB4"
               persontalka = Val(Right(���r��(0), 1)) + 10
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
            Case "TalkB5"
               persontalka = Val(Right(���r��(0), 1)) + 10
               VBEPerson(2, Index, 4, persontalka, 1) = ���r��(1)
        End Select
   End If
   If ���r��(0) = "EndTalk" Then
       ���L��T = False
   End If
���L����r��:
Loop
Close
End Sub
Sub �[�J�`�@�H���W�٦r��(ByVal name As String)
�`�@�H���W�� = �`�@�H���W�� & "=" & name
End Sub
Sub �[�J�`�@�H���ɮצW�r��(ByVal name As String)
�`�@�H���ɮצW = �`�@�H���ɮצW & "=" & name
End Sub
Sub ��s�H���M��_�ϥΪ̤�_��]()
Dim at() As String, i As Integer, j As Integer
at = Split(�`�@�H���W��, "=")
For i = 1 To 3
   FormMainMode.personnameus(i).Clear
   FormMainMode.personnameus(i).AddItem "�m�H���n"
   For j = 1 To UBound(at)
       FormMainMode.personnameus(i).AddItem at(j)
   Next
Next
End Sub
Sub ��s�H���M��_�q����_��]()
Dim at() As String, i As Integer, j As Integer
at = Split(�`�@�H���W��, "=")
For i = 1 To 3
   FormMainMode.personnamecom(i).Clear
   FormMainMode.personnamecom(i).AddItem "�m�H���n"
   For j = 1 To UBound(at)
       FormMainMode.personnamecom(i).AddItem at(j)
   Next
Next
End Sub
Sub ��s�H���M��_�ϥΪ̤�_�ܧ�(ByVal �{�b�Ҧb�� As Integer)
Dim at() As String
at = Split(�`�@�H���W��, "=")
Dim ag(1 To 3) As String  '�����ثe�ﶵ�Ȯɼ�
Dim ap As Integer, au As Integer, i As Integer, j As Integer, p As Integer, q As Integer, k As Integer '�Ȯ��ܼ�
For i = 1 To 3
   ag(i) = FormMainMode.personnameus(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnameus(i).Clear
   FormMainMode.personnameus(i).AddItem "�m�H���n"
   For j = 1 To UBound(at)
       FormMainMode.personnameus(i).AddItem at(j)
   Next
Next
 '===========================================
���ϥΪ̨ƥ� = False
'formmainmode.personnameus(�{�b�Ҧb��).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnameus(p).ListCount - 1
           If FormMainMode.personnameus(p).List(q) = ag(p) Then
               FormMainMode.personnameus(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnameus(p).ListIndex = -1
   End If
Next
���ϥΪ̨ƥ� = True
'========================================
For i = 1 To 3
   ap = FormMainMode.personnameus(i).ListCount - 1
   au = 0
        Do Until au > ap
            If FormMainMode.personnameus(i).List(au) <> "�m�H���n" Then
                Select Case i
                    Case 1
                       If FormMainMode.personnameus(2).Text = FormMainMode.personnameus(i).List(au) Or FormMainMode.personnameus(3).Text = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If FormMainMode.personnameus(1).Text = FormMainMode.personnameus(i).List(au) Or FormMainMode.personnameus(3).Text = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If FormMainMode.personnameus(2).Text = FormMainMode.personnameus(i).List(au) Or FormMainMode.personnameus(1).Text = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
            End If
            au = au + 1
        Loop
Next
'===========�ˬd���O�_�u���u�H���v�@��
For i = 1 To 3
   If FormMainMode.personnameus(i).ListCount = 1 Then
       FormMainMode.personnameus(i).Clear
   End If
Next
���ϥΪ̨ƥ� = False
'formmainmode.personnameus(�{�b�Ҧb��).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnameus(i).ListCount - 1
           If FormMainMode.personnameus(i).List(k) = ag(i) Then
               FormMainMode.personnameus(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnameus(i).ListIndex = -1
   End If
Next
���ϥΪ̨ƥ� = True
End Sub
Sub ��s�H���M��_�ϥΪ̤�_�ܧ�_�}�l�H��(ByVal �{�b�Ҧb�� As Integer, ByVal name1 As String, ByVal name2 As String, ByVal name3 As String)
Dim at() As String
at = Split(�`�@�H���W��, "=")
Dim ag(1 To 3) As String  '�����ثe�ﶵ�Ȯɼ�
Dim ap As Integer, au As Integer, i As Integer, j As Integer, p As Integer, q As Integer, k As Integer '�Ȯ��ܼ�
For i = 1 To 3
   ag(i) = FormMainMode.personnameus(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnameus(i).Clear
   FormMainMode.personnameus(i).AddItem "�m�H���n"
   For j = 1 To UBound(at)
       FormMainMode.personnameus(i).AddItem at(j)
   Next
Next
 '===========================================
���ϥΪ̨ƥ� = False
'formmainmode.personnameus(�{�b�Ҧb��).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnameus(p).ListCount - 1
           If FormMainMode.personnameus(p).List(q) = ag(p) Then
               FormMainMode.personnameus(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnameus(p).ListIndex = -1
   End If
Next
'========================================
For i = 1 To 3
   ap = FormMainMode.personnameus(i).ListCount - 1
   au = 0
        Do Until au > ap
'            If formmainmode.personnameus(i).List(au) <> "�m�H���n" Then
                Select Case i
                    Case 1
                       If name2 = FormMainMode.personnameus(i).List(au) Or name3 = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If name1 = FormMainMode.personnameus(i).List(au) Or name3 = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If name2 = FormMainMode.personnameus(i).List(au) Or name1 = FormMainMode.personnameus(i).List(au) Then
                           FormMainMode.personnameus(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
'            End If
            au = au + 1
        Loop
Next
'===========�ˬd���O�_�u���u�H���v�@��
For i = 1 To 3
   If FormMainMode.personnameus(i).ListCount = 1 Then
       FormMainMode.personnameus(i).Clear
   End If
Next
���ϥΪ̨ƥ� = False
'formmainmode.personnameus(�{�b�Ҧb��).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnameus(i).ListCount - 1
           If FormMainMode.personnameus(i).List(k) = ag(i) Then
               FormMainMode.personnameus(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnameus(i).ListIndex = -1
   End If
Next
End Sub

Sub ��s�H���M��_�q����_�ܧ�(ByVal �{�b�Ҧb�� As Integer)
Dim at() As String
at = Split(�`�@�H���W��, "=")
Dim ag(1 To 3) As String  '�����ثe�ﶵ�Ȯɼ�
Dim ap As Integer, au As Integer, i As Integer, j As Integer, p As Integer, q As Integer, k As Integer '�Ȯ��ܼ�
For i = 1 To 3
   ag(i) = FormMainMode.personnamecom(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnamecom(i).Clear
   FormMainMode.personnamecom(i).AddItem "�m�H���n"
   For j = 1 To UBound(at)
       FormMainMode.personnamecom(i).AddItem at(j)
   Next
Next
 '===========================================
���q���ƥ� = False
'formmainmode.personnamecom(�{�b�Ҧb��).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnamecom(p).ListCount - 1
           If FormMainMode.personnamecom(p).List(q) = ag(p) Then
               FormMainMode.personnamecom(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnamecom(p).ListIndex = -1
   End If
Next
���q���ƥ� = True
'========================================
For i = 1 To 3
   ap = FormMainMode.personnamecom(i).ListCount - 1
   au = 0
        Do Until au > ap
            If FormMainMode.personnamecom(i).List(au) <> "�m�H���n" Then
                Select Case i
                    Case 1
                       If FormMainMode.personnamecom(2).Text = FormMainMode.personnamecom(i).List(au) Or FormMainMode.personnamecom(3).Text = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If FormMainMode.personnamecom(1).Text = FormMainMode.personnamecom(i).List(au) Or FormMainMode.personnamecom(3).Text = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If FormMainMode.personnamecom(2).Text = FormMainMode.personnamecom(i).List(au) Or FormMainMode.personnamecom(1).Text = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
            End If
            au = au + 1
        Loop
Next
'===========�ˬd���O�_�u���u�H���v�@��
For i = 1 To 3
   If FormMainMode.personnamecom(i).ListCount = 1 Then
       FormMainMode.personnamecom(i).Clear
   End If
Next
���q���ƥ� = False
'formmainmode.personnamecom(�{�b�Ҧb��).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnamecom(i).ListCount - 1
           If FormMainMode.personnamecom(i).List(k) = ag(i) Then
               FormMainMode.personnamecom(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnamecom(i).ListIndex = -1
   End If
Next
���q���ƥ� = True
End Sub
Sub ��s�H���M��_�q����_�ܧ�_�}�l�H��(ByVal �{�b�Ҧb�� As Integer, ByVal name1 As String, ByVal name2 As String, ByVal name3 As String)
Dim at() As String
at = Split(�`�@�H���W��, "=")
Dim ag(1 To 3) As String  '�����ثe�ﶵ�Ȯɼ�
Dim ap As Integer, au As Integer, i As Integer, j As Integer, p As Integer, q As Integer, k As Integer '�Ȯ��ܼ�
For i = 1 To 3
   ag(i) = FormMainMode.personnamecom(i).Text
Next
'=====================
For i = 1 To 3
   FormMainMode.personnamecom(i).Clear
   FormMainMode.personnamecom(i).AddItem "�m�H���n"
   For j = 1 To UBound(at)
       FormMainMode.personnamecom(i).AddItem at(j)
   Next
Next
 '===========================================
���q���ƥ� = False
'formmainmode.personnamecom(�{�b�Ҧb��).ListIndex = ag
For p = 1 To 3
   If ag(p) <> "" Then
       For q = 0 To FormMainMode.personnamecom(p).ListCount - 1
           If FormMainMode.personnamecom(p).List(q) = ag(p) Then
               FormMainMode.personnamecom(p).ListIndex = q
           End If
       Next
   Else
       FormMainMode.personnamecom(p).ListIndex = -1
   End If
Next
'========================================
For i = 1 To 3
   ap = FormMainMode.personnamecom(i).ListCount - 1
   au = 0
        Do Until au > ap
'            If formmainmode.personnamecom(i).List(au) <> "�m�H���n" Then
                Select Case i
                    Case 1
                       If name2 = FormMainMode.personnamecom(i).List(au) Or name3 = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 2
                       If name1 = FormMainMode.personnamecom(i).List(au) Or name3 = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                    Case 3
                       If name2 = FormMainMode.personnamecom(i).List(au) Or name1 = FormMainMode.personnamecom(i).List(au) Then
                           FormMainMode.personnamecom(i).RemoveItem au
                           ap = ap - 1
                           au = au - 1
                       End If
                End Select
'            End If
            au = au + 1
        Loop
Next
'===========�ˬd���O�_�u���u�H���v�@��
For i = 1 To 3
   If FormMainMode.personnamecom(i).ListCount = 1 Then
       FormMainMode.personnamecom(i).Clear
   End If
Next
���q���ƥ� = False
'formmainmode.personnamecom(�{�b�Ҧb��).ListIndex = ag
For i = 1 To 3
   If ag(i) <> "" Then
       For k = 0 To FormMainMode.personnamecom(i).ListCount - 1
           If FormMainMode.personnamecom(i).List(k) = ag(i) Then
               FormMainMode.personnamecom(i).ListIndex = k
           End If
       Next
   Else
       FormMainMode.personnamecom(i).ListIndex = -1
   End If
Next
End Sub
Sub �d���H����T���_�ϥΪ�(ByVal Index As Integer)
FormMainMode.PEGFusbi1(Index).Caption = VBEPerson(1, Index, 1, 3, 1)
FormMainMode.PEGFusbi2(Index).Caption = VBEPerson(1, Index, 1, 3, 2)
FormMainMode.PEGFusbi3(Index).Caption = VBEPerson(1, Index, 1, 3, 3)
FormMainMode.PEGFcardus(Index).Picture = LoadPicture(VBEPerson(1, Index, 1, 5, 5))
End Sub
Sub �d���H����T���_�q��(ByVal Index As Integer)
FormMainMode.PEGFcardcompi1(Index).Caption = VBEPerson(2, Index, 1, 3, 1)
FormMainMode.PEGFcardcompi2(Index).Caption = VBEPerson(2, Index, 1, 3, 2)
FormMainMode.PEGFcardcompi3(Index).Caption = VBEPerson(2, Index, 1, 3, 3)
FormMainMode.PEGFcardcom(Index).Picture = LoadPicture(VBEPerson(2, Index, 1, 5, 5))
End Sub
Sub �����H��_�ϥΪ�(ByVal Index As Integer)
Dim i As Integer, j As Integer, k As Integer
For i = 1 To 4
    For j = 1 To 30
        For k = 1 To 10
           VBEPerson(1, Index, i, j, k) = ""
        Next
    Next
Next
'==============
VBEPerson(1, Index, 1, 5, 5) = App.Path & "\gif\personunknown.jpg"
VBEPerson(1, Index, 1, 3, 1) = "?"
VBEPerson(1, Index, 1, 3, 2) = "?"
VBEPerson(1, Index, 1, 3, 3) = "?"
VBEPerson(1, Index, 1, 1, 1) = "?"
VBEPerson(1, Index, 1, 1, 2) = "?"
VBEPerson(1, Index, 1, 1, 3) = "?"
VBEPerson(1, Index, 1, 2, 1) = "?"
VBEPerson(1, Index, 1, 2, 2) = "?"
VBEPerson(1, Index, 1, 4, 1) = "??????"
VBEPerson(1, Index, 2, 3, 5) = 1
VBEPerson(1, Index, 1, 3, 4) = "000000"
End Sub
Sub �����H��_�q��(ByVal Index As Integer)
Dim i As Integer, j As Integer, k As Integer
For i = 1 To 4
    For j = 1 To 30
        For k = 1 To 10
           VBEPerson(2, Index, i, j, k) = ""
        Next
    Next
Next
'==============
VBEPerson(2, Index, 1, 5, 5) = App.Path & "\gif\personunknown.jpg"
VBEPerson(2, Index, 1, 3, 1) = "?"
VBEPerson(2, Index, 1, 3, 2) = "?"
VBEPerson(2, Index, 1, 3, 3) = "?"
VBEPerson(2, Index, 1, 1, 1) = "?"
VBEPerson(2, Index, 1, 1, 2) = "?"
VBEPerson(2, Index, 1, 1, 3) = "?"
VBEPerson(2, Index, 1, 2, 1) = "?"
VBEPerson(2, Index, 1, 2, 2) = "?"
VBEPerson(2, Index, 1, 4, 1) = "??????"
VBEPerson(2, Index, 2, 3, 5) = 1
VBEPerson(2, Index, 1, 4, 3) = "?.?.?"
VBEPerson(2, Index, 1, 3, 4) = "000000"
End Sub
Function �H����ܿ��() As String
Dim personcomname As String '�q����H���W�ټȮɬ����ܼ�
Dim talkname() As String  '�C�y��ܤH���O�����O�ܼ�
Dim persontalkname As String  '�C�y��ܤH���O���`�ܼ�
Dim persontalkrec As String '�`�@�i��ܫ��w��ܬ����s����
Dim persontalkrecnum As Integer '�`�@�i��ܫ��w��ܬ�����
Dim at() As String '��ܹ�ܼȮ��ܼ�
Dim m As Integer, i As Integer, k As Integer '�Ȯ��ܼ�
Dim atbo(1 To 10) As Boolean '�H��������ܪťռаO�O����
personcomname = VBEPerson(2, 1, 1, 1, 1)

For i = 1 To 20
   persontalkname = VBEPerson(1, 1, 4, i, 2)
   talkname = Split(persontalkname, "&")
   For k = 0 To UBound(talkname)
       If talkname(k) = personcomname Then
           persontalkrec = persontalkrec & i & "="
           persontalkrecnum = persontalkrecnum + 1
           k = UBound(talkname) '�H�xExitFor
       End If
    Next
Next

If persontalkrecnum >= 1 Then
    m = Int(Rnd() * persontalkrecnum) + 1
    at = Split(persontalkrec, "=")
    �H����ܿ�� = VBEPerson(1, 1, 4, at(m - 1), 1)
Else
    Do
        Randomize
        m = Int(Rnd() * 10) + 1
        If atbo(m) = False Then
            �H����ܿ�� = VBEPerson(1, 1, 4, m + 20, 1)
            atbo(m) = True
        End If
        If �H����ܿ�� <> "" Then
            Exit Do
        ElseIf atbo(1) = True And atbo(2) = True And atbo(3) = True And atbo(4) = True And atbo(5) = True _
        And atbo(6) = True And atbo(7) = True And atbo(8) = True And atbo(9) = True And atbo(10) = True Then
            �H����ܿ�� = ""
            Exit Do
        Else
            atbo(m) = True
        End If
    Loop
End If
End Function
Sub �M������H����T�ܼ�(ByVal uscom As Integer, ByVal num As Integer)
Dim i As Integer, j As Integer, k As Integer
For i = 1 To UBound(VBEPerson, 3)
    For j = 1 To UBound(VBEPerson, 4)
        For k = 1 To UBound(VBEPerson, 5)
            VBEPerson(uscom, num, i, j, k) = ""
        Next
    Next
Next
End Sub
