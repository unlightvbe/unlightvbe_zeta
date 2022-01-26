VERSION 5.00
Begin VB.Form Formatkingcom 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '虫絬㏕﹚
   Caption         =   "UnlightVBE-м币笆い"
   ClientHeight    =   9945
   ClientLeft      =   9480
   ClientTop       =   1965
   ClientWidth     =   6780
   Icon            =   "Formatkingcom.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   6780
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   8040
   End
   Begin VB.PictureBox atkingcomjpg 
      Appearance      =   0  'キ
      BackColor       =   &H00000000&
      BorderStyle     =   0  '⊿Τ絬
      ForeColor       =   &H80000008&
      Height          =   11025
      Left            =   -840
      Picture         =   "Formatkingcom.frx":0CCA
      ScaleHeight     =   11025
      ScaleWidth      =   18600
      TabIndex        =   0
      Top             =   -600
      Width           =   18600
   End
End
Attribute VB_Name = "Formatkingcom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Integer

Private Sub Form_Activate()
t1.Enabled = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
   YesNo = MsgBox("絋﹚瞒秨笴栏?", 36, "UnlightVBE-╰参矗ボ")
   If YesNo = 6 Then
    End
   Else
    Cancel = 1
   End If
End If
End Sub

Private Sub t1_Timer()
If ヘ玡计(31) = 19 Then
   Formatkingcom.Visible = False
   t1.Enabled = False
   If Val(FormMainMode.atkingnumtot.Caption) > 0 Then
      If atkingno(Val(FormMainMode.atkingnumtot.Caption), 11) = 0 Then
          If Formsetting.checktest.Value = 1 Then Debug.Print "Formatkingcom If atkingno(,11) = 0 "
          '=======================
          FormMainMode.atkingnumtot.Caption = Val(FormMainMode.atkingnumtot.Caption) - 1
          FormMainMode.atkingtrtot.Interval = 20
          FormMainMode.atkingtrtot.Enabled = True
      End If
   End If
ElseIf ヘ玡计(31) = 10 Then
   FormMainMode.м磅︽い传瓜_Timer
   FormMainMode.м磅︽い币笆_Timer
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
ElseIf ヘ玡计(31) = 7 Then
   FormMainMode.wmpse5.Controls.play
   ╰参摸.浪琩贾冀 5
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
ElseIf ヘ玡计(31) = 5 Then
   atkingcomjpg.Visible = True
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
Else
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
End If
End Sub
