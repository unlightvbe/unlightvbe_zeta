VERSION 5.00
Begin VB.Form Formatkingus 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "UnlightVBE-�ޯ�Ұʤ�"
   ClientHeight    =   9195
   ClientLeft      =   5085
   ClientTop       =   1275
   ClientWidth     =   6135
   Icon            =   "Formatkingus.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   6135
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   8160
   End
   Begin VB.PictureBox atkingusjpg 
      Appearance      =   0  '����
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   11025
      Left            =   0
      Picture         =   "Formatkingus.frx":0CCA
      ScaleHeight     =   11025
      ScaleWidth      =   10680
      TabIndex        =   0
      Top             =   0
      Width           =   10680
   End
End
Attribute VB_Name = "Formatkingus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim d As Integer

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
�@��t����.���}�C������ Cancel, UnloadMode
End Sub

Private Sub t1_Timer()
If �ثe��(31) = 19 Then
   Formatkingus.Visible = False
   t1.Enabled = False
   If Val(FormMainMode.atkingnumtot.Caption) > 0 Then
      If atkingno(Val(FormMainMode.atkingnumtot.Caption), 11) = 0 Then
         If Formsetting.checktest.Value = 1 Then Debug.Print "Formatkingus If atkingno(,11) = 0 ��"
         '==============
         FormMainMode.atkingnumtot.Caption = Val(FormMainMode.atkingnumtot.Caption) - 1
         FormMainMode.atkingtrtot.Interval = 20
         FormMainMode.atkingtrtot.Enabled = True
      End If
   End If
ElseIf �ثe��(31) = 10 Then
   FormMainMode.�ޯ���椤�󴫹Ϥ�_Timer
   FormMainMode.�ޯ���椤�Ұ�_Timer
   �ثe��(31) = Val(�ثe��(31)) + 1
ElseIf �ثe��(31) = 7 Then
   FormMainMode.wmpse5.Controls.play
   �@��t����.�ˬd���ּ��� 5
   �ثe��(31) = Val(�ثe��(31)) + 1
ElseIf �ثe��(31) = 5 Then
   atkingusjpg.Visible = True
   �ثe��(31) = Val(�ثe��(31)) + 1
Else
   �ثe��(31) = Val(�ثe��(31)) + 1
End If
End Sub
