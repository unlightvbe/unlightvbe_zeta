VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form Formchangeperson 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "UnlightVBE-交換角色"
   ClientHeight    =   4845
   ClientLeft      =   6690
   ClientTop       =   2535
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Formchangeperson.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6450
   Begin VB.Timer 使用者方智慧型AI_自動控制選人 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6120
      Top             =   4320
   End
   Begin VB.PictureBox card 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3615
      Index           =   1
      Left            =   360
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      Begin ImageX.aicAlphaImage PEAFcardbackclick 
         Height          =   795
         Index           =   1
         Left            =   480
         Top             =   1320
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   1402
         Image           =   "Formchangeperson.frx":0CCA
         Props           =   13
      End
      Begin UnlightVBE.大人物形像 cardbackus 
         Height          =   3615
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   6376
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   8
         Left            =   1440
         TabIndex        =   17
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   9
         Left            =   1440
         TabIndex        =   18
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   10
         Left            =   1440
         TabIndex        =   19
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   11
         Left            =   1440
         TabIndex        =   20
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   12
         Left            =   1440
         TabIndex        =   21
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   13
         Left            =   1440
         TabIndex        =   22
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   14
         Left            =   1440
         TabIndex        =   23
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin VB.Label cardhp 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   555
         TabIndex        =   3
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label cardatk 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label carddef 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Top             =   3240
         Width           =   495
      End
   End
   Begin VB.PictureBox PEAFcardback 
      Appearance      =   0  '平面
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3615
      Index           =   2
      Left            =   3480
      Picture         =   "Formchangeperson.frx":33FF
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   10
         Left            =   2240
         TabIndex        =   74
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   9
         Left            =   1930
         TabIndex        =   73
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   8
         Left            =   1635
         TabIndex        =   72
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   7
         Left            =   1340
         TabIndex        =   71
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   6
         Left            =   1040
         TabIndex        =   70
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   10
         Left            =   2240
         TabIndex        =   69
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   9
         Left            =   1930
         TabIndex        =   68
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   8
         Left            =   1630
         TabIndex        =   67
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   7
         Left            =   1340
         TabIndex        =   66
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   6
         Left            =   1040
         TabIndex        =   65
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   10
         Left            =   2240
         TabIndex        =   64
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   9
         Left            =   1930
         TabIndex        =   63
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   8
         Left            =   1630
         TabIndex        =   62
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   7
         Left            =   1340
         TabIndex        =   61
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   6
         Left            =   1040
         TabIndex        =   60
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   10
         Left            =   2240
         TabIndex        =   59
         Top             =   600
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   9
         Left            =   1930
         TabIndex        =   58
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   8
         Left            =   1630
         TabIndex        =   57
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   7
         Left            =   1340
         TabIndex        =   56
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   6
         Left            =   880
         TabIndex        =   55
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   5
         Left            =   740
         TabIndex        =   54
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   4
         Left            =   580
         TabIndex        =   53
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   6
         Left            =   880
         TabIndex        =   52
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   5
         Left            =   740
         TabIndex        =   51
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   4
         Left            =   580
         TabIndex        =   50
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   6
         Left            =   885
         TabIndex        =   49
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   5
         Left            =   740
         TabIndex        =   48
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   4
         Left            =   580
         TabIndex        =   47
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   6
         Left            =   885
         TabIndex        =   46
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   5
         Left            =   740
         TabIndex        =   45
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   8
         Left            =   100
         TabIndex        =   44
         Top             =   1960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   7
         Left            =   100
         TabIndex        =   43
         Top             =   1530
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   6
         Left            =   100
         TabIndex        =   42
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   5
         Left            =   100
         TabIndex        =   41
         Top             =   630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   4
         Left            =   580
         TabIndex        =   40
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   6
         Left            =   1040
         TabIndex        =   39
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin VB.Label PEAFpersoncardback_main 
         BackStyle       =   0  '透明
         Caption         =   "DEF+7。防禦成功時，對手受到與所超過之防禦同值的傷害"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   8.25
            Charset         =   136
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   79
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   78
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   77
         Top             =   1245
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   76
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   105
         TabIndex        =   75
         Top             =   315
         Width           =   2295
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   5
         Left            =   120
         Top             =   340
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":8E34
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   6
         Left            =   120
         Top             =   800
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":8F09
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   7
         Left            =   120
         Top             =   1280
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":8FDE
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   8
         Left            =   120
         Top             =   1710
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":90B3
         Props           =   13
      End
   End
   Begin VB.PictureBox PEAFcardback 
      Appearance      =   0  '平面
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3615
      Index           =   1
      Left            =   240
      Picture         =   "Formchangeperson.frx":9188
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   80
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   121
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   120
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   1
         Left            =   100
         TabIndex        =   119
         Top             =   630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   2
         Left            =   100
         TabIndex        =   118
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   3
         Left            =   100
         TabIndex        =   117
         Top             =   1530
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_turn 
         Height          =   135
         Index           =   4
         Left            =   100
         TabIndex        =   116
         Top             =   1960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   238
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   115
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range1 
         Height          =   255
         Index           =   3
         Left            =   885
         TabIndex        =   114
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   113
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   112
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range2 
         Height          =   255
         Index           =   3
         Left            =   885
         TabIndex        =   111
         Top             =   1080
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   110
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   109
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range3 
         Height          =   255
         Index           =   3
         Left            =   880
         TabIndex        =   108
         Top             =   1520
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   1
         Left            =   580
         TabIndex        =   107
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   2
         Left            =   740
         TabIndex        =   106
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_range4 
         Height          =   255
         Index           =   3
         Left            =   880
         TabIndex        =   105
         Top             =   1950
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   104
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   103
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   102
         Top             =   600
         Width           =   290
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num1 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   101
         Top             =   600
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   100
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   99
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   98
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   97
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num2 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   96
         Top             =   1080
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   95
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   94
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   3
         Left            =   1630
         TabIndex        =   93
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   92
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num3 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   91
         Top             =   1520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   1
         Left            =   1040
         TabIndex        =   90
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   2
         Left            =   1340
         TabIndex        =   89
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   3
         Left            =   1635
         TabIndex        =   88
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   4
         Left            =   1930
         TabIndex        =   87
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin UnlightVBE.uc卡片背面 PEAFpersoncardback_num4 
         Height          =   255
         Index           =   5
         Left            =   2240
         TabIndex        =   86
         Top             =   1950
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   85
         Top             =   315
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   84
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   83
         Top             =   1245
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_text 
         BackStyle       =   0  '透明
         Caption         =   "精密射擊"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   82
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label PEAFpersoncardback_main 
         BackStyle       =   0  '透明
         Caption         =   "DEF+7。防禦成功時，對手受到與所超過之防禦同值的傷害"
         BeginProperty Font 
            Name            =   "Noto Sans T Chinese DemiLight"
            Size            =   8.25
            Charset         =   136
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   2280
         Width           =   2295
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   1
         Left            =   120
         Top             =   340
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":EBBD
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   2
         Left            =   120
         Top             =   800
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":EC92
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   3
         Left            =   120
         Top             =   1280
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":ED67
         Props           =   13
      End
      Begin ImageX.aicAlphaImage PEAFcardbackBR 
         Height          =   435
         Index           =   4
         Left            =   120
         Top             =   1710
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   767
         Image           =   "Formchangeperson.frx":EE3C
         Props           =   13
      End
   End
   Begin VB.Timer trchange 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   360
   End
   Begin VB.PictureBox card 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3615
      Index           =   2
      Left            =   3480
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   4
      Top             =   480
      Width           =   2535
      Begin ImageX.aicAlphaImage PEAFcardbackclick 
         Height          =   795
         Index           =   2
         Left            =   480
         Top             =   1320
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   1402
         Image           =   "Formchangeperson.frx":EF11
         Props           =   13
      End
      Begin UnlightVBE.大人物形像 cardbackus 
         Height          =   3615
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   6376
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   15
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   16
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   17
         Left            =   360
         TabIndex        =   26
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   18
         Left            =   360
         TabIndex        =   27
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   19
         Left            =   360
         TabIndex        =   28
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   20
         Left            =   360
         TabIndex        =   29
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   21
         Left            =   360
         TabIndex        =   30
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   22
         Left            =   1440
         TabIndex        =   31
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   23
         Left            =   1440
         TabIndex        =   32
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   24
         Left            =   1440
         TabIndex        =   33
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   25
         Left            =   1440
         TabIndex        =   34
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   26
         Left            =   1440
         TabIndex        =   35
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   27
         Left            =   1440
         TabIndex        =   36
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin UnlightVBE.uc異常狀態 personusspe 
         Height          =   375
         Index           =   28
         Left            =   1440
         TabIndex        =   37
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
      End
      Begin VB.Label carddef 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   7
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label cardatk 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label cardhp 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Bradley Gratis"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   555
         TabIndex        =   5
         Top             =   3240
         Width           =   375
      End
   End
   Begin VB.Image bnok 
      Height          =   345
      Index           =   2
      Left            =   3600
      Picture         =   "Formchangeperson.frx":11646
      Top             =   4200
      Width           =   2250
   End
   Begin VB.Image bnok 
      Height          =   345
      Index           =   1
      Left            =   480
      Picture         =   "Formchangeperson.frx":13F24
      Top             =   4200
      Width           =   2250
   End
End
Attribute VB_Name = "Formchangeperson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub bnok_Click(Index As Integer)
If cardhp(Index).Caption > 0 Then
    戰鬥系統類.人物交換_使用者_指定交換 Index + 1
    執行動作_交換人物角色_結束執行
End If
End Sub

Private Sub bnok_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bnok(Index).Picture = LoadPicture(App.Path & "\gif\changeok_2.bmp")
End Sub

Sub card_Click(Index As Integer)
Dim k As Integer
戰鬥系統類.技能說明載入_人物卡片背面_交換角色 Index
'======================================================
PEAFcardback(Index).Left = card(Index).Left
PEAFcardback(Index).Top = card(Index).Top
Select Case Index
      Case 1
            For k = 1 To 4
                 Formchangeperson.PEAFcardbackBR(k).Opacity = 0
            Next
      Case 2
            For k = 1 To 4
                 Formchangeperson.PEAFcardbackBR(k + 4).Opacity = 0
            Next
End Select
FormMainMode.wmpse9.Controls.stop
FormMainMode.wmpse9.Controls.play
一般系統類.檢查音樂播放 9
PEAFcardback(Index).Visible = True
PEAFcardback(Index).ZOrder
End Sub

Private Sub card_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PEAFcardbackclick(Index).Visible = True
End Sub

Sub cardhp_Change(Index As Integer)
If Val(cardhp(Index).Caption) = Val(liveusmax(角色待機人物紀錄數(1, Index + 1))) Then
   cardhp(Index).ForeColor = RGB(255, 255, 255)
   cardbackus(Index).Visible = False
End If
 If Val(cardhp(Index).Caption) < Val(liveusmax(角色待機人物紀錄數(1, Index + 1))) Then
   cardhp(Index).ForeColor = RGB(255, 255, 128)
   cardbackus(Index).Visible = False
 End If
 If Val(cardhp(Index).Caption) <= Val(liveus41(角色待機人物紀錄數(1, Index + 1))) Then
   cardhp(Index).ForeColor = RGB(255, 0, 0)
   cardbackus(Index).Visible = False
 End If
If Val(cardhp(Index).Caption) <= 0 Then
    cardhp(Index).Caption = 0
    cardbackus(Index).大人物圖片 = app_path & "gif\cardblack.png"
    cardbackus(Index).Visible = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bnok(1).Picture = LoadPicture(App.Path & "\gif\changeok_1.bmp")
bnok(2).Picture = LoadPicture(App.Path & "\gif\changeok_1.bmp")
PEAFcardbackclick(1).Visible = False
PEAFcardbackclick(2).Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim m As Integer
If FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption > 0 Then
    執行動作_交換人物角色_結束執行
Else
    Randomize
    m = Int(Rnd() * 2) + 1
    If cardhp(m).Caption > 0 Then
       戰鬥系統類.人物交換_使用者_指定交換 m + 1
       執行動作_交換人物角色_結束執行
    Else
       If m = 2 Then m = 3 Else m = 2
       戰鬥系統類.人物交換_使用者_指定交換 m + 1
       執行動作_交換人物角色_結束執行
    End If
End If
End Sub

Private Sub PEAFcardback_Click(Index As Integer)
PEAFcardback(Index).Visible = False
FormMainMode.wmpse9.Controls.stop
FormMainMode.wmpse9.Controls.play
一般系統類.檢查音樂播放 9
End Sub


Sub PEAFcardbackBR_Click(Index As Integer, ByVal Button As Integer)
Dim ahmt As String, i As Integer, k As Integer
Select Case Index
     Case Is <= 4
           ahmt = VBEPerson(1, 角色待機人物紀錄數(1, 2), 3, Index, 5)
            For i = 1 To Len(ahmt)
                If Mid(ahmt, i, 1) = "&" Then
                    Mid(ahmt, i, 1) = Chr(10)
                End If
            Next
           PEAFpersoncardback_main(1).Caption = ahmt
           PEAFcardbackBR(Index).Opacity = 100
           人物卡面背面編號紀錄數(6) = Index
           For k = 1 To 4
                 If k <> Index Then
                     PEAFcardbackBR(Index).Opacity = 0
                 End If
           Next
     Case Is >= 5
           ahmt = VBEPerson(1, 角色待機人物紀錄數(1, 3), 3, Index - 4, 5)
           For i = 1 To Len(ahmt)
                If Mid(ahmt, i, 1) = "&" Then
                    Mid(ahmt, i, 1) = Chr(10)
                End If
            Next
           PEAFpersoncardback_main(2).Caption = ahmt
           PEAFcardbackBR(Index).Opacity = 100
           人物卡面背面編號紀錄數(7) = Index
           For k = 5 To 8
                 If k <> Index Then
                     PEAFcardbackBR(Index).Opacity = 0
                 End If
           Next
End Select
End Sub

Sub PEAFcardbackBR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Integer
PEAFcardbackBR(Index).Opacity = 100
Select Case Index
     Case Is <= 4
           For k = 1 To 4
                If k <> Index And k <> 人物卡面背面編號紀錄數(6) Then
                    PEAFcardbackBR(k).Opacity = 0
                End If
           Next
     Case Is >= 5
           For k = 5 To 8
                If k <> Index And k <> 人物卡面背面編號紀錄數(7) Then
                    PEAFcardbackBR(k).Opacity = 0
                End If
           Next
End Select
End Sub


Private Sub PEAFcardbackclick_Click(Index As Integer, ByVal Button As Integer)
Formchangeperson.card_Click (Index)
End Sub

Private Sub PEAFpersoncardback_main_Click(Index As Integer)
PEAFcardback(Index).Visible = False
FormMainMode.wmpse9.Controls.stop
FormMainMode.wmpse9.Controls.play
一般系統類.檢查音樂播放 9
End Sub


Private Sub PEAFpersoncardback_text_Click(Index As Integer)
Call PEAFcardbackBR_Click(Index, 0)
End Sub


Private Sub PEAFpersoncardback_text_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PEAFcardbackBR_MouseMove(Index, 0, 0, 0, 0)
End Sub


Sub 使用者方智慧型AI_自動控制選人_Timer()
Dim i As Integer
For i = 1 To 2
    If Val(cardhp(i).Caption) > 0 Then
        Formchangeperson.bnok_Click (i)
        Formchangeperson.使用者方智慧型AI_自動控制選人.Enabled = False
        Exit Sub
    End If
Next
End Sub
