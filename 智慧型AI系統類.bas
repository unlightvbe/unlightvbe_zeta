Attribute VB_Name = "智慧型AI系統類"
Option Explicit
Public cardcountAInum() As String  '公用牌計算暫時基本資料(第x張,1.正面類型/2.正面數值/3.反面類型/4.反面數值/5.牌編號)
Public cardcountAInumMOV() As String  '公用牌計算暫時基本資料-移動階段續-原本(第x張,1.正面類型/2.正面數值/3.反面類型/4.反面數值/5.牌編號)
Dim cardAIn() As Integer '排列組合計算暫時變數
Dim cardAInumans As String '排列組合計算暫時變數
Public cardAInumnm() As String '排列組合計算最終數值
Public cardAInumFinal() As Integer '排列組合計算最終期望值
Public cardAInumFinal2() As Integer '排列組合計算最終期望值-排列後
Public cardAInumcase(1 To 5, 1 To 2) As Integer '公用牌計算統計資料(1.ATK-劍/2.DEF/3.MOV/4.SPE/5.ATK-槍,1.組合下最低數值/2.組合下最高數值)
Public cardAInumcaseperson() As Integer '公用牌計算統計暫時資料-個別組合
Public cardAInumuscom As Integer '手牌擁有者牌數記錄暫時變數
Public cardAInumcasepersonTER() As Integer '公用牌計算統計暫時資料-個別組合-個別卡面數值計數統計
Public cardAInumselect1 As Integer  '公用牌計算統計比序暫時變數-目前最高期望值
Public cardAInumselect4 As Integer  '公用牌計算統計比序暫時變數-目前最高個別加總期望值
Public cardAInumselect2 As String '公用牌計算統計比序暫時變數-目前最高期望值下編號串-初始
Public cardAInumselect3() As String '公用牌計算統計比序暫時變數-目前最高期望值下編號串-陣列
Public cardAInumchoose As Integer '公用牌計算最終選擇組合編號
Public cardAInumMOVmain(1 To 2, 1 To 15) As String 'AI-移動階段續-組合暫時紀錄
Public cardAInumMOVnm() As String 'AI-移動階段續-正向面-計算排列組合串暫時紀錄
Public cardAInumMOVnmtot() As String 'AI-移動階段續-正向面-總共排列組合串相關資料暫時紀錄
Public cardAInumMOVFinal(1 To 3) As String 'AI-移動階段續-正向面-最終結果紀錄數(1.最終排列組合串/2.最終排列組合編號/3.最終選定目標距離[1.近/2.遠])
Public 是否移動階段續估計判斷程序 As Boolean 'AI-移動階段續-是否為估計判斷程序標記數
Public cardAInumOvertenrecord(1 To 10) As Integer 'AI引導程序-超出牌張數-牌紀錄暫時變數(1~10.牌編號)
Public personatkingtfr(1 To 5) As Integer '計算個別技能-是否為Ex技(1~4.(1)有/(2)無,5.是否有封印)
Sub 智慧型AI系統計算_一階段_初始(ByVal pagenumber As Integer)
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
'=========計算正反面排列組合數值
智慧型AI系統類.排列組合計算 pagenumber
End Sub
Sub 智慧型AI系統計算_一階段_取得牌面資料(ByVal 是否一般 As Boolean, ByVal uscom As Integer)
Dim i As Integer
If 是否一般 = True Then
        '=========擷取目前牌面資料
        Select Case uscom
            Case 1
                戰鬥系統類.出牌順序計算_使用者_手牌
            Case 2
                戰鬥系統類.出牌順序計算_電腦_手牌
        End Select
        Dim w As Integer '暫時變數
        w = 2 * uscom '(2-使用者手牌/4-電腦手牌)
        For i = 1 To pageglead(uscom)
            cardcountAInum(i, 5) = 出牌順序統計暫時變數(w, i, 2)
            cardcountAInum(i, 1) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 1)
            cardcountAInum(i, 2) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 2)
            cardcountAInum(i, 3) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 3)
            cardcountAInum(i, 4) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 4)
        Next
End If
'======================
智慧型AI系統類.排列組合統計數值計算_手牌總計
智慧型AI系統類.排列組合統計數值計算_個別組合
End Sub
Sub 智慧型AI系統計算_二階段_計算期望值_初始(ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
Dim wnum As Integer, whnum As Integer, i As Integer, j As Integer '暫時變數
Select Case turn
    Case 1 '===攻擊階段
         If uscom = 1 Then whnum = atkus(角色人物對戰人數(1, 2)) Else whnum = atkcom(角色人物對戰人數(2, 2))
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
    Case 2  '===防禦階段
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
    Case 3  '===移動階段
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
Sub 智慧型AI系統計算_二階段_計算期望值_個別技能(ByVal name As String, ByVal turn As Integer, ByVal movecpre As Integer, ByVal uscom As Integer)
智慧型AI系統類.檢查人物技能是否有EX技 uscom, name
If personatkingtfr(5) = 1 Then
   Exit Sub '有封印狀態時無法發動技能
End If
Select Case name
     Case "艾伯李斯特"
           智慧型AI人物類.艾伯李斯特 turn, movecpre, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "雪莉"
           智慧型AI人物類.雪莉 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "艾茵"
           智慧型AI人物類.艾茵 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "古魯瓦爾多"
           智慧型AI人物類.古魯瓦爾多 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "帕茉"
           智慧型AI人物類.帕茉 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "史塔夏"
           智慧型AI人物類.史塔夏 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "C.C."
           智慧型AI人物類.CC turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "伊芙琳"
           智慧型AI人物類.伊芙琳 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "布勞"
           智慧型AI人物類.布勞 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "梅倫"
           智慧型AI人物類.梅倫 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "音音夢"
           智慧型AI人物類.音音夢 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "艾依查庫"
           智慧型AI人物類.艾依查庫 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "阿貝爾"
           智慧型AI人物類.阿貝爾 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "利恩"
           智慧型AI人物類.利恩 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "夏洛特"
           智慧型AI人物類.夏洛特 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "泰瑞爾"
           智慧型AI人物類.泰瑞爾 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "瑪格莉特"
           智慧型AI人物類.瑪格莉特 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "庫勒尼西"
           智慧型AI人物類.庫勒尼西 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "蕾格烈芙"
           智慧型AI人物類.蕾格烈芙 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "多妮妲"
           智慧型AI人物類.多妮妲 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "傑多"
           智慧型AI人物類.傑多 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "阿奇波爾多"
           智慧型AI人物類.阿奇波爾多 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "露緹亞"
           智慧型AI人物類.露緹亞 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "梅莉"
           智慧型AI人物類.梅莉 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "貝琳達"
           智慧型AI人物類.貝琳達 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "蕾"
           智慧型AI人物類.蕾 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "羅莎琳"
           智慧型AI人物類.羅莎琳 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "洛洛妮"
           智慧型AI人物類.洛洛妮 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "克頓"
           智慧型AI人物類.克頓 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "艾蕾可"
           智慧型AI人物類.艾蕾可 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
     Case "尤莉卡"
           智慧型AI人物類.尤莉卡 turn, movecpre, uscom, personatkingtfr(1), personatkingtfr(2), personatkingtfr(3), personatkingtfr(4)
End Select

End Sub
Sub 排列組合計算(ByVal qnum As Integer)
Dim i As Integer
'===========
ReDim cardAIn(1 To Val(qnum))
Erase cardAInumnm
cardAInumans = ""
Dim s As Integer
For i = 1 To qnum   '重設區塊數值
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
    智慧型AI系統類.排列組合計算_區塊進位 qnum '共[qnum]位數
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
Sub 排列組合計算_區塊進位(ByVal num As Integer)
Dim i As Integer
For i = 1 To num - 1
    If cardAIn(i) = 2 Then
        cardAIn(i + 1) = cardAIn(i + 1) + 1
        cardAIn(i) = 0
    End If
Next

End Sub
Sub 排列組合統計數值計算_手牌總計()
Dim we As Integer, i As Integer, j As Integer '暫時變數
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
Sub 排列組合統計數值計算_個別組合()
Dim we As Integer, i As Integer, j As Integer '暫時變數
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
Sub 智慧型AI系統計算_三階段_統計排列()
Dim i As Integer, j As Integer, k As Integer
'=================複製內容
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
Sub 智慧型AI系統計算_四階段_比序_1_初始()
Dim i As Integer
For i = 1 To 2 ^ cardAInumuscom
    If Val(cardAInumFinal2(i, 1)) > Val(cardAInumselect1) Then
        cardAInumselect1 = cardAInumFinal2(i, 1)
    End If
Next
'====================
If cardAInumselect1 < 0 Then cardAInumselect1 = 0 '去除總期望值為負數之組合
'====================
For i = 1 To 2 ^ cardAInumuscom
    If cardAInumFinal2(i, 1) = cardAInumselect1 Then
        cardAInumselect2 = cardAInumselect2 & "=" & cardAInumFinal2(i, 2)
    End If
Next
'====================
If cardAInumselect2 = "" Then  '沒有任何組合符合條件
    cardAInumselect2 = "-10=-10"
End If
End Sub
Sub 智慧型AI系統計算_四階段_比序_2_超額比序判斷_1()
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
Sub 智慧型AI系統計算_四階段_比序_2_超額比序判斷_2()
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
Sub 智慧型AI系統計算_四階段_比序_3_選擇組合()
If UBound(cardAInumselect3) > 1 Then
    Dim wtr As Integer '暫時變數
    wtr = Int(Rnd() * UBound(cardAInumselect3)) + 1
    cardAInumchoose = cardAInumselect3(wtr)
Else
    cardAInumchoose = cardAInumselect3(1)
End If
End Sub
Sub 智慧型AI系統計算_最後階段_實行選牌(ByVal choose As Integer, ByVal uscom As Integer)
Dim wer As Integer, i As Integer, cspce As String, cspme As String '暫時變數
If choose = 1 Then
    wer = 0
Else
    wer = 1
End If
'=================
Dim pu As Integer '暫時變數
'=====
If cardAInumchoose = -10 Then  '==沒有任何組合符合出牌條件
    Exit Sub
End If
'=======================如組合符合出牌條件的話
Select Case uscom
     Case 1 '==使用者方
            For i = 1 To cardAInumuscom
                    pu = cardcountAInum(i, 5)
                    If Mid(cardAInumnm(cardAInumchoose - 1), i, 1) = 1 And cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 4
                    ElseIf cardAInumcaseperson(cardAInumchoose, 2, i) >= wer Then
                        pagecardnum(pu, 11) = 3
                    End If
            Next
     Case 2 '==電腦方
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
Sub 智慧型AI系統計算_暫時匯出(ByVal uscom As Integer)
Dim i As Integer, k As Integer
If Formsetting.checktest.Value = 1 Then
'    Open App.Path & "\test\out1.txt" For Output As #1
    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & 戰鬥系統類.turn & "turn_" & 戰鬥系統類.turnatk & "_" & uscom & "_1.txt" For Output As #1
    For i = 1 To 2 ^ cardAInumuscom
        Print #1, cardAInumnm(Val(cardAInumFinal2(i, 2)) - 1) & "=" & cardAInumFinal2(i, 1) & "/" & cardAInumFinal2(i, 4) & "#" & cardAInumFinal2(i, 2) & "@";
        For k = 1 To cardAInumuscom
            Print #1, cardAInumcaseperson(Val(cardAInumFinal2(i, 2)), 2, k) & "=";
        Next
        Print #1,
    Next
    Close
    'MsgBox "已匯出完畢1"
End If
End Sub
Sub 智慧型AI系統計算_引導程序_試驗1(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer)
智慧型AI系統類.智慧型AI系統計算_一階段_初始 uscom
智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turn, movecpre, uscom
智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turn, movecpre, uscom
智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
智慧型AI系統類.智慧型AI系統計算_暫時匯出 uscom
End Sub
Sub 智慧型AI系統計算_引導程序_選擇(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer)
If Val(pageglead(uscom)) > 10 Then
    智慧型AI系統類.智慧型AI系統計算_引導程序_超出牌張數 uscom, turn, name, movecpre, choose
ElseIf Val(pageglead(uscom)) > 0 And Val(pageglead(uscom)) <= 10 Then
    智慧型AI系統類.智慧型AI系統計算_一階段_初始 pageglead(uscom)
    智慧型AI系統類.智慧型AI系統計算_一階段_取得牌面資料 True, uscom
    智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turn, movecpre, uscom
    智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turn, movecpre, uscom
    智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_1_初始
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_1
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_2
    智慧型AI系統類.智慧型AI系統計算_暫時匯出 uscom
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_3_選擇組合
    If turn = 3 And cardAInumchoose > 0 Then
        智慧型AI系統類.智慧型AI系統計算_引導程序_移動階段續 uscom, turn, name, movecpre, choose, pageglead(uscom)
    Else
        智慧型AI系統類.智慧型AI系統計算_最後階段_實行選牌 choose, uscom
    End If
End If
End Sub
Sub 智慧型AI系統計算_引導程序_移動階段續(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer, ByVal pagenumber As Integer)
If Val(pagenumber) > 0 Then
    Select Case 智慧型AI系統計算_移動階段續_判斷出牌資格(uscom)
        Case True
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_一階段_準備進行資料
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_二階段_進行估計排列組合串計算 pagenumber, uscom
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_三階段_進行估計期望值計算 uscom, name, choose, movecpre, pagenumber
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_四階段_統計估計期望值及判斷 uscom
            智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_五階段_實行選牌 choose, uscom, pagenumber
        Case False
            智慧型AI系統類.智慧型AI系統計算_移動階段續_否定面_一階段_重設期望值_個別
            智慧型AI系統類.智慧型AI系統計算_移動階段續_否定面_二階段_選擇行動 uscom
            智慧型AI系統類.智慧型AI系統計算_最後階段_實行選牌 choose, uscom
    End Select
End If
End Sub
Function 智慧型AI系統_目前可執行之人物判斷(ByVal name As String) As Boolean
If Formsetting.chkusenewai.Value = 0 Then
    智慧型AI系統_目前可執行之人物判斷 = False
    Exit Function
End If
Select Case name
    Case "艾伯李斯特"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "雪莉"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "艾茵"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "古魯瓦爾多"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "帕茉"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "史塔夏"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "C.C."
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "伊芙琳"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "布勞"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "梅倫"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "音音夢"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "艾依查庫"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "阿貝爾"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "利恩"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "夏洛特"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "泰瑞爾"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "瑪格莉特"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "庫勒尼西"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "蕾格烈芙"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "多妮妲"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "傑多"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "阿奇波爾多"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "露緹亞"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "梅莉"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "貝琳達"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "蕾"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "羅莎琳"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "洛洛妮"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "克頓"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "艾蕾可"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case "尤莉卡"
            智慧型AI系統_目前可執行之人物判斷 = True
    Case Else
            智慧型AI系統_目前可執行之人物判斷 = False
End Select
End Function
Function 階層數(ByVal num As Integer) As Single
Dim w As Double, i As Integer
w = 1
If num <> 0 Then
    For i = 1 To Val(num)
        w = Val(w) * Val(i)
    Next
Else
    w = 1
End If
階層數 = w
End Function
Function 階層數_取C(ByVal c1 As Integer, ByVal c2 As Integer) As Single

階層數_取C = (智慧型AI系統類.階層數(c1) / 智慧型AI系統類.階層數(Val(c1) - Val(c2))) / 智慧型AI系統類.階層數(c2)

End Function
Sub 智慧型AI系統計算_移動階段續_取得計算之排列組合(ByVal n1 As Integer, ByVal n2 As Integer)
Dim wtstr As String, wtall As Integer, wtpnum() As String, wtn As Integer, i As Integer, j As Integer
'===================
智慧型AI系統類.排列組合計算 n1
wtall = 智慧型AI系統類.階層數_取C(n1, n2)
ReDim cardAInumMOVnm(1 To wtall) As String
'====================
For i = 1 To 2 ^ n1
    wtn = 0
    For j = 1 To n1
        If Val(Mid(cardAInumnm(i - 1), j, 1)) = 1 Then
            wtn = wtn + 1
        End If
    Next
    If wtn = n2 Then '==有n2張出牌之組合
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
'    MsgBox "失敗"
'End If
For i = 1 To UBound(wtpnum)
    cardAInumMOVnm(i) = cardAInumnm(wtpnum(i) - 1)
Next
End Sub
Function 智慧型AI系統計算_移動階段續_判斷出牌資格(ByVal uscom As Integer) As Boolean
Erase cardAInumMOVmain
Erase cardAInumMOVnm
Erase cardAInumMOVnmtot
Dim wtmovnum As Integer, i As Integer '暫時變數
If cardAInumchoose = -10 Then
    智慧型AI系統計算_移動階段續_判斷出牌資格 = False
    Exit Function
End If
'============紀錄目前組合
cardAInumMOVmain(1, 1) = cardAInumselect1
cardAInumMOVmain(1, 2) = cardAInumselect4
cardAInumMOVmain(1, 3) = cardAInumnm(cardAInumchoose - 1)
cardAInumMOVmain(1, 4) = cardAInumcaseperson(cardAInumchoose, 1, 13)
cardAInumMOVmain(1, 5) = cardAInumchoose
For i = 1 To cardAInumuscom
    cardAInumMOVmain(2, i) = cardAInumcaseperson(cardAInumchoose, 2, i)
Next
'==============計算有效移動數
wtmovnum = cardAInumMOVmain(1, 4)
For i = 14 * (角色人物對戰人數(uscom, 2) - 1) + 1 To 14 * 角色人物對戰人數(uscom, 2)
    If (人物異常狀態資料庫(uscom, i, 3) = 6 And uscom = 2) Or (人物異常狀態資料庫(uscom, i, 3) = 12 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) - Val(人物異常狀態資料庫(uscom, i, 1))
    End If
    If (人物異常狀態資料庫(uscom, i, 3) = 3 And uscom = 2) Or (人物異常狀態資料庫(uscom, i, 3) = 9 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) + Val(人物異常狀態資料庫(uscom, i, 1))
    End If
    If (人物異常狀態資料庫(uscom, i, 3) = 16 And uscom = 1) Or (人物異常狀態資料庫(uscom, i, 3) = 17 And uscom = 2) Then
        wtmovnum = -100
    End If
Next
'=====================
If wtmovnum >= 2 Then
    智慧型AI系統計算_移動階段續_判斷出牌資格 = True
Else
    智慧型AI系統計算_移動階段續_判斷出牌資格 = False
End If
End Function
Sub 智慧型AI系統計算_移動階段續_否定面_一階段_重設期望值_個別()
Dim i As Integer
For i = 1 To cardAInumuscom
    If cardAInumcaseperson(cardAInumchoose, 2, i) < 10 Then
        cardAInumcaseperson(cardAInumchoose, 2, i) = 0
        cardAInumMOVmain(2, i) = 0
    End If
Next
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_一階段_準備進行資料()
Dim wercnum As Integer, werct As String, werpnum As Integer, k As Integer, q As Integer
ReDim cardcountAInumMOV(1 To cardAInumuscom, 1 To 5) As String
是否移動階段續估計判斷程序 = True
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
Sub 智慧型AI系統計算_移動階段續_正向面_二階段_進行估計排列組合串計算(ByVal pagenumber As Integer, ByVal uscom As Integer)
Dim weru As Integer, wernum As Integer, werqr As String
Dim werstru As String
Dim werpstr() As String
Dim wermovnm As Integer, wermovynm As Integer, i As Integer, k As Integer
'============進行估計之移動牌排列組合計算
For i = 1 To Val(cardAInumMOVnmtot(0, 3))
       智慧型AI系統計算_移動階段續_取得計算之排列組合 Val(cardAInumMOVnmtot(0, 3)), i
       weru = 1
       wernum = 階層數_取C(Val(cardAInumMOVnmtot(0, 3)), i)
        For k = Val(cardAInumMOVnmtot(0, 2)) To (Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)) - 1
             cardAInumMOVnmtot(k, 1) = cardAInumMOVnm(weru)
             weru = Val(weru) + 1
        Next
        cardAInumMOVnmtot(0, 2) = Val(cardAInumMOVnmtot(0, 2)) + Val(wernum)
Next
'=====================進行剩餘移動牌之排列組合串整合
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
'=========================測試用匯出
If Formsetting.checktest.Value = 1 Then
'    Open App.Path & "\test\out2.txt" For Output As #1
    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & 戰鬥系統類.turn & "turn_" & 戰鬥系統類.turnatk & "_" & uscom & "_2.txt" For Output As #1
    For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
        Print #1, cardAInumMOVnmtot(i, 2)
    Next
    Print #1, cardAInumMOVmain(1, 5) & "=" & cardAInumMOVmain(1, 3)
    Close
    'MsgBox "已匯出完畢2"
End If
'==============================
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_三階段_進行估計期望值計算(ByVal uscom As Integer, ByVal name As String, ByVal choose As Integer, ByVal movecpre As Integer, ByVal pagenumber As Integer)
Dim weru As Integer, wertp As Integer, movecpren As Integer, turnm As Integer, werucount As Boolean, i As Integer, k As Integer, q As Integer, wp As Integer, wds As Integer
For i = 1 To 2 ^ Val(cardAInumMOVnmtot(0, 3))
    For k = 1 To 2
         '===========將資料轉移至待運算資料
         weru = 0
         For wp = 1 To pagenumber
              If Mid(cardAInumMOVnmtot(i, 2), wp, 1) = "n" Then
                  weru = Val(weru) + 1
              End If
         Next
         If Val(weru) > 0 Then
                 智慧型AI系統類.智慧型AI系統計算_一階段_初始 weru
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
                智慧型AI系統類.智慧型AI系統計算_一階段_取得牌面資料 False, uscom
                智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turnm, movecpren, uscom
                智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turnm, movecpren, uscom
                智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
                智慧型AI系統類.智慧型AI系統計算_四階段_比序_1_初始
                智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_1
                智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_2
                智慧型AI系統類.智慧型AI系統計算_四階段_比序_3_選擇組合
        Else
                cardAInumselect1 = 0
        End If
        '=======================將重新估計後資料儲存
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
'=========================測試用匯出
If Formsetting.checktest.Value = 1 Then
'    Open App.Path & "\test\out3.txt" For Output As #1
    Open App.Path & "\test\AIout" & Format(Now, "_yyyy-m-d_hh-mm-ss_") & 戰鬥系統類.turn & "turn_" & 戰鬥系統類.turnatk & "_" & uscom & "_3.txt" For Output As #1
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
    'MsgBox "已匯出完畢3"
End If
'==============================
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_四階段_統計估計期望值及判斷(ByVal uscom As Integer)
Dim atk1max As Integer, atk2max As Integer, defmax As Integer, chemax As Integer, chestr As String
Dim wtmovnum As Integer, i As Integer
'==================篩選是否符合移動量
For i = 14 * (角色人物對戰人數(uscom, 2) - 1) + 1 To 14 * 角色人物對戰人數(uscom, 2)
    If (人物異常狀態資料庫(uscom, i, 3) = 6 And uscom = 2) Or (人物異常狀態資料庫(uscom, i, 3) = 12 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) - Val(人物異常狀態資料庫(uscom, i, 1))
    End If
    If (人物異常狀態資料庫(uscom, i, 3) = 3 And uscom = 2) Or (人物異常狀態資料庫(uscom, i, 3) = 9 And uscom = 1) Then
        wtmovnum = Val(wtmovnum) + Val(人物異常狀態資料庫(uscom, i, 1))
    End If
    If (人物異常狀態資料庫(uscom, i, 3) = 16 And uscom = 1) Or (人物異常狀態資料庫(uscom, i, 3) = 17 And uscom = 2) Then
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
           智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_確認實行_選擇最終組合 1, atk1max
     Case 2
           cardAInumMOVFinal(3) = 2
           cardAInumMOVFinal(2) = atk2max
           智慧型AI系統類.智慧型AI系統計算_移動階段續_正向面_確認實行_選擇最終組合 2, atk2max
     Case 3
           cardAInumMOVFinal(1) = cardAInumMOVnmtot(2 ^ Val(cardAInumMOVnmtot(0, 3)), 2)
           cardAInumMOVFinal(3) = 3
           cardAInumMOVFinal(2) = defmax
End Select
End Sub
Sub 智慧型AI系統計算_移動階段續_正向面_確認實行_選擇最終組合(ByVal movche As Integer, ByVal atkmax As Integer)
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
        werpagenum = 0 '==目的取最大之出牌數
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
                wermovmaxnum = 0 '==目的取最大之移動數
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
Sub 智慧型AI系統計算_移動階段續_正向面_五階段_實行選牌(ByVal choose As Integer, ByVal uscom As Integer, ByVal pagenumber As Integer)
Dim wer As Integer '暫時變數
If choose = 1 Then
    wer = 0
Else
    wer = 1
End If
'=================
Dim pu As Integer, i As Integer, cspce As String, cspme As String '暫時變數
'=======================如組合符合出牌條件的話
Select Case uscom
     Case 1 '==使用者方
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
            '===================選擇行動
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      目前數(33) = 3
                 Case 2
                      目前數(33) = 1
                 Case 3
                      目前數(33) = 2
            End Select
     Case 2 '==電腦方
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
            '===================選擇行動
            Select Case cardAInumMOVFinal(3)
                 Case 1
                      電腦方移動階段選擇數 = 3
                 Case 2
                      電腦方移動階段選擇數 = 1
                 Case 3
                      電腦方移動階段選擇數 = 2
            End Select
End Select

是否移動階段續估計判斷程序 = False
End Sub
Sub 智慧型AI系統計算_引導程序_超出牌張數(ByVal uscom As Integer, ByVal turn As Integer, ByVal name As String, ByVal movecpre As Integer, ByVal choose As Integer)
Dim i As Integer, w As Integer
If Val(pageglead(uscom)) > 10 Then
    Erase cardAInumOvertenrecord
    智慧型AI系統類.智慧型AI系統計算_一階段_初始 10
    '=========擷取目前牌面資料(前10張)
        Select Case uscom
            Case 1
                戰鬥系統類.出牌順序計算_使用者_手牌
            Case 2
                戰鬥系統類.出牌順序計算_電腦_手牌
        End Select
        w = 2 * uscom '(2-使用者手牌/4-電腦手牌)
        For i = 1 To 10
            cardcountAInum(i, 5) = 出牌順序統計暫時變數(w, i, 2)
            cardAInumOvertenrecord(i) = 出牌順序統計暫時變數(w, i, 2)
            cardcountAInum(i, 1) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 1)
            cardcountAInum(i, 2) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 2)
            cardcountAInum(i, 3) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 3)
            cardcountAInum(i, 4) = pagecardnum(出牌順序統計暫時變數(w, i, 2), 4)
        Next
     '========================
    智慧型AI系統類.智慧型AI系統計算_一階段_取得牌面資料 False, uscom
    智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_初始 turn, movecpre, uscom
    智慧型AI系統類.智慧型AI系統計算_二階段_計算期望值_個別技能 name, turn, movecpre, uscom
    智慧型AI系統類.智慧型AI系統計算_三階段_統計排列
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_1_初始
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_1
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_2_超額比序判斷_2
    智慧型AI系統類.智慧型AI系統計算_暫時匯出 uscom
    智慧型AI系統類.智慧型AI系統計算_四階段_比序_3_選擇組合
    If turn = 3 And cardAInumchoose > 0 Then
        智慧型AI系統類.智慧型AI系統計算_引導程序_移動階段續 uscom, turn, name, movecpre, choose, 10
    Else
        智慧型AI系統類.智慧型AI系統計算_最後階段_實行選牌 choose, uscom
    End If
    '==========================
    If turn <> 3 Then
        戰鬥系統類.comatk_智慧型AI引導程序_超出牌張數 turn, movecpre, choose
    End If
End If
End Sub
Sub 檢查人物技能是否有EX技(ByVal uscom As Integer, ByVal name As String)
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
                If (人物異常狀態資料庫(uscom, k, 3) = 22 And uscom = 1) Or _
                    (人物異常狀態資料庫(uscom, k, 3) = 23 And uscom = 2) Then
                    personatkingtfr(5) = 1
                End If
          Next
     End If
Next
End Sub
Sub 智慧型AI系統_使用者出牌階段判斷反轉()
Dim i As Integer
For i = 1 To 106
    If Val(pagecardnum(i, 11)) = 4 And Val(pagecardnum(i, 5)) = 1 And Val(pagecardnum(i, 6)) = 1 Then
        FormMainMode.cgen_Click (i)
        pagecardnum(i, 11) = 3
    End If
Next
End Sub
Sub 智慧型AI系統計算_移動階段續_否定面_二階段_選擇行動(ByVal uscom As Integer)
Select Case uscom
    Case 1
        目前數(33) = 2
    Case 2
         電腦方移動階段選擇數 = 2
End Select
End Sub
