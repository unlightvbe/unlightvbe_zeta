Attribute VB_Name = "戰鬥系統類"
Public Const a1a As String = "ATK-劍"
Public Const a2a As String = "DEF"
Public Const a3a As String = "MOV"
Public Const a4a As String = "SPE"
Public Const a5a As String = "ATK-槍"
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

Public atkingno(1 To 8, 1 To 11) As String '技能發動排序暫時圖片路徑儲存變數(技能發動順序8~1,1.圖片路徑/2.(1)使用者/(2)電腦方/3.Left/4.Top(座標)/5.視窗寬度(Width)/6.視窗高度(Height)/7.技能編號/8.技能執行中時啟動值/9.技能執行中換圖片檢查值/10.第2張圖片路徑)
Public goicheck(1 To 2) As Integer   '攻擊/防禦模式加骰數值檢查碼
Public pageonin(1 To 106) As Integer  '牌張正反面檢查碼
Public liveus(1 To 3) As Integer, livecom(1 To 3) As Integer, liveusmax(1 To 3) As Integer, livecommax(1 To 3) As Integer
Public turn As Integer, atkus(1 To 3) As Integer, atkcom(1 To 3) As Integer, defus(1 To 3) As Integer, defcom(1 To 3) As Integer, pagecheckus As Integer, pagecheckcom As Integer, pagegive As Integer, goidefus As Integer, movecom As Integer, moveus As Integer, movecp As Integer, chkcomck As Integer, uslevel(1 To 3) As Integer, comlevel(1 To 3) As Integer, liveus41(1 To 3) As Integer, livecom41(1 To 3) As Integer, movecheckcom As Integer, movecheckus As Integer
Public nameus(1 To 3) As String, namecom(1 To 3) As String
Public moveturn As Integer  '攻擊／防禦模式先後檢查碼(1.使用者先攻/2.電腦先攻)
Public atkinghelpxy(1 To 2, 1 To 4, 1 To 2) As Integer '技能說明欄座標指定資料(1.使用者方/2.電腦方,第1~4個技能,1.Left/2.Top(座標))
Public pageusleadmax(0 To 1) As Integer   '使用者牌順序計數表(0.手牌/1.出牌)
Public pagecomleadmax(0 To 1) As Integer   '電腦牌順序計數表(0.手牌/1.出牌)
Public pageqlead(1 To 2) As Integer   '出牌計數變數(1.使用者/2.電腦)
Public pageglead(1 To 2) As Integer   '手牌計數變數(1.使用者/2.電腦)
Public movedsus As Integer   '使用者移動階段決定值變數
Public turnpageonin As Integer  '階段是否可出牌變數(一般)
Public turnpageoninatking As Integer  '階段是否可出牌變數(技能使用)
Public goickus As Integer '牌值一次檢查碼
Public atkingck(1 To 161, 1 To 2) As Integer '技能階段啟動碼(x.人物技能編號,1.技能執行階段/2.技能啟動檢查值)
Public atkingckai(1 To 140, 1 To 2) As Integer 'AI技能階段啟動碼(x.人物技能編號,1.技能執行階段/2.技能啟動檢查值)
Public atkingtrn(1 To 4) As Integer '技能計數器暫時儲存變數(1.使用者(現)/2.電腦(現)/3.使用者(備份)/4.電腦(備份))
Public akhpnm As Integer  '技能說明暫時變數
Public turnatk As Integer  '攻擊／防禦階段變數(1.使用者攻擊、電腦防禦,2.使用者防禦、電腦攻擊,3.發牌、移動)
Public trend暫時變數 As Integer '結束階段計數器暫時變數
Public HP檢查變數 As Boolean 'HP檢查階段是否已檢查變數
Public HP檢查階段數 As Integer 'HP檢查階段變數(1.移動階段後,2.攻擊/防禦階段前,3.攻/防禦階段後)
Public 距離單位(1 To 2, 1 To 2, 1 To 2) As Integer  '距離單位暫時儲存資料(1.HP血條/2.牌移動,1.使用者/2.電腦,1.Left單位/2.Top單位)
Public personminixy(1 To 2, 1 To 3, 1 To 3, 1 To 2) As Integer '小人物圖片座標指定資料(1.使用者/2.電腦,第n位,1.近距離/2.中距離/3.遠距離,1.Left/2.Top(座標))
Public 人物異常狀態資料庫(1 To 2, 1 To 42, 1 To 3) As Integer '異常狀態資料(1.使用者/2.電腦,第x個異常狀態,1.狀態數值/2.狀態統計數(剩餘回合/累計)/3.狀態編號)
Public 異常狀態檢查數(1 To 40, 1 To 2) As Integer '異常狀態啟動碼(x.異常狀態編號,1.狀態執行階段/2.狀態啟動檢查值)
Public 技能動畫顯示階段數 As Integer '技能動畫計數器階段碼(1.攻擊/防禦階段-普通,2.移動階段-普通/3.發牌階段後、移動階段前/4.移動階段後/5.攻擊階段後/6.防禦階段後/7.回合結束時)
Public 攻擊防禦骰子總數(1 To 4) As Integer '攻擊/防禦模式骰子數量資料(1.使用者(總)/2.電腦(總)/3.使用者(原)/4.電腦(原))
Public atkingpagetot(1 To 2, 1 To 5) As Integer  '每階段出牌種類及數值統計資料(1.使用者/2.電腦,1.劍/2.防/3.移/4.特/5.槍)
Public 骰數零檢查值(1 To 2) As Boolean '當前階段骰子數量是否為零檢查數(1.使用者/2.電腦)
Public pagecardnum(1 To 106, 1 To 11) As String '公用牌資料(第x編號(1~70-公牌/71~88-使用者事件牌/89~106-電腦事件牌),1.正面類型/2.正面數值/3.反面類型/4.反面數值/5.(1)使用者-(2)電腦/6.(1)手牌-(2)出牌-(3)藏牌-(4)牌堆/7.出牌順序/8.圖片編號/9.目前Left(座標)/10.目前Top(座標)/11.(1)電腦方出牌(裏)-(2)電腦發出牌(外))
Public 牌總階段數(1 To 3) As Integer '牌擁有總階段數(1.使用者/2.電腦/3.總計)
Public 牌移動暫時變數(1 To 3) As Long '牌移動計數器暫時變數(1.Left單位/2.Top單位/3.牌張編號)
Public 目前數(1 To 33) As Integer '總暫時變數
Public 出牌順序統計暫時變數(1 To 4, 1 To 106, 1 To 2) As Integer '出牌順序統計總暫時資料(1.使用者出牌/2.使用者手牌/3.電腦出牌/4.電腦手牌,第x順序,1.目前牌出牌順序/2.牌張編號)
Public 距離單位_收牌暫時數(1 To 106, 1 To 3) As Integer  '收牌個別距離單位暫時儲存變數(第x順序,1.Left單位/2.Top單位/3.牌張編號)
Public 階段狀態數 As Integer '每階段開始結束狀態檢查數(1.開始階段(使用者)/2.結束階段(使用者)/3.開始階段(電腦)/4.結束階段(電腦)/5.交換角色)
Public 小人物頭像移動方向數(1 To 2) As Integer '小人物頭像移動方向狀態數(1.使用者/2.電腦[1.向內,2.向外])
Public 血量計數器動畫暫時變數(1 To 2, 1 To 2) As Integer '開始初始階段-血量動畫計數器暫時變數(1.使用者血條/2.電腦血條,1.每次移動量/2.是否已完成)
Public 時間軸顏色變化紀錄暫時變數(1 To 4, 1 To 3) As Integer '時間軸進行顏色變化階段紀錄暫時變數(1~3(1)單位變化量(1(1).時間軸(內外))/2.目前累計量/3.目前顏色(R,G,B),4.(1)時間軸(外)階段數-(1)黑變紅-(2)紅變黑/2.目前累計量/3.目前顏色(R))
Public 人物技能數字指示 As Integer '戰鬥系統表單-atkingnumtot.Caption的變數表示
Public 開始卡片移動動畫完成數(1 To 2, 1 To 4) As Integer   '開始時每張卡片移動動畫完成紀錄數(1.使用者/2.電腦,1~3.卡片/4.目前第幾張)
Public 交換角色紀錄暫時變數(1 To 4) As Integer '交換角色雙方紀錄暫時數(1.使用者/2.電腦/3.是否當下首次/4.交換角色完執行階段數)
Public pageeventnum(1 To 2, 1 To 18, 1 To 2) As String '事件卡排列紀錄資料(1.使用者/2.電腦,1~18-編號,1.事件卡名稱/2.事件卡檔案名稱)
Public 擲骰後骰傷害數 As Integer '戰鬥系統表單-fm2.Caption的變數表示
Public 戰鬥模式勝敗紀錄數 As Integer '戰鬥系統當前勝敗紀錄暫時變數(1.使用者方勝利/2.使用者方敗北/3.平手)
Public 電腦方移動階段選擇數 As Integer '移動階段電腦方選擇之行動暫時變數
Public 電腦方事件卡是否出完選擇數 As Boolean '電腦方先出事件卡是否出完暫時紀錄
Public 人物卡面背面編號紀錄數(1 To 7) As Integer '人物卡片背面技能說明人物編號暫時變數(1.(1).使用者/(2).電腦,2.第n位,3.目前使用者方使用人物編號/4.目前選擇之技能編號(使用者方使用人物)/5.目前選擇之技能編號(其他)/6~7.目前選擇之技能編號(交換角色)
Public 擲骰表單溝通暫時變數(1 To 4) As Integer 'Form6表單值溝通暫時變數(1.一回合中先後判斷(1.前/2.後),2.原始骰值(使用者)-擲骰後有效傷害數,3.原始骰值(電腦)-擲骰後傷害對象(1.使用者/2.電腦),4.(1.使用者先攻/2.電腦先攻))
Public 人物消失檢查暫時變數(1 To 3) As Integer '人物消失檢查計數器紀錄暫時變數(1.目前計數/2.使用者標記/3.電腦標記)
Public 公用牌各牌類型紀錄數(0 To 31, 1 To 2) As Integer '各場景公用牌牌類型紀錄暫時變數(0.(1)目前已發牌總數量/(2)目前場景牌總數量,1~31.(1)目前已使用之牌數/(2)該牌型能使用之總數量)
Public 卡片人物資訊檔案讀取失敗紀錄串 As String '卡片人物資訊檔案讀取失敗時檔案名紀錄暫時變數
Sub 人物技能欄燈開關(ByVal k As Boolean, ByVal n As Integer)
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
Sub 人物異常狀態表設定_初設(ByVal 雙方 As Integer, ByVal 第幾個 As Integer, ByVal 異常狀態編號 As Integer, ByVal ph As String, ByVal num1 As Integer, ByVal num2 As Integer)
If Formsetting.chknewdefferent.Value = 1 Then
    ph = 執行動作_路徑使用新式異常狀態圖案(ph)
End If
'===================================
Select Case 雙方
    Case 1
        FormMainMode.personusspe(第幾個).異常狀態圖片 = ph
        FormMainMode.personusspe(第幾個).person_num = num1
        FormMainMode.personusspe(第幾個).person_turn = num2
        人物異常狀態資料庫(1, 第幾個, 1) = num1
        人物異常狀態資料庫(1, 第幾個, 2) = num2
        人物異常狀態資料庫(1, 第幾個, 3) = 異常狀態編號
        FormMainMode.personusspe(第幾個).Visible = True
    Case 2
        FormMainMode.personcomspe(第幾個).異常狀態圖片 = ph
        FormMainMode.personcomspe(第幾個).person_num = num1
        FormMainMode.personcomspe(第幾個).person_turn = num2
        人物異常狀態資料庫(2, 第幾個, 1) = num1
        人物異常狀態資料庫(2, 第幾個, 2) = num2
        人物異常狀態資料庫(2, 第幾個, 3) = 異常狀態編號
        FormMainMode.personcomspe(第幾個).Visible = True
End Select

End Sub
Function 執行動作_路徑使用新式異常狀態圖案(ByVal ph As String) As String
For i = 1 To Len(ph)
    If Mid(ph, i, 1) = "." Then
        ph = Mid(ph, 1, i - 1) & "new" & Right(ph, 4)
        Exit For
    End If
Next
執行動作_路徑使用新式異常狀態圖案 = ph
End Function
Sub 自動捲軸捲動()
FormMainMode.messageus.ListIndex = FormMainMode.messageus.ListCount - 1
End Sub
Sub 傷害執行_技能直傷_使用者(ByVal tot As Integer, ByVal num As Integer)
'===============================
異常狀態檢查數(35, 1) = 1
異常狀態.庇護_使用者 num, tot '(階段1)
'===============================
If atking_蕾_守護模式狀態啟動值 = False Then
    Select Case num
       Case 1
          If tot > 0 And liveus(角色人物對戰人數(1, 2)) > 0 Then
              If tot >= liveus(角色人物對戰人數(1, 2)) Then
                 FormMainMode.messageus.AddItem "您受到了" & liveus(角色人物對戰人數(1, 2)) & "點傷害。"
                 戰鬥系統類.自動捲軸捲動
                 FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption = 0
                 FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption = 0
                 liveus(角色人物對戰人數(1, 2)) = 0
                 FormMainMode.bloodnumus1.Caption = 0
                 FormMainMode.bloodlineout1.Width = 0
                 牌總階段數(1) = 牌總階段數(1) + 1
              Else
                 FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption = Val(FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption) - tot
                 FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption = Val(FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption) - tot
                 liveus(角色人物對戰人數(1, 2)) = liveus(角色人物對戰人數(1, 2)) - tot
                 FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
                 FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (距離單位(1, 1, 1) * tot)
                 FormMainMode.messageus.AddItem "您受到了" & tot & "點傷害。"
                 戰鬥系統類.自動捲軸捲動
              End If
              戰鬥系統類.播放傷害音樂
           End If
       Case Is > 1
           If tot > 0 And liveus(角色待機人物紀錄數(1, num)) > 0 Then
              If tot >= liveus(角色待機人物紀錄數(1, num)) Then
                 liveus(角色待機人物紀錄數(1, num)) = 0
                 If FormMainMode.uspi1(角色待機人物紀錄數(1, num)).Caption = "" Then
                     FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption = -liveusmax(角色待機人物紀錄數(1, num))
                     FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption = -liveusmax(角色待機人物紀錄數(1, num))
                 Else
                     FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption = 0
                     FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption = 0
                 End If
                 牌總階段數(1) = 牌總階段數(1) + 1
              Else
                 FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption = Val(FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption) - tot
                 liveus(角色待機人物紀錄數(1, num)) = liveus(角色待機人物紀錄數(1, num)) - tot
                 FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption = Val(FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption) - tot
              End If
           End If
    End Select
End If
End Sub
Sub 播放傷害音樂()
Select Case movecp
    Case 1
        FormMainMode.wmpse2.Controls.play
        一般系統類.檢查音樂播放 2
    Case Is >= 2
        FormMainMode.wmpse8.Controls.play
        一般系統類.檢查音樂播放 8
End Select
End Sub
Sub 執行動作_技能手動結束()
FormMainMode.atkingnumtot.Caption = Val(FormMainMode.atkingnumtot) - 1
FormMainMode.atkingtrtot.Interval = 20
FormMainMode.atkingtrtot.Enabled = True
End Sub
Sub 回復執行_使用者(ByVal tot As Integer, ByVal num As Integer)
Select Case num
   Case 1
         If liveus(角色人物對戰人數(1, 2)) > 0 And tot > 0 Then
               If liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2)) >= tot Then
                    FormMainMode.messageus.AddItem "你的HP恢復了" & tot & "點。"
                    FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + 距離單位(1, 1, 1) * tot
                    liveus(角色人物對戰人數(1, 2)) = Val(liveus(角色人物對戰人數(1, 2))) + tot
                    FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption = liveus(角色人物對戰人數(1, 2))
                    FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption = liveus(角色人物對戰人數(1, 2))
                    FormMainMode.bloodnumus1.Caption = liveus(角色人物對戰人數(1, 2))
                    戰鬥系統類.自動捲軸捲動
              ElseIf liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2)) < tot Then
                    If liveusmax(角色人物對戰人數(1, 2)) - liveus(角色人物對戰人數(1, 2)) > 0 Then
                       FormMainMode.messageus.AddItem "你的HP恢復了" & Val(liveusmax(角色人物對戰人數(1, 2))) - Val(liveus(角色人物對戰人數(1, 2))) & "點。"
                       FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width + 距離單位(1, 1, 1) * (Val(liveusmax(角色人物對戰人數(1, 2))) - Val(liveus(角色人物對戰人數(1, 2))))
                       liveus(角色人物對戰人數(1, 2)) = Val(liveusmax(角色人物對戰人數(1, 2)))
                       FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption = liveus(角色人物對戰人數(1, 2))
                       FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption = liveus(角色人物對戰人數(1, 2))
                       FormMainMode.bloodnumus1.Caption = liveus(角色人物對戰人數(1, 2))
                       戰鬥系統類.自動捲軸捲動
                    End If
              End If
        End If
   Case Is > 1
        If liveus(角色待機人物紀錄數(1, num)) > 0 And tot > 0 Then
               If liveusmax(角色待機人物紀錄數(1, num)) - liveus(角色待機人物紀錄數(1, num)) >= tot Then
                    liveus(角色待機人物紀錄數(1, num)) = Val(liveus(角色待機人物紀錄數(1, num))) + tot
                    FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption = liveus(角色待機人物紀錄數(1, num))
                    FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption = liveus(角色待機人物紀錄數(1, num))
              ElseIf liveusmax(角色待機人物紀錄數(1, num)) - liveus(角色待機人物紀錄數(1, num)) < tot Then
                    If liveusmax(角色待機人物紀錄數(1, num)) - liveus(角色待機人物紀錄數(1, num)) > 0 Then
                       liveus(角色待機人物紀錄數(1, num)) = Val(liveusmax(角色待機人物紀錄數(1, num)))
                       FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption = liveus(角色待機人物紀錄數(1, num))
                       FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption = liveus(角色待機人物紀錄數(1, num))
                    End If
              End If
        End If
End Select
End Sub
Sub 回復執行_電腦(ByVal tot As Integer, ByVal num As Integer)
Select Case num
   Case 1
         If livecom(角色人物對戰人數(2, 2)) > 0 And tot > 0 Then
               If livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2)) >= tot Then
                    FormMainMode.messageus.AddItem "對方的HP恢復了" & tot & "點。"
                    FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - 距離單位(1, 2, 1) * tot
                    livecom(角色人物對戰人數(2, 2)) = Val(livecom(角色人物對戰人數(2, 2))) + tot
                    FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = livecom(角色人物對戰人數(2, 2))
                    FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = livecom(角色人物對戰人數(2, 2))
                    FormMainMode.bloodnumcom1.Caption = livecom(角色人物對戰人數(2, 2))
                    戰鬥系統類.自動捲軸捲動
              ElseIf livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2)) < tot Then
                    If livecommax(角色人物對戰人數(2, 2)) - livecom(角色人物對戰人數(2, 2)) > 0 Then
                       FormMainMode.messageus.AddItem "對方的HP恢復了" & Val(livecommax(角色人物對戰人數(2, 2))) - Val(livecom(角色人物對戰人數(2, 2))) & "點。"
                       FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left - 距離單位(1, 2, 1) * (Val(livecommax(角色人物對戰人數(2, 2))) - Val(livecom(角色人物對戰人數(2, 2))))
                       livecom(角色人物對戰人數(2, 2)) = Val(livecommax(角色人物對戰人數(2, 2)))
                       FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = livecom(角色人物對戰人數(2, 2))
                       FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = livecom(角色人物對戰人數(2, 2))
                       FormMainMode.bloodnumcom1.Caption = livecom(角色人物對戰人數(2, 2))
                       戰鬥系統類.自動捲軸捲動
                    End If
              End If
        End If
   Case Is > 1
        If livecom(角色待機人物紀錄數(2, num)) > 0 And tot > 0 Then
               If livecommax(角色待機人物紀錄數(2, num)) - livecom(角色待機人物紀錄數(2, num)) >= tot Then
                    livecom(角色待機人物紀錄數(2, num)) = Val(livecom(角色待機人物紀錄數(2, num))) + tot
                    FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = Val(FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption) + tot
                    FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = Val(FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption) + tot
              ElseIf livecommax(角色待機人物紀錄數(2, num)) - livecom(角色待機人物紀錄數(2, num)) < tot Then
                       livecom(角色待機人物紀錄數(2, num)) = Val(livecommax(角色待機人物紀錄數(2, num)))
                       If FormMainMode.compi1(角色待機人物紀錄數(2, num)).Caption = "" Then
                            FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = 0
                            FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = 0
                       Else
                            FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = livecom(角色待機人物紀錄數(2, num))
                            FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = livecom(角色待機人物紀錄數(2, num))
                       End If
              End If
        End If
End Select
End Sub
Function 傷害執行_使用者(ByVal tot As Integer)
'===============================
異常狀態檢查數(35, 1) = 1
異常狀態.庇護_使用者 1, tot '(階段1)
'===============================
If tot > 0 And liveus(角色人物對戰人數(1, 2)) > 0 Then
      If tot >= liveus(角色人物對戰人數(1, 2)) Then
         FormMainMode.messageus.AddItem "您受到了" & liveus(角色人物對戰人數(1, 2)) & "點傷害。"
         戰鬥系統類.自動捲軸捲動
         FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption = 0
         FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption = 0
         liveus(角色人物對戰人數(1, 2)) = 0
         FormMainMode.bloodnumus1.Caption = 0
         FormMainMode.bloodlineout1.Width = 0
         牌總階段數(1) = 牌總階段數(1) + 1
      Else
         FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption = Val(FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption) - tot
         FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption = Val(FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption) - tot
         liveus(角色人物對戰人數(1, 2)) = liveus(角色人物對戰人數(1, 2)) - tot
         FormMainMode.bloodnumus1.Caption = Val(FormMainMode.bloodnumus1.Caption) - tot
         FormMainMode.bloodlineout1.Width = FormMainMode.bloodlineout1.Width - (距離單位(1, 1, 1) * tot)
         FormMainMode.messageus.AddItem "您受到了" & tot & "點傷害。"
         戰鬥系統類.自動捲軸捲動
      End If
戰鬥系統類.播放傷害音樂
End If
End Function
Sub 傷害執行_技能直傷_電腦(ByVal tot As Integer, ByVal num As Integer)
'===============================
異常狀態檢查數(36, 1) = 1
異常狀態.庇護_電腦 num, tot '(階段1)
'===============================
If atking_AI_蕾_守護模式狀態啟動值 = False Then
    Select Case num
        Case 1
           If tot > 0 And livecom(角色人物對戰人數(2, 2)) > 0 Then
                    If tot >= livecom(角色人物對戰人數(2, 2)) Then
                       FormMainMode.messageus.AddItem "對方受到了" & livecom(角色人物對戰人數(2, 2)) & "點傷害。"
                       戰鬥系統類.自動捲軸捲動
                       FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = 0
                       FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = 0
                       FormMainMode.bloodnumcom1.Caption = 0
                       livecom(角色人物對戰人數(2, 2)) = 0
                       FormMainMode.bloodlineout2.Left = 11580
                       牌總階段數(2) = 牌總階段數(2) + 1
                    Else
                       FormMainMode.messageus.AddItem "對方受到了" & Val(tot) & "點傷害。"
                       戰鬥系統類.自動捲軸捲動
                       FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = Val(FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption) - tot
                       FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = Val(FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption) - tot
                       FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
                       livecom(角色人物對戰人數(2, 2)) = livecom(角色人物對戰人數(2, 2)) - tot
                       FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (距離單位(1, 2, 1) * tot)
                    End If
            戰鬥系統類.播放傷害音樂
            End If
        Case Is > 1
           If tot > 0 And livecom(角色待機人物紀錄數(2, num)) > 0 Then
                    If tot >= livecom(角色待機人物紀錄數(2, num)) Then
                       If FormMainMode.compi1(角色待機人物紀錄數(2, num)).Caption = "" Then
                           FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = -livecommax(角色待機人物紀錄數(2, num))
                           FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = -livecommax(角色待機人物紀錄數(2, num))
                       Else
                           FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = 0
                           FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = 0
                       End If
                       livecom(角色待機人物紀錄數(2, num)) = 0
                       牌總階段數(2) = 牌總階段數(2) + 1
                    Else
                       FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = Val(FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption) - tot
                       FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = Val(FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption) - tot
                       livecom(角色待機人物紀錄數(2, num)) = livecom(角色待機人物紀錄數(2, num)) - tot
                    End If
            End If
    End Select
End If
End Sub
Function 傷害執行_電腦(ByVal tot As Integer)
'===============================
異常狀態檢查數(36, 1) = 1
異常狀態.庇護_電腦 1, tot '(階段1)
'===============================
If tot > 0 And livecom(角色人物對戰人數(2, 2)) > 0 Then
        If tot >= livecom(角色人物對戰人數(2, 2)) Then
           FormMainMode.messageus.AddItem "對方受到了" & livecom(角色人物對戰人數(2, 2)) & "點傷害。"
           戰鬥系統類.自動捲軸捲動
           FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = 0
           FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = 0
           FormMainMode.bloodnumcom1.Caption = 0
           livecom(角色人物對戰人數(2, 2)) = 0
           FormMainMode.bloodlineout2.Left = 11580
           牌總階段數(2) = 牌總階段數(2) + 1
        Else
           FormMainMode.messageus.AddItem "對方受到了" & Val(tot) & "點傷害。"
           戰鬥系統類.自動捲軸捲動
           FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = Val(FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption) - tot
           FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = Val(FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption) - tot
           FormMainMode.bloodnumcom1.Caption = Val(FormMainMode.bloodnumcom1.Caption) - tot
           livecom(角色人物對戰人數(2, 2)) = livecom(角色人物對戰人數(2, 2)) - tot
           FormMainMode.bloodlineout2.Left = FormMainMode.bloodlineout2.Left + (距離單位(1, 2, 1) * tot)
        End If
戰鬥系統類.播放傷害音樂
End If
End Function
Sub 執行動作_使用者_棄牌(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) - 1
    目前數(5) = pagecardnum(n, 7)
    pagecardnum(n, 6) = 3
'    戰鬥系統類.座標計算_使用者手牌
    牌移動暫時變數(1) = 240
    牌移動暫時變數(2) = 960
    牌移動暫時變數(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '指定目前Left(座標)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
'    牌順序增加_手牌_使用者 n
    目前數(15) = 4
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
End Sub
Sub 執行動作_牌組_回牌_使用者(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
'    目前數(9) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 1
    pagecardnum(n, 6) = 1
    戰鬥系統類.座標計算_使用者手牌
    牌移動暫時變數(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '指定目前Left(座標)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.公用牌回復正面 n
    戰鬥系統類.計算牌移動距離單位
    牌順序增加_手牌_使用者 n
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
End Sub
Sub 執行動作_電腦牌_偷牌_使用者(ByVal n As Integer)
    FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    目前數(9) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 1
    pagecardnum(n, 6) = 1
    戰鬥系統類.座標計算_使用者手牌
    牌移動暫時變數(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '指定目前Left(座標)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    牌順序增加_手牌_使用者 n
    目前數(15) = 2
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
End Sub
Sub 執行動作_使用者牌_偷牌_電腦(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pageusglead = Val(FormMainMode.pageusglead) - 1
    目前數(5) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 2
    pagecardnum(n, 6) = 1
    戰鬥系統類.座標計算_電腦手牌
    牌移動暫時變數(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '指定目前Left(座標)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    牌順序增加_手牌_電腦 n
    目前數(15) = 20
    戰鬥系統類.公用牌變背面
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
End Sub
Sub 執行動作_牌組_回牌_電腦(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
'    目前數(5) = pagecardnum(n, 7)
    pagecardnum(n, 5) = 2
    pagecardnum(n, 6) = 1
    戰鬥系統類.座標計算_電腦手牌
    牌移動暫時變數(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '指定目前Left(座標)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    牌順序增加_手牌_電腦 n
    戰鬥系統類.公用牌變背面
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
End Sub
Sub 執行動作_翻牌(ByVal n As Integer)
    FormMainMode.card(n).Width = 810
    FormMainMode.card(n).Height = 1260
    FormMainMode.card(n).Picture = LoadPicture(app_path & "card\" & pagecardnum(n, 8) & "-" & pageonin(n) & ".bmp")
    FormMainMode.card(n).Visible = True
    FormMainMode.wmpse4.Controls.stop
    FormMainMode.wmpse4.Controls.play
    一般系統類.檢查音樂播放 4
End Sub
Sub 座標計算_電腦出牌()
Dim xy As Long  '暫時變數(首牌Left)
If pageqlead(2) = 1 Then
    牌移動暫時變數(1) = 5260
    牌移動暫時變數(2) = 1120
ElseIf pageqlead(2) > 1 Then
    xy = (pageqlead(2) - 1) * 460
    牌移動暫時變數(1) = (Val(5260) - xy) + ((pageqlead(2) - 1) * Val(960))
    牌移動暫時變數(2) = 1120
End If

End Sub
Sub 座標計算_電腦手牌()
牌移動暫時變數(1) = 10560 - 240 * (Val(FormMainMode.pagecomglead) - 1) '計算Left座標
牌移動暫時變數(2) = -600 '指定Top座標
End Sub
Sub 座標計算_使用者出牌()
Dim xy As Long   '暫時變數(首牌Left)
If pageqlead(1) = 1 Then
    牌移動暫時變數(1) = 5260
    牌移動暫時變數(2) = 4840
ElseIf pageqlead(1) > 1 Then
    xy = (pageqlead(1) - 1) * 460
    牌移動暫時變數(1) = (Val(5260) - xy) + ((pageqlead(1) - 1) * Val(960))
    牌移動暫時變數(2) = 4840
End If

End Sub
Sub 座標計算_使用者手牌()
If Val(FormMainMode.pageusglead) <= 9 Then
    牌移動暫時變數(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 1) '計算Left座標
Else
   牌移動暫時變數(1) = 2640 + 900 * (Val(FormMainMode.pageusglead) - 10)
End If

If Val(FormMainMode.pageusglead) <= 9 Then
   牌移動暫時變數(2) = 6700 '指定Top座標
Else
   牌移動暫時變數(2) = 7980 '指定Top座標
End If
End Sub
Sub 牌順序增加_出牌_電腦(ByRef m As Integer)
pagecardnum(m, 7) = pagecomleadmax(1) + 1
pagecomleadmax(1) = pagecomleadmax(1) + 1
End Sub
Sub 牌順序增加_手牌_電腦(ByRef m As Integer)
pagecardnum(m, 7) = pagecomleadmax(0) + 1
pagecomleadmax(0) = pagecomleadmax(0) + 1
End Sub
Sub 牌順序增加_手牌_使用者(ByVal m As Integer)
pagecardnum(m, 7) = pageusleadmax(0) + 1
pageusleadmax(0) = pageusleadmax(0) + 1
End Sub
Sub 牌順序增加_出牌_使用者(ByRef m As Integer)
pagecardnum(m, 7) = pageusleadmax(1) + 1
pageusleadmax(1) = pageusleadmax(1) + 1
End Sub
Sub 執行動作_電腦_棄牌(ByVal n As Integer)
    FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) - 1
    目前數(9) = pagecardnum(n, 7)
    pagecardnum(n, 6) = 3
    牌移動暫時變數(1) = 240
    牌移動暫時變數(2) = 960
    牌移動暫時變數(3) = n
    pagecardnum(n, 9) = FormMainMode.card(n).Left  '指定目前Left(座標)
    pagecardnum(n, 10) = FormMainMode.card(n).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    目前數(15) = 5
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
End Sub
Sub 執行動作_洗牌_舊()
For g = 1 To 57
     If pagecardnum(g, 6) = 3 Then
         pagegive = Val(pagegive) - 1
         pagecardnum(g, 6) = 4
     End If
Next
FormMainMode.pageul = 57 - Val(pagegive)
End Sub
Sub 執行動作_洗牌()
For g = 1 To 57
     If pagecardnum(g, 6) = 3 Then
         公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) - 1
         pagecardnum(g, 6) = 4
         Select Case pagecardnum(g, 8)
            Case "021"  '==移1槍1類
                 公用牌各牌類型紀錄數(1, 1) = Val(公用牌各牌類型紀錄數(1, 1)) - 1
            Case "019"  '==移1槍2類
                 公用牌各牌類型紀錄數(2, 1) = Val(公用牌各牌類型紀錄數(2, 1)) - 1
            Case "017"  '==移1槍3類
                 公用牌各牌類型紀錄數(3, 1) = Val(公用牌各牌類型紀錄數(3, 1)) - 1
            Case "025"  '==移1盾1類
                 公用牌各牌類型紀錄數(4, 1) = Val(公用牌各牌類型紀錄數(4, 1)) - 1
            Case "024"  '==移1盾2類
                 公用牌各牌類型紀錄數(5, 1) = Val(公用牌各牌類型紀錄數(5, 1)) - 1
            Case "023"  '==移1盾3類
                 公用牌各牌類型紀錄數(6, 1) = Val(公用牌各牌類型紀錄數(6, 1)) - 1
            Case "026"  '==移2特3類
                 公用牌各牌類型紀錄數(7, 1) = Val(公用牌各牌類型紀錄數(7, 1)) - 1
            Case "027"  '==移3移3類
                 公用牌各牌類型紀錄數(8, 1) = Val(公用牌各牌類型紀錄數(8, 1)) - 1
            Case "001"  '==劍6劍6類
                 公用牌各牌類型紀錄數(9, 1) = Val(公用牌各牌類型紀錄數(9, 1)) - 1
            Case "011"  '==劍1槍1類
                 公用牌各牌類型紀錄數(10, 1) = Val(公用牌各牌類型紀錄數(10, 1)) - 1
            Case "007"  '==劍2槍1類
                 公用牌各牌類型紀錄數(11, 1) = Val(公用牌各牌類型紀錄數(11, 1)) - 1
            Case "006"  '==劍2槍2類
                 公用牌各牌類型紀錄數(12, 1) = Val(公用牌各牌類型紀錄數(12, 1)) - 1
            Case "004"  '==劍3槍3類
                 公用牌各牌類型紀錄數(13, 1) = Val(公用牌各牌類型紀錄數(13, 1)) - 1
            Case "028"  '==劍5槍5類
                 公用牌各牌類型紀錄數(14, 1) = Val(公用牌各牌類型紀錄數(14, 1)) - 1
            Case "012"  '==劍1盾1類
                 公用牌各牌類型紀錄數(15, 1) = Val(公用牌各牌類型紀錄數(15, 1)) - 1
            Case "009"  '==劍2盾1類
                 公用牌各牌類型紀錄數(16, 1) = Val(公用牌各牌類型紀錄數(16, 1)) - 1
            Case "008"  '==劍2盾2類
                 公用牌各牌類型紀錄數(17, 1) = Val(公用牌各牌類型紀錄數(17, 1)) - 1
            Case "005"  '==劍3盾3類
                 公用牌各牌類型紀錄數(18, 1) = Val(公用牌各牌類型紀錄數(18, 1)) - 1
            Case "013"  '==劍1特1類
                 公用牌各牌類型紀錄數(19, 1) = Val(公用牌各牌類型紀錄數(19, 1)) - 1
            Case "010"  '==劍2特1類
                 公用牌各牌類型紀錄數(20, 1) = Val(公用牌各牌類型紀錄數(20, 1)) - 1
            Case "003"  '==劍4特1類
                 公用牌各牌類型紀錄數(21, 1) = Val(公用牌各牌類型紀錄數(21, 1)) - 1
            Case "002"  '==劍5特2類
                 公用牌各牌類型紀錄數(22, 1) = Val(公用牌各牌類型紀錄數(22, 1)) - 1
            Case "015"  '==槍4槍4類
                 公用牌各牌類型紀錄數(23, 1) = Val(公用牌各牌類型紀錄數(23, 1)) - 1
            Case "020"  '==槍2特1類
                 公用牌各牌類型紀錄數(24, 1) = Val(公用牌各牌類型紀錄數(24, 1)) - 1
            Case "018"  '==槍3特2類
                 公用牌各牌類型紀錄數(25, 1) = Val(公用牌各牌類型紀錄數(25, 1)) - 1
            Case "016"  '==槍4特1類
                 公用牌各牌類型紀錄數(26, 1) = Val(公用牌各牌類型紀錄數(26, 1)) - 1
            Case "014"  '==槍5特2類
                 公用牌各牌類型紀錄數(27, 1) = Val(公用牌各牌類型紀錄數(27, 1)) - 1
            Case "022"  '==盾5盾5類
                 公用牌各牌類型紀錄數(28, 1) = Val(公用牌各牌類型紀錄數(28, 1)) - 1
            Case "029"  '==盾3特5類
                 公用牌各牌類型紀錄數(29, 1) = Val(公用牌各牌類型紀錄數(29, 1)) - 1
         End Select
     End If
Next
FormMainMode.pageul = Val(公用牌各牌類型紀錄數(0, 2)) - Val(公用牌各牌類型紀錄數(0, 1))
End Sub
Sub 執行動作_清除所有異常狀態_電腦()
For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
   If 人物異常狀態資料庫(2, i, 2) > 0 Then
      異常狀態檢查數(人物異常狀態資料庫(2, i, 3), 2) = 0
      人物異常狀態資料庫(2, i, 2) = 0
   End If
Next
戰鬥系統類.異常狀態繼承_電腦
End Sub
Sub 執行動作_清除所有異常狀態_使用者()
For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
   If 人物異常狀態資料庫(1, i, 2) > 0 Then
      異常狀態檢查數(人物異常狀態資料庫(1, i, 3), 2) = 0
      人物異常狀態資料庫(1, i, 2) = 0
   End If
Next
戰鬥系統類.異常狀態繼承_使用者
End Sub
Sub 執行動作_距離變更(ByVal m As Integer)
Dim anw(1 To 2) As Integer
Dim anh(1 To 2) As Integer
anw(1) = Val(FormMainMode.personusminijpg.小人物圖片width) / 2
anw(2) = Val(FormMainMode.personcomminijpg.小人物圖片width) / 2
anh(1) = Val(FormMainMode.personusminijpg.小人物圖片height)
anh(2) = Val(FormMainMode.personcomminijpg.小人物圖片height)
Select Case m
  Case 1
    FormMainMode.movejpg.小人物圖片 = app_path & "\gif\short.png"
    FormMainMode.movejpg.Left = 4440
    FormMainMode.movejpg.Top = 2520
'    formmainmode.personusminijpg.Left = personminixy(1, 角色人物對戰人數(1, 2), 1, 1)
'    formmainmode.personcomminijpg.Left = personminixy(2, 角色人物對戰人數(2, 2), 1, 1)
    FormMainMode.personusminijpg.Left = 4320 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 7080 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
  Case 2
    FormMainMode.movejpg.小人物圖片 = app_path & "\gif\middle.png"
    FormMainMode.movejpg.Left = 2880
    FormMainMode.movejpg.Top = 2000
'    formmainmode.personusminijpg.Left = personminixy(1, 角色人物對戰人數(1, 2), 2, 1)
'    formmainmode.personcomminijpg.Left = personminixy(2, 角色人物對戰人數(2, 2), 2, 1)
    FormMainMode.personusminijpg.Left = 2640 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 8680 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
  Case 3
    FormMainMode.movejpg.小人物圖片 = app_path & "\gif\long.png"
    FormMainMode.movejpg.Left = 1080
    FormMainMode.movejpg.Top = 2360
'    formmainmode.personusminijpg.Left = personminixy(1, 角色人物對戰人數(1, 2), 3, 1)
'    formmainmode.personcomminijpg.Left = personminixy(2, 角色人物對戰人數(2, 2), 3, 1)
    FormMainMode.personusminijpg.Left = 1040 - anw(1)
    FormMainMode.personusminijpg.Top = 5880 - anh(1)
    FormMainMode.personcomminijpg.Left = 10320 - anw(2)
    FormMainMode.personcomminijpg.Top = 5880 - anh(2)
End Select
'============以下是異常狀態檢查及啟動
異常狀態檢查數(33, 1) = 1
異常狀態.咒縛_使用者 m  '(階段1)
'=====
異常狀態檢查數(34, 1) = 1
異常狀態.咒縛_電腦 m  '(階段1)
'============
movecp = m
End Sub
Sub 計算牌移動距離單位()
If 牌移動暫時變數(1) >= pagecardnum(牌移動暫時變數(3), 9) Then
   距離單位(2, 1, 1) = (牌移動暫時變數(1) - pagecardnum(牌移動暫時變數(3), 9)) \ 12
Else
   距離單位(2, 1, 1) = -((pagecardnum(牌移動暫時變數(3), 9) - 牌移動暫時變數(1)) \ 12)
End If

If 牌移動暫時變數(2) >= pagecardnum(牌移動暫時變數(3), 10) Then
   距離單位(2, 1, 2) = (牌移動暫時變數(2) - pagecardnum(牌移動暫時變數(3), 10)) \ 12
Else
   距離單位(2, 1, 2) = -((pagecardnum(牌移動暫時變數(3), 10) - 牌移動暫時變數(2)) \ 12)
End If
End Sub
Sub 異常狀態繼承_使用者()
For k = 1 To 3
    For i = 14 * (角色待機人物紀錄數(1, k) - 1) + 1 To (14 * 角色待機人物紀錄數(1, k)) - 1
         If 人物異常狀態資料庫(1, i, 2) = 0 Then
             If 人物異常狀態資料庫(1, i + 1, 2) > 0 Then
                  FormMainMode.personusspe(i).異常狀態圖片 = FormMainMode.personusspe(i + 1).異常狀態圖片
                  FormMainMode.personusspe(i).person_num = FormMainMode.personusspe(i + 1).person_num
                  FormMainMode.personusspe(i).person_turn = FormMainMode.personusspe(i + 1).person_turn
                  人物異常狀態資料庫(1, i, 2) = 人物異常狀態資料庫(1, i + 1, 2)
                  人物異常狀態資料庫(1, i, 3) = 人物異常狀態資料庫(1, i + 1, 3)
                  人物異常狀態資料庫(1, i, 1) = 人物異常狀態資料庫(1, i + 1, 1)
                  For j = 1 To 3
                     人物異常狀態資料庫(1, i + 1, j) = 0
                  Next
                  FormMainMode.personusspe(i + 1).Visible = False
                  FormMainMode.personusspe(i).Visible = True
             Else
                  For j = 1 To 3
                     人物異常狀態資料庫(1, i, j) = 0
                  Next
                  FormMainMode.personusspe(i).Visible = False
             End If
        End If
    Next
Next
End Sub
Sub 異常狀態繼承_電腦()
For k = 1 To 3
    For i = 14 * (角色待機人物紀錄數(2, k) - 1) + 1 To (14 * 角色待機人物紀錄數(2, k)) - 1
          If 人物異常狀態資料庫(2, i, 2) = 0 Then
              If 人物異常狀態資料庫(2, i + 1, 2) > 0 Then
                  FormMainMode.personcomspe(i).異常狀態圖片 = FormMainMode.personcomspe(i + 1).異常狀態圖片
                  FormMainMode.personcomspe(i).person_num = FormMainMode.personcomspe(i + 1).person_num
                  FormMainMode.personcomspe(i).person_turn = FormMainMode.personcomspe(i + 1).person_turn
                  人物異常狀態資料庫(2, i, 2) = 人物異常狀態資料庫(2, i + 1, 2)
                  人物異常狀態資料庫(2, i, 3) = 人物異常狀態資料庫(2, i + 1, 3)
                  人物異常狀態資料庫(2, i, 1) = 人物異常狀態資料庫(2, i + 1, 1)
                  For j = 1 To 3
                     人物異常狀態資料庫(2, i + 1, j) = 0
                  Next
                  FormMainMode.personcomspe(i + 1).Visible = False
                  FormMainMode.personcomspe(i).Visible = True
              Else
                  For j = 1 To 3
                     人物異常狀態資料庫(2, i, j) = 0
                  Next
                  FormMainMode.personcomspe(i).Visible = False
              End If
          End If
    Next
Next
End Sub
Sub 特殊_史塔夏_殺戮狀態_使用者()
Select Case atking_史塔夏_殺戮模式狀態數(1)
   Case 1
            If atking_史塔夏_殺戮模式狀態數(5) = 0 Then
                atking_史塔夏_殺戮模式狀態數(3) = 攻擊防禦骰子總數(1)
                atking_史塔夏_殺戮模式狀態數(4) = 攻擊防禦骰子總數(1) * 2
                atking_史塔夏_殺戮模式狀態數(5) = 1
                攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) * 2
            ElseIf atking_史塔夏_殺戮模式狀態數(5) = 1 Then
                atking_史塔夏_殺戮模式狀態數(3) = atking_史塔夏_殺戮模式狀態數(3) + (攻擊防禦骰子總數(1) - atking_史塔夏_殺戮模式狀態數(4))
                攻擊防禦骰子總數(1) = atking_史塔夏_殺戮模式狀態數(3) * 2
                atking_史塔夏_殺戮模式狀態數(4) = atking_史塔夏_殺戮模式狀態數(3) * 2
            End If
    Case 2
           atking_史塔夏_殺戮模式狀態數(3) = 0
           atking_史塔夏_殺戮模式狀態數(4) = 0
           atking_史塔夏_殺戮模式狀態數(5) = 0
    Case 3
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.小人物圖片 = app_path & "gif\史塔夏\一般\Staciamini1.png"
            FormMainMode.personusminijpg.小人物影子圖片 = app_path & "gif\史塔夏\一般\Staciaminidown1.png"
            FormMainMode.personusminijpg.小人物影子Left = 10
            FormMainMode.personusminijpg.小人物影子top差 = -50
            FormDice.jpgus.大人物圖片 = app_path & "gif\史塔夏\一般\Staciaperson1.png"
            FormMainMode.顯示列1.使用者方小人物圖片 = app_path & "gif\史塔夏\一般\Staciaf1.png"
            atking_史塔夏_殺戮模式狀態數(2) = 0
            atking_史塔夏_殺戮模式狀態數(3) = 0
            atking_史塔夏_殺戮模式狀態數(4) = 0
            atking_史塔夏_殺戮模式狀態數(5) = 0
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personusminijpg.Visible = True
    Case 4
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciamini1.png"
            FormMainMode.personusminijpg.小人物影子圖片 = app_path & "gif\史塔夏\殺戮\Staciaminidown1.png"
            FormMainMode.personusminijpg.小人物影子Left = -90
            FormMainMode.personusminijpg.小人物影子top差 = -60
            FormDice.jpgus.大人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaperson1.png"
            FormMainMode.顯示列1.使用者方小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaf1.png"
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personusminijpg.Visible = True
End Select
End Sub
Sub 特殊_史塔夏_殺戮狀態_電腦()
Select Case atking_AI_史塔夏_殺戮模式狀態數(1)
   Case 1
            If atking_AI_史塔夏_殺戮模式狀態數(5) = 0 Then
                atking_AI_史塔夏_殺戮模式狀態數(3) = 攻擊防禦骰子總數(2)
                atking_AI_史塔夏_殺戮模式狀態數(4) = 攻擊防禦骰子總數(2) * 2
                atking_AI_史塔夏_殺戮模式狀態數(5) = 1
                攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) * 2
            ElseIf atking_AI_史塔夏_殺戮模式狀態數(5) = 1 Then
                atking_AI_史塔夏_殺戮模式狀態數(3) = atking_AI_史塔夏_殺戮模式狀態數(3) + (攻擊防禦骰子總數(2) - atking_AI_史塔夏_殺戮模式狀態數(4))
                攻擊防禦骰子總數(2) = atking_AI_史塔夏_殺戮模式狀態數(3) * 2
                atking_AI_史塔夏_殺戮模式狀態數(4) = atking_AI_史塔夏_殺戮模式狀態數(3) * 2
            End If
    Case 2
           atking_AI_史塔夏_殺戮模式狀態數(3) = 0
           atking_AI_史塔夏_殺戮模式狀態數(4) = 0
           atking_AI_史塔夏_殺戮模式狀態數(5) = 0
    Case 3
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.小人物圖片 = app_path & "gif\史塔夏\一般\Staciamini2.png"
            FormMainMode.personcomminijpg.小人物影子圖片 = app_path & "gif\史塔夏\一般\Staciaminidown2.png"
            FormMainMode.personcomminijpg.小人物影子Left = 10
            FormMainMode.personcomminijpg.小人物影子top差 = -50
            FormDice.jpgcom.大人物圖片 = app_path & "gif\史塔夏\一般\Staciaperson2.png"
            FormMainMode.顯示列1.電腦方小人物圖片 = app_path & "gif\史塔夏\一般\Staciaf2.png"
            atking_AI_史塔夏_殺戮模式狀態數(2) = 0
            atking_AI_史塔夏_殺戮模式狀態數(3) = 0
            atking_AI_史塔夏_殺戮模式狀態數(4) = 0
            atking_AI_史塔夏_殺戮模式狀態數(5) = 0
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personcomminijpg.Visible = True
    Case 4
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciamini2.png"
            FormMainMode.personcomminijpg.小人物影子圖片 = app_path & "gif\史塔夏\殺戮\Staciaminidown2.png"
            FormMainMode.personcomminijpg.小人物影子Left = 90
            FormMainMode.personcomminijpg.小人物影子top差 = -60
            FormDice.jpgcom.大人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaperson2.png"
            FormMainMode.顯示列1.電腦方小人物圖片 = app_path & "gif\史塔夏\殺戮\Staciaf2.png"
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personcomminijpg.Visible = True
End Select
End Sub

Sub 特殊_音音夢_成長狀態_使用者()
Select Case atking_音音夢_成長模式狀態數(1)
   Case 1
            攻擊防禦骰子總數(1) = 10
    Case 2
           攻擊防禦骰子總數(1) = 10
           戰鬥系統類.直接寫入顯示列數值 1, 10
    Case 3
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.小人物圖片 = app_path & "gif\音音夢\一般\Nenemmini1.png"
            FormMainMode.personusminijpg.小人物影子圖片 = app_path & "gif\音音夢\一般\Nenemminidown1.png"
            FormMainMode.personusminijpg.小人物影子Left = 10
            FormMainMode.personusminijpg.小人物影子top差 = -20
            FormDice.jpgus.大人物圖片 = app_path & "gif\音音夢\一般\Nenemperson1.png"
            FormMainMode.顯示列1.使用者方小人物圖片 = app_path & "gif\音音夢\一般\Nenemf1.png"
            atking_音音夢_成長模式狀態數(2) = 0
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personusminijpg.Visible = True
    Case 4
            FormMainMode.personusminijpg.Visible = False
            FormMainMode.personusminijpg.小人物圖片 = app_path & "gif\音音夢\成長\Nenemmini1.png"
            FormMainMode.personusminijpg.小人物影子圖片 = app_path & "gif\音音夢\成長\Nenemminidown1.png"
            FormMainMode.personusminijpg.小人物影子Left = 20
            FormMainMode.personusminijpg.小人物影子top差 = -90
            FormDice.jpgus.大人物圖片 = app_path & "gif\音音夢\成長\Nenemperson1.png"
            FormMainMode.顯示列1.使用者方小人物圖片 = app_path & "gif\音音夢\成長\Nenemf1.png"
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personusminijpg.Visible = True
End Select
End Sub
Sub 特殊_音音夢_成長狀態_電腦()
Select Case atking_AI_音音夢_成長模式狀態數(1)
   Case 1
            攻擊防禦骰子總數(2) = 10
    Case 2
           攻擊防禦骰子總數(2) = 10
           戰鬥系統類.直接寫入顯示列數值 2, 10
    Case 3
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.小人物圖片 = app_path & "gif\音音夢\一般\Nenemmini2.png"
            FormMainMode.personcomminijpg.小人物影子圖片 = app_path & "gif\音音夢\一般\Nenemminidown2.png"
            FormMainMode.personcomminijpg.小人物影子Left = 10
            FormMainMode.personcomminijpg.小人物影子top差 = -20
            FormDice.jpgcom.大人物圖片 = app_path & "gif\音音夢\一般\Nenemperson2.png"
            FormMainMode.顯示列1.電腦方小人物圖片 = app_path & "gif\音音夢\一般\Nenemf2.png"
            atking_AI_音音夢_成長模式狀態數(2) = 0
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personcomminijpg.Visible = True
    Case 4
            FormMainMode.personcomminijpg.Visible = False
            FormMainMode.personcomminijpg.小人物圖片 = app_path & "gif\音音夢\成長\Nenemmini2.png"
            FormMainMode.personcomminijpg.小人物影子圖片 = app_path & "gif\音音夢\成長\Nenemminidown2.png"
            FormMainMode.personcomminijpg.小人物影子Left = 20
            FormMainMode.personcomminijpg.小人物影子top差 = -90
            FormDice.jpgus.大人物圖片 = app_path & "gif\音音夢\成長\Nenemperson2.png"
            FormMainMode.顯示列1.使用者方小人物圖片 = app_path & "gif\音音夢\成長\Nenemf2.png"
            戰鬥系統類.執行動作_距離變更 movecp
            FormMainMode.personcomminijpg.Visible = True
End Select
End Sub

Sub 特殊_布勞_一般立繪更換_使用者()
Dim m As Integer
Randomize
m = Int(Rnd() * 3) + 1
Select Case m
    Case 1
       FormDice.jpgus.大人物圖片 = app_path & "gif\布勞\Blauperson1-1.png"
    Case 2
       FormDice.jpgus.大人物圖片 = app_path & "gif\布勞\Blauperson1-2.png"
    Case 3
       FormDice.jpgus.大人物圖片 = app_path & "gif\布勞\Blauperson1-3.png"
End Select
End Sub
Sub 特殊_布勞_一般立繪更換_電腦()
Dim m As Integer
Randomize
m = Int(Rnd() * 3) + 1
Select Case m
    Case 1
       FormDice.jpgcom.大人物圖片 = app_path & "gif\布勞\Blauperson2-1.png"
    Case 2
       FormDice.jpgcom.大人物圖片 = app_path & "gif\布勞\Blauperson2-2.png"
    Case 3
       FormDice.jpgcom.大人物圖片 = app_path & "gif\布勞\Blauperson2-3.png"
End Select
End Sub
Function 特殊_尤莉卡_檢查超載是否啟動_使用者() As Boolean
If atkingck(49, 2) = 1 And atking_尤莉卡_超載目前階段紀錄數(3) > 0 Then
    特殊_尤莉卡_檢查超載是否啟動_使用者 = True
Else
    特殊_尤莉卡_檢查超載是否啟動_使用者 = False
End If
End Function
Function 特殊_尤莉卡_檢查超載是否啟動_電腦() As Boolean
If atkingckai(139, 2) = 1 And atking_AI_尤莉卡_超載目前階段紀錄數(3) > 0 Then
    特殊_尤莉卡_檢查超載是否啟動_電腦 = True
Else
    特殊_尤莉卡_檢查超載是否啟動_電腦 = False
End If
End Function
Sub comatk_AI_雪莉_巨大黑犬_劍(ByVal i As Integer)
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
Sub comatk_AI_雪莉_飛刃雨_移(j As Integer)
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
Sub comatk_AI_傑多_因果之幻_移(j As Integer)
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
Sub comatk_AI_雪莉_自殺傾向_特(ByVal a As Integer)
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
Sub comatk_AI_雪莉_多妮妲_異質者_特(ByVal a As Integer)
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
Sub comatk_AI_蕾_終曲_無盡輪迴的終結_特(ByVal a As Integer)
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
Sub 直接寫入顯示列數值(ByVal n As Integer, ByVal num As Integer)
If num < 0 Then num = 0
Select Case n
    Case 1
        FormMainMode.顯示列1.goi1 = num
    Case 2
        FormMainMode.顯示列1.goi2 = num
End Select
End Sub
Sub 小人物頭像執行完判斷_使用者()
If turnatk = 1 Or turnatk = 2 Then
   turnpageonin = 1
'   階段狀態數 = 1
End If
If turnatk = 3 Then
    FormMainMode.trtimeline.Enabled = True
End If
End Sub
Sub 小人物頭像執行完判斷_電腦()
If turnatk = 1 Or turnatk = 2 Or turnatk = 3 Then
   階段狀態數 = 3
   FormMainMode.電腦出牌.Enabled = True
End If
End Sub
Sub 公用牌變背面()
FormMainMode.card(牌移動暫時變數(3)).Width = 720
FormMainMode.card(牌移動暫時變數(3)).Height = 990
FormMainMode.card(牌移動暫時變數(3)).Picture = LoadPicture(app_path & "card\cardback.bmp")
End Sub
Sub 公用牌回復正面(ByVal num As Integer)
FormMainMode.card(num).Width = 810
FormMainMode.card(num).Height = 1260
FormMainMode.card(num).Picture = LoadPicture(app_path & "card\" & pagecardnum(num, 8) & "-" & pageonin(num) & ".bmp")
End Sub
Sub 出牌順序計算_使用者_手牌()
Dim pagegustot As Integer '暫時變數

For i = 1 To 106
   For j = 1 To 2
      出牌順序統計暫時變數(2, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 1 Then
    pagegustot = Val(pagegustot) + 1
    出牌順序統計暫時變數(2, pagegustot, 1) = Val(pagecardnum(i, 7))
    出牌順序統計暫時變數(2, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If 出牌順序統計暫時變數(2, o, 1) > 出牌順序統計暫時變數(2, i, 1) Then
    g = 出牌順序統計暫時變數(2, i, 1)
    h = 出牌順序統計暫時變數(2, i, 2)
    出牌順序統計暫時變數(2, i, 1) = 出牌順序統計暫時變數(2, o, 1)
    出牌順序統計暫時變數(2, i, 2) = 出牌順序統計暫時變數(2, o, 2)
    出牌順序統計暫時變數(2, o, 1) = g
    出牌順序統計暫時變數(2, o, 2) = h
   End If
  Next
Next
'MsgBox 123
End Sub
Sub 出牌順序計算_使用者_出牌()
Dim pagegustot As Integer '暫時變數

For i = 1 To 106
   For j = 1 To 2
      出牌順序統計暫時變數(1, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 1 Then
    pagegustot = Val(pagegustot) + 1
    出牌順序統計暫時變數(1, pagegustot, 1) = Val(pagecardnum(i, 7))
    出牌順序統計暫時變數(1, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If 出牌順序統計暫時變數(1, o, 1) > 出牌順序統計暫時變數(1, i, 1) Then
    g = 出牌順序統計暫時變數(1, i, 1)
    h = 出牌順序統計暫時變數(1, i, 2)
    出牌順序統計暫時變數(1, i, 1) = 出牌順序統計暫時變數(1, o, 1)
    出牌順序統計暫時變數(1, i, 2) = 出牌順序統計暫時變數(1, o, 2)
    出牌順序統計暫時變數(1, o, 1) = g
    出牌順序統計暫時變數(1, o, 2) = h
   End If
  Next
Next

End Sub
Sub 出牌順序計算_電腦_手牌()
Dim pagegustot As Integer '暫時變數

For i = 1 To 106
   For j = 1 To 2
      出牌順序統計暫時變數(4, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 1 And Val(pagecardnum(i, 5)) = 2 Then
       pagegustot = Val(pagegustot) + 1
       出牌順序統計暫時變數(4, pagegustot, 1) = Val(pagecardnum(i, 7))
       出牌順序統計暫時變數(4, pagegustot, 2) = i
   ElseIf Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 1 Then
       pagegustot = Val(pagegustot) + 1
       出牌順序統計暫時變數(4, pagegustot, 1) = Val(pagecardnum(i, 7))
       出牌順序統計暫時變數(4, pagegustot, 2) = i
   End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If 出牌順序統計暫時變數(4, o, 1) > 出牌順序統計暫時變數(4, i, 1) Then
    g = 出牌順序統計暫時變數(4, i, 1)
    h = 出牌順序統計暫時變數(4, i, 2)
    出牌順序統計暫時變數(4, i, 1) = 出牌順序統計暫時變數(4, o, 1)
    出牌順序統計暫時變數(4, i, 2) = 出牌順序統計暫時變數(4, o, 2)
    出牌順序統計暫時變數(4, o, 1) = g
    出牌順序統計暫時變數(4, o, 2) = h
   End If
  Next
Next
End Sub
Sub 出牌順序計算_電腦_出牌()
Dim pagegustot As Integer '暫時變數

For i = 1 To 106
   For j = 1 To 2
      出牌順序統計暫時變數(3, i, j) = 0
   Next
Next

For i = 1 To 106
   If Val(pagecardnum(i, 6)) = 2 And Val(pagecardnum(i, 5)) = 2 And Val(pagecardnum(i, 11)) = 2 Then
       pagegustot = Val(pagegustot) + 1
       出牌順序統計暫時變數(3, pagegustot, 1) = Val(pagecardnum(i, 7))
       出牌順序統計暫時變數(3, pagegustot, 2) = i
    End If
Next

For o = 1 To Val(pagegustot) - 1
  For i = o + 1 To Val(pagegustot)
   If 出牌順序統計暫時變數(3, o, 1) > 出牌順序統計暫時變數(3, i, 1) Then
    g = 出牌順序統計暫時變數(3, i, 1)
    h = 出牌順序統計暫時變數(3, i, 2)
    出牌順序統計暫時變數(3, i, 1) = 出牌順序統計暫時變數(3, o, 1)
    出牌順序統計暫時變數(3, i, 2) = 出牌順序統計暫時變數(3, o, 2)
    出牌順序統計暫時變數(3, o, 1) = g
    出牌順序統計暫時變數(3, o, 2) = h
   End If
  Next
Next
End Sub
Sub 收牌計算距離單位_使用者()
For i = 1 To 106
    距離單位_收牌暫時數(i, 1) = 0
    距離單位_收牌暫時數(i, 2) = 0
Next

戰鬥系統類.出牌順序計算_使用者_出牌
For i = 1 To pageqlead(1)
    牌移動暫時變數(1) = 240
    牌移動暫時變數(2) = 960
    牌移動暫時變數(3) = 出牌順序統計暫時變數(1, i, 2)
    pagecardnum(出牌順序統計暫時變數(1, i, 2), 9) = FormMainMode.card(出牌順序統計暫時變數(1, i, 2)).Left  '指定目前Left(座標)
    pagecardnum(出牌順序統計暫時變數(1, i, 2), 10) = FormMainMode.card(出牌順序統計暫時變數(1, i, 2)).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    距離單位_收牌暫時數(i, 1) = 距離單位(2, 1, 1)
    距離單位_收牌暫時數(i, 2) = 距離單位(2, 1, 2)
    距離單位_收牌暫時數(i, 3) = 出牌順序統計暫時變數(1, i, 2)
Next
End Sub
Sub 收牌計算距離單位_電腦()
For i = 1 To 106
    距離單位_收牌暫時數(i, 1) = 0
    距離單位_收牌暫時數(i, 2) = 0
Next

戰鬥系統類.出牌順序計算_電腦_出牌
For i = 1 To pageqlead(2)
    牌移動暫時變數(1) = 240
    牌移動暫時變數(2) = 960
    牌移動暫時變數(3) = 出牌順序統計暫時變數(3, i, 2)
    pagecardnum(出牌順序統計暫時變數(3, i, 2), 9) = FormMainMode.card(出牌順序統計暫時變數(3, i, 2)).Left  '指定目前Left(座標)
    pagecardnum(出牌順序統計暫時變數(3, i, 2), 10) = FormMainMode.card(出牌順序統計暫時變數(3, i, 2)).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    距離單位_收牌暫時數(i, 1) = 距離單位(2, 1, 1)
    距離單位_收牌暫時數(i, 2) = 距離單位(2, 1, 2)
    距離單位_收牌暫時數(i, 3) = 出牌順序統計暫時變數(3, i, 2)
Next
End Sub
Sub 技能啟動數量檢查()
FormMainMode.atkingnumtot.Caption = Val(atkingtrn(1)) + Val(atkingtrn(2))
Erase atkingno
End Sub
Sub 技能說明載入_使用者(ByVal n As Integer)
Dim ahmt As String
FormMainMode.atkinghelpt1.Caption = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 2)
FormMainMode.atkinghelpt2.Caption = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 3)
FormMainMode.atkinghelpt3.Caption = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 4)
ahmt = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 5)
For i = 1 To Len(ahmt)
    If Mid(ahmt, i, 1) = "&" Then
        Mid(ahmt, i, 1) = Chr(10)
    End If
Next
FormMainMode.atkinghelpt4.Caption = ahmt
If VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 6) <> "" Then
    FormMainMode.atkinghelpt3.FontSize = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 6))
Else
    FormMainMode.atkinghelpt3.FontSize = 10
End If
If VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 7) <> "" Then
    FormMainMode.atkinghelpt4.FontSize = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 7))
Else
    FormMainMode.atkinghelpt4.FontSize = 10
End If
End Sub
Sub 技能說明載入_電腦(ByVal n As Integer)
Dim ahmt As String
FormMainMode.atkinghelpt1.Caption = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 2)
FormMainMode.atkinghelpt2.Caption = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 3)
FormMainMode.atkinghelpt3.Caption = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 4)
ahmt = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 5)
For i = 1 To Len(ahmt)
    If Mid(ahmt, i, 1) = "&" Then
        Mid(ahmt, i, 1) = Chr(10)
    End If
Next
FormMainMode.atkinghelpt4.Caption = ahmt

If VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 6) <> "" Then
    FormMainMode.atkinghelpt3.FontSize = Val(VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 6))
Else
    FormMainMode.atkinghelpt3.FontSize = 10
End If
If VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 7) <> "" Then
    FormMainMode.atkinghelpt4.FontSize = Val(VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 7))
Else
    FormMainMode.atkinghelpt4.FontSize = 10
End If
End Sub
Sub 音量靜音調節設定()
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
Sub 時間軸_重設()
FormMainMode.timelineout1.X1 = 0
FormMainMode.timelineout2.X2 = 11310
時間軸顏色變化紀錄暫時變數(1, 1) = 23
時間軸顏色變化紀錄暫時變數(1, 2) = 77
時間軸顏色變化紀錄暫時變數(1, 3) = 0
時間軸顏色變化紀錄暫時變數(2, 1) = 0
時間軸顏色變化紀錄暫時變數(2, 2) = 0
時間軸顏色變化紀錄暫時變數(2, 3) = 0
時間軸顏色變化紀錄暫時變數(3, 1) = 111
時間軸顏色變化紀錄暫時變數(3, 2) = 251
時間軸顏色變化紀錄暫時變數(3, 3) = 50
FormMainMode.timelineout1.BorderColor = RGB(111, 251, 50)
FormMainMode.timelineout2.BorderColor = RGB(111, 251, 50)
End Sub
Sub 時間軸_停止()
FormMainMode.trtimeline.Enabled = False
FormMainMode.timelinein1.BorderColor = RGB(0, 0, 0)
FormMainMode.timelinein2.BorderColor = RGB(0, 0, 0)
End Sub
Sub 時間軸_隱藏()
FormMainMode.timeup.Visible = False
FormMainMode.timelinein1.Visible = False
FormMainMode.timelinein2.Visible = False
FormMainMode.timelineout1.Visible = False
FormMainMode.timelineout2.Visible = False
End Sub
Sub 時間軸_顯示()
FormMainMode.timeup.Visible = True
FormMainMode.timelinein1.Visible = True
FormMainMode.timelinein2.Visible = True
FormMainMode.timelineout1.Visible = True
FormMainMode.timelineout2.Visible = True
End Sub
Sub 階段執行判斷()
If Val(擲骰表單溝通暫時變數(4)) = 1 Then
   Select Case Val(擲骰表單溝通暫時變數(1))
    Case 1
       If 擲骰表單溝通暫時變數(4) = 1 Then
'           cn3.Visible = True
           擲骰表單溝通暫時變數(1) = 2
           目前數(22) = 14
           FormMainMode.等待時間.Enabled = True
       Else
'           cn4.Visible = True
           目前數(22) = 15
           FormMainMode.等待時間.Enabled = True
       End If
    Case 2
       If 擲骰表單溝通暫時變數(4) = 1 Then
'          cn4.Visible = True
          目前數(22) = 15
          FormMainMode.等待時間.Enabled = True
       Else
'          cn2.Visible = True
          擲骰表單溝通暫時變數(1) = 2
          目前數(22) = 13
          FormMainMode.等待時間.Enabled = True
       End If
    End Select
Else
   Select Case Val(擲骰表單溝通暫時變數(1))
    Case 1
       If 擲骰表單溝通暫時變數(4) = 1 Then
'          cn4.Visible = True
          目前數(22) = 15
          FormMainMode.等待時間.Enabled = True
       Else
'          cn2.Visible = True
          擲骰表單溝通暫時變數(1) = 2
          目前數(22) = 13
          FormMainMode.等待時間.Enabled = True
       End If
    Case 2
       If 擲骰表單溝通暫時變數(4) = 1 Then
'           cn3.Visible = True
           擲骰表單溝通暫時變數(1) = 2
           目前數(22) = 14
           FormMainMode.等待時間.Enabled = True
       Else
'           cn4.Visible = True
           目前數(22) = 15
           FormMainMode.等待時間.Enabled = True
       End If
    End Select
  End If
End Sub
Sub 電腦牌_模擬按牌(ByVal Index As Integer)
If pagecardnum(Index, 6) = 1 And pagecardnum(Index, 5) = 2 Then
   pagecardnum(Index, 6) = 2
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp = 1 Then
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp > 1 Then
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + Val(pagecardnum(Index, 2))
      If turnatk = 1 Then
         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2))
         攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) + Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) + Val(pagecardnum(Index, 2))
   End If
   '===================
    目前數(9) = pagecardnum(Index, 7)
    pagecardnum(Index, 7) = Val(pagecomleadmax(1)) + 1
    pagecomleadmax(1) = Val(pagecomleadmax(1)) + 1
    pageqlead(2) = Val(pageqlead(2)) + 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) - 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) + 1
    pagecardnum(Index, 11) = 2
   '===================以下是出牌對齊
    目前數(7) = 0
    戰鬥系統類.出牌順序計算_電腦_出牌
    FormMainMode.電腦出牌_出牌對齊_靠左.Enabled = True
    '============以下是技能檢查及啟動
    atkingckai(1, 1) = 2
    If turnatk = 2 Then
       AI技能.雪莉_自殺傾向 Index '(階段2)
       AI技能.音音夢_愉快抽血 Index '(階段1)
    End If
    If turnatk = 2 And atkingckai(26, 2) = 1 Then
        atkingckai(26, 1) = 2
        AI技能.艾依查庫_神速之劍 Index '(階段2)
        atkingckai(26, 1) = 1
    End If
    If turnatk = 2 And atkingckai(98, 2) = 1 Then
        atkingckai(98, 1) = 2
        AI技能.露緹亞_渦騎劍閃 Index  '(階段2)
        atkingckai(98, 1) = 1
    End If
   '=============以下是牌移動(出牌)(電腦)
    戰鬥系統類.座標計算_電腦出牌
    牌移動暫時變數(3) = Index
    pagecardnum(Index, 9) = FormMainMode.card(Index).Left  '指定目前Left(座標)
    pagecardnum(Index, 10) = FormMainMode.card(Index).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    目前數(15) = 0
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
   '================以下是手牌對齊
   目前數(8) = 0
   目前數(17) = 1
   '===================以下是事件卡檢查及啟動
   If pagecardnum(Index, 1) = a6a Then
       事件卡記錄暫時數(2, 3) = 1
       事件卡.機會_電腦 Index, pagecardnum(Index, 2)
   End If
   If turnatk = 1 Or turnatk = 2 Then
        If pagecardnum(Index, 1) = a7a Then
            事件卡記錄暫時數(2, 3) = 1
            事件卡.詛咒術_電腦 Index, pagecardnum(Index, 2)
        End If
   End If
   If pagecardnum(Index, 1) = a8a Then
       事件卡記錄暫時數(2, 3) = 1
       事件卡.HP回復_電腦 Index, pagecardnum(Index, 2)
   End If
   If pagecardnum(Index, 1) = a9a Then
       事件卡記錄暫時數(2, 3) = 1
       事件卡.聖水_電腦 Index, pagecardnum(Index, 2)
   End If
   '===================
End If

End Sub
Sub 電腦牌_模擬按牌_外(ByVal Index As Integer)
If pagecardnum(Index, 6) = 2 And pagecardnum(Index, 5) = 2 Then
   pagecardnum(Index, 6) = 1
   If pagecardnum(Index, 1) = a1a Then
      atkingpagetot(2, 1) = Val(atkingpagetot(2, 1)) - Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp = 1 Then
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - Val(pagecardnum(Index, 2))
      If turnatk = 2 And movecp > 1 Then
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - Val(pagecardnum(Index, 2))
      If turnatk = 1 Then
         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 2))
         攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - Val(pagecardnum(Index, 2))
   End If
   If pagecardnum(Index, 1) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - Val(pagecardnum(Index, 2))
   End If
   '================
   目前數(9) = pagecardnum(Index, 7)
    pagecardnum(Index, 7) = Val(pagecomleadmax(0)) + 1
    pagecomleadmax(0) = Val(pagecomleadmax(0)) + 1
    pageqlead(2) = Val(pageqlead(2)) - 1
    FormMainMode.pagecomglead = Val(FormMainMode.pagecomglead) + 1
    FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
    pagecardnum(Index, 11) = 0
   '============以下是技能檢查及啟動
    atkingckai(1, 1) = 2
    If turnatk = 2 Then
       AI技能.雪莉_自殺傾向 Index '(階段2)
       AI技能.音音夢_愉快抽血 Index '(階段1)
    End If
    If turnatk = 2 And atkingckai(26, 2) = 1 Then
        atkingckai(26, 1) = 2
        AI技能.艾依查庫_神速之劍 Index '(階段2)
        atkingckai(26, 1) = 1
    End If
    If turnatk = 2 And atkingckai(98, 2) = 1 Then
        atkingckai(98, 1) = 2
        AI技能.露緹亞_渦騎劍閃 Index  '(階段2)
        atkingckai(98, 1) = 1
    End If
   '=============以下是牌移動(回牌)(電腦)
    戰鬥系統類.座標計算_電腦手牌
    牌移動暫時變數(3) = Index
    pagecardnum(Index, 9) = FormMainMode.card(Index).Left  '指定目前Left(座標)
    pagecardnum(Index, 10) = FormMainMode.card(Index).Top  '指定目前Top(座標)
    戰鬥系統類.計算牌移動距離單位
    戰鬥系統類.公用牌變背面
    目前數(15) = 0
    FormMainMode.牌移動.Enabled = True
    FormMainMode.wmpse1.Controls.stop
    FormMainMode.wmpse1.Controls.play
    一般系統類.檢查音樂播放 1
   '================以下是出牌對齊
   目前數(7) = 0
   戰鬥系統類.出牌順序計算_電腦_出牌
   FormMainMode.電腦出牌_出牌對齊_靠右.Enabled = True
   '=====================以下是技能檢查及啟動(克頓-竊取資料)
   If turnatk = 2 And atkingck(157, 2) = 1 And atkingck(157, 1) = 5 Then
        技能.克頓_竊取資料 '(階段5)
   End If
    '====================
End If
End Sub
Sub 電腦牌_模擬轉牌_外(ByVal Index As Integer)
uspce = pagecardnum(Index, 1)
uspme = pagecardnum(Index, 2)
pagecardnum(Index, 1) = pagecardnum(Index, 3)
pagecardnum(Index, 2) = pagecardnum(Index, 4)
pagecardnum(Index, 3) = uspce
pagecardnum(Index, 4) = uspme
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
一般系統類.檢查音樂播放 3
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
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) + pagecardnum(Index, 2)
      If turnatk = 2 And movecp > 1 Then
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(pagecardnum(Index, 2))
      End If
   End If
   If pagecardnum(Index, 1) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) + pagecardnum(Index, 2)
      If turnatk = 1 Then
         攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + Val(pagecardnum(Index, 2))
         攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + Val(pagecardnum(Index, 2))
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
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 4))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a5a Then
      atkingpagetot(2, 5) = Val(atkingpagetot(2, 5)) - pagecardnum(Index, 4)
      If turnatk = 2 And movecp > 1 Then
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 4))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a2a Then
      atkingpagetot(2, 2) = Val(atkingpagetot(2, 2)) - pagecardnum(Index, 4)
      If turnatk = 1 Then
          攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) - Val(pagecardnum(Index, 4))
          攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) - Val(pagecardnum(Index, 4))
      End If
   End If
   If pagecardnum(Index, 3) = a3a Then
      atkingpagetot(2, 3) = Val(atkingpagetot(2, 3)) - pagecardnum(Index, 4)
   End If
   If pagecardnum(Index, 3) = a4a Then
      atkingpagetot(2, 4) = Val(atkingpagetot(2, 4)) - pagecardnum(Index, 4)
   End If
    '============以下是技能檢查及啟動
    If turnatk = 2 Then
        atkingckai(26, 1) = 3
        AI技能.艾依查庫_神速之劍 Index '(階段3)
        atkingckai(1, 1) = 3
        AI技能.雪莉_自殺傾向 Index  '(階段3)
        atkingckai(111, 1) = 2
        AI技能.音音夢_愉快抽血 Index '(階段2)
    End If
    If turnatk = 2 Then
        atkingckai(98, 1) = 3
        AI技能.露緹亞_渦騎劍閃 Index '(階段3)
    End If
    '=================
    atkingckai(1, 1) = 1
    atkingckai(111, 1) = 1
    Call FormMainMode.pagecomqlead_Change
End Sub
Sub 骰數零執行判斷()
FormDice.outprocess
End Sub
Sub 雙方HP檢查()
Dim inp As Integer 'RND暫時變數
Dim person(1 To 2) As Integer
Erase 人物消失檢查暫時變數
If livecom(角色人物對戰人數(2, 2)) <= 0 Then
   人物消失檢查暫時變數(3) = 1
   If livecom(角色待機人物紀錄數(2, 2)) > 0 Then
'       人物交換_電腦_指定交換 2
       person(2) = 2
       交換角色紀錄暫時變數(2) = 1
'       牌總階段數(2) = 牌總階段數(2) + 1
   ElseIf livecom(角色待機人物紀錄數(2, 3)) > 0 Then
'       人物交換_電腦_指定交換 3
       交換角色紀錄暫時變數(2) = 1
       person(2) = 2
'       牌總階段數(2) = 牌總階段數(2) + 1
   Else
       person(2) = 1
   End If
End If
If Val(FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption) <= 0 Then
   人物消失檢查暫時變數(2) = 1
   If Val(FormMainMode.usbi1(角色待機人物紀錄數(1, 2)).Caption) > 0 Or Val(FormMainMode.usbi1(角色待機人物紀錄數(1, 3)).Caption) > 0 Then
'       執行動作_交換人物角色_初始
       person(1) = 2
       交換角色紀錄暫時變數(1) = 1
'       牌總階段數(1) = 牌總階段數(1) + 1
   Else
       person(1) = 1
   End If
End If

If person(1) = 2 Or person(2) = 2 Then
   目前數(22) = 21
   FormMainMode.人物消失檢查.Enabled = True
   Exit Sub
ElseIf person(1) = 0 And person(2) = 1 Then
   戰鬥模式勝敗紀錄數 = 1
   目前數(22) = 36
   FormMainMode.人物消失檢查.Enabled = True
ElseIf person(1) = 1 And person(2) = 0 Then
   目前數(22) = 36
   戰鬥模式勝敗紀錄數 = 2
   FormMainMode.人物消失檢查.Enabled = True
ElseIf person(1) = 1 And person(2) = 1 Then
   Randomize
   inp = Int(Rnd() * 2) + 1
   Select Case inp
       Case 1
           戰鬥模式勝敗紀錄數 = 1
           目前數(22) = 36
           FormMainMode.人物消失檢查.Enabled = True
       Case 2
           戰鬥模式勝敗紀錄數 = 2
           目前數(22) = 36
           FormMainMode.人物消失檢查.Enabled = True
    End Select
End If

If FormMainMode.人物消失檢查.Enabled = False Then
  Select Case HP檢查階段數
     Case 1
       '----------以下為階段繼續實行（移動階段3）
        目前數(22) = 4
        FormMainMode.等待時間.Enabled = True
     Case 2
'         atkingnumtot = 0
          目前數(22) = 11
          FormMainMode.等待時間.Enabled = True
     Case 3
        戰鬥系統類.階段執行判斷
     Case 4
        FormMainMode.NextTurn_階段2.Enabled = True
'     Case 5
'        目前數(26) = 1
'        formmainmode.骰子執行完啟動.Enabled = True
  End Select
End If
End Sub
Function 雙方HP檢查_結束回合檢查() As Boolean
Dim num(1 To 2) As Integer '選擇人物暫時變數
If turn >= Val(Formsetting.ckendturnnum.Text) And Formsetting.ckendturn.Value = 1 Then
        雙方HP檢查_結束回合檢查 = True
        '==============
        For i = 1 To 3
            If liveus(角色待機人物紀錄數(1, i)) > 0 Then
                num(1) = Val(num(1)) + Val(liveus(角色待機人物紀錄數(1, i)))
            End If
            If livecom(角色待機人物紀錄數(2, i)) > 0 Then
                num(2) = Val(num(2)) + Val(livecom(角色待機人物紀錄數(2, i)))
            End If
         Next
        '==============
        If num(1) > num(2) Then
           戰鬥模式勝敗紀錄數 = 1
           FormMainMode.trend.Enabled = True
        ElseIf num(1) < num(2) Then
           戰鬥模式勝敗紀錄數 = 2
           FormMainMode.trend.Enabled = True
        ElseIf num(1) = num(2) Then
            '無條件敗北
            戰鬥模式勝敗紀錄數 = 2
            FormMainMode.trend.Enabled = True
        End If
Else
     雙方HP檢查_結束回合檢查 = False
End If
End Function

Sub checkpage()

For i = 1 To 目前數(11)
  If 目前數(10) = 1 Then
   FormMainMode.pageusqlead = Val(FormMainMode.pageusqlead) - 1
   pageqlead(1) = Val(pageqlead(1)) - 1
  ElseIf 目前數(10) = 2 Then
   FormMainMode.pagecomqlead = Val(FormMainMode.pagecomqlead) - 1
   pageqlead(2) = Val(pageqlead(2)) - 1
  End If
Next
End Sub
Sub chkcom()
If goicheck(2) = 0 Then
  If atkingpagetot(2, 1) > 0 And movecp = 1 Then
    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkcom(角色人物對戰人數(2, 2))
    攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
    goicheck(2) = 1
  ElseIf atkingpagetot(2, 5) > 0 And movecp > 1 Then
    攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + atkcom(角色人物對戰人數(2, 2))
    攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + atkcom(角色人物對戰人數(2, 2))
    goicheck(2) = 1
  End If
  If goicheck(2) = 1 Then
    '=========以下是技能檢查及發動
        異常狀態檢查數(1, 1) = 1
        異常狀態.ATK加_電腦 '(階段1)
        '=======
        異常狀態檢查數(26, 1) = 1
        異常狀態.聖痕_電腦 '(階段1)
        '=======
        異常狀態檢查數(4, 1) = 1
        異常狀態.ATK減_電腦 '(階段1)
        '=======
        異常狀態檢查數(25, 1) = 1
        異常狀態.能力低下_電腦 '(階段1)
     '==============
  End If
End If
End Sub
Sub chkdef()
If goidefus = 0 Then
 攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + defus(角色人物對戰人數(1, 2))
 攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + defus(角色人物對戰人數(1, 2))
 FormMainMode.顯示列1.goi1 = Val(FormMainMode.顯示列1.goi1) + defus(角色人物對戰人數(1, 2))
 goidefus = 1
   '=========以下是技能檢查及發動
'   If 異常狀態檢查數(8, 2) = 1 Then
      異常狀態檢查數(8, 1) = 1
      異常狀態.DEF加_使用者 '(階段1)
'   End If
'   If 異常狀態檢查數(11, 2) = 1 Then
      異常狀態檢查數(11, 1) = 1
      異常狀態.DEF減_使用者 '(階段1)
'   End If
   異常狀態檢查數(13, 1) = 1
   異常狀態.聖痕_使用者 '(階段1)
   '====
   異常狀態檢查數(24, 1) = 1
   異常狀態.能力低下_使用者 '(階段1)
   '====
   異常狀態檢查數(39, 1) = 1
   異常狀態.臨界_使用者 '(階段1)
   '==============
End If
End Sub
Sub chkdefcom()
If chkcomck = 0 Then
 攻擊防禦骰子總數(2) = 攻擊防禦骰子總數(2) + defcom(角色人物對戰人數(2, 2))
 攻擊防禦骰子總數(4) = 攻擊防禦骰子總數(4) + defcom(角色人物對戰人數(2, 2))
 FormMainMode.顯示列1.goi2 = Val(FormMainMode.顯示列1.goi2) + defcom(角色人物對戰人數(2, 2))
 chkcomck = 1
    '=========以下是技能檢查及發動
'   If 異常狀態檢查數(8, 2) = 1 Then
      異常狀態檢查數(2, 1) = 1
      異常狀態.DEF加_電腦  '(階段1)
'   End If
'   If 異常狀態檢查數(11, 2) = 1 Then
      異常狀態檢查數(5, 1) = 1
      異常狀態.DEF減_電腦 '(階段1)
'   End If
   異常狀態檢查數(26, 1) = 2
   異常狀態.聖痕_電腦 '(階段2)
   '===
   異常狀態檢查數(25, 1) = 1
   異常狀態.能力低下_電腦 '(階段1)
   '==============
End If
End Sub
Sub chkus1()
If goicheck(1) = 0 Then
 If atkingpagetot(1, 1) > 0 Then
   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkus(角色人物對戰人數(1, 2))
   攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
   goicheck(1) = 1
   '=========以下是技能檢查及發動
'   If 異常狀態檢查數(13, 2) = 1 Then
      異常狀態檢查數(13, 1) = 1
      異常狀態.聖痕_使用者 '(階段1)
'   End If
'   If 異常狀態檢查數(24, 2) = 1 Then
      異常狀態檢查數(24, 1) = 1
      異常狀態.能力低下_使用者 '(階段1)
'   End If
'   If 異常狀態檢查數(7, 2) = 1 Then
      異常狀態檢查數(7, 1) = 1
      異常狀態.ATK加_使用者 '(階段1)
'   End If
'   If 異常狀態檢查數(10, 2) = 1 Then
      異常狀態檢查數(10, 1) = 1
      異常狀態.ATK減_使用者 '(階段1)
'   End If
    '====
    異常狀態檢查數(39, 1) = 1
    異常狀態.臨界_使用者  '(階段1)
   '==============
  End If
End If
End Sub
Sub chkus2()
If goicheck(1) = 0 Then
  If atkingpagetot(1, 5) > 0 Then
   攻擊防禦骰子總數(1) = 攻擊防禦骰子總數(1) + atkus(角色人物對戰人數(1, 2))
   攻擊防禦骰子總數(3) = 攻擊防禦骰子總數(3) + atkus(角色人物對戰人數(1, 2))
   goicheck(1) = 1
   '=========以下是技能檢查及發動
'   If 異常狀態檢查數(13, 2) = 1 Then
      異常狀態檢查數(13, 1) = 1
      異常狀態.聖痕_使用者 '(階段1)
'   End If
'   If 異常狀態檢查數(24, 2) = 1 Then
      異常狀態檢查數(24, 1) = 1
      異常狀態.能力低下_使用者 '(階段1)
'   End If
'   If 異常狀態檢查數(7, 2) = 1 Then
      異常狀態檢查數(7, 1) = 1
      異常狀態.ATK加_使用者 '(階段1)
'   End If
'   If 異常狀態檢查數(10, 2) = 1 Then
      異常狀態檢查數(10, 1) = 1
      異常狀態.ATK減_使用者 '(階段1)
'   End If
    '====
    異常狀態檢查數(39, 1) = 1
    異常狀態.臨界_使用者  '(階段1)
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
Sub comatk_智慧型AI引導程序_超出牌張數(ByVal turn As Integer, ByVal movecpre As Integer, ByVal choose As Integer)
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
Sub getpage_舊(ByVal k As Integer, m As Integer)
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
                      Case 1 '使用者
                          pagecardnum(m, 11) = 0
                          FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                          FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
                          戰鬥系統類.座標計算_使用者手牌
                          牌移動暫時變數(3) = m
                          pagecardnum(m, 9) = 240 '指定目前Left(座標)
                          pagecardnum(m, 10) = 960 '指定目前Top(座標)
                          FormMainMode.card(m).Left = 240
                          FormMainMode.card(m).Top = 960
                          戰鬥系統類.計算牌移動距離單位
                          戰鬥系統類.公用牌回復正面 (牌移動暫時變數(3))
                          FormMainMode.card(m).Visible = True
                          戰鬥系統類.牌順序增加_手牌_使用者 m
                          FormMainMode.牌移動.Enabled = True
                          FormMainMode.wmpse1.Controls.stop
                          FormMainMode.wmpse1.Controls.play
                          一般系統類.檢查音樂播放 1
                      Case 2 '電腦
                          pagecardnum(m, 11) = 0
                          FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                          FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
                          戰鬥系統類.座標計算_電腦手牌
                          牌移動暫時變數(3) = m
                          pagecardnum(m, 9) = 240 '指定目前Left(座標)
                          pagecardnum(m, 10) = 960 '指定目前Top(座標)
                          FormMainMode.card(m).Left = 240
                          FormMainMode.card(m).Top = 960
                          戰鬥系統類.計算牌移動距離單位
                          戰鬥系統類.公用牌變背面
                          FormMainMode.card(m).Visible = True
                          戰鬥系統類.牌順序增加_手牌_電腦 m
                          FormMainMode.牌移動.Enabled = True
                          FormMainMode.wmpse1.Controls.stop
                          FormMainMode.wmpse1.Controls.play
                          一般系統類.檢查音樂播放 1
        End Select
End Sub
Sub moveatkin()
Do
    For j = 71 To 106
      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
         If pagecardnum(j, 1) = a3a And pagecardnum(j, 3) = a3a Then '移動單面事件卡優先
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            目前數(25) = 目前數(25) + Val(pagecardnum(j, 2))
         End If
         If 目前數(25) >= 2 Then Exit Do
      End If
    Next
    For j = 1 To 106
      If Val(pagecardnum(j, 6)) = 1 And Val(pagecardnum(j, 5)) = 2 And Val(pagecardnum(j, 11)) <> 1 Then
         If pagecardnum(j, 1) = a3a Then
           pagecardnum(j, 11) = 1
'           movecom = Val(movecom) + Val(pagecardnum(j, 2))
            目前數(25) = 目前數(25) + 1
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
            目前數(25) = 目前數(25) + Val(pagecardnum(j, 2))
         End If
         If 目前數(25) >= 2 Then Exit Do
      End If
    Next
    Exit Do
Loop
'movecheckcom = movecom
End Sub
Sub movetnus()
FormMainMode.messageus.AddItem "你有主動權。"
'formmainmode.messageus.AddItem "現在的距離" & movecp & "。"
戰鬥系統類.自動捲軸捲動
FormMainMode.move3.Picture = LoadPicture(app_path & "gif\atk1.gif")
FormMainMode.move4.Picture = LoadPicture(app_path & "gif\def1.gif")
FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\atk2.gif")
FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\def2.gif")
moveturn = 1
'cn2.Visible = True
FormMainMode.cnmove2.Visible = False
擲骰表單溝通暫時變數(1) = 1
End Sub
Sub movetncom()
FormMainMode.messageus.AddItem "對方有主動權。"
'formmainmode.messageus.AddItem "現在的距離" & movecp & "。"
戰鬥系統類.自動捲軸捲動
FormMainMode.move3.Picture = LoadPicture(app_path & "gif\def1.gif")
FormMainMode.move4.Picture = LoadPicture(app_path & "gif\atk1.gif")
FormMainMode.atkdef1.Picture = LoadPicture(app_path & "gif\def2.gif")
FormMainMode.atkdef2.Picture = LoadPicture(app_path & "gif\atk2.gif")
moveturn = 2
'cn3.Visible = True
FormMainMode.cnmove2.Visible = False
擲骰表單溝通暫時變數(1) = 1
End Sub
Sub 人物交換_使用者_指定交換(ByVal num As Integer)
Dim ae As Integer
ae = 角色人物對戰人數(1, 2)
角色人物對戰人數(1, 2) = 角色待機人物紀錄數(1, num)
角色待機人物紀錄數(1, 1) = 角色人物對戰人數(1, 2)
角色待機人物紀錄數(1, num) = ae
FormMainMode.uspiin(角色待機人物紀錄數(1, num)).Left = 2520 * (num - 1)
FormMainMode.uspiin(角色待機人物紀錄數(1, num)).Visible = True
FormMainMode.cardus(角色待機人物紀錄數(1, num)).Visible = False

FormMainMode.uspiin(角色人物對戰人數(1, 2)).Left = 0
FormMainMode.uspiin(角色人物對戰人數(1, 2)).Visible = False
FormMainMode.cardus(角色人物對戰人數(1, 2)).Left = 0
FormMainMode.cardus(角色人物對戰人數(1, 2)).Top = 6240
FormMainMode.cardus(角色人物對戰人數(1, 2)).ZOrder
FormMainMode.cardus(角色人物對戰人數(1, 2)).Visible = True
For n = 1 To 4
    If VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 1) = "" Then
       FormMainMode.personatk(n).Caption = ""
       FormMainMode.personatk(n).Visible = False
    Else
       FormMainMode.personatk(n).Caption = VBEPerson(1, 角色人物對戰人數(1, 2), 3, n, 1)
       If VBEPerson(1, 角色人物對戰人數(1, 2), 2, 3, 5) = 1 Then
           FormMainMode.personatk(n).FontSize = 12
       Else
           FormMainMode.personatk(n).FontSize = VBEPerson(1, 角色人物對戰人數(1, 2), 2, 3, n)
       End If
       FormMainMode.personatk(n).Visible = True
    End If
Next
FormMainMode.personusminijpg.Visible = False
FormMainMode.personusminijpg.小人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 1)
FormMainMode.personusminijpg.小人物影子圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 2)
FormMainMode.顯示列1.使用者方小人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 4)
FormMainMode.personusminijpg.小人物影子Left = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 1, 5))
FormMainMode.personusminijpg.小人物影子top差 = Val(VBEPerson(1, 角色人物對戰人數(1, 2), 2, 1, 6))
FormDice.jpgus.大人物圖片 = VBEPerson(1, 角色人物對戰人數(1, 2), 1, 5, 3)
FormMainMode.顯示列1.使用者方小人物圖片left = -(FormMainMode.顯示列1.使用者方小人物圖片width)
FormMainMode.personusminijpg.Visible = True
'--------------------------計算新距離單位(HP血條)
距離單位(1, 1, 1) = 5295 \ liveusmax(角色人物對戰人數(1, 2))
FormMainMode.bloodlineout1.Width = (距離單位(1, 1, 1) * liveus(角色人物對戰人數(1, 2)))
FormMainMode.bloodnumus1.Caption = liveus(角色人物對戰人數(1, 2))
FormMainMode.bloodnumus2.Caption = liveusmax(角色人物對戰人數(1, 2))
'========================
執行動作_距離變更 movecp
'========================以下是技能檢查及啟動
If FormMainMode.uspi1(角色人物對戰人數(1, 2)).Caption = "史塔夏" Then
    If atking_史塔夏_殺戮模式狀態數(2) = 1 Then
       atking_史塔夏_殺戮模式狀態數(1) = 4
       戰鬥系統類.特殊_史塔夏_殺戮狀態_使用者 '(階段4)
    End If
End If
If FormMainMode.uspi1(角色人物對戰人數(1, 2)).Caption = "音音夢" Then
    If atking_音音夢_成長模式狀態數(2) = 1 Then
       atking_音音夢_成長模式狀態數(1) = 4
       戰鬥系統類.特殊_音音夢_成長狀態_使用者 '(階段4)
    End If
End If
'=============================
For i = 1 To 4
    戰鬥系統類.人物技能欄燈開關 False, i
Next
'=============================
If FormMainMode.uspi1(角色人物對戰人數(1, 2)).Caption = "尤莉卡" And atking_尤莉卡_超載目前階段紀錄數(3) > 0 Then
    atkingck(49, 2) = 1
    atkingck(49, 1) = 7
    技能.尤莉卡_超載  '(階段7)
End If
'==========
End Sub

Sub 人物交換_電腦_指定交換(ByVal num As Integer)
Dim ae As Integer
ae = 角色人物對戰人數(2, 2)
角色人物對戰人數(2, 2) = 角色待機人物紀錄數(2, num)
角色待機人物紀錄數(2, num) = ae
角色待機人物紀錄數(2, 1) = 角色人物對戰人數(2, 2)
FormMainMode.compiin(角色待機人物紀錄數(2, num)).Left = 2520 * (num - 1)
FormMainMode.compiin(角色人物對戰人數(2, 2)).Left = 0
For n = 1 To 4
    If VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 1) = "" Then
       FormMainMode.comaiatk(n).Caption = ""
       FormMainMode.comaiatk(n).Visible = False
    Else
       FormMainMode.comaiatk(n).Caption = VBEPerson(2, 角色人物對戰人數(2, 2), 3, n, 1)
       If VBEPerson(2, 角色人物對戰人數(2, 2), 2, 3, 5) = 1 Then
           FormMainMode.comaiatk(n).FontSize = 12
       Else
           FormMainMode.comaiatk(n).FontSize = VBEPerson(2, 角色人物對戰人數(2, 2), 2, 3, n)
       End If
       FormMainMode.comaiatk(n).Visible = True
    End If
Next
FormMainMode.personcomminijpg.Visible = False
'====================
FormMainMode.personcomminijpg.小人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 1)
FormMainMode.personcomminijpg.小人物影子圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 2)
FormMainMode.顯示列1.電腦方小人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 4)
FormMainMode.personcomminijpg.小人物影子Left = VBEPerson(2, 角色人物對戰人數(2, 2), 2, 1, 5)
FormMainMode.personcomminijpg.小人物影子top差 = VBEPerson(2, 角色人物對戰人數(2, 2), 2, 1, 6)
FormMainMode.cardcom(角色人物對戰人數(2, 2)).Picture = LoadPicture(VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 5))
FormDice.jpgcom.大人物圖片 = VBEPerson(2, 角色人物對戰人數(2, 2), 1, 5, 3)
FormMainMode.顯示列1.電腦方小人物圖片left = FormMainMode.ScaleWidth
FormMainMode.personcomminijpg.Left = personminixy(2, 角色人物對戰人數(2, 2), movecp, 1)
FormMainMode.personcomminijpg.Top = personminixy(2, 角色人物對戰人數(2, 2), movecp, 2)
FormMainMode.personcomminijpg.Visible = True
FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = livecom(角色人物對戰人數(2, 2))
FormMainMode.cardcompi2(角色人物對戰人數(2, 2)).Caption = atkcom(角色人物對戰人數(2, 2))
FormMainMode.cardcompi3(角色人物對戰人數(2, 2)).Caption = defcom(角色人物對戰人數(2, 2))
FormMainMode.compi1(角色人物對戰人數(2, 2)).Caption = namecom(角色人物對戰人數(2, 2))
FormMainMode.compi2(角色人物對戰人數(2, 2)).Caption = comlevel(角色人物對戰人數(2, 2))
FormMainMode.compiatk(角色人物對戰人數(2, 2)).Caption = atkcom(角色人物對戰人數(2, 2))
FormMainMode.compidef(角色人物對戰人數(2, 2)).Caption = defcom(角色人物對戰人數(2, 2))
FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = livecom(角色人物對戰人數(2, 2))
FormMainMode.compi5(角色人物對戰人數(2, 2)).Caption = livecommax(角色人物對戰人數(2, 2))
'--------------------------計算新距離單位(HP血條)
距離單位(1, 2, 1) = (11340 - 6060) \ livecommax(角色人物對戰人數(2, 2))
FormMainMode.bloodlineout2.Left = 11340 - (距離單位(1, 2, 1) * livecom(角色人物對戰人數(2, 2)))
FormMainMode.bloodnumcom1.Caption = livecom(角色人物對戰人數(2, 2))
FormMainMode.bloodnumcom2.Caption = livecommax(角色人物對戰人數(2, 2))
'==============================
執行動作_距離變更 movecp
'=============================
If FormMainMode.compi1(角色人物對戰人數(2, 2)).Caption = "尤莉卡" And atking_AI_尤莉卡_超載目前階段紀錄數(3) > 0 Then
    atkingckai(139, 2) = 1
    atkingckai(139, 1) = 7
    AI技能.尤莉卡_超載  '(階段7)
End If
'==========
End Sub
Sub 執行動作_交換人物角色_使用者_初始()
Dim i As Integer
Dim ne As Integer
For i = 2 To 3
   Formchangeperson.card(i - 1).Picture = FormMainMode.cardus(角色待機人物紀錄數(1, i)).Picture
   Formchangeperson.cardhp(i - 1).Caption = FormMainMode.usbi1(角色待機人物紀錄數(1, i)).Caption
   Formchangeperson.cardatk(i - 1).Caption = FormMainMode.usbi2(角色待機人物紀錄數(1, i)).Caption
   Formchangeperson.carddef(i - 1).Caption = FormMainMode.usbi3(角色待機人物紀錄數(1, i)).Caption
Next
ne = 1
For k = 2 To 3
    For j = 14 * (角色待機人物紀錄數(1, k) - 1) + 1 To 14 * 角色待機人物紀錄數(1, k)
'        For i = 14 * (k - 2) + 1 To 14 * (k - 1)
            If 人物異常狀態資料庫(1, j, 2) > 0 Then
                Formchangeperson.personusspe(ne).person_turn = FormMainMode.personusspe(j).person_turn
                Formchangeperson.personusspe(ne).person_num = FormMainMode.personusspe(j).person_num
                Formchangeperson.personusspe(ne).異常狀態圖片 = FormMainMode.personusspe(j).異常狀態圖片
                Formchangeperson.personusspe(ne).Visible = True
            Else
                Formchangeperson.personusspe(ne).Visible = False
            End If
            ne = ne + 1
'        Next
    Next
Next
交換角色紀錄暫時變數(1) = 0
For k = 1 To 2
     Formchangeperson.PEAFcardback(k).Visible = False
Next
If Formsetting.chkusenewaipersonauto.Value = 1 Then
    Formchangeperson.使用者方智慧型AI_自動控制選人.Enabled = True
End If
Formchangeperson.Left = FormMainMode.Left + 2430
Formchangeperson.Top = FormMainMode.Top + 1655
Formchangeperson.Show 1
End Sub
Sub 執行動作_交換人物角色_電腦_初始()
Select Case 交換角色紀錄暫時變數(2)
    Case 1
       交換角色紀錄暫時變數(2) = 0
       目前數(22) = 18
       FormMainMode.等待時間.Enabled = True
    Case 0
       目前數(22) = 19
       FormMainMode.等待時間.Enabled = True
End Select

End Sub
Sub 執行動作_交換人物角色_電腦_交換()
If livecom(角色待機人物紀錄數(2, 2)) > 0 Then
       人物交換_電腦_指定交換 2
ElseIf livecom(角色待機人物紀錄數(2, 3)) > 0 Then
       人物交換_電腦_指定交換 3
End If
執行動作_交換人物角色_結束執行
End Sub
Sub 執行動作_交換人物角色_初始()
If (交換角色紀錄暫時變數(1) = 1 Or 交換角色紀錄暫時變數(2) = 1) And 交換角色紀錄暫時變數(3) = 0 Then
    turnatk = 6
    階段狀態數 = 5
    戰鬥系統類.時間軸_重設
    FormMainMode.顯示列1.顯示列圖片 = App.Path & "\gif\linechange.png"
    FormMainMode.顯示列1.Visible = True
    FormMainMode.顯示列1.goi1顯示 = False
    FormMainMode.顯示列1.goi2顯示 = False
    戰鬥系統類.時間軸_顯示
    FormMainMode.trtimeline.Enabled = True
    小人物頭像移動方向數(1) = 2
    小人物頭像移動方向數(2) = 2
    FormMainMode.小人物頭像移動_使用者.Enabled = True
    FormMainMode.小人物頭像移動_電腦.Enabled = True
    交換角色紀錄暫時變數(3) = 1
    FormMainMode.顯示列1.移動階段選擇值 = 0
    FormMainMode.顯示列1.移動階段圖顯示 = False
End If
If 交換角色紀錄暫時變數(1) = 1 Then
    執行動作_交換人物角色_使用者_初始
ElseIf 交換角色紀錄暫時變數(2) = 1 Then
    執行動作_交換人物角色_電腦_初始
End If
End Sub
Sub 執行動作_移動階段選擇執行()
'===========交換角色類
If 交換角色紀錄暫時變數(1) = 1 Or 交換角色紀錄暫時變數(2) = 1 Then
    執行動作_交換人物角色_初始
Else
    交換角色紀錄暫時變數(3) = 0
    目前數(22) = 17
    FormMainMode.等待時間.Enabled = True
End If
End Sub
Sub 執行動作_人物死亡交換階段選擇執行()
If 交換角色紀錄暫時變數(1) = 1 Or 交換角色紀錄暫時變數(2) = 1 Then
    執行動作_交換人物角色_初始
Else
    交換角色紀錄暫時變數(3) = 0
    目前數(22) = 20
    FormMainMode.等待時間.Enabled = True
End If
End Sub
Sub 執行動作_交換人物角色_結束執行()
   Formchangeperson.Hide
   戰鬥系統類.時間軸_停止
   Select Case 交換角色紀錄暫時變數(4)
      Case 1
         執行動作_移動階段選擇執行
      Case 2
         執行動作_人物死亡交換階段選擇執行
    End Select
End Sub
Sub 事件卡處理_指定_使用者方()
Dim kp(1 To 18)  As Integer '事件卡標記暫時數
Dim m, km As Integer
If 事件卡記錄暫時數(0, 1) = 18 Then
    Do
        Randomize
        m = Int(Rnd() * 18) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(1, km, 1) = Formsetting.personus(m).Text
            pageeventnum(1, km, 2) = 一般系統類.事件卡資料庫(Formsetting.personus(m).Text, 2)
        End If
    Loop Until km >= 18
ElseIf 事件卡記錄暫時數(0, 1) = 12 Then
    Do
        Randomize
        m = Int(Rnd() * 6) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(1, km, 1) = Formsetting.personus(m).Text
            pageeventnum(1, km, 2) = 一般系統類.事件卡資料庫(Formsetting.personus(m).Text, 2)
        End If
    Loop Until km >= 6
    For i = 7 To 12
        pageeventnum(1, i, 1) = Formsetting.personus(i).Text
        pageeventnum(1, i, 2) = 一般系統類.事件卡資料庫(Formsetting.personus(i).Text, 2)
    Next
End If
End Sub
Sub 事件卡處理_指定_電腦方()
Dim kp(1 To 18)  As Integer '事件卡標記暫時數
Dim m, km As Integer
If 事件卡記錄暫時數(0, 1) = 18 Then
    Do
        Randomize
        m = Int(Rnd() * 18) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(2, km, 1) = Formsetting.personcom(m).Text
            pageeventnum(2, km, 2) = 一般系統類.事件卡資料庫(Formsetting.personcom(m).Text, 2)
        End If
    Loop Until km >= 18
ElseIf 事件卡記錄暫時數(0, 1) = 12 Then
    Do
        Randomize
        m = Int(Rnd() * 6) + 1
        If kp(m) = 0 Then
            kp(m) = 1
            km = km + 1
            pageeventnum(2, km, 1) = Formsetting.personcom(m).Text
            pageeventnum(2, km, 2) = 一般系統類.事件卡資料庫(Formsetting.personcom(m).Text, 2)
        End If
    Loop Until km >= 6
    For i = 7 To 11
        pageeventnum(2, i, 1) = Formsetting.personcom(i).Text
        pageeventnum(2, i, 2) = 一般系統類.事件卡資料庫(Formsetting.personcom(i).Text, 2)
    Next
End If
End Sub
Sub 事件卡處理_初始_使用者方()
Dim ck As Boolean
Dim m As Integer
If Formsetting.persontgruonus(1).Value = True Then '=====(無)
    For i = 1 To 18
       Randomize
       m = Int(Rnd() * 3) + 1
       Select Case m
          Case 1
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "劍1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
          Case 2
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "槍1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
          Case 3
             For j = 0 To Formsetting.personus(i).ListCount - 1
                If Formsetting.personus(i).List(j) = "防1" Then
                    Formsetting.personus(i).ListIndex = j
                End If
             Next
       End Select
    Next
ElseIf Formsetting.persontgruonus(2).Value = True Then '=====自訂
   If 事件卡記錄暫時數(0, 1) = 18 Or Formsetting.persontgreus.Value = 0 Then
        For i = 1 To 18
'            If Formsetting.personus(i).Text = "(無)" Then
            If 一般系統類.事件卡資料庫(Formsetting.personus(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "劍1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "槍1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "防1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    ElseIf 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
         For i = 1 To 18
            If Formsetting.personus(i).Text = "(無)" Or i >= 7 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "劍1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "槍1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "防1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    End If
ElseIf Formsetting.persontgruonus(3).Value = True Then '===============選擇最大值
    If Formsetting.persontgreus.Value = 1 Then  '===遵守規則
         For i = 1 To 18
             Select Case Formsetting.persontgus(i).Caption
                 Case 0
                      Randomize
                      m = Int(Rnd() * 8) + 1
                      Select Case m
                          Case 1
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "劍3/槍1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "槍3/劍1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 3
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "防3/移1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 4
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "劍3/移1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 5
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "槍3/移1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 6
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "劍3/防1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 7
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "槍3/防1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                          Case 8
                               For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "特2" Then
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
                                   If Formsetting.personus(i).List(j) = "劍5/槍3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "劍5/移1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "劍8" Then
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
                                   If Formsetting.personus(i).List(j) = "槍5/劍3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "槍5/移1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "槍8" Then
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
                                   If Formsetting.personus(i).List(j) = "防5/移1" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "防7" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "HP回復3" Then
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
                                   If Formsetting.personus(i).List(j) = "移3/特3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "移5" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 5
                    For j = 0 To Formsetting.personus(i).ListCount - 1
                        If Formsetting.personus(i).List(j) = "機會5" Then
                            Formsetting.personus(i).ListIndex = j
                        End If
                     Next
                 Case 6
                    For j = 0 To Formsetting.personus(i).ListCount - 1
                        If Formsetting.personus(i).List(j) = "詛咒術5" Then
                            Formsetting.personus(i).ListIndex = j
                        End If
                     Next
                 Case 7
                      Randomize
                      m = Int(Rnd() * 2) + 1
                     Select Case m
                         Case 1
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "特3/防3" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personus(i).ListCount - 1
                                   If Formsetting.personus(i).List(j) = "特5" Then
                                       Formsetting.personus(i).ListIndex = j
                                   End If
                                Next
                      End Select
             End Select
         Next
         If 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgreus.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "劍1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "槍1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personus(i).ListCount - 1
                         If Formsetting.personus(i).List(j) = "防1" Then
                             Formsetting.personus(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else  '================================不遵守規則
        For i = 1 To 18
            Do
               Randomize
               m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
               '==============================
                    Select Case Formsetting.personus(i).List(m)
                        Case "劍8"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "槍8"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "防7"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "移5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "HP回復3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "機會5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "詛咒術5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "特5"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "劍5/槍3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "槍5/劍3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "防5/移1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "槍5/移1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "劍5/移1"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "移3/特3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                        Case "特3/防3"
                            Formsetting.personus(i).ListIndex = m
                            Exit Do
                    End Select
            Loop
        Next
    End If
ElseIf Formsetting.persontgruonus(4).Value = True Then '=====隨機
    If Formsetting.persontgreus.Value = 1 Then '===遵守規則
        For i = 1 To 18
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
                If 一般系統類.事件卡資料庫(Formsetting.personus(i).List(m), 1) = Formsetting.persontgus(i).Caption Or _
                   一般系統類.事件卡資料庫(Formsetting.personus(i).List(m), 1) = 0 Then
                   Formsetting.personus(i).ListIndex = m
                   Exit Do
                End If
             Loop
         Next
        If 事件卡記錄暫時數(0, 1) = 12 Then
            For i = 7 To 18
                   Randomize
                   m = Int(Rnd() * 3) + 1
                   Select Case m
                      Case 1
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "劍1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                      Case 2
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "槍1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                      Case 3
                         For j = 0 To Formsetting.personus(i).ListCount - 1
                            If Formsetting.personus(i).List(j) = "防1" Then
                                Formsetting.personus(i).ListIndex = j
                            End If
                         Next
                   End Select
            Next
        End If
    Else '=============================不遵守規則
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personus(i).ListCount - 1)) + 1
            Formsetting.personus(i).ListIndex = m
         Next
    End If
End If
End Sub
Sub 事件卡處理_初始_電腦方()
Dim m As Integer
Dim ay() As String
If Formsetting.persontgruoncom(1).Value = True Then '=====(無)
    For i = 1 To 18
       Randomize
       m = Int(Rnd() * 3) + 1
       Select Case m
          Case 1
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "劍1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
          Case 2
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "槍1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
          Case 3
             For j = 0 To Formsetting.personcom(i).ListCount - 1
                If Formsetting.personcom(i).List(j) = "防1" Then
                    Formsetting.personcom(i).ListIndex = j
                End If
             Next
       End Select
    Next
ElseIf Formsetting.persontgruoncom(2).Value = True Then '=====自訂
   If 事件卡記錄暫時數(0, 1) = 18 Or Formsetting.persontgrecom.Value = 0 Then
        For i = 1 To 18
'            If Formsetting.personcom(i).Text = "(無)" Then
            If 一般系統類.事件卡資料庫(Formsetting.personcom(i).Text, 1) = 99 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "劍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "槍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "防1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    ElseIf 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
         For i = 1 To 18
            If Formsetting.personcom(i).Text = "(無)" Or i >= 7 Then
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "劍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "槍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "防1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
             End If
         Next
    End If
ElseIf Formsetting.persontgruoncom(3).Value = True Then '=====選擇最大值
    If Formsetting.persontgrecom.Value = 1 Then  '===遵守規則
         For i = 1 To 18
             Select Case Formsetting.persontgcom(i).Caption
                 Case 0
                      Randomize
                      m = Int(Rnd() * 8) + 1
                      Select Case m
                          Case 1
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "劍3/槍1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "槍3/劍1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 3
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "防3/移1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 4
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "劍3/移1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 5
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "槍3/移1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 6
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "劍3/防1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 7
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "槍3/防1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                          Case 8
                               For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "特2" Then
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
                                   If Formsetting.personcom(i).List(j) = "劍5/槍3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "劍5/移1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "劍8" Then
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
                                   If Formsetting.personcom(i).List(j) = "槍5/劍3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "槍5/移1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                              For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "槍8" Then
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
                                   If Formsetting.personcom(i).List(j) = "防5/移1" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "防7" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 3
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "HP回復3" Then
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
                                   If Formsetting.personcom(i).List(j) = "移3/特3" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                         Case 2
                                For j = 0 To Formsetting.personcom(i).ListCount - 1
                                   If Formsetting.personcom(i).List(j) = "移5" Then
                                       Formsetting.personcom(i).ListIndex = j
                                   End If
                                Next
                      End Select
                 Case 5
                    For j = 0 To Formsetting.personcom(i).ListCount - 1
                        If Formsetting.personcom(i).List(j) = "機會5" Then
                            Formsetting.personcom(i).ListIndex = j
                        End If
                     Next
                 Case 6
                    For j = 0 To Formsetting.personcom(i).ListCount - 1
                        If Formsetting.personcom(i).List(j) = "詛咒術5" Then
                            Formsetting.personcom(i).ListIndex = j
                        End If
                     Next
                 Case 7
                        For j = 0 To Formsetting.personcom(i).ListCount - 1
                           If Formsetting.personcom(i).List(j) = "特3/防3" Then
                               Formsetting.personcom(i).ListIndex = j
                           End If
                        Next
             End Select
         Next
         If 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "劍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "槍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "防1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else  '================================不遵守規則
        For i = 1 To 18
            Do
               Randomize
               m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
               '==============================
                    Select Case Formsetting.personcom(i).List(m)
                        Case "劍8"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "槍8"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "防7"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "移5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "HP回復3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "機會5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "詛咒術5"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "劍5/槍3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "槍5/劍3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "防5/移1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "槍5/移1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "劍5/移1"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "移3/特3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                        Case "特3/防3"
                            Formsetting.personcom(i).ListIndex = m
                            Exit Do
                    End Select
            Loop
        Next
    End If
ElseIf Formsetting.persontgruoncom(4).Value = True Then '=====隨機
    If Formsetting.persontgrecom.Value = 1 Then '===遵守規則
        For i = 1 To 18
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                If 一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 1) = Formsetting.persontgcom(i).Caption Or _
                   一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 1) = 0 Then
                   Formsetting.personcom(i).ListIndex = m
                   Exit Do
                End If
             Loop
         Next
         If 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "劍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "槍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "防1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else '=============================不遵守規則
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
            Formsetting.personcom(i).ListIndex = m
         Next
    End If
ElseIf Formsetting.persontgruoncom(5).Value = True Then '=====隨機(不含特)
    If Formsetting.persontgrecom.Value = 1 Then '===遵守規則
        For i = 1 To 18
             Do
                Randomize
                m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
                If 一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 1) = Formsetting.persontgcom(i).Caption Or _
                   一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 1) = 0 Then
                   ay = Split(一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 3), "=")
                   If ay(0) = a4a And ay(2) = a4a Then
                   Else
                        Formsetting.personcom(i).ListIndex = m
                        Exit Do
                   End If
                End If
             Loop
         Next
         If 事件卡記錄暫時數(0, 1) = 12 And Formsetting.persontgrecom.Value = 1 Then
            For i = 7 To 18
                Randomize
                m = Int(Rnd() * 3) + 1
                Select Case m
                   Case 1
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "劍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 2
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "槍1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                   Case 3
                      For j = 0 To Formsetting.personcom(i).ListCount - 1
                         If Formsetting.personcom(i).List(j) = "防1" Then
                             Formsetting.personcom(i).ListIndex = j
                         End If
                      Next
                End Select
            Next
        End If
    Else '=============================不遵守規則
         For i = 1 To 18
            Randomize
            m = Int(Rnd() * (Formsetting.personcom(i).ListCount - 1)) + 1
            ay = Split(一般系統類.事件卡資料庫(Formsetting.personcom(i).List(m), 3), "=")
            If ay(0) = a4a And ay(2) = a4a Then
                 i = i - 1
            Else
                 Formsetting.personcom(i).ListIndex = m
            End If
         Next
    End If
End If
End Sub
Sub 事件卡處理_分派_使用者方()
Dim tn As Integer
Dim ay() As String
tn = Val(FormMainMode.turni.Caption)
If tn <= 18 Then
    If tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgreus.Value = 0 Then
        If pageeventnum(1, tn, 1) <> "" Then
            ay = Split(一般系統類.事件卡資料庫(pageeventnum(1, tn, 1), 3), "=")
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
            戰鬥系統類.座標計算_使用者手牌
            牌移動暫時變數(3) = 70 + tn
            戰鬥系統類.牌順序增加_手牌_使用者 70 + tn
            pagecardnum(70 + tn, 9) = 牌移動暫時變數(1) '指定目前Left(座標)
            pagecardnum(70 + tn, 10) = 牌移動暫時變數(2) '指定目前Top(座標)
            FormMainMode.card(70 + tn).Left = 牌移動暫時變數(1)
            FormMainMode.card(70 + tn).Top = 牌移動暫時變數(2)
            FormMainMode.card(70 + tn).ZOrder
            FormMainMode.card(70 + tn).Visible = True
        End If
    End If
End If
End Sub
Sub 事件卡處理_分派_電腦方()
Dim tn As Integer
Dim ay() As String
tn = Val(FormMainMode.turni.Caption)
If tn <= 18 Then
    If tn <= 事件卡記錄暫時數(0, 1) Or Formsetting.persontgrecom.Value = 0 Then
        If pageeventnum(2, tn, 1) <> "" Then
            ay = Split(一般系統類.事件卡資料庫(pageeventnum(2, tn, 1), 3), "=")
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
            戰鬥系統類.座標計算_電腦手牌
            牌移動暫時變數(3) = 88 + tn
            戰鬥系統類.公用牌變背面
            戰鬥系統類.牌順序增加_手牌_電腦 88 + tn
            pagecardnum(88 + tn, 9) = 牌移動暫時變數(1) '指定目前Left(座標)
            pagecardnum(88 + tn, 10) = 牌移動暫時變數(2) '指定目前Top(座標)
            FormMainMode.card(88 + tn).Left = 牌移動暫時變數(1)
            FormMainMode.card(88 + tn).Top = 牌移動暫時變數(2)
            FormMainMode.card(88 + tn).ZOrder
            FormMainMode.card(88 + tn).Visible = True
            For i = 1 To 3
                FormMainMode.compiin(i).ZOrder
            Next
        End If
    End If
End If
End Sub
Sub 事件卡處理_計算張數()
If 角色人物對戰人數(1, 1) > 1 Or 角色人物對戰人數(2, 1) > 1 Then
    事件卡記錄暫時數(0, 1) = 18
Else
    事件卡記錄暫時數(0, 1) = 12
End If
End Sub
Function 執行動作_檢查是否有指定異常狀態(ByVal uscom As Integer, ByVal num As Integer) As Boolean
執行動作_檢查是否有指定異常狀態 = False
Select Case uscom
   Case 1
        For i = 14 * (角色人物對戰人數(1, 2) - 1) + 1 To 14 * 角色人物對戰人數(1, 2)
           If 人物異常狀態資料庫(1, i, 3) = num Then
               執行動作_檢查是否有指定異常狀態 = True
           End If
        Next
   Case 2
        For i = 14 * (角色人物對戰人數(2, 2) - 1) + 1 To 14 * 角色人物對戰人數(2, 2)
            If 人物異常狀態資料庫(2, i, 3) = num Then
                執行動作_檢查是否有指定異常狀態 = True
            End If
        Next
End Select
End Function
Sub 執行動作_防禦階段結束時技能啟動()
atkingtrn(1) = 0
atkingtrn(2) = 0
'=================以下是技能檢查及啟動(回合結束階段1)
If turnatk = 2 And atkingck(64, 2) = 1 Then
   atkingck(64, 1) = 3
   技能.梅倫_Jackpot  '(階段3)
End If
If turnatk = 2 And atkingck(146, 2) = 1 Then
   atkingck(146, 1) = 3
   技能.傑多_因果之刻  '(階段3)
End If
If turnatk = 2 And atkingck(100, 2) = 1 Then
   atkingck(100, 1) = 2
   技能.露緹亞_暗影之翼  '(階段2)
End If
If turnatk = 2 And atkingck(111, 2) = 1 Then
   atkingck(111, 1) = 3
   技能.貝琳達_水晶幻鏡  '(階段3)
End If
'=================
技能動畫顯示階段數 = 9
戰鬥系統類.技能啟動數量檢查
'===================
If turnatk = 2 And atkingck(64, 2) = 1 Then
   atkingck(64, 1) = 4
   技能.梅倫_Jackpot  '(階段4)
End If
If turnatk = 2 And atkingck(146, 2) = 1 Then
   atkingck(146, 1) = 4
   技能.傑多_因果之刻  '(階段4)
End If
If turnatk = 2 And atkingck(100, 2) = 1 Then
   atkingck(100, 1) = 3
   技能.露緹亞_暗影之翼  '(階段3)
End If
If turnatk = 2 And atkingck(111, 2) = 1 Then
   atkingck(111, 1) = 4
   技能.貝琳達_水晶幻鏡  '(階段4)
End If
'================
FormMainMode.atkingtrtot.Interval = 600
FormMainMode.atkingtrtot.Enabled = True
End Sub
Sub 執行動作_攻擊階段結束時技能啟動()
atkingtrn(1) = 0
atkingtrn(2) = 0
'=================以下是技能檢查及啟動(回合結束階段1)
If turnatk = 1 And atkingckai(31, 2) = 1 Then
   atkingckai(31, 1) = 3
   AI技能.梅倫_Jackpot  '(階段3)
End If
If turnatk = 1 And atkingckai(97, 2) = 1 Then
   atkingckai(97, 1) = 2
   AI技能.露緹亞_暗影之翼  '(階段2)
End If
If turnatk = 1 And atkingckai(121, 2) = 1 Then
   atkingckai(121, 1) = 3
   AI技能.傑多_因果之刻  '(階段3)
End If
If turnatk = 1 And atkingckai(123, 2) = 1 Then
   atkingckai(123, 1) = 3
   AI技能.貝琳達_水晶幻鏡  '(階段3)
End If
'=================
技能動畫顯示階段數 = 9
戰鬥系統類.技能啟動數量檢查
'===================
If turnatk = 1 And atkingckai(31, 2) = 1 Then
   atkingckai(31, 1) = 4
   AI技能.梅倫_Jackpot  '(階段4)
End If
If turnatk = 1 And atkingckai(97, 2) = 1 Then
   atkingckai(97, 1) = 3
   AI技能.露緹亞_暗影之翼  '(階段3)
End If
If turnatk = 1 And atkingckai(121, 2) = 1 Then
   atkingckai(121, 1) = 4
   AI技能.傑多_因果之刻  '(階段4)
End If
If turnatk = 1 And atkingckai(123, 2) = 1 Then
   atkingckai(123, 1) = 4
   AI技能.貝琳達_水晶幻鏡  '(階段4)
End If
'=================
FormMainMode.atkingtrtot.Interval = 600
FormMainMode.atkingtrtot.Enabled = True
End Sub
Sub 技能說明載入_人物卡片背面_使用者(ByVal n As Integer)
Dim strw() As String
If 角色人物對戰人數(1, 2) = n Then
    For i = 5 To 8
        FormMainMode.PEAFpersoncardback_text(i) = VBEPerson(1, n, 3, i - 4, 1)
        '========
        FormMainMode.PEAFpersoncardback_turn(i).物件類別 = 3
        FormMainMode.PEAFpersoncardback_turn(i).圖片 = app_path & "gif\卡片背面\CBturn.png"
        FormMainMode.PEAFpersoncardback_turn(i).項目編號 = Val(VBEPerson(1, n, 3, i - 4, 8))
        '============================
        Select Case i - 4
            Case 1
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range1(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range1(k + 3).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range1(k + 3).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range1(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range1(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range1(k + 3).項目編號 = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range2(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range2(k + 3).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range2(k + 3).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range2(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range2(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range2(k + 3).項目編號 = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range3(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range3(k + 3).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range3(k + 3).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range3(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range3(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range3(k + 3).項目編號 = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, n, 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range4(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range4(k + 3).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range4(k + 3).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range4(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k + 3).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range4(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range4(k + 3).項目編號 = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, n, 3, i - 4, 10), "&")
        Select Case i - 4
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num1(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num2(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num3(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num4(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
        FormMainMode.PEAFpersoncardback_turn(i).物件類別 = 3
        FormMainMode.PEAFpersoncardback_turn(i).圖片 = app_path & "gif\卡片背面\CBturn.png"
        FormMainMode.PEAFpersoncardback_turn(i).項目編號 = Val(VBEPerson(1, n, 3, i, 8))
        '============================
        Select Case i
            Case 1
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range1(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range1(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range1(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range2(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range2(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range2(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range3(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range3(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range3(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, n, 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range4(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, n, 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 1
                                 Else
                                     FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 3
                                 End If
                             Else
                                 FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             FormMainMode.PEAFpersoncardback_range4(k).物件類別 = 2
                             FormMainMode.PEAFpersoncardback_range4(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, n, 3, i, 10), "&")
        Select Case i
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num1(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num2(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num3(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).物件類別 = 1
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   FormMainMode.PEAFpersoncardback_num4(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
Sub 技能說明載入_人物卡片背面_電腦(ByVal n As Integer)
Dim strw() As String
For i = 1 To 4
    FormMainMode.PEAFpersoncardback_text(i) = VBEPerson(2, n, 3, i, 1)
    '========
    FormMainMode.PEAFpersoncardback_turn(i).物件類別 = 3
    FormMainMode.PEAFpersoncardback_turn(i).圖片 = app_path & "gif\卡片背面\CBturn.png"
    FormMainMode.PEAFpersoncardback_turn(i).項目編號 = Val(VBEPerson(2, n, 3, i, 8))
    '============================
    Select Case i
        Case 1
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range1(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range1(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range1(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range1(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range1(k).項目編號 = 2
                    Next
              End If
        Case 2
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range2(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range2(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range2(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range2(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range2(k).項目編號 = 2
                    Next
              End If
        Case 3
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range3(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range3(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range3(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range3(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range3(k).項目編號 = 2
                    Next
              End If
        Case 4
              If Len(VBEPerson(2, n, 3, i, 9)) = 3 Then
                     For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range4(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range4(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         If Mid(VBEPerson(2, n, 3, i, 9), k, 1) = 1 Then
                             If k < 3 Then
                                 FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 1
                             Else
                                 FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 3
                             End If
                         Else
                             FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 2
                         End If
                    Next
              Else
                    For k = 1 To 3
                         FormMainMode.PEAFpersoncardback_range4(k).物件類別 = 2
                         FormMainMode.PEAFpersoncardback_range4(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                         FormMainMode.PEAFpersoncardback_range4(k).項目編號 = 2
                    Next
              End If
    End Select
    '=========================================
    strw = Split(VBEPerson(2, n, 3, i, 10), "&")
    Select Case i
          Case 1
                For k = 0 To UBound(strw)
                        If Len(strw(k)) = 3 Then
                               FormMainMode.PEAFpersoncardback_num1(k + 1).物件類別 = 1
                               FormMainMode.PEAFpersoncardback_num1(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num1(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                               FormMainMode.PEAFpersoncardback_num2(k + 1).物件類別 = 1
                               FormMainMode.PEAFpersoncardback_num2(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num2(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                               FormMainMode.PEAFpersoncardback_num3(k + 1).物件類別 = 1
                               FormMainMode.PEAFpersoncardback_num3(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num3(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                               FormMainMode.PEAFpersoncardback_num4(k + 1).物件類別 = 1
                               FormMainMode.PEAFpersoncardback_num4(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                               FormMainMode.PEAFpersoncardback_num4(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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

Sub 執行動作_人物卡片背面解除亮光(ByVal n As Integer)
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
Sub 技能說明載入_人物卡片背面_交換角色(ByVal n As Integer)
Dim strw() As String
If n = 2 Then
    For i = 5 To 8
        Formchangeperson.PEAFpersoncardback_text(i) = VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 1)
        '========
        Formchangeperson.PEAFpersoncardback_turn(i).物件類別 = 3
        Formchangeperson.PEAFpersoncardback_turn(i).圖片 = app_path & "gif\卡片背面\CBturn.png"
        Formchangeperson.PEAFpersoncardback_turn(i).項目編號 = Val(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 8))
        '============================
        Select Case i - 4
            Case 1
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range1(k + 3).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range1(k + 3).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range1(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range1(k + 3).項目編號 = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range2(k + 3).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range2(k + 3).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range2(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range2(k + 3).項目編號 = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range3(k + 3).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range3(k + 3).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range3(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range3(k + 3).項目編號 = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range4(k + 3).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range4(k + 3).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range4(k + 3).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range4(k + 3).項目編號 = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i - 4, 10), "&")
        Select Case i - 4
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num1(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num2(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num3(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num4(k + 6).項目編號 = Val(Mid(strw(k), 2, 1))
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
    人物卡面背面編號紀錄數(7) = 0
'===================================================================
Else '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'===================================================================
    For i = 1 To 4
        Formchangeperson.PEAFpersoncardback_text(i) = VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 1)
        '========
        Formchangeperson.PEAFpersoncardback_turn(i).物件類別 = 3
        Formchangeperson.PEAFpersoncardback_turn(i).圖片 = app_path & "gif\卡片背面\CBturn.png"
        Formchangeperson.PEAFpersoncardback_turn(i).項目編號 = Val(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 8))
        '============================
        Select Case i
            Case 1
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range1(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range1(k).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range1(k).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range1(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range1(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range1(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range1(k).項目編號 = 2
                        Next
                  End If
            Case 2
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range2(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range2(k).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range2(k).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range2(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range2(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range2(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range2(k).項目編號 = 2
                        Next
                  End If
            Case 3
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range3(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range3(k).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range3(k).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range3(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range3(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range3(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range3(k).項目編號 = 2
                        Next
                  End If
            Case 4
                  If Len(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9)) = 3 Then
                         For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range4(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             If Mid(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 9), k, 1) = 1 Then
                                 If k < 3 Then
                                     Formchangeperson.PEAFpersoncardback_range4(k).項目編號 = 1
                                 Else
                                     Formchangeperson.PEAFpersoncardback_range4(k).項目編號 = 3
                                 End If
                             Else
                                 Formchangeperson.PEAFpersoncardback_range4(k).項目編號 = 2
                             End If
                        Next
                  Else
                        For k = 1 To 3
                             Formchangeperson.PEAFpersoncardback_range4(k).物件類別 = 2
                             Formchangeperson.PEAFpersoncardback_range4(k).圖片 = app_path & "gif\卡片背面\CBrge.png"
                             Formchangeperson.PEAFpersoncardback_range4(k).項目編號 = 2
                        Next
                  End If
        End Select
        '=========================================
        strw = Split(VBEPerson(1, 角色待機人物紀錄數(1, n + 1), 3, i, 10), "&")
        Select Case i
              Case 1
                    For k = 0 To UBound(strw)
                            If Len(strw(k)) = 3 Then
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num1(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num2(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num3(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).物件類別 = 1
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).圖片 = app_path & "gif\卡片背面\CB" & Mid(strw(k), 1, 1) & "-" & Mid(strw(k), 3, 1) & ".png"
                                   Formchangeperson.PEAFpersoncardback_num4(k + 1).項目編號 = Val(Mid(strw(k), 2, 1))
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
    人物卡面背面編號紀錄數(6) = 0
End If
End Sub
Sub getpage(ByVal k As Integer, m As Integer)
Dim qwp As Integer, n As Integer, uspce As String, uspme As String, yne As Boolean
If Val(公用牌各牌類型紀錄數(0, 1)) < Val(公用牌各牌類型紀錄數(0, 2)) Then
    yne = False
    Do
            Randomize
            qwp = Int(Rnd() * 29) + 1
            Select Case qwp
                    Case 1  '==移1槍1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 2  '==移1槍2類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 3  '==移1槍3類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 4  '==移1盾1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 5  '==移1盾2類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 6  '==移1盾3類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 7  '==移2特3類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 8  '==移3移3類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 9  '==劍6劍6類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 10  '==劍1槍1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 11  '==劍2槍1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 12  '==劍2槍2類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 13  '==劍3槍3類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 14  '==劍5槍5類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 15  '==劍1盾1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 16  '==劍2盾1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 17  '==劍2盾2類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 18  '==劍3盾3類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 19  '==劍1特1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 20  '==劍2特1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 21  '==劍4特1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 22  '==劍5特2類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 23  '==槍4槍4類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 24  '==槍2特1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 25  '==槍3特2類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 26  '==槍4特1類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 27  '==槍5特2類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 28  '==盾5盾5類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
                    Case 29  '==盾3特5類
                         If Val(公用牌各牌類型紀錄數(qwp, 1)) < Val(公用牌各牌類型紀錄數(qwp, 2)) Then
                            公用牌各牌類型紀錄數(qwp, 1) = Val(公用牌各牌類型紀錄數(qwp, 1)) + 1
                            公用牌各牌類型紀錄數(0, 1) = Val(公用牌各牌類型紀錄數(0, 1)) + 1
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
     '==================================隨機轉牌
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
            Case 1 '使用者
                pagecardnum(m, 11) = 0
                FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                FormMainMode.pageusglead.Caption = Val(FormMainMode.pageusglead) + 1
                戰鬥系統類.座標計算_使用者手牌
                牌移動暫時變數(3) = m
                pagecardnum(m, 9) = 240 '指定目前Left(座標)
                pagecardnum(m, 10) = 960 '指定目前Top(座標)
                FormMainMode.card(m).Left = 240
                FormMainMode.card(m).Top = 960
                戰鬥系統類.計算牌移動距離單位
                戰鬥系統類.公用牌回復正面 (牌移動暫時變數(3))
                FormMainMode.card(m).Visible = True
                戰鬥系統類.牌順序增加_手牌_使用者 m
                FormMainMode.牌移動.Enabled = True
                FormMainMode.wmpse1.Controls.stop
                FormMainMode.wmpse1.Controls.play
                一般系統類.檢查音樂播放 1
            Case 2 '電腦
                pagecardnum(m, 11) = 0
                FormMainMode.pageul = Val(FormMainMode.pageul) - 1
                FormMainMode.pagecomglead.Caption = Val(FormMainMode.pagecomglead) + 1
                戰鬥系統類.座標計算_電腦手牌
                牌移動暫時變數(3) = m
                pagecardnum(m, 9) = 240 '指定目前Left(座標)
                pagecardnum(m, 10) = 960 '指定目前Top(座標)
                FormMainMode.card(m).Left = 240
                FormMainMode.card(m).Top = 960
                戰鬥系統類.計算牌移動距離單位
                戰鬥系統類.公用牌變背面
                FormMainMode.card(m).Visible = True
                戰鬥系統類.牌順序增加_手牌_電腦 m
                FormMainMode.牌移動.Enabled = True
                FormMainMode.wmpse1.Controls.stop
                FormMainMode.wmpse1.Controls.play
                一般系統類.檢查音樂播放 1
        End Select
End If
End Sub
Sub 公用牌地圖牌種類配置(ByVal name As String)
Select Case name
     Case "萊丁貝魯格城堡"
           公用牌各牌類型紀錄數(0, 2) = 57
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 0
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 0
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 0
    Case "誘惑森林"
           公用牌各牌類型紀錄數(0, 2) = 50
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 0
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 2
           公用牌各牌類型紀錄數(16, 2) = 0
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 0
           公用牌各牌類型紀錄數(20, 2) = 0
           公用牌各牌類型紀錄數(21, 2) = 1
           公用牌各牌類型紀錄數(22, 2) = 1
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 0
    Case "垃圾之街"
           公用牌各牌類型紀錄數(0, 2) = 55
           公用牌各牌類型紀錄數(1, 2) = 2
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 1
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 0
    Case "冰封湖畔(新)"
           公用牌各牌類型紀錄數(0, 2) = 53
           公用牌各牌類型紀錄數(1, 2) = 4
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 2
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 0
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 1
    Case "人魂墓地"
           公用牌各牌類型紀錄數(0, 2) = 50
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 4
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 0
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 0
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 2
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 0
    Case "盡頭之村"
           公用牌各牌類型紀錄數(0, 2) = 54
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 1
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 0
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 0
           公用牌各牌類型紀錄數(26, 2) = 1
           公用牌各牌類型紀錄數(27, 2) = 0
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 0
    Case "風暴荒野"
           公用牌各牌類型紀錄數(0, 2) = 52
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 1
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 2
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 2
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 0
           公用牌各牌類型紀錄數(20, 2) = 0
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 1
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 0
           公用牌各牌類型紀錄數(29, 2) = 1
    Case "藩骸兒的遺跡"
           公用牌各牌類型紀錄數(0, 2) = 49
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 1
           公用牌各牌類型紀錄數(3, 2) = 1
           公用牌各牌類型紀錄數(4, 2) = 3
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 2
           公用牌各牌類型紀錄數(8, 2) = 0
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 1
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 2
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 1
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 1
    Case "魔都羅占布爾克"
           公用牌各牌類型紀錄數(0, 2) = 42
           公用牌各牌類型紀錄數(1, 2) = 0
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 2
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 1
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 0
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 0
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 1
    Case "瘋狂山脈"
           公用牌各牌類型紀錄數(0, 2) = 47
           公用牌各牌類型紀錄數(1, 2) = 2
           公用牌各牌類型紀錄數(2, 2) = 0
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 2
           公用牌各牌類型紀錄數(5, 2) = 0
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 2
           公用牌各牌類型紀錄數(8, 2) = 1
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 1
    Case "魔女山谷"
           公用牌各牌類型紀錄數(0, 2) = 52
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 3
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 1
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 3
           公用牌各牌類型紀錄數(11, 2) = 1
           公用牌各牌類型紀錄數(12, 2) = 1
           公用牌各牌類型紀錄數(13, 2) = 0
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 1
    Case "烏波斯的黑湖"
           公用牌各牌類型紀錄數(0, 2) = 50
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 1
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 2
           公用牌各牌類型紀錄數(8, 2) = 1
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 2
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 1
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 0
           公用牌各牌類型紀錄數(21, 2) = 1
           公用牌各牌類型紀錄數(22, 2) = 1
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 0
           公用牌各牌類型紀錄數(26, 2) = 1
           公用牌各牌類型紀錄數(27, 2) = 1
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 1
    Case "白魔的圓環石陣"
           公用牌各牌類型紀錄數(0, 2) = 50
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 2
           公用牌各牌類型紀錄數(8, 2) = 0
           公用牌各牌類型紀錄數(9, 2) = 0
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 0
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 1
           公用牌各牌類型紀錄數(22, 2) = 1
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 1
           公用牌各牌類型紀錄數(27, 2) = 1
           公用牌各牌類型紀錄數(28, 2) = 0
           公用牌各牌類型紀錄數(29, 2) = 0
    Case Else
           公用牌各牌類型紀錄數(0, 2) = 57
           公用牌各牌類型紀錄數(1, 2) = 6
           公用牌各牌類型紀錄數(2, 2) = 2
           公用牌各牌類型紀錄數(3, 2) = 2
           公用牌各牌類型紀錄數(4, 2) = 6
           公用牌各牌類型紀錄數(5, 2) = 2
           公用牌各牌類型紀錄數(6, 2) = 1
           公用牌各牌類型紀錄數(7, 2) = 3
           公用牌各牌類型紀錄數(8, 2) = 0
           公用牌各牌類型紀錄數(9, 2) = 1
           公用牌各牌類型紀錄數(10, 2) = 4
           公用牌各牌類型紀錄數(11, 2) = 2
           公用牌各牌類型紀錄數(12, 2) = 2
           公用牌各牌類型紀錄數(13, 2) = 2
           公用牌各牌類型紀錄數(14, 2) = 0
           公用牌各牌類型紀錄數(15, 2) = 4
           公用牌各牌類型紀錄數(16, 2) = 2
           公用牌各牌類型紀錄數(17, 2) = 2
           公用牌各牌類型紀錄數(18, 2) = 2
           公用牌各牌類型紀錄數(19, 2) = 1
           公用牌各牌類型紀錄數(20, 2) = 1
           公用牌各牌類型紀錄數(21, 2) = 2
           公用牌各牌類型紀錄數(22, 2) = 2
           公用牌各牌類型紀錄數(23, 2) = 1
           公用牌各牌類型紀錄數(24, 2) = 1
           公用牌各牌類型紀錄數(25, 2) = 1
           公用牌各牌類型紀錄數(26, 2) = 2
           公用牌各牌類型紀錄數(27, 2) = 2
           公用牌各牌類型紀錄數(28, 2) = 1
           公用牌各牌類型紀錄數(29, 2) = 0
End Select
End Sub
Sub 公用牌未使用檢查()
For i = Val(公用牌各牌類型紀錄數(0, 2)) + 1 To 70
     pagecardnum(i, 6) = 5
Next
End Sub
Sub 傷害執行_立即死亡_使用者(ByVal num As Integer)
Select Case num
   Case 1
        FormMainMode.messageus.AddItem "您受到了" & liveus(角色人物對戰人數(1, 2)) & "點傷害。"
        戰鬥系統類.自動捲軸捲動
        FormMainMode.usbi1(角色人物對戰人數(1, 2)).Caption = 0
        FormMainMode.uspi4(角色人物對戰人數(1, 2)).Caption = 0
        liveus(角色人物對戰人數(1, 2)) = 0
        FormMainMode.bloodnumus1.Caption = 0
        FormMainMode.bloodlineout1.Width = 0
        牌總階段數(1) = 牌總階段數(1) + 1
        戰鬥系統類.播放傷害音樂
   Case Is > 1
        liveus(角色待機人物紀錄數(1, num)) = 0
        If FormMainMode.uspi1(角色待機人物紀錄數(1, num)).Caption = "" Then
            FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption = -liveusmax(角色待機人物紀錄數(1, num))
            FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption = -liveusmax(角色待機人物紀錄數(1, num))
        Else
            FormMainMode.usbi1(角色待機人物紀錄數(1, num)).Caption = 0
            FormMainMode.uspi4(角色待機人物紀錄數(1, num)).Caption = 0
        End If
        牌總階段數(1) = 牌總階段數(1) + 1
End Select
End Sub
Sub 傷害執行_立即死亡_電腦(ByVal num As Integer)
Select Case num
    Case 1
        FormMainMode.messageus.AddItem "對方受到了" & livecom(角色人物對戰人數(2, 2)) & "點傷害。"
        戰鬥系統類.自動捲軸捲動
        FormMainMode.compi4(角色人物對戰人數(2, 2)).Caption = 0
        FormMainMode.cardcompi1(角色人物對戰人數(2, 2)).Caption = 0
        FormMainMode.bloodnumcom1.Caption = 0
        livecom(角色人物對戰人數(2, 2)) = 0
        FormMainMode.bloodlineout2.Left = 11580
        牌總階段數(2) = 牌總階段數(2) + 1
        戰鬥系統類.播放傷害音樂
    Case Is > 1
        If FormMainMode.compi1(角色待機人物紀錄數(2, num)).Caption = "" Then
            FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = -livecommax(角色待機人物紀錄數(2, num))
            FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = -livecommax(角色待機人物紀錄數(2, num))
        Else
            FormMainMode.compi4(角色待機人物紀錄數(2, num)).Caption = 0
            FormMainMode.cardcompi1(角色待機人物紀錄數(2, num)).Caption = 0
        End If
        livecom(角色待機人物紀錄數(2, num)) = 0
        牌總階段數(2) = 牌總階段數(2) + 1
End Select
End Sub

