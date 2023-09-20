Public Class Form1
    Dim Box(29), J, flag, n1, t, sum, ck, ck1, ck2, pick1, pick2, timesum, ps As Integer 'ps=讀歷史紀錄
    'J=判斷  flag=拿來判別有沒有重複的數字  n1 = 測試用數字 t=打開的時間 sum=蓋幾張了 ck=翻第幾張牌 ck1=圖片編號1 ck2圖片編號2 pick1=第一張圖的位子 pick2=第二張圖的位子
    Dim Box1(30) As Button '全部的按紐
    Dim timer As String '總共花的時間

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim str As String
        FileOpen(1, "1.txt", OpenMode.Input)
        Do Until EOF(1)
            str = LineInput(1)
            TextBox1.Text += str & "秒全部完成" & vbNewLine
        Loop
        FileClose()

        Box1(1) = Button1 : Box1(2) = Button2 : Box1(3) = Button3 : Box1(4) = Button4 : Box1(5) = Button5 : Box1(6) = Button6 : Box1(7) = Button7 : Box1(8) = Button8 : Box1(9) = Button9 : Box1(10) = Button10
        Box1(11) = Button11 : Box1(12) = Button12 : Box1(13) = Button13 : Box1(14) = Button14 : Box1(15) = Button15 : Box1(16) = Button16 : Box1(17) = Button17 : Box1(18) = Button18 : Box1(19) = Button19 : Box1(20) = Button20
        Box1(21) = Button21 : Box1(22) = Button22 : Box1(23) = Button23 : Box1(24) = Button24 : Box1(25) = Button25 : Box1(26) = Button26 : Box1(27) = Button27 : Box1(28) = Button28 : Box1(29) = Button29 : Box1(30) = Button30
        For i = 1 To 30
            Box1(i).Enabled = False
        Next
    End Sub

    Public Sub button_click(ByVal sender As System.Object, e As System.EventArgs) Handles _
        Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, Button10.Click,
        Button11.Click, Button12.Click, Button13.Click, Button14.Click, Button15.Click, Button16.Click, Button17.Click, Button18.Click, Button19.Click, Button20.Click,
        Button21.Click, Button22.Click, Button23.Click, Button24.Click, Button25.Click, Button26.Click, Button27.Click, Button28.Click, Button29.Click, Button30.Click
        If ck = 0 Then
            pick1 = sender.accessiblename '第一張圖的位子
            sender.image = New Bitmap(sender.tag.ToString & ".png") '打開圖片 
            ck = 1 ' 1 = 第一張
            ck1 = sender.tag 'ck1=圖片編號1
        Else
            sender.image = New Bitmap(sender.tag.ToString & ".png") : pick2 = sender.accessiblename : ck2 = sender.tag 'ck2=圖片編號2 pick2=第二張圖的位子
            If pick1 = pick2 Then '避免重複按到同一個圖片
                ck = 0
            Else
                time(0.5) '暫停
                If ck2 = ck1 Then '判定兩張圖是否一樣
                    Box1(pick1).Visible = False : Box1(pick2).Visible = False
                    sender.visible = False '如果是的話就把第二張蓋起來
                    sum += 1 'sum =蓋幾張了
                Else
                    redo()
                End If
                ck = 0
            End If
        End If
        If sum = 15 Then '偵測遊戲結束沒 總共要蓋15張
            Timer1.Enabled = False
            MsgBox("所花的時間" & timesum & "秒", , "遊戲結束")
        End If
    End Sub

    Public Sub time(ByVal t) '時間暫停
        Application.DoEvents() '暫停
        System.Threading.Thread.Sleep(t * 1000) '暫停t=幾秒
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer += 1 '秒數
        ToolStripStatusLabel1.Text = "已經花了" & timer & "秒"
        timesum = timer '總共花的時間
    End Sub

    Public Sub md() '洗牌
        Randomize()
        For i = 0 To 29
            n1 = Int((Rnd() * 30) + 1)
            J = 1
            Do While J = 1
                For x = 0 To 29
                    If n1 = Box(x) Then 'n1如果重複了 就執行下面的
                        flag = 1
                        Exit For
                    Else
                        flag = 0
                    End If
                Next
                If flag = 1 Then
                    n1 = Int((Rnd() * 30) + 1) '如果flag=1 那就=重複了 重取亂數
                    J = 1
                Else
                    J = 0
                End If
            Loop
            Box(i) = n1
        Next
        For i = 0 To 29
            If Box(i) >= 16 Then
                Box(i) -= 15 '總共有30個 超過15的就-15   n1=box(?)圖片編號
            End If
            Box1(i + 1).Image = New Bitmap(Box(i) & ".png") '對應的圖片
            Box1(i + 1).Tag = (Box(i)) 'tag=編號 全部+上邊號
        Next
    End Sub

    Sub redo() '全變成問號
        For i = 1 To 30
            Box1(i).Image = New Bitmap("問號.png")
        Next
    End Sub

    Sub redo2() '記住按鈕的位子 1到30
        For i = 1 To 30
            Box1(i).AccessibleName = i
        Next
    End Sub

    Private Sub 開始ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 開始ToolStripMenuItem.Click
        If TextBox2.Text = "" Then
            MsgBox("請輸入名稱")
        Else
            For i = 1 To 30
                Box1(i).Enabled = True
            Next
            md()
            time(3) '打開幾秒
            redo()
            redo2()
            Timer1.Interval = 1000 '千分之一秒 速率
            Timer1.Enabled = True '可執行timer1_tick
        End If
    End Sub

    Private Sub 存入歷史紀錄ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 存入歷史紀錄ToolStripMenuItem.Click
        Dim str As String
        str = TextBox2.Text
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("1.txt", True)
        file.Write(str & " " & timesum)
        file.Close()
    End Sub

    Private Sub 讀取歷史紀錄ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 讀取歷史紀錄ToolStripMenuItem.Click
        TextBox1.Clear()
        Dim str As String
        FileOpen(1, "1.txt", OpenMode.Input)
        Do Until EOF(1)
            str = LineInput(1)
            TextBox1.Text += str & "秒全部完成" & vbNewLine
        Loop
        FileClose()
    End Sub
End Class
