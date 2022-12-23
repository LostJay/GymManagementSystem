Imports System.Data.SqlClient

Public Class Form2
    Dim SystemWidth As Integer = SystemInformation.PrimaryMonitorSize.Width, SystemHeight As Integer = SystemInformation.PrimaryMonitorSize.Height
    Dim s_con As String = "server=" & "Zephyrus-LOSER" & ";database=" & "GymManagementSystemUsers" & ";Trusted_Connection=SSPI"
    Dim con As SqlConnection = New SqlConnection(s_con)
    Dim ListB As ListBox = New ListBox(), Lab As Label = New Label(), Lab2 As Label = New Label(), Text1 As TextBox = New TextBox(), btn As Button = New Button, vipuser As String = ""

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "健身房管理系统——管理员界面"
        Me.Size = New Size(SystemWidth * 4 / 5, SystemHeight * 4 / 5)
        Me.AutoSizeMode = False
        MenuStrip1.Size = New Size(SystemWidth * 4 / 5, SystemHeight * 1 / 10)
        MenuStrip1.Font = New Font("黑体", 15)
        con.Open()
        Lab.Text = "健身房管理员"
        Lab.Font = New Font("黑体"， 20)
        Lab.Left = Me.Width * 9 / 20
        Lab.Top = Me.Height / 8
        Lab.Width = Me.Width / 2
        Lab.Height = Me.Height / 16
        Me.Controls.Add(Lab)
        ListB.Size = New Size(Me.Width * 2 / 5, Me.Height * 7 / 10)
        ListB.Left = Me.Width / 10
        ListB.Top = Me.Height / 5
        ListB.Visible = False
        Me.Controls.Add(ListB)
        Lab2.Left = Me.Width * 3 / 5
        Text1.Left = Me.Width * 3 / 5
        btn.Left = Me.Width * 3 / 5
        Lab2.Top = Me.Height * 2 / 5
        Text1.Top = Me.Height * 1 / 2
        btn.Top = Me.Height * 3 / 5
        Lab2.Font = New Font("黑体", 12)
        Text1.Font = New Font("黑体", 12)
        btn.Text = "确  认"
        btn.Font = New Font("黑体", 12)
        Lab2.Width = Me.Width / 5
        Text1.Width = Me.Width / 5
        btn.Width = Me.Width / 5
        Me.Controls.Add(Lab2)
        Me.Controls.Add(Text1)
        Me.Controls.Add(btn)
        Lab2.Visible = False
        btn.Visible = False
        Text1.Visible = False
        AddHandler btn.Click, AddressOf btn_click
    End Sub

    Private Sub 安全退出ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 安全退出ToolStripMenuItem.Click
        Dim close As DialogResult = MsgBox("确认安全退出吗？", MessageBoxButtons.YesNo, "警告")
        If close = Windows.Forms.DialogResult.Yes Then
            con.Close()
            Application.Exit()
        End If
    End Sub

    Private Sub 系统成员管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 系统成员管理ToolStripMenuItem.Click
        Lab.Text = "系统用户列表"
        ListB.Items.Clear()
        ListB.Visible = True
        Dim query_cmd As SqlCommand = New SqlCommand()
        query_cmd.Connection = con
        query_cmd.CommandText = "Select * From UserInfo"
        Dim query_reader As SqlDataReader = query_cmd.ExecuteReader()
        If query_reader.HasRows() = True Then
            Do While query_reader.Read() = True
                Dim res As String = " "
                For i As Integer = 1 To 6
                    res += query_reader.Item(i - 1).ToString + "    "
                Next
                ListB.Items.Add(res)
            Loop
        End If
        query_reader.Close()
        Lab2.Visible = False
        btn.Visible = False
        Text1.Visible = False
    End Sub

    Private Sub 添加系统管理员ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 添加系统管理员ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要添加的系统管理人员账号:"
    End Sub

    Private Sub 删除用户ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除用户ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要删除的系统人员账号:"
    End Sub

    Private Sub 删除管理员ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除管理员ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要删除的系统管理人员账号:"
    End Sub

    Private Sub 添加会员信息ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 添加会员信息ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要添加会员的用户账号:"
    End Sub

    Private Sub 修改会员信息ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 修改会员信息ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要删除的会员卡号:"
    End Sub

    Private Sub 添加器材ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 添加器材ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要添加的器材ID:"
    End Sub

    Private Sub 维护器材ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 维护器材ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您需要维护的器材ID:"
    End Sub

    Private Sub 删除器材ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除器材ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要删除的器材ID:"
    End Sub

    Private Sub 添加商品ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 添加商品ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要添加的商品ID:"
    End Sub

    Private Sub 修改商品ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 修改商品ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要添修改的商品ID:"
    End Sub

    Private Sub 下架商品ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下架商品ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您要下架的商品ID:"
    End Sub

    Private Sub 会员信息管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 会员信息管理ToolStripMenuItem.Click
        Lab.Text = "会员信息列表"
        ListB.Items.Clear()
        ListB.Visible = True
        Dim query_cmd As SqlCommand = New SqlCommand()
        query_cmd.Connection = con
        query_cmd.CommandText = "Select * From VIPInfo"
        Dim query_reader As SqlDataReader = query_cmd.ExecuteReader()
        If query_reader.HasRows() = True Then
            Do While query_reader.Read() = True
                Dim res As String = " "
                For i As Integer = 1 To 3
                    res += query_reader.Item(i - 1).ToString + "    "
                Next
                ListB.Items.Add(res)
            Loop
        End If
        query_reader.Close()
        Lab2.Visible = False
        btn.Visible = False
        Text1.Visible = False
    End Sub

    Private Sub 取消维护器材ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 取消维护器材ToolStripMenuItem.Click
        Lab2.Visible = True
        btn.Visible = True
        Text1.Visible = True
        Lab2.Text = "请输入您取消维护的器材ID:"
    End Sub

    Private Sub 运动器材管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 运动器材管理ToolStripMenuItem.Click
        Lab.Text = "运动器材列表"
        ListB.Items.Clear()
        ListB.Visible = True
        Dim query_cmd As SqlCommand = New SqlCommand()
        query_cmd.Connection = con
        query_cmd.CommandText = "Select * From FacilityInfo"
        Dim query_reader As SqlDataReader = query_cmd.ExecuteReader()
        If query_reader.HasRows() = True Then
            Do While query_reader.Read() = True
                Dim res As String = " "
                For i As Integer = 1 To 3
                    res += query_reader.Item(i - 1).ToString + "    "
                Next
                ListB.Items.Add(res)
            Loop
        End If
        query_reader.Close()
        Lab2.Visible = False
        btn.Visible = False
        Text1.Visible = False
    End Sub

    Private Sub 系统管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 系统管理ToolStripMenuItem.Click
        Lab.Text = "系统管理"
        ListB.Visible = False
    End Sub

    Private Sub autosize1(sender As Object, e As EventArgs) Handles Me.Resize
        MenuStrip1.Size = New Size(SystemWidth * 4 / 5, SystemHeight * 1 / 10)
        Lab.Left = Me.Width * 9 / 20
        Lab.Top = Me.Height / 8
        Lab.Width = Me.Width / 2
        Lab.Height = Me.Height / 16
        ListB.Size = New Size(Me.Width * 2 / 5, Me.Height * 7 / 10)
        ListB.Left = Me.Width / 10
        ListB.Top = Me.Height / 5
        Lab2.Left = Me.Width * 3 / 5
        Text1.Left = Me.Width * 3 / 5
        btn.Left = Me.Width * 3 / 5
        Lab2.Top = Me.Height * 2 / 5
        Text1.Top = Me.Height * 1 / 2
        btn.Top = Me.Height * 3 / 5
        Lab2.Font = New Font("黑体", 12)
        Text1.Font = New Font("黑体", 12)
        btn.Text = "确  认"
        btn.Font = New Font("黑体", 12)
        Lab2.Width = Me.Width / 5
        Text1.Width = Me.Width / 5
        btn.Width = Me.Width / 5
    End Sub
    Private Sub 库存管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 库存管理ToolStripMenuItem.Click
        Lab.Text = "库存列表"
        ListB.Items.Clear()
        ListB.Visible = True
        Dim query_cmd As SqlCommand = New SqlCommand()
        query_cmd.Connection = con
        query_cmd.CommandText = "Select * From Stock"
        Dim query_reader As SqlDataReader = query_cmd.ExecuteReader()
        If query_reader.HasRows() = True Then
            Do While query_reader.Read() = True
                Dim res As String = " "
                For i As Integer = 1 To 4
                    res += query_reader.Item(i - 1).ToString + "    "
                Next
                ListB.Items.Add(res)
            Loop
        End If
        query_reader.Close()
        Lab2.Visible = False
        btn.Visible = False
        Text1.Visible = False
    End Sub
    Private Sub btn_click(ByVal s As Object, e As EventArgs)
        Dim text_lab2 As String = Lab2.Text, text_t As String = Text1.Text, text_l As Boolean = True, check As Char
        Select Case text_lab2
            Case "请输入您要添加的系统管理人员账号:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from UserInfo Where Users=" + text_t
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的账号不存在！",, "错误")
                        query_reult.Close()
                    Else
                        query_reult.Read()
                        Dim ifa As Boolean = query_reult.Item(2)
                        query_reult.Close()
                        If ifa = True Then
                            MsgBox("对不起，您输入的账号已经是管理员！",, "错误")
                        Else
                            query_cmd.CommandText = "update UserInfo Set IfAdmin = 1 where Users = " & text_t
                            query_cmd.ExecuteNonQuery()
                            MsgBox("用户" + text_t + "成功设置为管理员！",, "信息")
                            系统成员管理ToolStripMenuItem.PerformClick()
                        End If
                    End If
                End If

            Case "请输入您要删除的系统管理人员账号:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from UserInfo Where Users=" + text_t
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的账号不存在！",, "错误")
                    Else
                        query_reult.Read()
                        Dim ifa As Boolean = query_reult.Item(2)
                        query_reult.Close()
                        If ifa = False Then
                            MsgBox("对不起，您输入的账号不是管理员！",, "错误")
                        Else
                            query_cmd.CommandText = "update UserInfo Set IfAdmin = 0 where Users = " & text_t
                            query_cmd.ExecuteNonQuery()
                            MsgBox("用户" + text_t + "成功取消管理员权限！",, "信息")
                            系统成员管理ToolStripMenuItem.PerformClick()
                        End If
                    End If
                End If
            Case "请输入您需要维护的器材ID:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") And Asc(check) < Asc("A") Or Asc(check) > Asc("Z") And Asc(check) < Asc("a") Or Asc(check) > Asc("z") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from FacilityInfo Where FacilityID = '" + text_t + "'"
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的器材不存在！",, "错误")
                        query_reult.Close()
                    Else
                        query_reult.Read()
                        Dim ifa As Boolean = query_reult.Item(1)
                        query_reult.Close()
                        If ifa = True Then
                            MsgBox("对不起，该器材已经在维护中！",, "错误")
                        Else
                            query_cmd.CommandText = "update FacilityInfo Set IfMaintained = 1 where FacilityID = '" & text_t + "'"
                            query_cmd.ExecuteNonQuery()
                            MsgBox("器材" + text_t + "成功设置为维护状态！",, "信息")
                            运动器材管理ToolStripMenuItem.PerformClick()
                        End If
                    End If
                End If
            Case "请输入您取消维护的器材ID:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") And Asc(check) < Asc("A") Or Asc(check) > Asc("Z") And Asc(check) < Asc("a") Or Asc(check) > Asc("z") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from FacilityInfo Where FacilityID = '" + text_t + "'"
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的器材不存在！",, "错误")
                        query_reult.Close()
                    Else
                        query_reult.Read()
                        Dim ifa As Boolean = query_reult.Item(1)
                        query_reult.Close()
                        If ifa = False Then
                            MsgBox("对不起，该器材不在维护中！",, "错误")
                        Else
                            query_cmd.CommandText = "update FacilityInfo Set IfMaintained = 0 where FacilityID = '" & text_t + "'"
                            query_cmd.ExecuteNonQuery()
                            MsgBox("器材" + text_t + "成功设置为正常状态！",, "信息")
                            运动器材管理ToolStripMenuItem.PerformClick()
                        End If
                    End If
                End If
            Case "请输入您要删除的系统人员账号:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from UserInfo Where Users=" + text_t
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的账号不存在！",, "错误")
                    Else
                        query_reult.Close()
                        query_cmd.CommandText = "Delete from UserInfo where Users = " + text_t
                        query_cmd.ExecuteNonQuery()
                        MsgBox("账户" + text_t + "已成功删除！",, "信息")
                        系统成员管理ToolStripMenuItem.PerformClick()
                    End If
                End If
            Case "请输入您要删除的器材ID:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") And Asc(check) < Asc("A") Or Asc(check) > Asc("Z") And Asc(check) < Asc("a") Or Asc(check) > Asc("z") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from FacilityInfo Where FacilityID = '" + text_t + "'"
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的器材不存在！",, "错误")
                        query_reult.Close()
                    Else
                        query_reult.Close()
                        query_cmd.CommandText = "Delete from FacilityInfo where FacilityID = '" + text_t + "'"
                        query_cmd.ExecuteNonQuery()
                        MsgBox("设备" + text_t + "已成功删除！",, "信息")
                        运动器材管理ToolStripMenuItem.PerformClick()
                    End If
                End If
            Case "请输入您要下架的商品ID:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") And Asc(check) < Asc("A") Or Asc(check) > Asc("Z") And Asc(check) < Asc("a") Or Asc(check) > Asc("z") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from Stock Where CommodityID = '" + text_t + "'"
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的商品不存在！",, "错误")
                        query_reult.Close()
                    Else
                        query_reult.Close()
                        query_cmd.CommandText = "Delete from Stock where CommodityID = '" + text_t + "'"
                        query_cmd.ExecuteNonQuery()
                        MsgBox("商品" + text_t + "已成功下架！",, "信息")
                        下架商品ToolStripMenuItem.PerformClick()
                    End If
                End If
            Case "请输入您要添加会员的用户账号:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from UserInfo Where Users=" + text_t
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的账号不存在！",, "错误")
                    Else
                        Lab2.Text = "请输入注册会员卡号："
                        MsgBox("请输入注册会员卡号：",, "信息")
                        vipuser = text_t
                    End If
                    query_reult.Close()
                End If
            Case "请输入您要删除的会员卡号:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    query_cmd.CommandText = "Select * from VIPInfo Where VIPID = " + text_t
                    Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
                    If query_reult.HasRows = False Then
                        MsgBox("对不起，您输入的会员卡号不存在！",, "错误")
                    Else
                        query_reult.Close()
                        query_cmd.CommandText = "Delete from VIPInfo Where VIPID = " + text_t
                        query_cmd.ExecuteNonQuery()
                        MsgBox("会员卡号" + text_t + "已成功失效！",, "信息")
                        修改会员信息ToolStripMenuItem.PerformClick()
                    End If
                End If
            Case "请输入注册会员卡号："
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) < Asc("0") Or Asc(check) > Asc("9") Then
                        text_l = False
                    End If
                Next
                If text_l = True Then
                    If text_t.Length <> 19 Then text_l = False
                End If
                If text_l = True Then
                    MsgBox(Convert.ToInt32(Mid(text_t, 13, 2)))
                    If Convert.ToInt32(Mid(text_t, 13, 2)) > 12 Or Convert.ToInt32(Mid(text_t, 13, 2)) = 0 Then text_l = False
                End If
                If text_l = True Then
                    Select Case Mid(text_t, 13, 2)
                        Case "01", "03", "05", "07", "08", "10", "12"
                            If Convert.ToInt32(Mid(text_t, 15, 2)) > 31 Or Convert.ToInt32(Mid(text_t, 15, 2)) = 0 Then text_l = False
                        Case "02"
                            If Convert.ToInt32(Mid(text_t, 9, 4)) Mod 400 = 0 Or Convert.ToInt32(Mid(text_t, 9, 4)) Mod 4 = 0 And Convert.ToInt32(Mid(text_t, 9, 4)) Mod 100 <> 0 Then
                                If Convert.ToInt32(Mid(text_t, 15, 2)) > 29 Or Convert.ToInt32(Mid(text_t, 15, 2)) = 0 Then text_l = False
                            Else
                                If Convert.ToInt32(Mid(text_t, 15, 2)) > 28 Or Convert.ToInt32(Mid(text_t, 15, 2)) = 0 Then text_l = False
                            End If
                        Case Else
                            If Convert.ToInt32(Mid(text_t, 15, 2)) > 30 Or Convert.ToInt32(Mid(text_t, 15, 2)) = 0 Then text_l = False
                    End Select
                End If
                If text_l = False Then
                    MsgBox("会员卡号格式不正确,请重新输！",, "错误")
                Else
                    Dim query_cmd As SqlCommand = New SqlCommand()
                    query_cmd.Connection = con
                    Dim vipdate As String = Mid(text_t, 9, 4) + "/" + Mid(text_t, 13, 2) + "/" + Mid(text_t, 15, 2)
                    query_cmd.CommandText = "insert into VIPInfo (VIPID,Users,ExpirationTime) values(" + text_t + "," + vipuser + ",'" + vipdate + "')"
                    query_cmd.ExecuteNonQuery()
                    MsgBox("会员卡号" + text_t + "已成功添加到" + vipuser + "的卡包！",, "信息")
                    会员信息管理ToolStripMenuItem.PerformClick()
                End If
            Case "请输入您要添加的器材ID:"
                For i As Integer = 1 To text_t.Length
                    check = Mid(text_t, i)
                    If Asc(check) <Asc("0") Or Asc(check) > Asc("9") And Asc(check) <Asc("A") Or Asc(check) > Asc("Z") And Asc(check) <Asc("a") Or Asc(check) > Asc("z") Then
                        text_l = False
                    End If
                Next
                If text_l = False Then
                    MsgBox("请输入正确的格式！",, "错误")
                Else

                End If
        End Select
        Text1.Text = ""
    End Sub
End Class

