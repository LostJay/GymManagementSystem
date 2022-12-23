Imports System.Data.SqlClient
Public Class Form1
    Dim L_User As New Label, L_Password As New Label, T_User As New TextBox, T_Password As New TextBox, B_Login As New Button
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximizeBox = False
        Dim SystemWidth As Integer = SystemInformation.PrimaryMonitorSize.Width, SystemHeight As Integer = SystemInformation.PrimaryMonitorSize.Height
        MyBase.Width = SystemWidth * 0.26
        MyBase.Height = SystemHeight * 0.32
        MyBase.AutoSizeMode = False
        MyBase.Text = "健身房管理系统"
        MyBase.Left = SystemWidth / 2 - MyBase.Width / 2
        MyBase.Top = SystemHeight / 2 - MyBase.Height / 2
        Dim w As Integer = MyBase.Width, h As Integer = MyBase.Height
        L_User.Text = "账号："
        L_Password.Text = "密码："
        L_User.Left = w * 0.17
        L_User.Top = h * 0.4
        L_User.Font = New Font("黑体", 12)
        L_Password.Left = w * 0.17
        L_Password.Top = L_User.Top + L_User.Height * 2
        L_Password.Font = New Font("黑体", 12)
        T_User.Left = L_User.Left + L_User.Width
        T_User.Top = L_User.Top
        T_User.Font = New Font("黑体", 12)
        T_User.Width = w * 0.45
        T_User.MaxLength = 12
        T_Password.Left = L_Password.Left + L_Password.Width
        T_Password.Top = L_Password.Top
        T_Password.Font = New Font("黑体", 12)
        T_Password.Width = w * 0.45
        T_Password.PasswordChar = "*"
        T_Password.MaxLength = 12
        B_Login.Text = "登    录"
        B_Login.Font = New Font("黑体", 16)
        B_Login.Left = w * 0.175
        B_Login.Top = h * 0.7
        B_Login.Width = w * 0.6
        B_Login.Height = h * 0.1
        Me.Controls.Add(L_User)
        Me.Controls.Add(L_Password)
        Me.Controls.Add(T_User)
        Me.Controls.Add(T_Password)
        Me.Controls.Add(B_Login)
        AddHandler B_Login.Click, AddressOf Login_Click
        AddHandler T_User.KeyPress, AddressOf UserInputLimit
    End Sub

    Private Sub Login_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim u As String = T_User.Text, p As String = T_Password.Text
        Dim S_Database As String = "server=" & "ZEPHYRUS-LOSER" & ";database=" & "GymManagementSystemUsers" & ";Trusted_Connection=SSPI"
        If u = "" Then
            MsgBox("账号不能为空！",, "错误")
        Else
            Dim con As SqlConnection = New SqlConnection(S_Database)
            con.Open()
            Dim query_cmd As SqlCommand = New SqlCommand()
            query_cmd.Connection = con
            query_cmd.CommandText = "Select * from UserInfo Where Users=" + u
            Dim query_reult As SqlDataReader = query_cmd.ExecuteReader()
            If query_reult.HasRows = False Then
                MsgBox("对不起，您输入的账号不存在！",, "错误")
            Else
                If p = "" Then
                    MsgBox("密码不能为空！",, "错误")
                Else
                    query_reult.Read()
                    Dim p1 As String = query_reult.Item(1)
                    If p1 <> p Then
                        MsgBox("对不起，您输入的密码有误！",, "错误")
                    Else
                        Dim admin As Boolean = query_reult.Item(2), nickn As String = query_reult.Item(4)
                        Me.Visible = False
                        MsgBox("欢迎进入健身房管理系统，" + nickn,， "消息")
                        If admin = True Then
                            Form2.Visible = True
                        Else
                            Form3.Visible = True
                        End If
                    End If
                End If
            End If
            query_reult.Close()
            con.Close()
        End If
    End Sub

    Private Sub UserInputLimit(ByVal sender As Object, e As KeyPressEventArgs)
        If Not (Char.IsNumber(e.KeyChar) Or e.KeyChar = Chr(Keys.Back)) Then
            e.Handled = True
        End If
    End Sub
End Class
