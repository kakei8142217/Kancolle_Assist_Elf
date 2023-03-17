Public Class KAE_main_form
    Private Sub KAE_main_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        main_form = Me
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ship_locked_form = New KAE_ship_locked_form
        ship_locked_form.Show()
        Button1.Enabled = False
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("鼠标左键微动寿命-2")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("你以为会有二维码弹出来么，其实并没有，至少现在没有" & vbCrLf & "如果你认为本软件比较有用或很有趣的话，我很开心" & vbCrLf & "后续功能更新得看缘分，感谢你的支持，TZFM")
    End Sub
End Class
