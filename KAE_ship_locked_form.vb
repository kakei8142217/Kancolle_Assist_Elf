Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar

Public Class KAE_ship_locked_form
    Dim ship_locked As New ship_locked_class
    Dim file_system_dialog As New KPL.File.File_system_dialog_class


    Private Sub KAE_ship_locked_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ship_locked_form = Me

        Dim debug_mode As Boolean = 0
        If debug_mode = False Then
            Button4.Visible = False
            Button5.Visible = False
            Button6.Visible = False
            Button7.Visible = False
            Button8.Visible = False
        End If

        For Each dir As String In IO.Directory.GetDirectories(Application.StartupPath & "\extra_operation")
            ListBox1.Items.Add(IO.Path.GetFileName(dir))
        Next

        ComboBox1.Items.Add("全舰娘")
        ComboBox1.Items.Add("多号机优先")

        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub KAE_ship_locked_form_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If ship_locked.extra_selected = True Then
            file_handle.export_main_ship_group_to_xml(Application.StartupPath & "\extra_operation\23_spring\ship.xml")
            file_handle.export_extra_group_to_xml(Application.StartupPath & "\extra_operation\23_spring\extra.xml")
        End If
        main_form.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox1.SelectedIndex > -1 Then
            Dim return_mode As Integer = ship_locked.load_extra_and_user_data(ListBox1)
            If return_mode = 0 Then
                MsgBox("加载活动失败T_T未找到[extra.xml]，请检查->")
            ElseIf return_mode = 2 Then
                MsgBox("未检测到用户舰娘数据[ship.xml]，点击确定开始由csv转换[ship.xml]")
                Dim ship_csv_path As String = file_system_dialog.open_file_dialog("csv")
                If ship_csv_path <> "0" Then
                    file_handle.csv_to_ship_xml(ship_csv_path, Application.StartupPath & "\extra_operation\23_spring\ship.xml")
                    MsgBox("[ship.xml]转换完成,重启本程序生效")
                    Label5.Text = "重启本程序以读取[ship.xml]"
                End If
            ElseIf return_mode = 3 Then
                Label3.Text = "板凳舰娘【***[限定装备格数][是否打孔]筛选不可用***】"
                Label5.Text = "选择EX"
                MsgBox("检测到你的舰娘文件[ship.xml]由poi导出的csv文件转换生成" & vbCrLf & "以下筛选功能" & vbCrLf & "[限定装备格数][是否打孔]" & vbCrLf & "将无法使用" & vbCrLf & "(所有舰娘默认有99个装备格并已打孔)")
            Else
                Label5.Text = "选择EX"
            End If
            Button1.Enabled = False
            Button1.Visible = False
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ship_locked.extra_selected = True Then
            ship_locked.refurbish_jion_ship_listbox(ListBox1, ListBox2, ListBox3, Label1, Label2)
        End If
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

    End Sub

    Private Sub ListBox2_Click(sender As Object, e As EventArgs) Handles ListBox2.Click
        If ship_locked.extra_selected = True Then
            If ship_locked.refurbuish_blank_ship_listbox1(ListBox1, ListBox2, ListBox3, ListBox4, ComboBox1, 0) Then ListBox3.SelectedIndex = -1
        End If
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged

    End Sub

    Private Sub ListBox3_Click(sender As Object, e As EventArgs) Handles ListBox3.Click
        If ship_locked.extra_selected = True Then
            If ship_locked.refurbuish_blank_ship_listbox1(ListBox1, ListBox3, ListBox3, ListBox4, ComboBox1, 1) Then ListBox2.SelectedIndex = -1
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ship_locked.extra_selected = True Then
            ship_locked.refurbuish_blank_ship_listbox1(ListBox1, ListBox2, ListBox3, ListBox4, ComboBox1, ship_locked.last_fleet)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ship_locked.ship_join_extra(ListBox1, ListBox2, ListBox3, ListBox4) Then
            Dim listbox_selectedindex As Integer
            If ship_locked.last_fleet = 0 Then
                listbox_selectedindex = ListBox2.SelectedIndex
            Else
                listbox_selectedindex = ListBox3.SelectedIndex
            End If
            ship_locked.refurbish_jion_ship_listbox(ListBox1, ListBox2, ListBox3, Label1, Label2)
            If ship_locked.last_fleet = 0 Then
                ListBox2.SelectedIndex = listbox_selectedindex
            Else
                ListBox3.SelectedIndex = listbox_selectedindex
            End If
            ship_locked.refurbuish_blank_ship_listbox1(ListBox1, ListBox2, ListBox3, ListBox4, ComboBox1, ship_locked.last_fleet)
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ship_locked.ship_exit_extra(ListBox1, ListBox2, ListBox3) Then
            Dim listbox_selectedindex As Integer
            If ship_locked.last_fleet = 0 Then
                listbox_selectedindex = ListBox2.SelectedIndex
            Else
                listbox_selectedindex = ListBox3.SelectedIndex
            End If
            ship_locked.refurbish_jion_ship_listbox(ListBox1, ListBox2, ListBox3, Label1, Label2)
            If ship_locked.last_fleet = 0 Then
                ListBox2.SelectedIndex = listbox_selectedindex
            Else
                ListBox3.SelectedIndex = listbox_selectedindex
            End If
            ship_locked.refurbuish_blank_ship_listbox1(ListBox1, ListBox2, ListBox3, ListBox4, ComboBox1, ship_locked.last_fleet)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '从舰娘特殊装备模板CSV转换为XML
        Dim ship_special_equip_path As String = file_system_dialog.open_file_dialog("csv")
        If ship_special_equip_path <> "0" Then
            file_handle.csv_to_special_equip_xml(ship_special_equip_path, Application.StartupPath & "\system_data\ship_special_equip.xml")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '从用户舰娘CSV转换为XML
        Dim ship_csv_path As String = file_system_dialog.open_file_dialog("csv")
        If ship_csv_path <> "0" Then
            file_handle.csv_to_ship_xml(ship_csv_path, Application.StartupPath & "\extra_operation\23_spring\ship.xml")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'If ship_locked.ship_join_extra(1, ListBox1, ListBox3, ListBox4) Then
        '    Dim listbox3_selectedindex As Integer = ListBox3.SelectedIndex
        '    ship_locked.refurbish_jion_ship_listbox(ListBox1, ListBox2, ListBox3, Label1, Label2)
        '    ListBox3.SelectedIndex = listbox3_selectedindex
        'End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'If ship_locked.ship_exit_extra(1, ListBox1, ListBox3) Then
        '    Dim listbox3_selectedindex As Integer = ListBox3.SelectedIndex
        '    ship_locked.refurbish_jion_ship_listbox(ListBox1, ListBox2, ListBox3, Label1, Label2)
        '    ListBox3.SelectedIndex = listbox3_selectedindex
        'End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        file_handle.export_main_ship_group_to_xml(Application.StartupPath & "\extra_operation\23_spring\ship1.xml")
        file_handle.export_extra_group_to_xml(Application.StartupPath & "\extra_operation\23_spring\extra1.xml")
    End Sub


End Class

Public Class ship_locked_class

    Dim extra_selected_value As Boolean = False
    Dim last_select_fleet As Integer

    Dim join_extra_ship As ship_group_class
    Dim blank_ship_group1 As ship_group_class

    Public Property extra_selected As Boolean
        Get
            extra_selected = extra_selected_value
        End Get
        Set(value As Boolean)
            extra_selected_value = value
        End Set
    End Property

    Public ReadOnly Property last_fleet As Integer
        Get
            last_fleet = last_select_fleet
        End Get
    End Property

    Public Function load_extra_and_user_data(ByVal Listbox1 As ListBox) As Integer
        Dim extra_name As String = Listbox1.Items(Listbox1.SelectedIndex).ToString

        If IO.File.Exists(Application.StartupPath & "\extra_operation\" & extra_name & "\extra.xml") = False Then
            Return 0
            Exit Function
        End If

        If IO.File.Exists(Application.StartupPath & "\extra_operation\" & extra_name & "\ship.xml") = False Then
            Return 2
            Exit Function
        End If

        Listbox1.Items.Clear()

        file_handle.load_ship_to_main_ship_group(Application.StartupPath & "\extra_operation\" & extra_name & "\ship.xml")
        file_handle.load_extra_to_extra_group(Application.StartupPath & "\extra_operation\" & extra_name & "\extra.xml")
        file_handle.load_type_to_type_short_group(Application.StartupPath & "\system_data\type_short.xml")

        For a = 0 To extra_group.Length - 1
            Listbox1.Items.Add(extra_group.extra(a).Name)
        Next

        extra_selected_value = True

        If main_ship_group.Name = "poi" Then
            Return 3
        Else
            Return 1
        End If

    End Function

    Public Sub refurbish_jion_ship_listbox(ByRef extra_listbox1 As ListBox, ByRef join_listbox1 As ListBox, ByRef join_listbox2 As ListBox， ByRef fleet1_name_label1 As Label, ByRef fleet1_name_label2 As Label)
        If extra_listbox1.SelectedIndex >= 0 Then
            join_listbox1.Items.Clear()
            join_listbox2.Items.Clear()

            Dim fleet As fleet_class = extra_group.extra(extra_listbox1.SelectedIndex).get_fleet_group.fleet(0)
            If fleet IsNot Nothing Then
                fleet1_name_label1.Text = fleet.Name
                Dim showstr As String() = get_join_listbox_showstr(fleet)
                For a = 0 To showstr.Length - 1
                    join_listbox1.Items.Add(showstr(a))
                Next
            Else
                fleet1_name_label2.Text = "舰队2"
            End If

            fleet = extra_group.extra(extra_listbox1.SelectedIndex).get_fleet_group.fleet(1)
            If fleet IsNot Nothing Then
                fleet1_name_label2.Text = fleet.Name
                Dim showstr As String() = get_join_listbox_showstr(fleet)
                For a = 0 To showstr.Length - 1
                    join_listbox2.Items.Add(showstr(a))
                Next
            Else
                fleet1_name_label2.Text = "舰队2"
            End If
        End If
    End Sub

    Private Function get_join_listbox_showstr(ByVal fleet As fleet_class) As String()
        Dim return_str As String()
        ReDim return_str(-1)

        For a = 0 To fleet.Length - 1
            Dim showstr As String = ""
            If fleet.hole(a).ship_in_hole Then
                showstr = showstr & "[" & fleet.hole(a).get_ship.ship_name & "]--[" & fleet.hole(a).get_ship.Sort_value & "][" & fleet.hole(a).get_ship.get_type.Name & "][" & fleet.hole(a).get_ship.get_sticker.Name & "]"
            Else
                showstr = showstr & "[" & fleet.hole(a).get_ship.ship_name & "][" & fleet.hole(a).get_ship.Sort_value & "+]["
                For b = 0 To fleet.hole(a).get_type_short_group.Length - 1
                    showstr = showstr & fleet.hole(a).get_type_short_group.type_short(b).Name
                    If b <> fleet.hole(a).get_type_short_group.Length - 1 Then
                        showstr = showstr & "/"
                    Else
                        showstr = showstr & "]"
                    End If
                Next

                If fleet.hole(a).get_special_equip.Get_dna.Value("speed") = "" Then
                    showstr = showstr & "[不限速]"
                Else
                    showstr = showstr & "[" & fleet.hole(a).get_special_equip.Get_dna.Value("speed") & "]"
                End If

                If fleet.hole(a).Get_dna.Value("tear") = 0 Then
                    showstr = showstr & "[不限定打孔]"
                Else
                    showstr = showstr & "[必须打孔]"
                End If

                If fleet.hole(a).Get_dna.Value("hole_count") = 0 Then
                    showstr = showstr & "[不限定装备格数]"
                Else
                    showstr = showstr & "[至少" & fleet.hole(a).Get_dna.Value("hole_count") & "个装备格]"
                End If

                Dim have_special_equip_demand As Boolean = False
                For b = 3 To fleet.hole(a).get_special_equip.Get_dna.Length - 1
                    If fleet.hole(a).get_special_equip.Get_dna.Value(b) = "1" Then
                        If have_special_equip_demand = False Then
                            showstr = showstr & "["
                        End If
                        If have_special_equip_demand = True Then
                            showstr = showstr & "+"
                        End If
                        showstr = showstr & fleet.hole(a).get_special_equip.Get_dna.Key(b)
                        If have_special_equip_demand = False Then
                            have_special_equip_demand = True
                        End If
                    End If
                    If have_special_equip_demand = True And b = fleet.hole(a).get_special_equip.Get_dna.Length - 1 Then
                        showstr = showstr & "]"
                    End If
                Next
            End If
            If fleet.hole(a).Name <> "" Then
                showstr = showstr & "[" & fleet.hole(a).Name & "]"
            End If
            ReDim Preserve return_str(return_str.Length)
            return_str(return_str.Length - 1) = showstr
        Next

        Return return_str
    End Function

    Public Function refurbuish_blank_ship_listbox1(ByRef extra_listbox1 As ListBox, ByRef join_listbox1 As ListBox, ByVal join_listbox2 As ListBox, ByRef blank_listbox1 As ListBox, ByRef select_mode_combobox1 As ComboBox, ByVal fleet_index As Integer) As Boolean
        last_select_fleet = fleet_index
        Dim join_listbox As ListBox
        If last_select_fleet = 0 Then
            join_listbox = join_listbox1
        Else
            join_listbox = join_listbox2
        End If
        If extra_listbox1.SelectedIndex >= 0 And join_listbox.SelectedIndex >= 0 Then
            blank_listbox1.Items.Clear()

            blank_ship_group1 = main_ship_group.get_suit_hole_ship(extra_group.extra(extra_listbox1.SelectedIndex).get_fleet_group.fleet(fleet_index).hole(join_listbox.SelectedIndex), extra_group.extra(extra_listbox1.SelectedIndex), select_mode_combobox1.SelectedIndex)

            Dim show_str As String
            For a = 0 To blank_ship_group1.Length - 1
                show_str = blank_ship_group1.ship(a).ship_name & "[" & blank_ship_group1.ship(a).Sort_value & "]" & "[" & blank_ship_group1.ship(a).get_type.Name & "]"
                If blank_ship_group1.ship(a).Get_dna.Value("tear") <> "0" Then
                    show_str = show_str & "[已经打孔]"
                Else
                    show_str = show_str & "[未打孔]"
                End If
                show_str = show_str & "[" & blank_ship_group1.ship(a).Get_dna.Value("hole_count") & "个装备格]"
                If blank_ship_group1.ship(a).Length = 1 Then
                    show_str = show_str & "[" & blank_ship_group1.ship(a).get_sticker.Name & "]"
                Else
                    show_str = show_str & "[id-" & Replace(blank_ship_group1.ship(a).Name, "id", "") & "]"
                End If
                blank_listbox1.Items.Add(show_str)
            Next

            Return True
        Else
            Return False
        End If
    End Function

    Public Function ship_join_extra(ByVal extra_listbox As ListBox, ByVal join_listbox1 As ListBox, ByVal join_listbox2 As ListBox, ByVal blank_listbox As ListBox) As Boolean
        Dim return_boolean As Boolean = False
        Dim join_listbox As ListBox
        If last_select_fleet = 0 Then
            join_listbox = join_listbox1
        Else
            join_listbox = join_listbox2
        End If

        If extra_listbox.SelectedIndex > -1 And join_listbox.SelectedIndex > -1 And blank_listbox.SelectedIndex > -1 Then
            blank_ship_group1.ship(blank_listbox.SelectedIndex).set_sticker(extra_group.extra(extra_listbox.SelectedIndex).get_sticker_group.sticker(0))

            If extra_group.extra(extra_listbox.SelectedIndex).get_fleet_group.fleet(last_select_fleet).hole(join_listbox.SelectedIndex).ship_in_hole Then
                main_ship_group.ship(extra_group.extra(extra_listbox.SelectedIndex).get_fleet_group.fleet(last_select_fleet).hole(join_listbox.SelectedIndex).get_ship.Name).remove_sticker()
                extra_group.extra(extra_listbox.SelectedIndex).get_fleet_group.fleet(last_select_fleet).hole(join_listbox.SelectedIndex).remove_ship()
            End If

            extra_group.extra(extra_listbox.SelectedIndex).get_fleet_group.fleet(last_select_fleet).hole(join_listbox.SelectedIndex).set_ship(blank_ship_group1.ship(blank_listbox.SelectedIndex))
            return_boolean = True
        End If
        Return return_boolean
    End Function

    Public Function ship_exit_extra(ByVal extra_listbox As ListBox, ByVal join_listbox1 As ListBox, ByVal join_listbox2 As ListBox) As Boolean
        Dim return_boolean As Boolean = False
        Dim join_listbox As ListBox
        If last_select_fleet = 0 Then
            join_listbox = join_listbox1
        Else
            join_listbox = join_listbox2
        End If
        If extra_listbox.SelectedIndex > -1 And join_listbox.SelectedIndex > -1 Then
            If extra_group.extra(extra_listbox.SelectedIndex).get_fleet_group.fleet(last_select_fleet).hole(join_listbox.SelectedIndex).ship_in_hole Then
                main_ship_group.ship(extra_group.extra(extra_listbox.SelectedIndex).get_fleet_group.fleet(last_select_fleet).hole(join_listbox.SelectedIndex).get_ship.Name).remove_sticker()
                extra_group.extra(extra_listbox.SelectedIndex).get_fleet_group.fleet(last_select_fleet).hole(join_listbox.SelectedIndex).remove_ship()
                return_boolean = True
            End If
        End If
        Return return_boolean
    End Function

End Class