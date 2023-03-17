
Imports System.Xml.Schema

Public Class file_handle_class
    Public Sub csv_to_ship_xml(ByVal csv_file_path As String, ByVal xml_file_path As String)
        Dim special_equip_group As special_equip_group_class = load_ship_special_equip()

        Dim string_handle As New KPL.String_handle.String_handle_class
        Dim csv_string_array As String() = string_handle.CSV_file_to_array(csv_file_path)
        Dim xml_file As New Xml.XmlDocument
        Dim root_node As Xml.XmlElement = xml_file.CreateElement("kancolle_ship")
        xml_file.AppendChild(root_node)

        Dim mode As String = "74eo"
        Dim csv_field_array As String() = string_handle.CSV_string_to_array(csv_string_array(0))
        If csv_field_array(0) = "舰 ID" Then
            mode = "poi"
        End If
        root_node.SetAttribute("name", mode)

        For a = 1 To csv_string_array.Length - 1
            Dim node As Xml.XmlElement = xml_file.CreateElement("ship")
            csv_field_array = string_handle.CSV_string_to_array(csv_string_array(a))
            If mode = "74eo" Then
                node.SetAttribute("name", "id" & csv_field_array(0))
                node.SetAttribute("ship_name", csv_field_array(2))
                node.SetAttribute("sort_value", csv_field_array(3))

                Dim hole_count As Integer = 0
                For b = 12 To 16
                    If csv_field_array(b) <> "" Then
                        hole_count = hole_count + 1
                    End If
                Next
                node.SetAttribute("hole_count", hole_count)

                If csv_field_array(17) = "" Then
                    node.SetAttribute("tear", 0)
                Else
                    node.SetAttribute("tear", 1)
                End If

                Dim type_node As Xml.XmlElement = xml_file.CreateElement("type")
                type_node.SetAttribute("name", csv_field_array(1))
                node.AppendChild(type_node)
            ElseIf mode = "poi" Then
                node.SetAttribute("name", "id" & csv_field_array(0))
                node.SetAttribute("ship_name", csv_field_array(1))
                node.SetAttribute("sort_value", csv_field_array(7))

                node.SetAttribute("hole_count", 99)
                node.SetAttribute("tear", 1)

                Dim type_node As Xml.XmlElement = xml_file.CreateElement("type")
                type_node.SetAttribute("name", csv_field_array(5))
                node.AppendChild(type_node)
            End If

            Dim special_equip_node As Xml.XmlElement = xml_file.CreateElement("special_equip")
            Dim special_equip As special_equip_class = special_equip_group.special_equip(node.Attributes("ship_name").Value)
            special_equip.Get_dna.export_dna_to_node(special_equip_node)
            node.AppendChild(special_equip_node)

            root_node.AppendChild(node)
        Next

        xml_file.Save(xml_file_path)
    End Sub

    Public Sub csv_to_special_equip_xml(ByVal csv_file_path As String, ByVal xml_file_path As String)
        Dim string_handle As New KPL.String_handle.String_handle_class
        Dim csv_string_array As String() = string_handle.CSV_file_to_array(csv_file_path)
        Dim xml_file As New Xml.XmlDocument
        Dim root_node As Xml.XmlElement = xml_file.CreateElement("kancolle_special_equip")
        xml_file.AppendChild(root_node)
        For a = 1 To csv_string_array.Length - 1
            Dim node As Xml.XmlElement = xml_file.CreateElement("ship")
            Dim csv_field_array As String() = string_handle.CSV_string_to_array(csv_string_array(a))

            node.SetAttribute("name", csv_field_array(4))
            node.SetAttribute("speed", csv_field_array(36))
            node.SetAttribute("上陸用舟艇", "0")
            node.SetAttribute("特型内火艇", "0")
            node.SetAttribute("追加装甲", "0")
            node.SetAttribute("司令部施設", "0")
            node.SetAttribute("大型電探", "0")
            node.SetAttribute("特殊潜航艇", "0")

            root_node.AppendChild(node)
        Next
        xml_file.Save(xml_file_path)
        MsgBox("转换完成")
    End Sub

    Private Function load_ship_special_equip() As special_equip_group_class
        Dim return_group As New special_equip_group_class

        Dim xml_file As New Xml.XmlDocument
        xml_file.Load(Application.StartupPath & "\system_data\ship_special_equip.xml")

        For Each node As Xml.XmlElement In xml_file.DocumentElement.ChildNodes
            Dim special_equip As New special_equip_class
            special_equip.Get_dna.load_node_to_dna(node)
            return_group.set_special_equip(special_equip)
        Next

        Return return_group
    End Function

    Public Sub load_ship_to_main_ship_group(ByVal xml_file_path As String)
        Dim xml_ship As New Xml.XmlDocument
        xml_ship.Load(xml_file_path)
        main_ship_group.Get_dna.load_node_to_dna(xml_ship.DocumentElement)

        Dim ship As ship_class
        For Each node As Xml.XmlElement In xml_ship.DocumentElement.ChildNodes
            ship = New ship_class
            ship.Get_dna.load_node_to_dna(node)


            For Each child_node As Xml.XmlElement In node.ChildNodes
                If child_node.Name = "type" Then
                    ship.get_type.Get_dna.load_node_to_dna(child_node)
                ElseIf child_node.Name = "special_equip" Then
                    ship.get_special_equip.Get_dna.load_node_to_dna(child_node)
                ElseIf child_node.Name = "sticker" Then
                    Dim sticker As New sticker_class
                    sticker.Get_dna.load_node_to_dna(child_node)
                    ship.set_sticker(sticker)
                End If
            Next

            main_ship_group.set_ship(ship)
        Next

        main_ship_group.sort(False)
    End Sub

    Public Sub export_main_ship_group_to_xml(ByVal xml_file_path As String)
        Dim xml_ship As New Xml.XmlDocument
        Dim root_node As Xml.XmlElement = xml_ship.CreateElement("kancolle_ship")
        main_ship_group.Get_dna.export_dna_to_node(root_node)
        xml_ship.AppendChild(root_node)

        For a = 0 To main_ship_group.Length - 1
            Dim ship_node As Xml.XmlElement = xml_ship.CreateElement("ship")
            root_node.AppendChild(ship_node)
            main_ship_group.ship(a).Get_dna.export_dna_to_node(ship_node)
            Dim type_node As Xml.XmlElement = xml_ship.CreateElement("type")
            main_ship_group.ship(a).get_type.Get_dna.export_dna_to_node(type_node)
            ship_node.AppendChild(type_node)
            Dim special_equip_node As Xml.XmlElement = xml_ship.CreateElement("special_equip")
            main_ship_group.ship(a).get_special_equip.Get_dna.export_dna_to_node(special_equip_node)
            ship_node.AppendChild(special_equip_node)
            If main_ship_group.ship(a).Length > 0 Then
                Dim sticker_node As Xml.XmlElement = xml_ship.CreateElement("sticker")
                main_ship_group.ship(a).get_sticker.Get_dna.export_dna_to_node(sticker_node)
                ship_node.AppendChild(sticker_node)
            End If
        Next

        xml_ship.Save(xml_file_path)
    End Sub

    Public Sub load_type_to_type_short_group(ByVal xml_file_path)
        Dim xml_type As New Xml.XmlDocument
        xml_type.Load(xml_file_path)

        Dim type_short As type_short_class
        For Each node As Xml.XmlElement In xml_type.DocumentElement.ChildNodes
            type_short = New type_short_class
            type_short.Get_dna.load_node_to_dna(node)
            main_type_short_group.set_type_short(type_short)

            For Each child_node As Xml.XmlElement In node.ChildNodes
                Dim type As New type_class
                type.Get_dna.load_node_to_dna(child_node)
                type_short.set_type(type)
            Next
        Next
    End Sub

    Public Sub load_extra_to_extra_group(ByVal xml_file_path)
        Dim xml_extra As New Xml.XmlDocument
        xml_extra.Load(xml_file_path)

        For Each node As Xml.XmlElement In xml_extra.DocumentElement.ChildNodes
            Dim extra As New extra_class
            extra.Get_dna.load_node_to_dna(node)

            For Each child_node As Xml.XmlElement In node.ChildNodes
                If child_node.Name = "sticker_group" Then
                    For Each sticker_node As Xml.XmlElement In child_node.ChildNodes
                        Dim sticker As New sticker_class
                        sticker.Get_dna.load_node_to_dna(sticker_node)
                        extra.get_sticker_group.set_sticker(sticker)
                    Next
                ElseIf child_node.Name = "fleet_group" Then
                    For Each fleet_node As Xml.XmlElement In child_node.ChildNodes
                        Dim fleet As New fleet_class
                        fleet.Get_dna.load_node_to_dna(fleet_node)

                        For Each hole_node As Xml.XmlElement In fleet_node.ChildNodes
                            Dim hole As New hole_class
                            hole.Get_dna.load_node_to_dna(hole_node)

                            For Each hole_child_node As Xml.XmlElement In hole_node.ChildNodes
                                If hole_child_node.Name = "type_short_group" Then
                                    For Each type_short_node As Xml.XmlElement In hole_child_node.ChildNodes
                                        Dim type_short As New type_short_class
                                        type_short.Get_dna.load_node_to_dna(type_short_node)
                                        hole.get_type_short_group.set_type_short(type_short)
                                    Next
                                ElseIf hole_child_node.Name = "special_equip" Then
                                    hole.get_special_equip.Get_dna.load_node_to_dna(hole_child_node)
                                ElseIf hole_child_node.Name = "ship" Then
                                    Dim ship As New ship_class
                                    ship.Get_dna.load_node_to_dna(hole_child_node)
                                    For Each ship_child_node As Xml.XmlElement In hole_child_node.ChildNodes
                                        If ship_child_node.Name = "type" Then
                                            ship.get_type.Get_dna.load_node_to_dna(ship_child_node)
                                        ElseIf ship_child_node.Name = "special_equip" Then
                                            ship.get_special_equip.Get_dna.load_node_to_dna(ship_child_node)
                                        ElseIf ship_child_node.Name = "sticker" Then
                                            Dim sticker As New sticker_class
                                            sticker.Get_dna.load_node_to_dna(ship_child_node)
                                            ship.set_sticker(sticker)
                                        End If
                                    Next
                                    hole.set_ship(ship)
                                End If
                            Next

                            fleet.set_hole(hole)
                        Next

                        extra.get_fleet_group.set_fleet(fleet)
                    Next
                End If
            Next

            extra_group.set_extra(extra)
        Next
    End Sub

    Public Sub export_extra_group_to_xml(ByVal xml_file_path As String)
        Dim xml_extra As New Xml.XmlDocument
        Dim root_node As Xml.XmlElement = xml_extra.CreateElement("kancolle_extra")
        xml_extra.AppendChild(root_node)

        For a = 0 To extra_group.Length - 1
            Dim extra_node As Xml.XmlElement = xml_extra.CreateElement("extra")
            extra_group.extra(a).Get_dna.export_dna_to_node(extra_node)
            root_node.AppendChild(extra_node)

            Dim sticker_group_node As Xml.XmlElement = xml_extra.CreateElement("sticker_group")
            extra_node.AppendChild(sticker_group_node)
            For b = 0 To extra_group.extra(a).get_sticker_group.Length - 1
                Dim sticker_node As Xml.XmlElement = xml_extra.CreateElement("sticker")
                extra_group.extra(a).get_sticker_group.sticker(b).Get_dna.export_dna_to_node(sticker_node)
                sticker_group_node.AppendChild(sticker_node)
            Next

            Dim fleet_group_node As Xml.XmlElement = xml_extra.CreateElement("fleet_group")
            extra_node.AppendChild(fleet_group_node)
            For b = 0 To extra_group.extra(a).get_fleet_group.Length - 1
                Dim fleet_node As Xml.XmlElement = xml_extra.CreateElement("fleet")
                extra_group.extra(a).get_fleet_group.fleet(b).Get_dna.export_dna_to_node(fleet_node)
                fleet_group_node.AppendChild(fleet_node)

                For c = 0 To extra_group.extra(a).get_fleet_group.fleet(b).Length - 1
                    Dim hole_node As Xml.XmlElement = xml_extra.CreateElement("hole")
                    extra_group.extra(a).get_fleet_group.fleet(b).hole(c).Get_dna.export_dna_to_node(hole_node)
                    fleet_node.AppendChild(hole_node)
                    Dim type_short_group_node As Xml.XmlElement = xml_extra.CreateElement("type_short_group")
                    hole_node.AppendChild(type_short_group_node)

                    For d = 0 To extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_type_short_group.Length - 1
                        Dim type_short_node As Xml.XmlElement = xml_extra.CreateElement("type_short")
                        extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_type_short_group.type_short(d).Get_dna.export_dna_to_node(type_short_node)
                        type_short_group_node.AppendChild(type_short_node)
                    Next

                    Dim special_equip_node As Xml.XmlElement = xml_extra.CreateElement("special_equip")
                    extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_special_equip.Get_dna.export_dna_to_node(special_equip_node)
                    hole_node.AppendChild(special_equip_node)

                    If extra_group.extra(a).get_fleet_group.fleet(b).hole(c).ship_in_hole Then
                        Dim ship_node As Xml.XmlElement = xml_extra.CreateElement("ship")
                        extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_ship.Get_dna.export_dna_to_node(ship_node)
                        hole_node.AppendChild(ship_node)
                        Dim type_node As Xml.XmlElement = xml_extra.CreateElement("type")
                        extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_ship.get_type.Get_dna.export_dna_to_node(type_node)
                        ship_node.AppendChild(type_node)
                        Dim ship_special_equip_node As Xml.XmlElement = xml_extra.CreateElement("special_equip")
                        extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_ship.get_special_equip.Get_dna.export_dna_to_node(ship_special_equip_node)
                        ship_node.AppendChild(ship_special_equip_node)
                        If extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_ship.Length > 0 Then
                            Dim sticker_node As Xml.XmlElement = xml_extra.CreateElement("sticker")
                            extra_group.extra(a).get_fleet_group.fleet(b).hole(c).get_ship.get_sticker.Get_dna.export_dna_to_node(sticker_node)
                            ship_node.AppendChild(sticker_node)
                        End If
                    End If
                Next
            Next
        Next
        xml_extra.Save(xml_file_path)
    End Sub
End Class
