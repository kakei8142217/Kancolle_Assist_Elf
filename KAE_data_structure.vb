Public Class special_equip_class
    Inherits KPL.Data_volume.Larva_class
End Class

Public Class special_equip_group_class
    Inherits KPL.Data_volume.Larva_class

    Public Sub set_special_equip(ByVal special_equip As special_equip_class)
        Set_I_Larva(special_equip)
    End Sub

    Public Overloads ReadOnly Property special_equip(ByVal index As Integer) As special_equip_class
        Get
            special_equip = Get_I_Larva(index)
        End Get
    End Property

    Public Overloads ReadOnly Property special_equip(ByVal name As String) As special_equip_class
        Get
            special_equip = Get_I_Larva(name)
        End Get
    End Property
End Class

Public Class sticker_class
    Inherits KPL.Data_volume.Larva_class
End Class

Public Class sticker_group_class
    Inherits KPL.Data_volume.Larva_class
    Public Sub set_sticker(ByVal sticker As sticker_class)
        Set_I_Larva(sticker)
    End Sub

    Public Overloads ReadOnly Property sticker(ByVal index As Integer) As sticker_class
        Get
            sticker = Get_I_Larva(index)
        End Get
    End Property

    Public Overloads ReadOnly Property sticker(ByVal name As String) As sticker_class
        Get
            sticker = Get_I_Larva(name)
        End Get
    End Property

    Public Overloads Sub remove_sticker(ByVal index As Integer)
        Remove_I_Larva(index)
    End Sub

    Public Overloads Sub remove_sticker(ByVal name As String)
        Remove_I_Larva(name)
    End Sub
End Class

Public Class type_class
    Inherits KPL.Data_volume.Larva_class
End Class

Public Class type_short_class
    Inherits KPL.Data_volume.Larva_class

    Public Sub set_type(ByVal type As type_class)
        Set_I_Larva(type)
    End Sub

    Public Overloads ReadOnly Property type(ByVal index As Integer) As type_class
        Get
            type = Get_I_Larva(index)
        End Get
    End Property

    Public Overloads ReadOnly Property type(ByVal name As String) As type_class
        Get
            type = Get_I_Larva(name)
        End Get
    End Property
End Class

Public Class type_short_group_class
    Inherits KPL.Data_volume.Larva_class

    Public Sub set_type_short(ByVal type_short As type_short_class)
        Set_I_Larva(type_short)
    End Sub

    Public Overloads ReadOnly Property type_short(ByVal index As Integer) As type_short_class
        Get
            type_short = Get_I_Larva(index)
        End Get
    End Property

    Public Overloads ReadOnly Property type_short(ByVal name As String) As type_short_class
        Get
            type_short = Get_I_Larva(name)
        End Get
    End Property
End Class

Public Class ship_class
    Inherits KPL.Data_volume.Larva_class

    Dim type As New type_class
    Dim special_equip As New special_equip_class

    Public Property ship_name As String
        Get
            ship_name = dna.Value("ship_name")
        End Get
        Set(value As String)
            dna.Value("ship_name") = value
        End Set
    End Property

    Public ReadOnly Property get_type As type_class
        Get
            get_type = type
        End Get
    End Property

    Public ReadOnly Property get_special_equip As special_equip_class
        Get
            get_special_equip = special_equip
        End Get
    End Property

    Public Function set_sticker(ByVal sticker As sticker_class) As Boolean
        Dim return_boolean As Boolean = False
        If Length = 0 Then
            Set_I_Larva(sticker)
            return_boolean = True
        End If
        Return return_boolean
    End Function

    Public ReadOnly Property get_sticker As sticker_class
        Get
            get_sticker = Get_I_Larva(0)
        End Get
    End Property

    Public Function remove_sticker() As Boolean
        Dim return_boolean As Boolean = False
        If Length > 0 Then
            Remove_I_Larva(0)
            return_boolean = True
        End If
        Return return_boolean
    End Function

    Public ReadOnly Property suit_hole(ByVal hole As hole_class, ByVal extra As extra_class) As Boolean
        Get
            Dim return_boolean As Boolean = True

            Dim ship_group As ship_group_class = extra.get_hole_ship_group
            For a = 0 To ship_group.Length - 1
                If ship_group.ship(a).Name = Name Then
                    Return False
                    Exit Property
                End If
            Next

            Dim bingo As Boolean = False

            If Length > 0 Then
                For a = 0 To extra.get_sticker_group.Length - 1
                    If get_sticker.Name = extra.get_sticker_group.sticker(a).Name Then
                        bingo = True
                    End If
                Next

                If bingo = False Then
                    Return False
                    Exit Property
                End If
            End If

            Dim type_short_group As type_short_group_class = main_type_short_group.discern_type_short(get_type)
            bingo = False
            For a = 0 To type_short_group.Length - 1
                For b = 0 To hole.get_type_short_group.Length - 1
                    If type_short_group.type_short(a).Name = hole.get_type_short_group.type_short(b).Name Then
                        bingo = True
                        Exit For
                    End If
                Next
                If bingo = True Then
                    Exit For
                End If
            Next

            If bingo = False Then
                Return False
                Exit Property
            End If

            If hole.get_special_equip.Get_dna.Value("name") <> "" Then
                If hole.get_special_equip.Get_dna.Value("name") <> ship_name Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("sort_value") <> "" Then
                If Val(hole.get_special_equip.Get_dna.Value("sort_value")) > Val(dna.Value("sort_value")) Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.Get_dna.Value("tear") <> "0" Then
                If Val(hole.Get_dna.Value("tear")) > Val(dna.Value("tear")) Then
                    Return False
                    Exit Property
                End If
            End If
            If hole.Get_dna.Value("hole_count") <> "0" Then
                If Val(hole.Get_dna.Value("hole_count")) > Val(dna.Value("hole_count")) Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("speed") <> "" Then
                If get_special_equip.Get_dna.Value("speed") <> hole.get_special_equip.Get_dna.Value("speed") Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("上陸用舟艇") = 1 Then
                If get_special_equip.Get_dna.Value("上陸用舟艇") = 0 Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("特型内火艇") = 1 Then
                If get_special_equip.Get_dna.Value("特型内火艇") = 0 Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("追加装甲") = 1 Then
                If get_special_equip.Get_dna.Value("追加装甲") = 0 Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("司令部施設") = 1 Then
                If get_special_equip.Get_dna.Value("司令部施設") = 0 Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("大型電探") = 1 Then
                If get_special_equip.Get_dna.Value("大型電探") = 0 Then
                    Return False
                    Exit Property
                End If
            End If

            If hole.get_special_equip.Get_dna.Value("特殊潜航艇") = 1 Then
                If get_special_equip.Get_dna.Value("特殊潜航艇") = 0 Then
                    Return False
                    Exit Property
                End If
            End If

            Return return_boolean
        End Get
    End Property


End Class

Public Class ship_group_class
    Inherits KPL.Data_volume.Larva_class

    Public Sub set_ship(ByVal input_ship As ship_class)
        Set_I_Larva(input_ship)
    End Sub

    Public Overloads ReadOnly Property ship(ByVal index As Integer) As ship_class
        Get
            ship = Get_I_Larva(index)
        End Get
    End Property

    Public Overloads ReadOnly Property ship(ByVal name As String) As ship_class
        Get
            ship = Get_I_Larva(name)
        End Get
    End Property

    Public Overloads Sub remove_ship(ByVal index As Integer)
        Remove_I_Larva(index)
    End Sub

    Public Overloads Sub remove_ship(ByVal name As String)
        Remove_I_Larva(name)
    End Sub

    Public Function get_suit_hole_ship(ByVal hole As hole_class, ByVal extra As extra_class, Optional ByVal select_mode As Integer = 0) As ship_group_class
        Dim return_group As New ship_group_class
        For a = 0 To Length - 1
            If ship(a).suit_hole(hole, extra) Then
                If select_mode = 0 Then
                    return_group.set_ship(ship(a))
                ElseIf select_mode = 1 Then
                    If Insert_or_not_same_name_ship(ship(a), return_group) Then
                        return_group.set_ship(ship(a))
                    End If
                End If
            End If
        Next
        Return return_group
    End Function

    Private Function Insert_or_not_same_name_ship(ByVal ship As ship_class, ByRef return_group As ship_group_class) As Boolean
        Dim return_boolean As Boolean = False

        Dim bingo As Boolean = False
        Dim remove_id As String = ""
        For a = 0 To return_group.Length - 1
            If return_group.ship(a).ship_name = ship.ship_name Then
                If return_group.ship(a).Sort_value > ship.Sort_value Then
                    remove_id = return_group.ship(a).Name
                    bingo = True
                    Exit For
                ElseIf return_group.ship(a).Sort_value = ship.Sort_value Then
                    If Val(Replace(return_group.ship(a).Name, "id", "")) < Val(Replace(ship.Name, "id", "")) Then
                        remove_id = return_group.ship(a).Name
                        bingo = True
                        Exit For
                    End If
                End If
            End If
        Next

        If remove_id <> "" Then
            return_group.remove_ship(remove_id)
            return_boolean = True
        End If

        If bingo = False Then
            return_boolean = True
        End If

        Return return_boolean
    End Function

End Class

Public Class hole_class
    Inherits KPL.Data_volume.Larva_class

    Dim ship_group As New ship_group_class
    Dim type_short_group As New type_short_group_class
    Dim special_equip As New special_equip_class

    Public Function set_ship(ByVal ship As ship_class) As Boolean
        Dim return_boolean As Boolean = False
        If ship_group.Length = 0 Then
            ship_group.set_ship(ship)
            return_boolean = True
        End If
        Return return_boolean
    End Function

    Public ReadOnly Property get_ship As ship_class
        Get
            Dim return_ship As ship_class

            If ship_group.Length > 0 Then
                return_ship = ship_group.ship(0)
            Else
                return_ship = New ship_class
                return_ship.ship_name = "未选定舰娘"
                return_ship.Sort_value = special_equip.Sort_value
                return_ship.get_type.Name = get_type_show
                Dim sticker As New sticker_class
                sticker.Name = "无贴条"
                return_ship.set_sticker(sticker)
            End If

            Return return_ship
        End Get
    End Property

    Public Function remove_ship() As Boolean
        Dim return_boolean As Boolean = False
        If ship_group.Length > 0 Then
            ship_group.remove_ship(0)
            return_boolean = True
        End If
        Return return_boolean
    End Function

    Public ReadOnly Property ship_in_hole As Boolean
        Get
            Dim return_boolean As Boolean = False
            If ship_group.Length > 0 Then
                return_boolean = True
            End If
            Return return_boolean
        End Get
    End Property

    Public ReadOnly Property get_type_short_group As type_short_group_class
        Get
            get_type_short_group = type_short_group
        End Get
    End Property

    Public ReadOnly Property get_type_show As String
        Get
            Dim return_str As String = ""
            For a = 0 To type_short_group.Length - 1
                return_str = return_str & type_short_group.type_short(a).Name
                If a <> type_short_group.Length - 1 Then
                    return_str = return_str & "/"
                End If
            Next
            Return return_str
        End Get
    End Property

    Public ReadOnly Property get_special_equip As special_equip_class
        Get
            get_special_equip = special_equip
        End Get
    End Property

End Class

Public Class fleet_class
    Inherits KPL.Data_volume.Larva_class

    Public Sub set_hole(ByVal hole As hole_class)
        Set_I_Larva(hole)
    End Sub

    Public ReadOnly Property hole(ByVal index As Integer) As hole_class
        Get
            hole = Get_I_Larva(index)
        End Get
    End Property
End Class

Public Class fleet_group_class
    Inherits KPL.Data_volume.Larva_class

    Public Sub set_fleet(ByVal fleet As fleet_class)
        Set_I_Larva(fleet)
    End Sub

    Public ReadOnly Property fleet(ByVal index As Integer) As fleet_class
        Get
            fleet = Get_I_Larva(index)
        End Get
    End Property
End Class

Public Class extra_class
    Inherits KPL.Data_volume.Larva_class

    Dim sticker_group As New sticker_group_class
    Dim fleet_group As New fleet_group_class

    Public ReadOnly Property get_sticker_group As sticker_group_class
        Get
            get_sticker_group = sticker_group
        End Get
    End Property

    Public ReadOnly Property get_fleet_group As fleet_group_class
        Get
            get_fleet_group = fleet_group
        End Get
    End Property

    Public Function get_hole_ship_group() As ship_group_class
        Dim return_group As New ship_group_class
        For a = 0 To fleet_group.Length - 1
            For b = 0 To fleet_group.fleet(a).Length - 1
                If fleet_group.fleet(a).hole(b).get_ship.Name <> "未选定出击舰娘" Then
                    return_group.set_ship(fleet_group.fleet(a).hole(b).get_ship)
                End If
            Next
        Next
        Return return_group
    End Function

End Class

Public Class extra_group_class
    Inherits KPL.Data_volume.Larva_class

    Public Sub set_extra(ByVal extra As extra_class)
        Set_I_Larva(extra)
    End Sub

    Public Overloads ReadOnly Property extra(ByVal index As Integer) As extra_class
        Get
            extra = Get_I_Larva(index)
        End Get
    End Property

    Public Overloads ReadOnly Property extra(ByVal name As String) As extra_class
        Get
            extra = Get_I_Larva(name)
        End Get
    End Property
End Class

Public Class main_type_short_group_class
    Inherits type_short_group_class

    Public Overloads Function discern_type_short(ByVal ship As ship_class) As type_short_group_class
        Dim return_type_short_group As New type_short_group_class

        For a = 0 To Length - 1
            For b = 0 To Me.type_short(a).Length - 1
                If ship.get_type.Name = Me.type_short(a).type(b).Name Then
                    return_type_short_group.set_type_short(type_short(a))
                    Exit For
                End If
            Next
        Next

        Return return_type_short_group
    End Function

    Public Overloads Function discern_type_short(ByVal type As type_class) As type_short_group_class
        Dim return_type_short_group As New type_short_group_class

        For a = 0 To Length - 1
            For b = 0 To Me.type_short(a).Length - 1
                If type.Name = Me.type_short(a).type(b).Name Then
                    return_type_short_group.set_type_short(type_short(a))
                    Exit For
                End If
            Next
        Next

        Return return_type_short_group
    End Function
End Class



