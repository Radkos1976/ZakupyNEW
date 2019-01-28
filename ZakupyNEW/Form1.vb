Imports Npgsql
Imports System.Windows.Forms.DataVisualization.Charting
Imports DgvFilterPopup
Public Class Form1
    Dim V As DgvFilterManager
    Dim m As DGVColumnSelector.DataGridViewColumnSelector
    Dim g_focus As String
    Dim wrk As Boolean = False
    Private Sub COMB()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("Select KOOR from demands where KOOR<>'LUCPRZ' group by KOOR", conB)
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            For Each r In kors.Rows
                                ComboBox2.Items.Add(r("KOOR"))
                                ComboBox3.Items.Add(r("KOOR"))
                            Next
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " refr_graf ex")
        End Try
    End Sub
    Private Sub Refr_graf()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select b.koor,sum(b.brak)/a.al as bilans from (select sum(brak) AL from (select part_no,min(bal_stock) brak from demands a where work_day<=current_date and koor!='*' And koor!='LUCPRZ' group by part_no) a where brak<0) a,(select part_no,min(bal_stock) brak,koor from demands a where work_day<=current_date group by part_no,koor) b where b.brak<0 and b.koor!='*' And b.koor!='LUCPRZ' group by b.koor,a.al", conB)
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            For Each r In kors.Rows
                                If CInt(r("bilans") * 100) = 0 Then
                                    r.Delete()
                                Else
                                    r.BeginEdit()
                                    r("koor") = r("Koor") & " " & CInt(r("bilans") * 100).ToString & "%"
                                End If
                            Next
                            Chart1.DataSource = kors
                            Dim CArea As ChartArea = Chart1.ChartAreas(0)
                            CArea.BackColor = Color.Azure           '~~> Changing the Back Color of the Chart Area 
                            CArea.ShadowColor = Color.Red           '~~> Changing the Shadow Color of the Chart Area 
                            CArea.Area3DStyle.Enable3D = True
                            CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                            CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                            CArea.AxisY.LabelStyle.Format = "0%" '~~> Formatting Y axis to display values in %age
                            CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount

                            Dim Series1 As Series = Chart1.Series(0)
                            '~~> Setting the series Name
                            Series1.Name = "BRAKI"
                            'Chart1.Titles(0).Text = i & "%"
                            '~~> Assigning values to X and Y Axis
                            Chart1.Series(Series1.Name).XValueMember = "KOOR"
                            Chart1.Series(Series1.Name).YValueMembers = "Bilans"
                            '~~> Setting Font, Font Size and Bold
                            '~~> Setting Value Type
                            Chart1.Series(Series1.Name).XValueType = ChartValueType.Date
                            Chart1.Series(Series1.Name).YValueType = ChartValueType.Double
                            '~~> Setting the Chart Type for Display 
                            'Chart1.Series(Series1.Name).ChartType = SeriesChartType.Column
                            '~~> Display Data Labels
                            'Chart1.Series(Series1.Name).IsValueShownAsLabel = True
                            Chart1.Series(Series1.Name).SmartLabelStyle.Enabled = True
                            '~~> Setting label's Fore Color
                            Chart1.Series(Series1.Name).LabelForeColor = Color.Black
                            '~~> Setting label's Format to %age
                            'Chart1.Series(Series1.Name).LabelFormat = "0%"
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " refr_graf ex")
        End Try
        Refr_grafs()
    End Sub
    Private Sub Refr_grafs()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select b.type rodzaj,1+sum(b.brak)/a.al as bilans,sum(b.brak)*-1  BRAK, a.al SUMALL from (select type,sum(qty_demand) AL from demands a where work_day<=current_date and koor!='*' And koor!='LUCPRZ' group by type) a,(select part_no,min(bal_stock) brak,koor,type from demands a where work_day<=current_date group by part_no,koor,type) b where a.type=b.type and b.brak<0 and b.koor!='*' And b.koor!='LUCPRZ' group by b.type,a.al order by b.type desc ", conB)
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            Dim BR As Double
                            Dim ALL As Double
                            For Each r In kors.Rows
                                BR = BR + r("bRAK")
                                ALL = ALL + r("SUMALL")
                            Next
                            Label3.Text = "Procent dostaw " & Math.Round((1 - (BR / ALL)) * 100, 2) & "%"
                            Chart3.DataSource = kors
                            Dim CArea As ChartArea = Chart3.ChartAreas(0)
                            CArea.BackColor = Color.Azure           '~~> Changing the Back Color of the Chart Area 
                            CArea.ShadowColor = Color.Red         '~~> Changing the Shadow Color of the Chart Area 
                            CArea.Area3DStyle.Enable3D = True
                            CArea.AxisX.MajorGrid.Enabled = True   '~~> Removed the X axis major grids
                            CArea.AxisY.MajorGrid.Enabled = True  '~~> Removed the Y axis major grids
                            CArea.AxisY.LabelStyle.Format = "0.00%"
                            CArea.AxisY.Minimum = 0.7
                            CArea.AxisY.Maximum = 1
                            CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                            CArea.AxisY.IsLabelAutoFit = True
                            '~~> Formatting Y axis to display values in %age
                            CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                            Dim Series1 As Series = Chart3.Series(0)
                            '~~> Setting the series Name
                            Series1.Name = "BRAKI"
                            '~~> Assigning values to X and Y Axis
                            Chart3.Series(Series1.Name).XValueMember = "Rodzaj"
                            Chart3.Series(Series1.Name).YValueMembers = "Bilans"
                            '~~> Setting Font, Font Size and Bold
                            '~~> Setting Value Type
                            Chart3.Series(Series1.Name).XValueType = ChartValueType.String
                            Chart3.Series(Series1.Name).YValueType = ChartValueType.Double
                            '~~> Setting the Chart Type for Display 
                            Chart3.Series(Series1.Name).ChartType = SeriesChartType.Bar
                            '~~> Display Data Labels
                            Chart3.Series(Series1.Name).IsValueShownAsLabel = True
                            'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                            '~~> Setting label's Fore Color
                            Chart3.Series(Series1.Name).LabelForeColor = Color.DarkRed
                            '~~> Setting label's Format to %age
                            Chart3.Series(Series1.Name).LabelFormat = "0.00%"
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " refr_grafs ex")
        End Try
        Chrt2()
    End Sub
    Private Sub Chrt2()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("Select work_day,1+(Sum(brak)/sum(sumal)) as bilans from aktual_hist where sumal>0 and work_day<=@work_day and case when @type!='Wszystkie' then type=@type else type is not null end and case when @koor!='WSZYSCY' then koor=@koor else koor is not null end group by work_day order by work_day", conB)
                        con.Parameters.Add("work_day", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTimePicker1.Value.Date
                        con.Parameters.Add("type", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox1.SelectedItem.ToString
                        con.Parameters.Add("koor", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox2.SelectedItem.ToString
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            If kors.Rows.Count > 0 Then
                                Chart2.Visible = True
                                Chart2.DataSource = kors
                                Dim CArea As ChartArea = Chart2.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age
                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart2.Series(0)

                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart2.Series(Series1.Name).XValueMember = "Work_Day"
                                Chart2.Series(Series1.Name).YValueMembers = "Bilans"

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart2.Series(Series1.Name).XValueType = ChartValueType.Date
                                Chart2.Series(Series1.Name).YValueType = ChartValueType.Double
                                '~~> Setting the Chart Type for Display 
                                Chart2.Series(Series1.Name).ChartType = SeriesChartType.Column
                                '~~> Display Data Labels
                                Chart2.Series(Series1.Name).IsValueShownAsLabel = True
                                Chart2.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart2.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart2.Series(Series1.Name).LabelFormat = "0.0%"
                            Else
                                Chart2.Visible = False
                                MsgBox("Brak danych dla wprowadzonych kryteriów")
                            End If
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
            Chart2.Visible = False
            MsgBox("Brak danych dla wprowadzonych kryteriów")
        End Try
        Chrt3()
    End Sub
    Private Sub Chrt3()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("Select work_day,1+(Sum(brak_mag)/sum(sumal)) as bilans from aktual_hist where sumal>0 and work_day<=@work_day and case when @type!='Wszystkie' then type=@type else type is not null end and case when @koor!='WSZYSCY' then koor=@koor else koor is not null end group by work_day order by work_day", conB)
                        con.Parameters.Add("work_day", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTimePicker2.Value.Date
                        con.Parameters.Add("type", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox4.SelectedItem.ToString
                        con.Parameters.Add("koor", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox3.SelectedItem.ToString
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            con.Prepare()
                            kors.Load(rs)
                            If kors.Rows.Count > 0 Then
                                Chart4.Visible = True
                                Chart4.DataSource = kors
                                Dim CArea As ChartArea = Chart4.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age
                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart4.Series(0)
                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart4.Series(Series1.Name).XValueMember = "Work_Day"
                                Chart4.Series(Series1.Name).YValueMembers = "Bilans"

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart4.Series(Series1.Name).XValueType = ChartValueType.Date
                                Chart4.Series(Series1.Name).YValueType = ChartValueType.Double
                                '~~> Setting the Chart Type for Display 
                                Chart4.Series(Series1.Name).ChartType = SeriesChartType.Column
                                '~~> Display Data Labels
                                Chart4.Series(Series1.Name).IsValueShownAsLabel = True
                                Chart4.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart4.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart4.Series(Series1.Name).LabelFormat = "0.0%"
                            Else
                                Chart4.Visible = False
                                MsgBox("Brak danych dla wprowadzonych kryteriów")
                            End If
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
            Chart4.Visible = False
            MsgBox("Brak danych dla wprowadzonych kryteriów")
        End Try
        Chrt4()
    End Sub
    Private Sub Chrt4()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select work_day,typ,sum(qty_all),sum(brak),1-(sum(brak)/sum(qty_all)) as Gotowość from day_qty where typ='MRP' group by work_day,typ order by work_day,typ", conB)
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            If kors.Rows.Count > 0 Then
                                Chart5.DataSource = kors
                                Dim CArea As ChartArea = Chart5.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age
                                CArea.AxisY.Minimum = 0.3
                                CArea.AxisY.Maximum = 1
                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart5.Series(0)
                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart5.Series(Series1.Name).XValueMember = "Work_Day"
                                Chart5.Series(Series1.Name).YValueMembers = "Gotowość"

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart5.Series(Series1.Name).XValueType = ChartValueType.Date
                                Chart5.Series(Series1.Name).YValueType = ChartValueType.Double
                                '~~> Setting the Chart Type for Display 
                                Chart5.Series(Series1.Name).ChartType = SeriesChartType.Bar
                                '~~> Display Data Labels
                                Chart5.Series(Series1.Name).IsValueShownAsLabel = True
                                'Chart5.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart5.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart5.Series(Series1.Name).LabelFormat = "0.0%"
                            Else

                                MsgBox("Brak danych dla wprowadzonych kryteriów")

                            End If
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
        End Try
        Chrt5()
    End Sub
    Private Sub Chrt5()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select work_day,typ,sum(qty_all),sum(brak),1-(sum(brak)/sum(qty_all)) as Gotowość from day_qty where typ='DOP' group by work_day,typ order by work_day,typ", conB)
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            If kors.Rows.Count > 0 Then
                                Chart6.DataSource = kors
                                Dim CArea As ChartArea = Chart6.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age
                                CArea.AxisY.Minimum = 0.3
                                CArea.AxisY.Maximum = 1
                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart6.Series(0)
                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart6.Series(Series1.Name).XValueMember = "Work_Day"
                                Chart6.Series(Series1.Name).YValueMembers = "Gotowość"

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart6.Series(Series1.Name).XValueType = ChartValueType.Date
                                Chart6.Series(Series1.Name).YValueType = ChartValueType.Double
                                '~~> Setting the Chart Type for Display 
                                Chart6.Series(Series1.Name).ChartType = SeriesChartType.Bar
                                '~~> Display Data Labels
                                Chart6.Series(Series1.Name).IsValueShownAsLabel = True
                                'Chart6.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart6.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart6.Series(Series1.Name).LabelFormat = "0.0%"
                            Else

                                MsgBox("Brak danych dla wprowadzonych kryteriów")
                            End If
                        End Using
                    End Using
                End Using
            End Using

        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
        End Try
        Chrt6()
    End Sub
    Private Sub Chrt6()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select koor,typ,sum(prod_qty ) as Bilans from (select get_koor(part_no) koor,int_ord,case when dop=0 then 'MRP' else 'DOP' end typ,prod_qty from ord_lack where case when @date_required=current_date then date_required<=current_date else date_required=@date_required end group by get_koor(part_no),int_ord,case when ord_date<current_date then current_date else ord_date end,case when dop=0 then 'MRP' else 'DOP' end,prod_qty) a where typ='MRP' group by koor,typ", conB)
                        con.Parameters.Add("date_required", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTimePicker4.Value
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            Dim r As Data.DataRow
                            For Each r In kors.Rows
                                If r("bilans") = 0 Then
                                    r.Delete()
                                Else
                                    r.BeginEdit()
                                    r("koor") = r("Koor") & " " & r("bilans").ToString
                                End If
                            Next
                            Chart7.DataSource = kors
                            Dim CArea As ChartArea = Chart7.ChartAreas(0)
                            CArea.BackColor = Color.Azure           '~~> Changing the Back Color of the Chart Area 
                            CArea.ShadowColor = Color.Red           '~~> Changing the Shadow Color of the Chart Area 
                            CArea.Area3DStyle.Enable3D = True
                            CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                            CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                            CArea.AxisY.LabelStyle.Format = "0" '~~> Formatting Y axis to display values in %age
                            CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount

                            Dim Series1 As Series = Chart7.Series(0)
                            '~~> Setting the series Name
                            Series1.Name = "BRAKI"
                            'Chart1.Titles(0).Text = i & "%"
                            '~~> Assigning values to X and Y Axis
                            Chart7.Series(Series1.Name).XValueMember = "KOOR"
                            Chart7.Series(Series1.Name).YValueMembers = "Bilans"
                            '~~> Setting Font, Font Size and Bold
                            '~~> Setting Value Type
                            Chart7.Series(Series1.Name).XValueType = ChartValueType.Date
                            Chart7.Series(Series1.Name).YValueType = ChartValueType.Double
                            '~~> Setting the Chart Type for Display 
                            'Chart1.Series(Series1.Name).ChartType = SeriesChartType.Column
                            '~~> Display Data Labels
                            'Chart1.Series(Series1.Name).IsValueShownAsLabel = True
                            Chart7.Series(Series1.Name).SmartLabelStyle.Enabled = True
                            '~~> Setting label's Fore Color
                            Chart7.Series(Series1.Name).LabelForeColor = Color.Black
                            '~~> Setting label's Format to %age
                            'Chart1.Series(Series1.Name).LabelFormat = "0%"
                        End Using
                    End Using
                End Using
            End Using

        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " refr_graf ex")
        End Try
        Chrt7()
    End Sub
    Private Sub Chrt7()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select koor,typ,sum(prod_qty ) as Bilans from (select get_koor(part_no) koor,int_ord,case when dop=0 then 'MRP' else 'DOP' end typ,prod_qty from ord_lack where case when @date_required=current_date then date_required<=current_date else date_required=@date_required end group by get_koor(part_no),int_ord,case when ord_date<current_date then current_date else ord_date end,case when dop=0 then 'MRP' else 'DOP' end,prod_qty) a where typ='DOP' group by koor,typ ", conB)
                        con.Parameters.Add("date_required", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTimePicker6.Value
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            Dim r As Data.DataRow
                            For Each r In kors.Rows
                                If r("bilans") = 0 Then
                                    r.Delete()
                                Else
                                    r.BeginEdit()
                                    r("koor") = r("Koor") & " " & r("bilans").ToString
                                End If
                            Next
                            Chart8.DataSource = kors
                            Dim CArea As ChartArea = Chart8.ChartAreas(0)
                            CArea.BackColor = Color.Azure           '~~> Changing the Back Color of the Chart Area 
                            CArea.ShadowColor = Color.Red           '~~> Changing the Shadow Color of the Chart Area 
                            CArea.Area3DStyle.Enable3D = True
                            CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                            CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                            CArea.AxisY.LabelStyle.Format = "0" '~~> Formatting Y axis to display values in %age
                            CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount

                            Dim Series1 As Series = Chart8.Series(0)
                            '~~> Setting the series Name
                            Series1.Name = "BRAKI"
                            'Chart1.Titles(0).Text = i & "%"
                            '~~> Assigning values to X and Y Axis
                            Chart8.Series(Series1.Name).XValueMember = "KOOR"
                            Chart8.Series(Series1.Name).YValueMembers = "Bilans"
                            '~~> Setting Font, Font Size and Bold
                            '~~> Setting Value Type
                            Chart8.Series(Series1.Name).XValueType = ChartValueType.Date
                            Chart8.Series(Series1.Name).YValueType = ChartValueType.Double
                            '~~> Setting the Chart Type for Display 
                            'Chart1.Series(Series1.Name).ChartType = SeriesChartType.Column
                            '~~> Display Data Labels
                            'Chart1.Series(Series1.Name).IsValueShownAsLabel = True
                            Chart8.Series(Series1.Name).SmartLabelStyle.Enabled = True
                            '~~> Setting label's Fore Color
                            Chart8.Series(Series1.Name).LabelForeColor = Color.Black
                            '~~> Setting label's Format to %age
                            'Chart1.Series(Series1.Name).LabelFormat = "0%"
                        End Using
                    End Using
                End Using
            End Using
            Chrt9()
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " refr_graf ex")
        End Try
    End Sub
    Private Sub Chrt9()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select wrkc,1-(sum(brak)/sum(qty_all)) bilans from braki_gniazd where work_day=@work_day and substring(wrkc,1,1)='5' group by wrkc", conB)
                        con.Parameters.Add("work_day", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTimePicker3.Value
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            If kors.Rows.Count > 0 Then
                                Chart9.DataSource = kors
                                Dim CArea As ChartArea = Chart9.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age
                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart9.Series(0)
                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart9.Series(Series1.Name).XValueMember = "wrkc"
                                Chart9.Series(Series1.Name).YValueMembers = "bilans"

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart9.Series(Series1.Name).XValueType = ChartValueType.String
                                Chart9.Series(Series1.Name).YValueType = ChartValueType.Double
                                '~~> Setting the Chart Type for Display 
                                Chart9.Series(Series1.Name).ChartType = SeriesChartType.Bar
                                '~~> Display Data Labels
                                Chart9.Series(Series1.Name).IsValueShownAsLabel = True
                                'Chart5.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart9.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart9.Series(Series1.Name).LabelFormat = "0.0%"
                            Else

                                MsgBox("Brak danych dla wprowadzonych kryteriów")

                            End If
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " refr_grafs ex")
        End Try
        Chrt10()
    End Sub
    Private Sub Chrt10()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select wrkc,1-(sum(brak)/sum(qty_all)) bilans from braki_gniazd where work_day=@work_day and substring(wrkc,1,1) in ('4','1','2') group by wrkc", conB)
                        con.Parameters.Add("work_day", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTimePicker5.Value
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            If kors.Rows.Count > 0 Then
                                Chart10.DataSource = kors
                                Dim CArea As ChartArea = Chart10.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age

                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart10.Series(0)
                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart10.Series(Series1.Name).XValueMember = "wrkc"
                                Chart10.Series(Series1.Name).YValueMembers = "bilans"

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart10.Series(Series1.Name).XValueType = ChartValueType.Date
                                Chart10.Series(Series1.Name).YValueType = ChartValueType.Double
                                '~~> Setting the Chart Type for Display 
                                Chart10.Series(Series1.Name).ChartType = SeriesChartType.Bar
                                '~~> Display Data Labels
                                Chart10.Series(Series1.Name).IsValueShownAsLabel = True
                                'Chart6.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart10.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart10.Series(Series1.Name).LabelFormat = "0.0%"
                            Else

                                MsgBox("Brak danych dla wprowadzonych kryteriów")
                            End If
                        End Using
                    End Using
                End Using
            End Using

        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
        End Try
        Chrt11()
    End Sub
    Private Sub Chrt11()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select wrkc,1-(sum(brak)/sum(qty_all)) bilans from braki_gniazd where work_day=@work_day and substring(wrkc,1,1) not in ('5','4','1','2') group by wrkc", conB)
                        con.Parameters.Add("work_day", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTimePicker7.Value
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            If kors.Rows.Count > 0 Then
                                Chart11.DataSource = kors
                                Dim CArea As ChartArea = Chart11.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age

                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart11.Series(0)
                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart11.Series(Series1.Name).XValueMember = "wrkc"
                                Chart11.Series(Series1.Name).YValueMembers = "bilans"

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart11.Series(Series1.Name).XValueType = ChartValueType.Date
                                Chart11.Series(Series1.Name).YValueType = ChartValueType.Double
                                '~~> Setting the Chart Type for Display 
                                Chart11.Series(Series1.Name).ChartType = SeriesChartType.Bar
                                '~~> Display Data Labels
                                Chart11.Series(Series1.Name).IsValueShownAsLabel = True
                                'Chart6.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart11.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart11.Series(Series1.Name).LabelFormat = "0.0%"
                            Else

                                MsgBox("Brak danych dla wprowadzonych kryteriów")
                            End If
                        End Using
                    End Using
                End Using
            End Using

        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
        End Try
        Chrt12()
    End Sub
    Private Sub Chrt12()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select planner_buyer,a.in_plus/b.in_plus in_plus,a.in_minus/b.in_minus*-1 in_minus from (select planner_buyer,coalesce(sum(case when wartość>0 then wartość end ),0) in_plus,coalesce(sum(case when wartość<0 then wartość end ),0) in_minus from bilans_val group by planner_buyer) a,(select coalesce(sum(case when wartość>0 then wartość end ),0) in_plus,coalesce(sum(case when wartość<0 then wartość end ),0) in_minus from bilans_val ) b where a.in_plus/b.in_plus>0.0001 and a.in_minus/b.in_minus>0.0001", conB)
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            Dim seria As String
                            If ComboBox5.Text = "In Plus - nadwyżka" Then seria = "in_plus" Else seria = "in_minus"

                            If kors.Rows.Count > 0 Then
                                Chart12.DataSource = kors
                                Dim CArea As ChartArea = Chart12.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.0%" '~~> Formatting Y axis to display values in %age
                                'CArea.AxisY.Minimum = 0.3
                                'CArea.AxisY.Maximum = 1
                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart12.Series(0)

                                '~~> Setting the series Name
                                Series1.Name = "BRAKI"
                                '~~> Assigning values to X and Y Axis
                                Chart12.Series(Series1.Name).XValueMember = "planner_buyer"
                                Chart12.Series(Series1.Name).YValueMembers = seria

                                'Chart2.Annotations(ann.Name).
                                '~~> Setting Font, Font Size and Bold
                                '~~> Setting Value Type
                                Chart12.Series(Series1.Name).XValueType = ChartValueType.String
                                Chart12.Series(Series1.Name).YValueType = ChartValueType.Double

                                '~~> Setting the Chart Type for Display 
                                'Chart12.Series(Series1.Name).ChartType = SeriesChartType.Bar
                                '~~> Display Data Labels
                                Chart12.Series(Series1.Name).IsValueShownAsLabel = True

                                'Chart6.Series(Series1.Name).LabelAngle = 90
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart12.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart12.Series(Series1.Name).LabelFormat = "0.0%"
                            Else

                                MsgBox("Brak danych dla wprowadzonych kryteriów")
                            End If
                        End Using
                    End Using
                End Using
            End Using

        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
        End Try

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        wrk = True
        V = New DgvFilterManager(DataGridView1)
        m = New DGVColumnSelector.DataGridViewColumnSelector(DataGridView1)
        Dim day As Date
        Using conB As New NpgsqlConnection(NpA)
            conB.Open()
            Using con As New NpgsqlCommand("select date_fromnow(9)", conB)
                day = con.ExecuteScalar()
            End Using
        End Using
        With Me
            DateTimePicker1.MinDate = Now.Date
            DateTimePicker1.MaxDate = Now.AddDays(30).Date
            DateTimePicker1.Value = Now.AddDays(14).Date
            DateTimePicker2.MinDate = Now.Date
            DateTimePicker2.MaxDate = Now.AddDays(30).Date
            DateTimePicker2.Value = Now.AddDays(10).Date
            DateTimePicker3.MinDate = Now.Date
            DateTimePicker3.MaxDate = day
            DateTimePicker3.Value = Now.Date
            DateTimePicker4.MinDate = Now.Date
            DateTimePicker4.MaxDate = day
            DateTimePicker4.Value = Now.Date
            DateTimePicker5.MinDate = Now.Date
            DateTimePicker5.MaxDate = day
            DateTimePicker5.Value = Now.Date
            DateTimePicker6.MinDate = Now.Date
            DateTimePicker6.MaxDate = day
            DateTimePicker6.Value = Now.Date
            DateTimePicker7.MinDate = Now.Date
            DateTimePicker7.MaxDate = day
            DateTimePicker7.Value = Now.Date
            ComboBox1.SelectedIndex = 0
            ComboBox2.SelectedIndex = 0
            ComboBox3.SelectedIndex = 0
            ComboBox4.SelectedIndex = 0
            ComboBox5.SelectedIndex = 0
            Refresh()
        End With
        COMB()
        Reorg()
        Refr_graf()
        wrk = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Chrt2()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Chrt3()
    End Sub
    Private Sub Chart4_MouseClick(sender As Object, e As MouseEventArgs) Handles Chart4.MouseClick
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint
        HTR = Chart4.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart4.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT", Date.FromOADate(SelectDataPoint.XValue))
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub
    Private Sub Chart2_MouseClick(sender As Object, e As MouseEventArgs) Handles Chart2.MouseClick
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint
        HTR = Chart2.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart2.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("PROG", Date.FromOADate(SelectDataPoint.XValue))
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Panel1.Visible = False Then
            Panel1.Visible = True
            Panel2.Visible = True
            Panel4.Visible = True
            Panel7.Visible = True
            FlowLayoutPanel1.Controls.Add(Panel1)
            FlowLayoutPanel1.Controls.Add(Panel2)
            FlowLayoutPanel1.Controls.Add(Panel4)
            FlowLayoutPanel1.Controls.Add(Panel7)

        Else
            Panel3.Visible = True
            FlowLayoutPanel1.Controls.Add(Panel3)
            Panel6.Visible = True
            FlowLayoutPanel1.Controls.Add(Panel6)
            Panel8.Visible = True
            FlowLayoutPanel1.Controls.Add(Panel8)

        End If
        Me.SplitContainer1.Panel2Collapsed = True
        Me.SplitContainer1.SplitterDistance = Me.SplitContainer1.Size.Height
        Reorg()
    End Sub
    Private Sub FILL_TBL(col As String, dat As Date, Optional strg As String = "")
        Dim tblFields As String
        Dim f As String

        If col = "AKT" Then
            tblFields = "Select * from Formatka a where a.work_day=@dat and case when @rodzaj4!='Wszystkie' then a.rodzaj=@rodzaj4 else a.rodzaj is not null end and case when @koor3!='WSZYSCY' then a.koor=@koor3 else a.koor is not null end"
        ElseIf col = "PROG" Then
            tblFields = "Select * from Formatka_bil a where a.work_day=@dat and case when @rodzaj1!='Wszystkie' then a.rodzaj=@rodzaj1 else a.rodzaj is not null end and case when @koor2!='WSZYSCY' then a.koor=@koor2 else a.koor is not null end "
        ElseIf col = "VAL" Then
            tblFields = "select * from bilans_val where planner_buyer=@strg and case when @wart='In Plus - nadwyżka' then wartość>0 else wartość<0 end"
        ElseIf col = "brak_d" Then
            tblFields = "select * from braki_poreal where date_required=@dat and typ='DOP'"
        ElseIf col = "brak_m" Then
            tblFields = "select * from braki_poreal where date_required=@dat and typ='MRP'"
        ElseIf col = "AKT_d" Then
            tblFields = "Select get_koor(part_no) koor,Case When dop=0 Then 'MRP' else 'DOP' end typ,part_no,descr,sum(prod_qty) zagrożenie_prod,sum(qty_demand) Potrzeba,wrkc,next_wrkc from ord_lack where case when @dat=current_date then date_required<=current_date else date_required=@dat end and get_koor(part_no)=@strg and  dop!=0 and order_supp_dmd!='Zam. zakupu' group by get_koor(part_no),Case When dop=0 Then 'MRP' else 'DOP' end ,part_no,descr,wrkc,next_wrkc"
        ElseIf col = "AKT_m" Then
            tblFields = "Select get_koor(part_no) koor,Case When dop=0 Then 'MRP' else 'DOP' end typ,part_no,descr,sum(prod_qty) zagrożenie_prod,sum(qty_demand) Potrzeba,wrkc,next_wrkc from ord_lack where case when @dat=current_date then date_required<=current_date else date_required=@dat end and get_koor(part_no)=@strg and  dop=0 and order_supp_dmd!='Zam. zakupu' group by get_koor(part_no),Case When dop=0 Then 'MRP' else 'DOP' end ,part_no,descr,wrkc,next_wrkc"
        ElseIf col = "AKT_a" Then
            tblFields = "Select * from Formatka a where a.work_day=@dat and a.koor=@strg"
        ElseIf col = "AKT_p" Then
            tblFields = "Select * from Formatka a where a.work_day=@dat and  a.rodzaj=@strg"
        ElseIf col = "AKT_g" Then
            tblFields = "Select get_koor(part_no) koor,Case When dop=0 Then 'MRP' else 'DOP' end typ,part_no,descr,sum(prod_qty) zagrożenie_prod,sum(qty_demand) Potrzeba,CASE WHEN substring(wrkc, 1, 1) = '4' THEN CASE WHEN substring(wrkc, 1, 2) = '40' THEN wrkc ELSE '400TAP' END ELSE wrkc END AS wrkc,CASE WHEN substring(next_wrkc, 1, 1) = '4' THEN CASE WHEN substring(next_wrkc, 1, 2) = '40' THEN next_wrkc ELSE '400TAP' END ELSE next_wrkc END AS next_wrkc from ord_lack where case when @dat=current_date then date_required<=current_date else date_required=@dat end and (CASE WHEN substring(wrkc, 1, 1) = '4' THEN CASE WHEN substring(wrkc, 1, 2) = '40' THEN wrkc ELSE '400TAP' END ELSE wrkc END =@strg or CASE WHEN substring(next_wrkc, 1, 1) = '4' THEN CASE WHEN substring(next_wrkc, 1, 2) = '40' THEN next_wrkc ELSE '400TAP' END ELSE next_wrkc END=@strg) and order_supp_dmd!='Zam. zakupu' group by get_koor(part_no),Case When dop=0 Then 'MRP' else 'DOP' end ,part_no,descr,CASE WHEN substring(wrkc, 1, 1) = '4' THEN CASE WHEN substring(wrkc, 1, 2) = '40' THEN wrkc ELSE '400TAP' END ELSE wrkc END,CASE WHEN substring(next_wrkc, 1, 1) = '4' THEN CASE WHEN substring(next_wrkc, 1, 2) = '40' THEN next_wrkc ELSE '400TAP' END ELSE next_wrkc END"
        Else
            tblFields = ""
        End If
        Using kors As New DataTable()
            Using conB As New NpgsqlConnection(NpA)
                conB.Open()
                Using con As New NpgsqlCommand(tblFields, conB)
                    con.Parameters.Add("dat", NpgsqlTypes.NpgsqlDbType.Date).Value = dat.Date
                    con.Parameters.Add("rodzaj4", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox4.SelectedItem.ToString
                    con.Parameters.Add("rodzaj1", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox1.SelectedItem.ToString
                    con.Parameters.Add("koor3", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox3.SelectedItem.ToString
                    con.Parameters.Add("koor2", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox2.SelectedItem.ToString
                    con.Parameters.Add("strg", NpgsqlTypes.NpgsqlDbType.Varchar).Value = strg
                    con.Parameters.Add("wart", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ComboBox5.SelectedItem
                    con.Prepare()
                    Using rs As NpgsqlDataReader = con.ExecuteReader
                        kors.Load(rs)
                        f = V.BaseFilter
                        V.ActivateAllFilters(False)
                        DataGridView1.DataSource = kors
                        V.BaseFilter = f
                        V.RebuildFilter()
                    End Using
                End Using
            End Using
        End Using
        Reorg()
    End Sub
    Private Sub DataGridView1_RowHeaderMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        Process.Start("http://ifsvapp1.sits.local:59080/client/runtime/Ifs.Fnd.Explorer.application?url=ifsapf%3AfrmAvailabilityPlanning%3Faction%3Dget%26key1%3D*%255EST%255E" & DataGridView1.Rows(e.RowIndex).Cells("Part_no").Value.ToString & "%255E*%26COMPANY%3DSITS")
    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If ToolStripButton1.Checked = True Then
            Panel1.Visible = True
            Panel2.Visible = True
            Panel4.Visible = True
            Panel3.Visible = True
            FlowLayoutPanel1.Controls.Add(Panel1)
            FlowLayoutPanel1.Controls.Add(Panel2)
            FlowLayoutPanel1.Controls.Add(Panel3)
            FlowLayoutPanel1.Controls.Add(Panel4)
        Else
            Panel1.Visible = False
            Panel2.Visible = False
            Panel4.Visible = False
            Panel3.Visible = False
            FlowLayoutPanel1.Controls.Remove(Panel1)
            FlowLayoutPanel1.Controls.Remove(Panel2)
            FlowLayoutPanel1.Controls.Remove(Panel4)
            FlowLayoutPanel1.Controls.Remove(Panel3)
            If Not SplitContainer1.Panel2Collapsed Then SplitContainer1.Panel2Collapsed = True
        End If
        reorg()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If ToolStripButton3.Checked = True Then
            Panel6.Visible = True
            Panel7.Visible = True
            Panel8.Visible = True
            FlowLayoutPanel1.Controls.Add(Panel6)
            FlowLayoutPanel1.Controls.Add(Panel7)
            FlowLayoutPanel1.Controls.Add(Panel8)

        Else
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            FlowLayoutPanel1.Controls.Remove(Panel6)
            FlowLayoutPanel1.Controls.Remove(Panel7)
            FlowLayoutPanel1.Controls.Remove(Panel8)
        End If
        Reorg()
    End Sub
    Private Sub Reorg()
        With FlowLayoutPanel1.Controls
            If Panel1.Visible Then .SetChildIndex(Panel1, 0)
            If Panel2.Visible Then .SetChildIndex(Panel2, IIf(Panel7.Visible, 1, 6))
            If Panel4.Visible Then .SetChildIndex(Panel4, 2)
            If Panel7.Visible Then .SetChildIndex(Panel7, 3)
            If Panel6.Visible Then .SetChildIndex(Panel6, 4)
            If Panel3.Visible Then .SetChildIndex(Panel3, 5)
            If Panel8.Visible Then .SetChildIndex(Panel8, 7)
            If Panel5.Visible Then .SetChildIndex(Panel5, 8)
            If Panel9.Visible Then .SetChildIndex(Panel9, 9)
            If Panel10.Visible Then .SetChildIndex(Panel10, 10)
            If Panel11.Visible Then .SetChildIndex(Panel11, 11)
        End With
        FlowLayoutPanel1.Refresh()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not wrk Then Refr_graf()
    End Sub

    Private Sub DateTimePicker6_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker6.ValueChanged
        If Not wrk Then Refr_graf()
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        If Not wrk Then Refr_graf()
    End Sub

    Private Sub DateTimePicker7_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker7.ValueChanged
        If Not wrk Then Refr_graf()
    End Sub

    Private Sub DateTimePicker5_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker5.ValueChanged
        If Not wrk Then Refr_graf()
    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        If Not wrk Then Refr_graf()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If Not wrk Then Refr_graf()
    End Sub

    Private Sub Chart12_Click(sender As Object, e As MouseEventArgs) Handles Chart12.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart12.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart12.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("VAL", Now.Date, SelectDataPoint.AxisLabel)
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub
    Private Sub Chart6_Click(sender As Object, e As MouseEventArgs) Handles Chart6.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart6.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart6.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("brak_d", Date.FromOADate(SelectDataPoint.XValue))
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart5_Click(sender As Object, e As MouseEventArgs) Handles Chart5.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint
        HTR = Chart5.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart5.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("brak_m", Date.FromOADate(SelectDataPoint.XValue))
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart8_Click(sender As Object, e As MouseEventArgs) Handles Chart8.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart8.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart8.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT_d", DateTimePicker6.Value, SelectDataPoint.AxisLabel.Substring(0, 7).Trim())
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart7_Click(sender As Object, e As MouseEventArgs) Handles Chart7.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart7.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart7.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT_m", DateTimePicker4.Value, SelectDataPoint.AxisLabel.Substring(0, 7).Trim())
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart1_Click(sender As Object, e As MouseEventArgs) Handles Chart1.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart1.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart1.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT_a", Now.Date, SelectDataPoint.AxisLabel.Substring(0, 7).Trim())
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart3_Click(sender As Object, e As MouseEventArgs) Handles Chart3.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart3.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart3.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT_p", Now.Date, SelectDataPoint.AxisLabel)
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart9_Click(sender As Object, e As MouseEventArgs) Handles Chart9.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart9.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart9.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT_g", DateTimePicker3.Value.Date, SelectDataPoint.AxisLabel)
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart10_Click(sender As Object, e As MouseEventArgs) Handles Chart10.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart10.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart10.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT_g", DateTimePicker5.Value.Date, SelectDataPoint.AxisLabel)
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

    Private Sub Chart11_Click(sender As Object, e As MouseEventArgs) Handles Chart11.Click
        Dim HTR As HitTestResult
        Dim SelectDataPoint As DataPoint

        HTR = Chart11.HitTest(e.X, e.Y)
        If HTR.ChartElementType.ToString = "DataPoint" Or HTR.ChartElementType.ToString = "ChartArea" Or HTR.ChartElementType.ToString = "DataPointLabel" Then
            SelectDataPoint = Chart11.Series(0).Points(HTR.PointIndex)
            'MsgBox(SelectDataPoint.ToString)
            FILL_TBL("AKT_g", DateTimePicker7.Value.Date, SelectDataPoint.AxisLabel)
            Me.SplitContainer1.Panel2Collapsed = False
            Me.SplitContainer1.SplitterDistance = CInt(Me.SplitContainer1.Size.Height * 0.5)
        End If
    End Sub

End Class