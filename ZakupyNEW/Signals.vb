Imports System.IO
Imports System.Threading
Imports Npgsql
Imports DgvFilterPopup
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Net.WebClient

Public Class Signals
    Private Const WM_SETREDRAW As Integer = 11
    Dim upd_now As Boolean = False
    Public src As New DataTable
    Dim Braki As New DataTable
    Public runnow As Boolean
    Public rase_evn As Boolean
    Dim V As DgvFilterManager
    Dim m As DGVColumnSelector.DataGridViewColumnSelector
    Dim tree_busy As Boolean
    Dim txtdat(5, 1) As String
    Dim EVNTdat(5) As String
    Public col_war As Boolean = True
    Friend NotInheritable Class NativeMethods
        Public Declare Function SendMessage Lib "user32" _
            Alias "SendMessageA" _
            (ByVal hWnd As Integer, ByVal wMsg As Integer,
            ByVal wParam As Integer, ByRef lParam As Object) _
            As Integer
    End Class
    Public Sub New()
        ' This call is required by the designer.
        'Application.Run(LoginForm1)
        InitializeComponent()
        DataGridView1.GetType.InvokeMember("DoubleBuffered", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.SetProperty, Nothing, DataGridView1, New Object() {True})
        Timer1.Stop()
        LoginForm1.Hide()
        rase_evn = True
        If Not Directory.Exists("D:\") Then
            driveD = False
        Else
            driveD = True
        End If
        If Not Directory.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES") Then Directory.CreateDirectory(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES")
        Dim myImageList As New ImageList()
        myImageList.Images.Add(ToolStripButton6.Image) 'IFS 0
        myImageList.Images.Add(ToolStripButton9.Image) 'APPICO 1 
        myImageList.Images.Add(ToolStripStatusLabel1.Image) 'OFFLINE 2 
        myImageList.Images.Add(ToolStripButton12.Image) 'Bez informacji 3 
        myImageList.Images.Add(ToolStripButton9.Image)  'Potwierdzone braki 4 
        myImageList.Images.Add(ToolStripButton10.Image) 'Czekam na informację 5
        myImageList.Images.Add(ToolStripButton13.Image) 'Nie używam 6
        myImageList.Images.Add(ToolStripButton11.Image) 'Komponent alternatywny 7
        myImageList.Images.Add(ToolStripButton2.Image) ' filtr daty
        Me.TreeView1.ImageList = myImageList
        Fil_txtdat()
        If GetSetting("Zakupy", "Main", "online", "0") = 1 Then
            ToolStripStatusLabel1.Text = "ONLINE"
            ToolStripButton3.Checked = True
        Else
            ToolStripStatusLabel1.Text = "OFFLINE"
            ToolStripButton3.Checked = False
        End If
        InitializeCheckTreeView()
        If Not Directory.Exists("Z:\katpryw\Radkos") Then ToolStripStatusLabel2.Text = "Brak wymiany z siecią SITS" Else ToolStripStatusLabel2.Text = "sieć SITS dostępna"
        ' Add any initialization after the InitializeComponent() call.
        Get_datset()

        DataGridView1.PerformLayout()
        V = New DgvFilterManager(DataGridView1)
        m = New DGVColumnSelector.DataGridViewColumnSelector(DataGridView1)
        V.AutoCreateFilters = True
        rase_evn = False
        Refr_SETT()
        ToolStripStatusLabel5.Text = "Data aktualizacji danych : " & Refr_date.ToString
        DataGridView1.Refresh()
        Timer1.Start()
        If GetSetting("ZakupyNEW", "Main", "UID", "NoPostgresql") = "NoPostgresql" Then
            ToolStrip3.Enabled = False
            ToolStrip2.Enabled = False
            StatusStrip1.Items(0).Enabled = False
        Else
            ToolStrip3.Enabled = True
            ToolStrip2.Enabled = True
            StatusStrip1.Items(0).Enabled = True
        End If
        If GetSetting("Zakupy", "Main", "online", "0") = 1 Then Timer1.Start()

        'DataGridView1.RowTemplate.Height = 20
    End Sub
    Private Sub Fil_txtdat()
        txtdat(0, 0) = "Tydzień"
        txtdat(0, 1) = 7
        txtdat(1, 0) = "Dwa Tygodnie"
        txtdat(1, 1) = 14
        txtdat(2, 0) = "Trzy tygodnie"
        txtdat(2, 1) = 21
        txtdat(3, 0) = "30 dni"
        txtdat(3, 1) = 30
        txtdat(4, 0) = "60 dni"
        txtdat(4, 1) = 60
        txtdat(5, 0) = "90 dni"
        txtdat(5, 1) = 90
        EVNTdat(0) = "Opóźniona dostawa"
        EVNTdat(1) = "Dzisiejsza dostawa"
        EVNTdat(2) = "Dostawa na dzisiejsze ilości"
        EVNTdat(3) = "Braki w gwarantowanej dacie"
        EVNTdat(4) = "Brakujące ilości"
        EVNTdat(5) = "Brak zamówień zakupu"
    End Sub
    Private Sub InitializeCheckTreeView()
        ' Show check boxes for the TreeView.
        ' Create the StateImageList and add two images.
        TreeView1.StateImageList = New ImageList()
        TreeView1.StateImageList.Images.Add(My.Resources.ToggleDownAltgreen1)
        TreeView1.StateImageList.Images.Add(My.Resources.ToggleRightAltred1)

    End Sub
    Private Sub Refr_SETT()
        tree_busy = True
        Get_changes()
        If TreeView1.Nodes.Count > 0 Then Save_check()
        Pop_tree()
        Refr_graf()
        Retr_nodes()
        Do_TREE_CHOISE()
        GC.Collect()
        tree_busy = False
    End Sub
    Private Sub Get_changes()
        If src.Rows.Count > 0 Then
            Dim rRow() As Data.DataRow
            Try
                Using kors As New DataTable()
                    Using conB As New NpgsqlConnection(NpA)
                        conB.Open()
                        Try
                            Application.DoEvents()
                            Using con As New NpgsqlCommand("select * from public.potw where date_created>=current_timestamp - interval '10 minute'", conB)
                                Using rs As NpgsqlDataReader = con.ExecuteReader
                                    Using Res1 As New DataTable()
                                        Res1.Load(rs)
                                        If Res1.Rows.Count > 0 Then
                                            For Each Rs1 As DataRow In Res1.Rows
                                                If Rs1("rodzaj_potw") <> "NIE ZAMAWIAM" Then
                                                    'MsgBox(Mid(Rs1("Data_dost").Value.ToString, 5, 2) & "/" & Microsoft.VisualBasic.Right(Rs1("Data_dost").Value.ToString, 2) & "/" & Microsoft.VisualBasic.Left(Rs1("Data_dost").Value.ToString, 4))
                                                    rRow = src.Select("indeks='" & Rs1("indeks") & "' and wlk_dost='" & Convert.ToDouble(Rs1("dost_ilosc")).ToString.Replace(".", ",") & "' and Data_dost='" & Rs1("data_dost") & "'")
                                                Else
                                                    rRow = src.Select("indeks='" & Rs1("indeks") & "' and status_Informacji<>''")
                                                End If
                                                If rRow.Count > 0 Then
                                                    For Each rek As DataRow In rRow
                                                        If rek("status_informacji").ToString.Length > 3 Then
                                                            If rek("Status_Informacji") <> Rs1("Rodzaj_potw") Then
                                                                rek("Status_Informacji") = Rs1("Rodzaj_potw")
                                                                rek("Informacja") = Rs1("info")
                                                                rek.AcceptChanges()
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    End Using
                                End Using
                            End Using

                        Catch ex As Exception
                            If SHOW_err Then MsgBox(ex.Message & " get_changes ex")
                        End Try
                    End Using
                End Using
            Catch ex1 As Exception
                If SHOW_err Then MsgBox(ex1.Message & " get_changes ex1")
            End Try
        End If
    End Sub
    Private Sub Save_check()
        Dim tmp_txt As String
        Dim nod As TreeNode
        Dim nod1 As TreeNode
        Dim nod2 As TreeNode
        TreeView1.BeginUpdate()
        Try
            If File.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\NODES.txt") Then File.Delete(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\NODES.txt")
            tmp_txt = IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\NODES.txt"
            Dim objWriter As New System.IO.StreamWriter(tmp_txt)
            For Each nod In TreeView1.Nodes
                If nod.Checked Then
                    objWriter.WriteLine(nod.Name & "_" & nod.FullPath & "_")
                    If nod.Nodes.Count > 0 Then
                        For Each nod1 In nod.Nodes
                            If nod1.Checked Then
                                objWriter.WriteLine(nod1.Name & "_" & nod1.FullPath & "_")
                                If nod1.Nodes.Count > 0 Then
                                    For Each nod2 In nod1.Nodes
                                        If nod2.Checked Then
                                            objWriter.WriteLine(nod2.Name & "_" & nod2.FullPath & "_")
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                End If
            Next
            objWriter.Close()
            objWriter = Nothing
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " save_check ex")
        End Try
        TreeView1.EndUpdate()
    End Sub
    Private Sub Retr_nodes()
        Dim z As String
        Dim tmp_txt As String
        Dim nod As TreeNode
        Dim nod1 As TreeNode
        Dim nod2 As TreeNode
        TreeView1.BeginUpdate()
        Try
            If File.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\NODES.txt") Then
                tmp_txt = IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\NODES.txt"
                Dim objReader As New System.IO.StreamReader(tmp_txt)
                z = objReader.ReadToEnd
                objReader.Close()
                objReader = Nothing
                For Each nod In TreeView1.Nodes
                    If InStr(z, "_" & nod.FullPath & "_") > 0 Then
                        nod.Checked = True
                        If nod.Nodes.Count > 0 Then
                            For Each nod1 In nod.Nodes
                                If InStr(z, "_" & nod1.FullPath & "_") > 0 Then
                                    nod1.Checked = True
                                    If nod1.Nodes.Count > 0 Then
                                        For Each nod2 In nod1.Nodes
                                            If InStr(z, "_" & nod2.FullPath & "_") > 0 Then
                                                nod2.Checked = True
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " retr_nodes ex")
        End Try
        TreeView1.EndUpdate()
        nod = Nothing
        nod1 = Nothing
        nod2 = Nothing
    End Sub
    Private Sub Pop_tree()
        Try
            Try
                Using _dataTable As New DataTable()
                    Using conB As New NpgsqlConnection(NpA)
                        conB.Open()
                        Using con As New NpgsqlCommand("Select PLANNER_BUYER,coalesce(Status_Informacji,'INFORMACJA') Status_Informacji,TYP_zdarzenia from DATA group by PLANNER_BUYER,coalesce(Status_Informacji,'INFORMACJA'),TYP_zdarzenia order by PLANNER_BUYER,coalesce(Status_Informacji,'INFORMACJA'),TYP_zdarzenia ", conB)
                            Using rs As NpgsqlDataReader = con.ExecuteReader
                                _dataTable.Load(rs)
                                TreeView1.BeginUpdate()
                                TreeView1.Nodes.Clear()
                                Dim Rootnode As TreeNode = Nothing
                                Dim Mainnode As TreeNode = Nothing
                                Dim Childnode As TreeNode = Nothing
                                TreeView1.CheckBoxes = True
                                Dim r_nam As String = String.Empty
                                Dim MainName As String = String.Empty
                                Rootnode = TreeView1.Nodes.Add(key:="CHOISE", text:="FILTR Użytkownicy",
                             imageIndex:=5, selectedImageIndex:=4)
                                For Each row As DataRow In _dataTable.Rows
                                    If r_nam <> row(0).ToString Then
                                        Rootnode = TreeView1.Nodes.Add(key:="PLANNER_BUYER", text:=row(0).ToString,
                             imageIndex:=0, selectedImageIndex:=0)
                                        r_nam = row(0).ToString
                                    End If
                                    If MainName <> row(1).ToString Then
                                        Mainnode = Rootnode.Nodes.Add(key:="Status_Informacji", text:="Status_Informacji: " & row(1).ToString,
                                        imageIndex:=GET_ICON(row(1)), selectedImageIndex:=5)
                                        MainName = row(1).ToString
                                    End If
                                    Childnode = Mainnode.Nodes.Add(key:="TYP_zdarzenia", text:=row(2).ToString,
                            imageIndex:=2, selectedImageIndex:=4)
                                Next
                                Rootnode = TreeView1.Nodes.Add(key:="MAIN", text:="Kolorowanie warunkowe",
                             imageIndex:=5, selectedImageIndex:=4)
                                Rootnode = TreeView1.Nodes.Add(key:="DATES", text:="Filtry daty",
                             imageIndex:=8, selectedImageIndex:=5)
                                Mainnode = Rootnode.Nodes.Add(key:="DATA_BRAKU", text:="Data braku",
                                        imageIndex:=5, selectedImageIndex:=4)
                                For i = 0 To 5
                                    Childnode = Mainnode.Nodes.Add(key:=txtdat(i, 1), text:=txtdat(i, 0),
                                imageIndex:=2, selectedImageIndex:=4)
                                Next i
                                Mainnode = Rootnode.Nodes.Add(key:="DATA_dost", text:="Data dostawy",
                                        imageIndex:=5, selectedImageIndex:=4)
                                For i = 0 To 5
                                    Childnode = Mainnode.Nodes.Add(key:=txtdat(i, 1), text:=txtdat(i, 0),
                                imageIndex:=2, selectedImageIndex:=4)
                                Next i
                                Rootnode = TreeView1.Nodes.Add(key:="EVNT", text:="Typy zdarzeń",
                                imageIndex:=2, selectedImageIndex:=4)
                                For i = 0 To 5
                                    Mainnode = Rootnode.Nodes.Add(key:="TYP_zdarzenia", text:=EVNTdat(i),
                                            imageIndex:=5, selectedImageIndex:=4)
                                Next
                                Rootnode = TreeView1.Nodes.Add(key:="CHK", text:="Status Informacji",
                                imageIndex:=2, selectedImageIndex:=4)
                                Mainnode = Rootnode.Nodes.Add(key:="Status_Informacji", text:="Informacyjne",
                                            imageIndex:=5, selectedImageIndex:=4)
                                Mainnode = Rootnode.Nodes.Add(key:="Status_Informacji", text:="Brak potwierdzenia",
                                            imageIndex:=5, selectedImageIndex:=4)
                                Mainnode = Rootnode.Nodes.Add(key:="Status_Informacji", text:="Informacja przesłana",
                                            imageIndex:=5, selectedImageIndex:=4)
                                Mainnode = Rootnode.Nodes.Add(key:="Status_Informacji", text:="INNE",
                                            imageIndex:=5, selectedImageIndex:=4)

                                TreeView1.Nodes(0).EnsureVisible()
                            End Using
                        End Using
                    End Using
                End Using
            Catch ex1 As Exception
                If SHOW_err Then MsgBox(ex1.Message & " pop_tree ex1")
            End Try
            TreeView1.EndUpdate()
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " pop_tree ex")
        End Try
    End Sub
    Private Shadows Function GET_ICON(NAME As String) As Integer
        Select Case NAME
            Case "BRAK"
                Return 3
            Case "NIE ZAMAWIAM"
                Return 6
            Case "POTWIERDZONE"
                Return 4
            Case "Użyj FR"
                Return 7
            Case "Czekam na INFO"
                Return 5
            Case Else
                Return 0
        End Select
    End Function
    Private Sub ToolStripStatusLabel1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripStatusLabel1.Click
        changepass.ShowDialog()
        If GetSetting("Zakupy", "Main", "online", "0") = 1 Then ToolStripStatusLabel1.Text = "ONLINE" Else ToolStripStatusLabel1.Text = "OFFLINE"
        changepass.Dispose()
    End Sub
    Private Sub Form2_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Panel1.TopLevelControl.BringToFront()
        Me.Panel1.Left = CInt(Me.Width / 2) - SplitContainer1.SplitterDistance - CInt(Panel1.Width / 2)
        Me.Panel1.Top = CInt(Me.Height / 2) - CInt(Panel1.Height)
        Me.Panel1.Visible = True
        Me.Label1.Text = "Zapisuję ustawienia"
        Me.Refresh()
        Save_check()
        SaveSettings()
        ToolStripManager.SaveSettings(Me)
        My.Settings.FRM2WS = Me.WindowState
        My.Settings.FRM2SIZEW = Me.Width
        My.Settings.FRM2SIZEH = Me.Height
        My.Settings.SplitCDIST1 = SplitContainer1.SplitterDistance
        My.Settings.SplitCDIST2 = SplitContainer2.SplitterDistance
        My.Settings.SplitCDIST3 = SplitContainer3.SplitterDistance
        Me.Label1.Text = "Proszę czekać zamykam aktywne procesy"
        Me.Refresh()
    End Sub
    Private Sub Refr_graf()
        Try
            Using kors As New DataTable()
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select b.koor,sum(b.brak)/a.al as bilans from (select sum(brak) AL from (select part_no,min(bal_stock) brak from demands a where work_day<=current_date and koor!='*' And koor!='LUCPRZ' group by part_no) a where brak<0) a,(select part_no,min(bal_stock) brak,koor from demands a where work_day<=current_date group by part_no,koor) b where b.brak<0 and b.koor!='*' And b.koor!='LUCPRZ' group by b.koor,a.al", conB)
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                            Dim r As Data.DataRow
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
                            Label2.Text = "Poziom dostaw dla prod. " & Math.Round((1 - (BR / ALL)) * 100, 2) & "%"
                            Chart3.DataSource = kors
                            Dim CArea As ChartArea = Chart3.ChartAreas(0)
                            CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                            CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                            CArea.Area3DStyle.Enable3D = True
                            CArea.AxisX.MajorGrid.Enabled = True   '~~> Removed the X axis major grids
                            CArea.AxisY.MajorGrid.Enabled = True  '~~> Removed the Y axis major grids
                            CArea.AxisY.LabelStyle.Format = "0.00%"
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
    End Sub
    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            If Not IsDBNull(DataGridView1.SelectedRows(0).Cells("Informacja").Value) Then
                ToolStripComboBox1.Text = DataGridView1.SelectedRows(0).Cells("Informacja").Value.ToString().Substring(InStr(DataGridView1.SelectedRows(0).Cells("Informacja").Value.ToString(), "::",) + 1)
            Else
                ToolStripComboBox1.Text = ""
            End If
            If Me.SplitContainer3.SplitterDistance + 4 <> Me.SplitContainer3.Size.Height Then
                Using kors As New DataTable()
                    Using conB As New NpgsqlConnection(NpA)
                        conB.Open()
                        Using con As New NpgsqlCommand("select a.work_day,case when a.balance<0 then a.balance*-1 else null end bilans from demands a left join (select part_no,min(work_day) fir,max(work_day) lir from demands where balance<0 group by part_no) b on b.part_no=a.part_no where (a.work_day between b.fir and b.lir) and a.part_no=@part_no order by a.work_day", conB)
                            con.Parameters.Add("part_no", NpgsqlTypes.NpgsqlDbType.Varchar).Value = DataGridView1.SelectedRows(0).Cells("INDEKS").Value.ToString
                            con.Prepare()

                            Using rs As NpgsqlDataReader = con.ExecuteReader
                                kors.Load(rs)
                                Chart2.DataSource = kors
                                Dim CArea As ChartArea = Chart2.ChartAreas(0)
                                CArea.BackColor = Color.FromArgb(180, Color.Azure)           '~~> Changing the Back Color of the Chart Area 
                                CArea.ShadowColor = Color.FromArgb(180, Color.Red)          '~~> Changing the Shadow Color of the Chart Area 
                                CArea.Area3DStyle.Enable3D = True
                                CArea.AxisX.MajorGrid.Enabled = False   '~~> Removed the X axis major grids
                                CArea.AxisY.MajorGrid.Enabled = False   '~~> Removed the Y axis major grids
                                CArea.AxisY.LabelStyle.Format = "0.00" '~~> Formatting Y axis to display values in %age
                                CArea.AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                                Dim Series1 As Series = Chart2.Series(0)
                                Dim ann As Annotation = Chart2.Annotations(0)

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
                                'Chart2.Series(Series1.Name).Color = Color.FromArgb(180, Color.Blue)
                                '~~> Setting label's Fore Color
                                Chart2.Series(Series1.Name).LabelForeColor = Color.DarkRed
                                '~~> Setting label's Format to %age
                                Chart2.Series(Series1.Name).LabelFormat = "0.00"
                            End Using
                        End Using
                    End Using
                End Using
            Else
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowHeaderMouseClick ex")
        End Try
    End Sub
    Private Sub DataGridView1_RowHeaderMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        Process.Start("rundll32.exe", "dfshim.dll,ShOpenVerbApplication " & "http://ifsvapp1.sits.local:59080/client/runtime/Ifs.Fnd.Explorer.application?url=ifsapf%3AfrmAvailabilityPlanning%3Faction%3Dget%26key1%3D*%255EST%255E" & DataGridView1.Rows(e.RowIndex).Cells("Indeks").Value.ToString & "%255E*%26COMPANY%3DSITS")
    End Sub
    Private Sub DataGridView1_CellPainting(ByVal sender As Object, ByVal e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If e.ColumnIndex = -1 AndAlso e.RowIndex >= 0 Then
            Dim dgvRow As DataGridViewRow = Me.DataGridView1.Rows(e.RowIndex)
            Dim dtdos As Date = CDate(dgvRow.Cells("Data_Dost").Value)
            Dim dta_lt As Date = CDate(dgvRow.Cells("date").Value)
            Dim status As String = dgvRow.Cells("typ_zdarzenia").Value
            If status = "Brakujące ilości" Or status = "Dostawa na dzisiejsze ilości" Then
                If dta_lt < dtdos Then
                    If (dta_lt - Date.Now) + (dta_lt - Date.Now) < (dtdos - Date.Now) Then
                        Dim bm As New Bitmap(My.Resources.Danger)
                        With e.Graphics
                            .FillRectangle(Brushes.Yellow, New Rectangle(e.CellBounds.Width - e.CellBounds.Size.Width + 1, e.CellBounds.Top, e.CellBounds.Size.Width, e.CellBounds.Size.Height))
                            .DrawRectangle(New Pen(Brushes.SaddleBrown, 1), New Rectangle(e.CellBounds.Width - e.CellBounds.Size.Width + 1, e.CellBounds.Top + 1, e.CellBounds.Size.Width - 2, e.CellBounds.Size.Height - 3))
                            .DrawImage(bm, e.CellBounds.Width - e.CellBounds.Size.Width + 4, e.CellBounds.Top, e.CellBounds.Size.Height - 1, e.CellBounds.Size.Height - 1)
                        End With
                        e.Handled = True
                    Else
                        Dim bm As New Bitmap(My.Resources.Help)
                        With e.Graphics
                            .FillRectangle(Brushes.WhiteSmoke, New Rectangle(e.CellBounds.Width - e.CellBounds.Size.Width + 1, e.CellBounds.Top, e.CellBounds.Size.Width, e.CellBounds.Size.Height))
                            .DrawRectangle(New Pen(Brushes.LightGray, 1), New Rectangle(e.CellBounds.Width - e.CellBounds.Size.Width + 1, e.CellBounds.Top + 1, e.CellBounds.Size.Width - 2, e.CellBounds.Size.Height - 3))
                            .DrawImage(bm, e.CellBounds.Width - e.CellBounds.Size.Width + 4, e.CellBounds.Top, e.CellBounds.Size.Height - 1, e.CellBounds.Size.Height - 1)
                        End With
                        e.Handled = True
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim dgvRow As DataGridViewRow = Me.DataGridView1.Rows(e.RowIndex)
        Dim bls As Double = CDbl(dgvRow.Cells("bilans").Value)
        Dim dtdos As Date = CDate(dgvRow.Cells("Data_Dost").Value)
        Try
            If col_war Then
                If bls < 0 And dtdos < Now.Date Then
                    With dgvRow.DefaultCellStyle
                        .BackColor = Color.RosyBrown                  '
                        .ForeColor = Color.White
                        .Font = New Font(DataGridView1.Font, FontStyle.Bold)
                    End With
                Else
                    If bls < 0 And dtdos = Now.Date Then
                        With dgvRow.DefaultCellStyle
                            .BackColor = Color.LightPink
                            '.ForeColor = Color.White
                            .Font = New Font(DataGridView1.Font, FontStyle.Bold)
                        End With
                    Else
                        Dim bil_dost As Double = CDbl(dgvRow.Cells("bil_dost_dzień").Value)
                        If bil_dost < 0 And dtdos <= Now.Date Then
                            With dgvRow.DefaultCellStyle
                                .BackColor = Color.Peru                  'ładny na opóżnienia jeden dzień plus Color.Peru
                                .ForeColor = Color.White
                                .Font = New Font(DataGridView1.Font, FontStyle.Italic)
                            End With
                        Else
                            If bls < 0 Then
                                Dim tp_zd As String = dgvRow.Cells("typ_zdarzenia").Value.ToString
                                If tp_zd = "Brakujące ilości" Then
                                    With dgvRow.DefaultCellStyle
                                        .BackColor = Color.Red                  '
                                        .ForeColor = Color.White
                                        .Font = New Font(DataGridView1.Font, FontStyle.Bold)
                                    End With
                                ElseIf tp_zd = "Braki w gwarantowanej dacie" Then
                                    With dgvRow.DefaultCellStyle
                                        .BackColor = Color.Yellow
                                        '.ForeColor = Color.White
                                        .Font = New Font(DataGridView1.Font, FontStyle.Italic)
                                    End With
                                Else
                                    With dgvRow.DefaultCellStyle
                                        .BackColor = Color.Gray                 '
                                        .ForeColor = Color.White
                                        .Font = New Font(DataGridView1.Font, FontStyle.Bold)
                                    End With
                                End If
                            Else
                                If bil_dost < 0 Then
                                    With dgvRow.DefaultCellStyle
                                        .BackColor = Color.LightBlue               '
                                        '.ForeColor = Color.White
                                        .Font = New Font(DataGridView1.Font, FontStyle.Bold)
                                    End With
                                Else

                                End If
                            End If
                        End If

                    End If
                End If
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " DataGridView1_RowPostPaint")
        End Try
    End Sub
    Private Sub Save_fil()
        Try
            If File.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\Filter.xml") Then File.Delete(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\Filter.xml")
            'Where Col.Visible AndAlso V(Col.Name).FilterExpression.Trim.Length > 0
            Dim Cols =
       <Columns>
           <%=
               From Col In DataGridView1.Columns.Cast(Of DataGridViewColumn)()
               Where V(Col.Name).Active = True Select
               <Item ColumnName=<%= Col.Name %> Filter=<%= V(Col.Name).FilterExpression %> Caption=<%= V(Col.Name).FilterCaption %> HEADER=<%= V(Col.Name).DataGridViewColumn.HeaderText %>/>
           %>
       </Columns>
            My.Computer.FileSystem.WriteAllText(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\Filter.xml", Cols.ToString, False)
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " save_fil")
        End Try
    End Sub
    Private Sub Load_fil(Optional fil As String = "")
        Dim tmp_txt As String
        Try
            If File.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\Filter.xml") Then
                tmp_txt = IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\Filter.xml"
                Dim Filters = From Items In XDocument.Load(tmp_txt)...<Item>
                              Where Not String.IsNullOrEmpty(Items.<Filter>.ToString)
                              Select Name = Items.@ColumnName, Filter = Items.@Filter, Caption = Items.@Caption, Header = Items.@HEADER
                If Filters.Count > 0 Then
                    For Each i In Filters
                        With V.Item(i.Name)
                            .FilterExpression = i.Filter
                            .FilterCaption = i.Caption
                            .DataGridViewColumn.HeaderText = i.Header
                            .Active = True
                            .FilterApplySoon = True
                        End With
                        V.BaseFilter = fil
                        V.RebuildFilter()
                    Next
                Else
                    V.BaseFilter = fil
                    V.RebuildFilter()
                End If
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " load_fil")
        End Try
    End Sub
    Private Sub ToolStripStatusLabel2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripStatusLabel2.Click
        If Not Directory.Exists("Z:\katpryw\Radkos") Then ToolStripStatusLabel2.Text = "Brak wymiany z siecią SITS" Else ToolStripStatusLabel2.Text = "sieć SITS dostępna"
    End Sub
    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        If Not rase_evn Then
            Timer1.Stop()
            Me.Panel1.TopLevelControl.BringToFront()
            Me.Panel1.Left = CInt(Me.Width / 2) - SplitContainer1.SplitterDistance - CInt(Panel1.Width / 2)
            Me.Panel1.Top = CInt(Me.Height / 2) - CInt(Panel1.Height)
            Me.Panel1.Visible = True
            Me.Label1.Text = "Proszę czekać pobieram dane"
            Me.Refresh()
            upd_now = True
            BackgroundWorker1.RunWorkerAsync()
            Thread.Sleep(20)
            Timer1.Start()
        Else
            MsgBox("System automatycznie zaczął pobierać dane o godz :" & start_cal.TimeOfDay.ToString)
        End If
    End Sub
    Private Sub SplitContainer1_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles SplitContainer1.MouseDoubleClick
        If Me.SplitContainer1.Panel1.AutoScrollMinSize.Height = Me.SplitContainer1.SplitterDistance Then
            Me.SplitContainer1.SplitterDistance = Me.TreeView1.PreferredSize.Width + Me.SplitContainer1.Panel1.PreferredSize.Width + Me.SplitContainer1.SplitterWidth + 50
        Else
            Me.SplitContainer1.SplitterDistance = Me.SplitContainer1.Panel1.AutoScrollMinSize.Width
        End If
    End Sub
    Private Sub Get_datset(Optional runnow As Boolean = False)
        rase_evn = True
        Dim wyn As String
        wyn = Signal_data.CRE_Soures(runnow)
        If wyn = "SQLREAD" Then
            src = Save_ADO(runnow)
            QWERTY()
        ElseIf wyn = "GETDATA" Then
            src = Save_ADO(runnow)
            'src = GETDATA()
            QWERTY()
        Else
            rase_evn = False
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        If Not rase_evn Then
            If upd_now Then
                Dim Thread1 As New System.Threading.Thread(AddressOf Task_2)
                Thread1.Start()
                Get_datset(True)
            Else
                Get_datset()
            End If
        End If
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If rase_evn Then
            ToolStripStatusLabel5.Text = "Data aktualizacji danych : " & Refr_date.ToString
            Refr_SETT()
            Me.Panel1.Visible = False
            Me.Refresh()
            If upd_now Then upd_now = False
            rase_evn = False
        End If
    End Sub
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        If Not rase_evn Then
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub
    Private Sub TreeView1_AfterCheck(sender As Object, e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterCheck
        'MsgBox(e.Node.Text & " " & e.Node.Name)
        If Not tree_busy Then
            tree_busy = True
            Do_TREE_CHOISE()
            tree_busy = False
        End If
    End Sub
    Private Sub Do_TREE_CHOISE()
        Dim fil_ter As String = String.Empty
        Dim tmp_fil As String
        Dim tm_dat As String = String.Empty
        Dim nod As TreeNode
        Dim nod1 As TreeNode
        Dim nod2 As TreeNode
        Dim nodT As TreeNode
        Dim chk1 As Boolean = False
        Dim chk2 As Boolean = False
        Dim chk3 As Boolean = False
        Dim chk4 As Boolean = False
        TreeView1.BeginUpdate()
        Try
            For Each nod In TreeView1.Nodes
                If nod.Text = "FILTR Użytkownicy" Then
                    If nod.Checked <> True Then
                        For Each nodT In TreeView1.Nodes
                            If nodT.Name = "PLANNER_BUYER" Then
                                nodT.Collapse()
                                nodT.ForeColor = Color.Gray
                            End If
                        Next
                    Else
                        For Each nodT In TreeView1.Nodes
                            If nodT.Name = "PLANNER_BUYER" Then
                                nodT.ForeColor = Color.Black
                            End If
                        Next
                    End If
                End If
                If nod.Text = "Kolorowanie warunkowe" Then
                    col_war = nod.Checked
                End If
                If nod.Nodes.Count > 0 And Color.Gray <> nod.ForeColor Then
                    For Each nod1 In nod.Nodes
                        If nod1.Nodes.Count > 0 Then
                            For Each nod2 In nod1.Nodes
                                If nod2.Checked Then
                                    chk2 = True
                                End If
                            Next
                        End If
                        If chk2 Or nod1.Checked Then
                            nod1.Checked = True
                            nod1.Expand()
                            chk2 = False
                            chk1 = True
                        End If
                    Next
                End If
                If chk1 Then
                    nod.Expand()
                    nod.Checked = True
                    chk1 = False
                End If
            Next '
            tmp_fil = String.Empty
            'Pobierz dane do filtrowania tabeli
            chk1 = False
            For Each nod In TreeView1.Nodes
                If nod.Checked And Color.Gray <> nod.ForeColor Then
                    chk2 = False
                    chk3 = False
                    If nod.Name = "PLANNER_BUYER" Then
                        nod.Expand()
                        tmp_fil = tmp_fil & IIf(chk1, " or (", "(") & nod.Name & "='" & nod.Text & "'"
                        chk1 = True
                        For Each nod1 In nod.Nodes
                            If nod1.Checked Then
                                nod1.Expand()
                                tmp_fil = tmp_fil & IIf(chk2, " or ", " And (") & nod1.Name & IIf(Replace(nod1.Text, "Status_Informacji: ", "") = "INFORMACJA", " is NULL", "='" & Replace(nod1.Text, "Status_Informacji: ", "") & "'")
                                chk2 = True
                                For Each nod2 In nod1.Nodes
                                    If nod2.Checked Then
                                        tmp_fil = tmp_fil & IIf(chk3, " or ", " And (") & nod2.Name & "='" & nod2.Text & "'"
                                        chk3 = True
                                    End If
                                Next
                                If chk3 Then tmp_fil = tmp_fil & ")"
                            Else
                                nod1.Collapse()
                            End If
                        Next
                        If chk2 Then tmp_fil = tmp_fil & ")"
                        tmp_fil = tmp_fil & ")"
                    Else
                        If nod.Name = "DATES" Then
                            nod.Expand()
                            For Each nod1 In nod.Nodes
                                If nod1.Checked Then
                                    nod1.Expand()
                                    For Each nod2 In nod1.Nodes
                                        If nod2.Checked Then
                                            chk4 = True
                                            tm_dat = nod1.Name & "<=#" & Now.AddDays(nod2.Name).ToString & "#"
                                        End If
                                    Next
                                End If
                                If chk4 Then tmp_fil = IIf(Len(tmp_fil) > 0, tmp_fil & " and ", "") & tm_dat
                            Next
                        End If
                        tm_dat = String.Empty
                        If nod.Name = "EVNT" Then
                            nod.Expand()
                            For Each nod1 In nod.Nodes
                                If nod1.Checked Then
                                    tm_dat = IIf(Len(tm_dat) > 0, tm_dat & " or ", "") & nod1.Name & "='" & nod1.Text & "'"
                                End If
                            Next
                            If Len(tm_dat) > 0 Then tmp_fil = IIf(Len(tmp_fil) > 0, tmp_fil & " and (", "") & tm_dat & IIf(Len(tmp_fil) > 0, ")", "")
                        End If
                        tm_dat = String.Empty
                        If nod.Name = "CHK" Then
                            nod.Expand()
                            For Each nod1 In nod.Nodes
                                If nod1.Checked Then
                                    tm_dat = IIf(Len(tm_dat) > 0, tm_dat & " or (", "(") & nod1.Name & IIf(nod1.Text = "Informacyjne", " is null ", IIf(nod1.Text = "Brak potwierdzenia", "='BRAK'", IIf(nod1.Text = "INNE", "='NIE ZAMAWIAM' or " & nod1.Name & "='Użyj FR' or " & nod1.Name & "='Czekam na INFO'", "='POTWIERDZONE'"))) & ")"
                                End If
                            Next
                            If Len(tm_dat) > 0 Then tmp_fil = IIf(Len(tmp_fil) > 0, tmp_fil & " and (", "") & tm_dat & IIf(Len(tmp_fil) > 0, ")", "")
                        End If

                    End If
                Else
                    nod.Collapse()
                End If
            Next
            Dim no_sor As Boolean = False
            Dim visible_row As Integer = DataGridView1.FirstDisplayedScrollingRowIndex
            Dim visibl_col As Integer = DataGridView1.FirstDisplayedScrollingColumnIndex
            Dim sorcol As DataGridViewColumn
            Dim sor_typ As SortOrder
            Dim i As Integer = -1
            Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
            Dim a(selectedRowCount - 1) As String

            If Not DataGridView1.SortedColumn Is Nothing Then
                sorcol = DataGridView1.SortedColumn
                sor_typ = DataGridView1.SortOrder
            Else
                sorcol = Nothing
                sor_typ = 0
                no_sor = True
            End If
            If selectedRowCount > 0 Then

                For i = 0 To selectedRowCount - 1
                    a(i) = DataGridView1.SelectedRows(i).Index.ToString
                Next i
            End If
            SuspendDrawing(DataGridView1)
            Save_fil()
            V.ActivateAllFilters(False)
            With DataGridView1
                .DataSource = src
                Load_fil(tmp_fil)
                Try
                    If Not no_sor Then .Sort(.Columns(sorcol.Name), IIf(sor_typ.ToString = "Ascending", System.ComponentModel.ListSortDirection.Ascending, System.ComponentModel.ListSortDirection.Descending))
                    If .Rows.Count > 0 Then
                        If visible_row > -1 Then .FirstDisplayedScrollingRowIndex = visible_row
                        If visibl_col > -1 Then .FirstDisplayedScrollingColumnIndex = visibl_col
                    End If
                    If i > -1 Then
                        i = i - 1
                        .CurrentCell = Me.DataGridView1(1, CInt(a(0)))
                        For j = i To 0 Step -1
                            .Rows(a(j)).Selected = True
                        Next
                    End If
                Catch ex1 As Exception
                    If SHOW_err Then MsgBox(ex1.Message & " do_TREE_CHOISE ex1")
                End Try
            End With
            ResumeDrawing(DataGridView1)
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " do_TREE_CHOISE ex")
        End Try
        TreeView1.EndUpdate()
        nod = Nothing
        nod1 = Nothing
        nod2 = Nothing
        GC.Collect()
    End Sub
    Private Sub TreeView1_BeforeCheck(sender As Object, e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeCheck
        If Color.Gray = e.Node.ForeColor Then e.Cancel = True
    End Sub
    Private Sub TreeView1_BeforeExpand(sender As Object, e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand
        If Color.Gray = e.Node.ForeColor Then e.Cancel = True
    End Sub
    Private Sub TreeView1_BeforeSelect(sender As Object, e As System.Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeSelect
        If Color.Gray = e.Node.ForeColor Then e.Cancel = True
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        If ToolStripButton3.Checked Then
            Timer1.Start()
        Else
            Timer1.Stop()
        End If
        SaveSetting("Zakupy", "Main", "online", ToolStripButton3.CheckState)
        If GetSetting("Zakupy", "Main", "online", "0") = 1 Then ToolStripStatusLabel1.Text = "ONLINE" Else ToolStripStatusLabel1.Text = "OFFLINE"
    End Sub
    Private Sub LoadSettings()
        Try
            If Not IsNothing(My.Settings.SETCOL) Then

                If Not IsNothing(My.Settings.MySettingTypes) Then

                    Dim s As MySettingTypes.DataGridViewColumnSetting = My.Settings.MySettingTypes
                    Dim pos As Integer = 0
                    For Each ColumnName As String In s.ColumnNames
                        Try
                            Me.DataGridView1.Columns(ColumnName).DisplayIndex = s.ColumnDisplayIndex(pos)
                            Me.DataGridView1.Columns(ColumnName).Width = s.ColumnSize(pos)
                            Me.DataGridView1.Columns(ColumnName).Visible = s.ColumnVisiblility(pos)
                        Catch ex As Exception
                            If SHOW_err Then MsgBox(ex.Message & " loadSettings ex")
                        End Try
                        pos = pos + 1
                    Next

                Else
                    My.Settings.MySettingTypes = New MySettingTypes.DataGridViewColumnSetting
                    Me.SaveSettings()
                End If
            Else
                My.Settings.MySettingTypes = New MySettingTypes.DataGridViewColumnSetting
                My.Settings.SETCOL = "USTAWIONE"
                Me.SaveSettings()
            End If
        Catch ex1 As Exception
            If SHOW_err Then MsgBox(ex1.Message & " loadSettings ex")
        End Try
    End Sub
    Private Sub SaveSettings()
        Dim x As New MySettingTypes.DataGridViewColumnSetting
        Try
            For Each c As DataGridViewColumn In DataGridView1.Columns
                x.ColumnNames.Add(c.Name)
                x.ColumnDisplayIndex.Add(c.DisplayIndex)
                x.ColumnSize.Add(c.Width)
                x.ColumnVisiblility.Add(c.Visible)
                My.Settings.MySettingTypes = x
                My.Settings.Save()
            Next
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " saveSettings ex")
        End Try
    End Sub
    Private Sub ToolStripButton9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton9.Click
        'potwierdzenie zamówień
        Try
            Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
            If selectedRowCount > 0 Then
                Try
                    Using kors As New DataTable()

                        Using conB As New NpgsqlConnection(NpA)
                            conB.Open()
                            Using con As New NpgsqlCommand("select confirm_purch(@indeks,@dost_ilosc,@data_dost,@rodzaj_potw,@termin_wazn,@koor,@sum_dost)", conB)
                                con.Parameters.Add("indeks", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("dost_ilosc", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Parameters.Add("data_dost", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("rodzaj_potw", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("termin_wazn", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("koor", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("sum_dost", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Prepare()
                                'Pobierz dane do column
                                For i = 0 To selectedRowCount - 1
                                    If (DataGridView1.SelectedRows(i).Cells("TYP_zdarzenia").Value.ToString = "Brakujące ilości" Or DataGridView1.SelectedRows(i).Cells("TYP_zdarzenia").Value.ToString = "Dostawa na dzisiejsze ilości") And DataGridView1.SelectedRows(i).Cells("Status_Informacji").Value.ToString = "BRAK" Then
                                        con.Parameters(0).Value = DataGridView1.SelectedRows(i).Cells("INDEKS").Value.ToString
                                        con.Parameters(1).Value = CDbl(DataGridView1.SelectedRows(i).Cells("Wlk_dost").Value)
                                        con.Parameters(2).Value = CDate(DataGridView1.SelectedRows(i).Cells("Data_dost").Value.ToString)
                                        con.Parameters(3).Value = "POTWIERDZONE"
                                        con.Parameters(4).Value = CDate(DataGridView1.SelectedRows(i).Cells("Data_dost").Value.ToString)
                                        con.Parameters(5).Value = DataGridView1.SelectedRows(i).Cells("PLANNER_BUYER").Value.ToString
                                        con.Parameters(6).Value = CDbl(DataGridView1.SelectedRows(i).Cells("Wlk_dost").Value)
                                        Dim wys = con.ExecuteScalar()
                                    End If
                                Next i
                            End Using
                        End Using
                    End Using
                Catch ex1 As Exception
                    If SHOW_err Then MsgBox(ex1.Message)
                End Try
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select updt_dta_potw()", conB)
                        Dim wys = con.ExecuteScalar()
                    End Using
                End Using
                If Not rase_evn Then src = Save_ADO(runnow)
                'If Not rase_evn Then src = GETDATA(True)
                Refr_SETT()
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Form2_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not My.Settings.FRM2WS = FormWindowState.Minimized Then
            Me.WindowState = My.Settings.FRM2WS
            If Not My.Settings.FRM2SIZEW = 0 Then Me.Width = My.Settings.FRM2SIZEW
            If Not My.Settings.FRM2SIZEH = 0 Then Me.Height = My.Settings.FRM2SIZEH
        End If
        Application.DoEvents()
        LoadSettings()
        If Not My.Settings.FRM2WS = FormWindowState.Minimized Then
            ToolStripManager.LoadSettings(Me)
            If My.Settings.SplitCDIST1 = 0 Then Me.SplitContainer1.SplitterDistance = Me.TreeView1.PreferredSize.Width + Me.SplitContainer1.Panel1.PreferredSize.Width + Me.SplitContainer1.SplitterWidth + 50 Else SplitContainer1.SplitterDistance = My.Settings.SplitCDIST1
            If My.Settings.SplitCDIST3 = 0 Then Me.SplitContainer3.SplitterDistance = CInt(Me.SplitContainer3.Size.Height * 0.8) Else SplitContainer3.SplitterDistance = My.Settings.SplitCDIST3
            If My.Settings.SplitCDIST2 <> 0 Then Me.SplitContainer2.SplitterDistance = My.Settings.SplitCDIST2
        End If
    End Sub
    Private Sub ToolStripButton13_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton13.Click
        'Komponent nie zamawiany
        Try
            Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
            If selectedRowCount > 0 Then
                Try
                    Using kors As New DataTable()
                        Using conB As New NpgsqlConnection(NpA)
                            conB.Open()
                            Using con As New NpgsqlCommand("select confirm_purch(@indeks,@dost_ilosc,@data_dost,@rodzaj_potw,@termin_wazn,@koor,@sum_dost)", conB)
                                con.Parameters.Add("indeks", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("dost_ilosc", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Parameters.Add("data_dost", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("rodzaj_potw", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("termin_wazn", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("koor", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("sum_dost", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Prepare()
                                'Pobierz dane do column
                                For i = 0 To selectedRowCount - 1
                                    If DataGridView1.SelectedRows(i).Cells("Status_Informacji").Value.ToString = "BRAK" And DataGridView1.SelectedRows(i).Cells("Status_Informacji").Value.ToString.Length > 3 Then
                                        con.Parameters(0).Value = DataGridView1.SelectedRows(i).Cells("INDEKS").Value.ToString
                                        con.Parameters(1).Value = CDbl(0)
                                        con.Parameters(2).Value = Now.Date.AddYears(2)
                                        con.Parameters(3).Value = "NIE ZAMAWIAM"
                                        con.Parameters(4).Value = Now.Date.AddYears(2)
                                        con.Parameters(5).Value = DataGridView1.SelectedRows(i).Cells("PLANNER_BUYER").Value.ToString
                                        con.Parameters(6).Value = CDbl(0)
                                        Dim wys = con.ExecuteScalar()
                                    End If
                                Next i
                            End Using
                        End Using
                    End Using
                Catch ex1 As Exception
                    If SHOW_err Then MsgBox(ex1.Message)
                End Try
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select updt_dta_potw()", conB)
                        Dim wys = con.ExecuteScalar()
                    End Using
                End Using
                If Not rase_evn Then src = Save_ADO(runnow)
                'If Not rase_evn Then src = GETDATA(True)
                Refr_SETT()
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ToolStripButton12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton12.Click
        'Kasowanie potwierdzeń
        Try
            Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
            If selectedRowCount > 0 Then
                Try
                    Using kors As New DataTable()
                        Using conB As New NpgsqlConnection(NpA)
                            conB.Open()
                            Using con As New NpgsqlCommand("select confirm_purch(@indeks,@dost_ilosc,@data_dost,@rodzaj_potw,@termin_wazn,@koor,@sum_dost)", conB)
                                con.Parameters.Add("indeks", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("dost_ilosc", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Parameters.Add("data_dost", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("rodzaj_potw", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("termin_wazn", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("koor", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("sum_dost", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Prepare()
                                'Pobierz dane do column
                                For i = 0 To selectedRowCount - 1
                                    If DataGridView1.SelectedRows(i).Cells("Status_Informacji").Value.ToString <> "BRAK" And DataGridView1.SelectedRows(i).Cells("Status_Informacji").Value.ToString.Length > 3 Then
                                        con.Parameters(0).Value = DataGridView1.SelectedRows(i).Cells("INDEKS").Value.ToString
                                        con.Parameters(1).Value = CDbl(0)
                                        con.Parameters(2).Value = CDate(DataGridView1.SelectedRows(i).Cells("Data_dost").Value.ToString)
                                        con.Parameters(3).Value = "BRAK"
                                        con.Parameters(4).Value = Now.Date.AddYears(2)
                                        con.Parameters(5).Value = DataGridView1.SelectedRows(i).Cells("PLANNER_BUYER").Value.ToString
                                        con.Parameters(6).Value = CDbl(0)
                                        Dim wys = con.ExecuteScalar()
                                    End If
                                Next i
                            End Using
                        End Using
                    End Using
                Catch ex1 As Exception
                    If SHOW_err Then MsgBox(ex1.Message)
                End Try
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select updt_dta_potw()", conB)
                        Dim wys = con.ExecuteScalar()
                    End Using
                End Using
                If Not rase_evn Then src = Save_ADO(runnow)
                'If Not rase_evn Then src = GETDATA(True)
                Refr_SETT()
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ToolStripButton11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton11.Click
        'Komponent alter
        Try
            Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
            If selectedRowCount > 0 Then
                Try
                    Using kors As New DataTable()
                        Using conB As New NpgsqlConnection(NpA)
                            conB.Open()
                            Using con As New NpgsqlCommand("select confirm_purch(@indeks,@dost_ilosc,@data_dost,@rodzaj_potw,@termin_wazn,@koor,@sum_dost)", conB)
                                con.Parameters.Add("indeks", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("dost_ilosc", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Parameters.Add("data_dost", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("rodzaj_potw", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("termin_wazn", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("koor", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("sum_dost", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Prepare()
                                'Pobierz dane do column
                                For i = 0 To selectedRowCount - 1
                                    If LSet(DataGridView1.SelectedRows(i).Cells("Indeks").Value.ToString, 1) = "5" And (DataGridView1.SelectedRows(i).Cells("TYP_zdarzenia").Value.ToString = "Brakujące ilości" Or DataGridView1.SelectedRows(i).Cells("TYP_zdarzenia").Value.ToString = "Dostawa na dzisiejsze ilości") And DataGridView1.SelectedRows(i).Cells("Status_Informacji").Value.ToString = "BRAK" Then
                                        con.Parameters(0).Value = DataGridView1.SelectedRows(i).Cells("INDEKS").Value.ToString
                                        con.Parameters(1).Value = CDbl(DataGridView1.SelectedRows(i).Cells("Wlk_dost").Value)
                                        con.Parameters(2).Value = CDate(DataGridView1.SelectedRows(i).Cells("Data_dost").Value.ToString)
                                        con.Parameters(3).Value = "Użyj FR"
                                        con.Parameters(4).Value = CDate(DataGridView1.SelectedRows(i).Cells("Data_dost").Value.ToString)
                                        con.Parameters(5).Value = DataGridView1.SelectedRows(i).Cells("PLANNER_BUYER").Value.ToString
                                        con.Parameters(6).Value = CDbl(DataGridView1.SelectedRows(i).Cells("Wlk_dost").Value)
                                        Dim wys = con.ExecuteScalar()
                                    End If
                                Next i
                            End Using
                        End Using
                    End Using
                Catch ex1 As Exception
                    If SHOW_err Then MsgBox(ex1.Message)
                End Try
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select updt_dta_potw()", conB)
                        Dim wys = con.ExecuteScalar()
                    End Using
                End Using
                If Not rase_evn Then src = Save_ADO(runnow)
                'If Not rase_evn Then src = GETDATA(True)
                Refr_SETT()
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click
        'Czekam na info
        Try
            Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
            If selectedRowCount > 0 Then
                Try
                    Using kors As New DataTable()
                        Using conB As New NpgsqlConnection(NpA)
                            conB.Open()
                            Using con As New NpgsqlCommand("select confirm_purch(@indeks,@dost_ilosc,@data_dost,@rodzaj_potw,@termin_wazn,@koor,@sum_dost)", conB)
                                con.Parameters.Add("indeks", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("dost_ilosc", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Parameters.Add("data_dost", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("rodzaj_potw", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("termin_wazn", NpgsqlTypes.NpgsqlDbType.Date)
                                con.Parameters.Add("koor", NpgsqlTypes.NpgsqlDbType.Varchar)
                                con.Parameters.Add("sum_dost", NpgsqlTypes.NpgsqlDbType.Double)
                                con.Prepare()
                                'Pobierz dane do column
                                For i = 0 To selectedRowCount - 1
                                    If (DataGridView1.SelectedRows(i).Cells("TYP_zdarzenia").Value.ToString = "Brakujące ilości" Or DataGridView1.SelectedRows(i).Cells("TYP_zdarzenia").Value.ToString = "Dostawa na dzisiejsze ilości") And DataGridView1.SelectedRows(i).Cells("Status_Informacji").Value.ToString = "BRAK" Then
                                        con.Parameters(0).Value = DataGridView1.SelectedRows(i).Cells("INDEKS").Value.ToString
                                        con.Parameters(1).Value = CDbl(DataGridView1.SelectedRows(i).Cells("Wlk_dost").Value)
                                        con.Parameters(2).Value = CDate(DataGridView1.SelectedRows(i).Cells("Data_dost").Value.ToString)
                                        con.Parameters(3).Value = "Czekam na INFO"
                                        con.Parameters(4).Value = Now.AddDays(3)
                                        con.Parameters(5).Value = DataGridView1.SelectedRows(i).Cells("PLANNER_BUYER").Value.ToString
                                        con.Parameters(6).Value = CDbl(DataGridView1.SelectedRows(i).Cells("Wlk_dost").Value)
                                        Dim wys = con.ExecuteScalar()
                                    End If
                                Next i
                            End Using
                        End Using
                    End Using
                Catch ex As Exception
                    If SHOW_err Then MsgBox(ex.Message)
                End Try
                If Not rase_evn Then src = Save_ADO(runnow)
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select updt_dta_potw()", conB)
                        Dim wys = con.ExecuteScalar()
                    End Using
                End Using
                If Not rase_evn Then src = Save_ADO(runnow)
                'If Not rase_evn Then src = GETDATA(True)
                Refr_SETT()
            End If
        Catch ex1 As Exception
            If SHOW_err Then MsgBox(ex1.Message)
        End Try
    End Sub
    Private Sub SplitContainer3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles SplitContainer3.MouseDoubleClick
        If Me.SplitContainer3.Panel2.AutoScrollMinSize.Height = Me.SplitContainer3.SplitterDistance Then
            Me.SplitContainer3.SplitterDistance = Me.Chart2.PreferredSize.Width + Me.SplitContainer3.Panel2.PreferredSize.Width + Me.SplitContainer3.SplitterWidth + 50
        Else
            Me.SplitContainer3.SplitterDistance = Me.SplitContainer3.Panel2.AutoScrollMinSize.Width
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        If Me.SplitContainer3.SplitterDistance < CInt(Me.SplitContainer3.Size.Height * 0.8) Then
            Me.SplitContainer3.SplitterDistance = CInt(Me.SplitContainer3.Size.Height * 0.8)
        Else
            'Me.SplitContainer3.SplitterDistance = Me.Chart2.PreferredSize.Width + Me.SplitContainer3.Panel2.PreferredSize.Width + Me.SplitContainer3.SplitterWidth + 50
            Me.SplitContainer3.SplitterDistance = Me.SplitContainer3.Size.Height
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        If Me.SplitContainer3.SplitterDistance > CInt(Me.SplitContainer3.Size.Height * 0.8) Then
            Me.SplitContainer3.SplitterDistance = CInt(Me.SplitContainer3.Size.Height * 0.8)
        Else
            Me.SplitContainer3.SplitterDistance = 0
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Me.SplitContainer1.SplitterDistance = Me.SplitContainer1.Panel1.AutoScrollMinSize.Width
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Me.SplitContainer1.SplitterDistance = Me.TreeView1.PreferredSize.Width + Me.SplitContainer1.Panel1.PreferredSize.Width + Me.SplitContainer1.SplitterWidth + 50
    End Sub
    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click
        Chart1.Visible = False
        Chart3.Visible = True
    End Sub
    Private Sub Chart3_Click(sender As Object, e As EventArgs) Handles Chart3.Click
        Chart3.Visible = False
        Chart1.Visible = True
    End Sub
    Private Sub Task_2()
        Dim i, j As Integer
        Dim pb As New Progress_Bar With {
            .WindowTitle = "Odświeżam dane z ORACLE",
            .TimeOut = 60,
            .CallerThreadSet = Threading.Thread.CurrentThread
        }
        j = 1
        Thread.Sleep(2000)
        Do While DateDiff(DateInterval.Second, start_cal, Now) < 150
            If Not rase_evn Then
                With pb
                    i = i + 1
                    If Now.Millisecond < 100 And Len(PROC) > 4 Then
                        .PartialProgressText = PROC
                        .PartialProgressValue = 1
                    Else
                        .PartialProgressText = PROC
                        .PartialProgressValue = ((Now.Millisecond / 999) * 100)
                    End If
                    .OverallProgressText = "Koniec..."
                    If i > 100 Then i = 100
                    .OverallProgressValue = i
                End With
            Else
                With pb
                    If Now.Millisecond < 100 And Len(PROC) > 4 Then
                        .PartialProgressText = PROC
                        .PartialProgressValue = 1
                    Else
                        .PartialProgressText = PROC
                        .PartialProgressValue = ((Now.Millisecond / 999) * 100)
                    End If
                    .OverallProgressText = "Koniec obliczeń za około " & 150 - DateDiff(DateInterval.Second, start_cal, Now) & " sek"
                    i = CInt(((DateDiff(DateInterval.Second, start_cal, Now) / 150) * 100))
                    .OverallProgressValue = IIf(i > 100, 100, IIf(i < 0, 0, i))
                End With
            End If
            Threading.Thread.Sleep(j)
        Loop
        With pb
            .OverallProgressText = "Blisko końca..."
            .OverallProgressValue = ((DateDiff(DateInterval.Second, start_cal, Now) / 150) * 100)
            .Dispose()
        End With
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Indicators.Show()
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Clipboard.SetDataObject(DataGridView1.GetClipboardContent())
        DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim chan As Boolean = False
        If ToolStripComboBox1.Text <> "" Then
            Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
            If selectedRowCount > 0 Then
                For i = 0 To selectedRowCount - 1
                    If IsDBNull(DataGridView1.SelectedRows(i).Cells("Informacja").Value) And DataGridView1.SelectedRows(i).Cells("status_informacji").Value <> "BRAK" Then
                        DataGridView1.SelectedRows(i).Cells("Informacja").Value = Environment.MachineName & "::" & ToolStripComboBox1.Text
                        chan = True
                        Try
                            Using kors As New DataTable()
                                Using conB As New NpgsqlConnection(NpA)
                                    conB.Open()
                                    Using con As New NpgsqlCommand("select addinfo_purch(@indeks,@data_dost,@info,@dost_ilosc)", conB)
                                        con.Parameters.Add("indeks", NpgsqlTypes.NpgsqlDbType.Varchar)
                                        con.Parameters.Add("data_dost", NpgsqlTypes.NpgsqlDbType.Date)
                                        con.Parameters.Add("info", NpgsqlTypes.NpgsqlDbType.Varchar)
                                        con.Parameters.Add("dost_ilosc", NpgsqlTypes.NpgsqlDbType.Double)
                                        con.Prepare()
                                        con.Parameters(0).Value = DataGridView1.SelectedRows(i).Cells("indeks").Value.ToString()
                                        con.Parameters(1).Value = Convert.ToDateTime(DataGridView1.SelectedRows(i).Cells("data_dost").Value)
                                        con.Parameters(2).Value = Environment.MachineName & "::" & ToolStripComboBox1.Text
                                        con.Parameters(3).Value = Convert.ToDouble(DataGridView1.SelectedRows(i).Cells("wlk_dost").Value)
                                        Dim wys = con.ExecuteScalar()
                                    End Using
                                End Using
                            End Using
                        Catch ex As Exception
                            If SHOW_err Then MsgBox(ex.Message)
                        End Try

                    End If
                Next i
                If chan Then
                    Using conB As New NpgsqlConnection(NpA)
                        conB.Open()
                        Using con As New NpgsqlCommand("select updt_dta_potw()", conB)
                            Dim wys = con.ExecuteScalar()
                        End Using
                    End Using
                    If Not rase_evn Then src = Save_ADO(runnow)
                    'If Not rase_evn Then src = GETDATA(True)
                    Refr_SETT()
                End If
            End If
        End If
    End Sub

    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles ToolStripButton14.Click
        Dim chan As Boolean = False
        Dim selectedRowCount As Integer =
                   DataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected)
        If selectedRowCount > 0 Then
            For i = 0 To selectedRowCount - 1
                If Not IsDBNull(DataGridView1.SelectedRows(i).Cells("Informacja").Value) Then
                    chan = True
                    Try
                        Using kors As New DataTable()
                            Using conB As New NpgsqlConnection(NpA)
                                conB.Open()
                                Using con As New NpgsqlCommand("select addinfo_purch(@indeks,@data_dost,null,@dost_ilosc)", conB)
                                    con.Parameters.Add("indeks", NpgsqlTypes.NpgsqlDbType.Varchar)
                                    con.Parameters.Add("data_dost", NpgsqlTypes.NpgsqlDbType.Date)
                                    con.Parameters.Add("dost_ilosc", NpgsqlTypes.NpgsqlDbType.Double)
                                    con.Prepare()
                                    con.Parameters(0).Value = DataGridView1.SelectedRows(i).Cells("indeks").Value.ToString()
                                    con.Parameters(1).Value = Convert.ToDateTime(DataGridView1.SelectedRows(i).Cells("data_dost").Value)
                                    con.Parameters(2).Value = Convert.ToDouble(DataGridView1.SelectedRows(i).Cells("wlk_dost").Value)
                                    Dim wys = con.ExecuteScalar()
                                End Using
                            End Using
                        End Using
                    Catch ex As Exception
                        If SHOW_err Then MsgBox(ex.Message)
                    End Try
                End If
            Next i
            If chan Then
                Using conB As New NpgsqlConnection(NpA)
                    conB.Open()
                    Using con As New NpgsqlCommand("select updt_dta_potw()", conB)
                        Dim wys = con.ExecuteScalar()
                    End Using
                End Using
                If Not rase_evn Then src = Save_ADO(runnow)
                'If Not rase_evn Then src = GETDATA(True)
                Refr_SETT()
            End If
        End If
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Process.Start("microsoft-edge:http://10.0.1.29:3000")
    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim nod, nod1, nod2 As TreeNode
        Dim chk As Boolean = False
        tree_busy = True
        If Not e.Node.Parent Is Nothing Then
            If e.Node.Parent.Name = "DATA_BRAKU" Or e.Node.Parent.Name = "DATA_dost" Then
                For Each nod In e.Node.Parent.Nodes
                    If nod IsNot e.Node And nod.Checked Then
                        nod.Checked = False
                    End If
                Next
            End If
            If e.Node.Parent.Name = "DATES" Then
                If e.Node.Checked Then
                    For Each nod In e.Node.Nodes
                        If nod.Checked Then
                            chk = True
                        End If
                    Next
                End If
                If Not chk Then e.Node.Collapse()
            End If
            If e.Node.Parent.Name = "Status_Informacji" Or e.Node.Parent.Name = "PLANNER_BUYER" Or e.Node.Parent.Name = "EVNT" Or e.Node.Parent.Name = "CHK" Then
                If e.Node.Parent.Parent IsNot Nothing Then
                    If e.Node.Checked Then
                        For Each nod In TreeView1.Nodes("EVNT").Nodes
                            If nod.Checked Then
                                nod.Checked = False
                            End If
                        Next
                        For Each nod In TreeView1.Nodes("CHK").Nodes
                            If nod.Checked Then
                                nod.Checked = False
                            End If
                        Next
                    End If
                Else
                    If e.Node.Checked Then
                        For Each nod In TreeView1.Nodes("PLANNER_BUYER").Nodes
                            If nod.Checked Then
                                For Each nod1 In nod.Nodes
                                    If nod1.Checked Then
                                        For Each nod2 In nod1.Nodes
                                            If nod2.Checked Then
                                                nod2.Checked = False
                                            End If
                                        Next
                                        nod1.Checked = False
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
        End If
        tree_busy = False
    End Sub
    Private Sub SuspendDrawing(Control As Control)
        NativeMethods.SendMessage(Control.Handle, WM_SETREDRAW, False, 0)
    End Sub
    Private Sub ResumeDrawing(Control As Control)
        NativeMethods.SendMessage(Control.Handle, WM_SETREDRAW, True, 0)
    End Sub
End Class
