Imports Npgsql
Public Class Dialog1
    Public partno As String
    Public datdost As Date
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click
        Using kors As New DataTable()
            Using conB As New NpgsqlConnection(NpA)
                conB.Open()
                DataGridView1.DataSource = ""
                If ListView1.Items(0).Selected = True Then
                    Using con As New NpgsqlCommand("select b.dop,b.zlec,b.prod_qty,b.koor,b.order_no,b.line_no,b.rel_no,b.part_no,b.descr,b.prom_week,b.ship_date,b.country,b.prod_date,b.indeks,b.opis,b.qty_demand from (select id,indeks,data_dost from public.data) a left join (select * from braki) b  on b.indeks||'_'||b.data_dost=a.indeks||'_'||a.data_dost where b.ordid is not null and a.indeks=@partno and a.data_dost=@datdost", conB)
                        con.Parameters.Add("partno", NpgsqlTypes.NpgsqlDbType.Varchar).Value = partno
                        con.Parameters.Add("datdost", NpgsqlTypes.NpgsqlDbType.Date).Value = datdost
                        con.Prepare()
                        Using rs As NpgsqlDataReader = con.ExecuteReader
                            kors.Load(rs)
                        End Using
                        If kors.Rows.Count > 0 Then
                            With Me
                                .Width = 857
                                .Height = 380
                            End With
                            DataGridView1.DataSource = kors
                        Else
                            With Me
                                .Width = 288
                                .Height = 380
                            End With
                        End If
                    End Using
                Else
                    If ListView1.Items(1).Selected = True Then
                        Using con As New NpgsqlCommand("select b.dop,b.zlec,b.prod_qty,b.indeks,b.opis,b.qty_demand,b.data_dost,b.date_reuired from (select a.id,a.indeks,a.data_dost,b.ordid,b.prod_date from (select id,indeks,data_dost from public.data) a left join (select ordid,indeks,data_dost,prod_date from braki) b on b.indeks||'_'||b.data_dost=a.indeks||'_'||a.data_dost where b.ordid is not null) a,(select * from braki )b where b.ordid=a.ordid and b.prod_date>a.prod_date and b.indeks!=a.indeks and (b.status_informacji='BRAK' or b.status_informacji is null) and a.indeks=@partno and a.data_dost=@datdost group by b.dop,b.zlec,b.prod_qty,b.indeks,b.opis,b.qty_demand,b.data_dost,b.date_reuired", conB)
                            con.Parameters.Add("partno", NpgsqlTypes.NpgsqlDbType.Varchar).Value = partno
                            con.Parameters.Add("datdost", NpgsqlTypes.NpgsqlDbType.Date).Value = datdost
                            con.Prepare()
                            Using rs As NpgsqlDataReader = con.ExecuteReader
                                kors.Load(rs)
                            End Using
                            If kors.Rows.Count > 0 Then
                                With Me
                                    .Width = 857
                                    .Height = 380
                                End With
                                DataGridView1.DataSource = kors
                            Else
                                With Me
                                    .Width = 288
                                    .Height = 380
                                End With
                            End If
                        End Using
                    Else
                        If ListView1.Items(2).Selected = True Then
                            Using con As New NpgsqlCommand("select a.dop,a.zlec,a.prod_qty,a.indeks,a.opis,a.qty_demand,a.koor,a.order_no,a.line_no,a.rel_no,a.part_no,a.descr,a.prom_week,a.ship_date,a.country,a.prod_date from braki a,cust_ord a_1_1 LEFT JOIN ( SELECT cust_ord_1.order_no FROM cust_ord cust_ord_1 WHERE cust_ord_1.state_conf::text = 'Wydrukow.'::text AND cust_ord_1.last_mail_conf IS NOT NULL GROUP BY cust_ord_1.order_no) c ON c.order_no::text = a_1_1.order_no::text WHERE (a_1_1.state_conf::text = 'Nie wydruk.'::text OR a_1_1.last_mail_conf IS NULL) AND is_refer(a_1_1.addr1) = true AND substring(a_1_1.order_no::text, 1, 1) = 'S'::text AND (a_1_1.cust_order_state::text <> ALL (ARRAY['Częściowo dostarczone'::character varying::text, 'Zablok. kredyt'::character varying::text, 'Zaplanowane'::character varying::text])) AND (substring(a_1_1.part_no::text, 1, 3) <> ALL (ARRAY['633'::text, '628'::text, '1K1'::text, '1U2'::text, '632'::text])) AND (c.order_no IS NOT NULL AND a_1_1.dop_connection_db::text <> 'MAN'::text OR c.order_no IS NULL) and a.cust_id=a_1_1.id and a.indeks=@partno and a.data_dost=@datdost", conB)
                                con.Parameters.Add("partno", NpgsqlTypes.NpgsqlDbType.Varchar).Value = partno
                                con.Parameters.Add("datdost", NpgsqlTypes.NpgsqlDbType.Date).Value = datdost
                                con.Prepare()
                                Using rs As NpgsqlDataReader = con.ExecuteReader
                                    kors.Load(rs)
                                End Using
                                If kors.Rows.Count > 0 Then
                                    With Me
                                        .Width = 857
                                        .Height = 380
                                    End With
                                    DataGridView1.DataSource = kors
                                Else
                                    With Me
                                        .Width = 288
                                        .Height = 380
                                    End With
                                End If
                            End Using
                        End If
                    End If
                End If
            End Using
        End Using
    End Sub
End Class
