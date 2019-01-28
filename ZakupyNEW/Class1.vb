
'  *****************************************
'  ** DataGridViewColumnSelector ver 1.0  **
'  **                                     **
'  ** Author : Vincenzo Rossi             **
'  ** Country: Naples, Italy              **
'  ** Year   : 2008                       **
'  ** Mail   : redmaster@tiscali.it       **
'  **                                     **
'  ** Released under                      **
'  **   The Code Project Open License     **
'  **                                     **
'  **   Please do not remove this header, **
'  **   I will be grateful if you mention **
'  **   me in your credits. Thank you     **
'  **                                     **
'  *****************************************
Imports Npgsql

Namespace DGVColumnSelector
    ''' <summary>
    ''' Add column show/hide capability to a DataGridView. When user right-clicks 
    ''' the cell origin a popup, containing a list of checkbox and column names, is
    ''' shown. 
    ''' </summary>
    Class DataGridViewColumnSelector : Implements IDisposable

        ' the DataGridView to which the DataGridViewColumnSelector is attached
        Private mDataGridView As DataGridView = Nothing
        ' a CheckedListBox containing the column header text and checkboxes
        Private mCheckedListBox As CheckedListBox
        ' a ToolStripDropDown object used to show the popup
        Private mPopup As ToolStripDropDown

        ''' <summary>
        ''' The max height of the popup
        ''' </summary>
        Public MaxHeight As Integer = 300
        ''' <summary>
        ''' The width of the popup
        ''' </summary>
        Public Width As Integer = 200

        ''' <summary>
        ''' Gets or sets the DataGridView to which the DataGridViewColumnSelector is attached
        ''' </summary>
        Public Property DataGridView() As DataGridView
            Get
                Return mDataGridView
            End Get
            Set(value As DataGridView)
                ' If any, remove handler from current DataGridView 
                If mDataGridView IsNot Nothing Then
                    RemoveHandler mDataGridView.CellMouseClick, AddressOf Me.MDataGridView_CellMouseClick
                End If
                ' Set the new DataGridView
                mDataGridView = value
                ' Attach CellMouseClick handler to DataGridView
                If mDataGridView IsNot Nothing Then
                    AddHandler mDataGridView.CellMouseClick, AddressOf Me.MDataGridView_CellMouseClick
                End If
            End Set
        End Property

        ' When user right-clicks the cell origin, it clears and fill the CheckedListBox with
        ' columns header text. Then it shows the popup. 
        ' In this way the CheckedListBox items are always refreshed to reflect changes occurred in 
        ' DataGridView columns (column additions or name changes and so on).
        Private Sub MDataGridView_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)

            If e.Button = MouseButtons.Right AndAlso e.RowIndex = -1 AndAlso e.ColumnIndex = -1 Then
                mCheckedListBox.Items.Clear()
                For Each c As DataGridViewColumn In mDataGridView.Columns
                    mCheckedListBox.Items.Add(c.HeaderText, c.Visible)
                Next
                Dim PreferredHeight As Integer = (mCheckedListBox.Items.Count * 16) + 7
                mCheckedListBox.Height = If((PreferredHeight < MaxHeight), PreferredHeight, MaxHeight)
                mCheckedListBox.Width = Me.Width
                mPopup.Show(mDataGridView.PointToScreen(New Point(e.X, e.Y)))
            ElseIf e.Button = MouseButtons.Right AndAlso e.RowIndex <> -1 AndAlso e.ColumnIndex = -1 Then
                'Dialog1.Tag = e.RowIndex
                'Dialog1.ShowDialog()
                Dim aza = New Inform(mDataGridView, e.RowIndex)
            End If
        End Sub

        ' The constructor creates an instance of CheckedListBox and ToolStripDropDown.
        ' the CheckedListBox is hosted by ToolStripControlHost, which in turn is
        ' added to ToolStripDropDown.
        Public Sub New()
            mCheckedListBox = New CheckedListBox With {
                .CheckOnClick = True
            }
            AddHandler mCheckedListBox.ItemCheck, AddressOf Me.MCheckedListBox_ItemCheck

            Dim mControlHost As New ToolStripControlHost(mCheckedListBox) With {
                .Padding = Padding.Empty,
                .Margin = Padding.Empty,
                .AutoSize = False
            }

            mPopup = New ToolStripDropDown With {
                .Padding = Padding.Empty
            }
            mPopup.Items.Add(mControlHost)
        End Sub

        Public Sub New(dgv As DataGridView)
            Me.New()
            Me.DataGridView = dgv
        End Sub

        ' When user checks / unchecks a checkbox, the related column visibility is 
        ' switched.
        Private Sub MCheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
            mDataGridView.Columns(e.Index).Visible = (e.NewValue = CheckState.Checked)
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' Aby wykryć nadmiarowe wywołania

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    mCheckedListBox.Dispose()
                    mPopup.Dispose()
                    ' TODO: wyczyść stan zarządzany (obiekty zarządzane).
                End If

                ' TODO: zwolnij niezarządzane zasoby (niezarządzane obiekty) i przesłoń poniższą metodę Finalize().
                ' TODO: ustaw wartość null dla dużych pól.
            End If
            disposedValue = True
        End Sub

        ' TODO: przesłoń metodę Finalize() tylko w sytuacji, gdy powyższa metoda Dispose(disposing As Boolean) ma kod umożliwiający zwolnienie niezarządzanych zasobów.
        'Protected Overrides Sub Finalize()
        '    ' Nie zmieniaj tego kodu. Umieść kod porządkujący w powyższej funkcji Dispose(disposing As Boolean).
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' Język Visual Basic dodaje ten kod, aby prawidłowo zaimplementować wzorzec rozporządzający.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Nie zmieniaj tego kodu. Umieść kod porządkujący w powyższej funkcji Dispose(disposing As Boolean).
            Dispose(True)
            ' TODO: usuń komentarz z poniższego wiersza, jeśli powyższa metoda Finalize() została przesłonięta.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Namespace
Class Inform
    Public i_datadost As Date
    Public i_partno As String
    Public i_typ As String
    Public status As String
    Public Sub New(i_crtlnam As Object, i_rwind As Long)
        If i_crtlnam.FindForm.Name = "Form2" Then
            i_partno = i_crtlnam.Rows(i_rwind).Cells("Indeks").Value.ToString
            i_datadost = i_crtlnam.Rows(i_rwind).Cells("data_dost").Value
            i_typ = i_crtlnam.Rows(i_rwind).Cells("typ_zdarzenia").Value
            status = i_crtlnam.Rows(i_rwind).Cells("status_informacji").Value.ToString
        Else
            i_partno = i_crtlnam.Rows(i_rwind).Cells("part_no").Value.ToString
            i_datadost = i_crtlnam.Rows(i_rwind).Cells("work_day").Value
            i_typ = "none"
            status = ""
        End If
        Dialog1.partno = i_partno
        Dialog1.datdost = i_datadost
        Dialog1.GroupBox1.Text = "Indeks : " & i_partno
        Dialog1.Width = 288
        Dialog1.Height = 380
        Dialog1.Label1.Text = i_crtlnam.Rows(i_rwind).Cells("opis").Value.ToString
        Dialog1.Label2.Text = "Data dostawy : " & i_datadost.ToString
        Dialog1.Label5.Text = i_typ
        If i_typ = "Dostawa na dzisiejsze ilości" Then
            Dialog1.Label3.Text = "Możliwe działania :" & Chr(10) & "- Sygnał informacyjny" & Chr(10) & "- Realne przyśpieszenie dostawy" & Chr(10) & "- Modyfikacja daty dostawy w IFS"
        ElseIf i_typ = "Brak zamówień zakupu" Then
            Dialog1.Label3.Text = "Możliwe działania :" & Chr(10) & "- Utworzenie zamówienia w IFS" & Chr(10) & "- Określenie komponentu jako wycofany i " & Chr(10) & " zmiana leadtime'u na komponencie"
        ElseIf i_typ = "Braki w gwarantowanej dacie" Then
            Dialog1.Label3.Text = "Możliwe działania :" & Chr(10) & "- Utworzenie zamówienia w IFS" & Chr(10) & "- Określenie komponentu jako wycofany"
        ElseIf i_typ = "Dzisiejsza dostawa" Then
            Dialog1.Label3.Text = "Możliwe działania :" & Chr(10) & "- Sygnał informacyjny" & Chr(10) & "- Modyfikacja dostawy w IFS"
        ElseIf i_typ = "Opóźniona dostawa" Then
            Dialog1.Label3.Text = "Możliwe działania :" & Chr(10) & "- Sygnał informacyjny" & Chr(10) & "- Urealnienie daty dostawy w IFS"
        ElseIf i_typ = "Brakujące ilości" Then
            Dialog1.Label3.Text = "Możliwe działania :" & Chr(10) & "- Uruchomienie komunikatu" & Chr(10) & "- Modyfikacja dostawy w IFS"
        Else
            Dialog1.Label3.Text = ""
        End If
        Dialog1.Label4.Text = "Status:" & status
        Using kors As New DataTable()
            Using zamK As New DataTable()
                Using confK As New DataTable()
                    Using conB As New NpgsqlConnection(NpA)
                        conB.Open()
                        Using con As New NpgsqlCommand("select b.* from (select a.id,a.indeks,a.data_dost,b.ordid,b.prod_date from (select id,indeks,data_dost from public.data) a left join (select ordid,indeks,data_dost,prod_date from braki) b on b.indeks||'_'||b.data_dost=a.indeks||'_'||a.data_dost where b.ordid is not null) a,(select * from braki )b where b.ordid=a.ordid and b.prod_date>a.prod_date and b.indeks!=a.indeks and (b.status_informacji='BRAK' or b.status_informacji is null) and a.indeks=@i_partno and a.data_dost=@i_datadost", conB)
                            con.Parameters.Add("i_partno", NpgsqlTypes.NpgsqlDbType.Varchar).Value = i_partno
                            con.Parameters.Add("i_datadost", NpgsqlTypes.NpgsqlDbType.Date).Value = i_datadost
                            con.Prepare()
                            Using rs As NpgsqlDataReader = con.ExecuteReader
                                kors.Load(rs)
                            End Using
                        End Using
                        Using con As New NpgsqlCommand("select b.dop,b.zlec,b.prod_qty,b.indeks,b.opis,b.qty_demand,b.data_dost,b.date_reuired from (select id,indeks,data_dost from public.data) a left join (select * from braki) b  on b.indeks||'_'||b.data_dost=a.indeks||'_'||a.data_dost where b.ordid is not null and a.indeks=@i_partno and a.data_dost=@i_datadost group by b.dop,b.zlec,b.prod_qty,b.indeks,b.opis,b.qty_demand,b.data_dost,b.date_reuired", conB)
                            'i_partno & "' and a.data_dost='" & i_datadost.ToString()
                            con.Parameters.Add("i_partno", NpgsqlTypes.NpgsqlDbType.Varchar).Value = i_partno
                            con.Parameters.Add("i_datadost", NpgsqlTypes.NpgsqlDbType.Date).Value = i_datadost
                            con.Prepare()
                            Using rs As NpgsqlDataReader = con.ExecuteReader
                                zamK.Load(rs)
                            End Using
                        End Using
                        Using con As New NpgsqlCommand("select a.* from braki a,cust_ord a_1_1 LEFT JOIN ( SELECT cust_ord_1.order_no FROM cust_ord cust_ord_1 WHERE cust_ord_1.state_conf::text = 'Wydrukow.'::text AND cust_ord_1.last_mail_conf IS NOT NULL GROUP BY cust_ord_1.order_no) c ON c.order_no::text = a_1_1.order_no::text WHERE (a_1_1.state_conf::text = 'Nie wydruk.'::text OR a_1_1.last_mail_conf IS NULL) AND is_refer(a_1_1.addr1) = true AND substring(a_1_1.order_no::text, 1, 1) = 'S'::text AND (a_1_1.cust_order_state::text <> ALL (ARRAY['Częściowo dostarczone'::character varying::text, 'Zablok. kredyt'::character varying::text, 'Zaplanowane'::character varying::text])) AND (substring(a_1_1.part_no::text, 1, 3) <> ALL (ARRAY['633'::text, '628'::text, '1K1'::text, '1U2'::text, '632'::text])) AND (c.order_no IS NOT NULL AND a_1_1.dop_connection_db::text <> 'MAN'::text OR c.order_no IS NULL) and a.cust_id=a_1_1.id and a.indeks=@i_partno and a.data_dost=@i_datadost", conB)
                            con.Parameters.Add("i_partno", NpgsqlTypes.NpgsqlDbType.Varchar).Value = i_partno
                            con.Parameters.Add("i_datadost", NpgsqlTypes.NpgsqlDbType.Date).Value = i_datadost
                            con.Prepare()
                            Using rs As NpgsqlDataReader = con.ExecuteReader
                                confK.Load(rs)
                            End Using
                        End Using
                        Dialog1.ListView1.Items(0).Text = "Lista zamówień/zleceń dotyczących sygnału (" & zamK.Rows.Count() & ")"
                        Dialog1.ListView1.Items(1).Text = "Sygnały w konflikcie (" & kors.Rows.Count() & ")"
                        Dialog1.ListView1.Items(2).Text = "Zamówienia nie potwierdzone związane z sygnałem (" & confK.Rows.Count() & ")"
                    End Using
                End Using
            End Using
        End Using
        Dialog1.ShowDialog()
    End Sub
End Class


