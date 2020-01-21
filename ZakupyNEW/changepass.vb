Imports Npgsql
Public Class changepass

    ' TODO: Wstaw kod, aby wykonaæ niestandardowe uwierzytelnianie przy u¿yciu podanej nazwy u¿ytkownika i has³a 
    ' (Zobacz https://go.microsoft.com/fwlink/?LinkId=35339).  
    ' Niestandardowy podmiot zabezpieczeñ mo¿e nastêpnie byæ do³¹czony do podmiotu zabezpieczeñ bie¿¹cego w¹tku g³ównego w nastêpuj¹cy sposób: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' gdzie klasa CustomPrincipal jest implementacj¹ interfejsu IPrincipal u¿ywan¹ do przeprowadzania uwierzytelniania. 
    ' Nastêpnie obiekt My.User zwróci informacjê o to¿samoœci hermetyzowan¹ w obiekcie CustomPrincipal
    ' takie jak nazwa u¿ytkownika, nazwa wyœwietlana, itp.
    Dim count_pol As Int32 = 0
    Dim txt1, txt2, txt3, txt4 As Int32

    Private Sub PasswordTextBox_TextChanged(sender As Object, e As EventArgs) Handles PasswordTextBox.TextChanged
        If Len(PasswordTextBox.Text) > 0 Then txt2 = 1 Else txt2 = 0
        chk_ok()
    End Sub
    Private Sub UsernameTextBox_TextChanged(sender As Object, e As EventArgs) Handles UsernameTextBox.TextChanged
        If Len(UsernameTextBox.Text) > 0 Then txt1 = 1 Else txt1 = 0
        chk_ok()
    End Sub
    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated
        Dim txt1 As String = TextBox1.Text.ToString
        If Len(TextBox1.Text) > 0 And txt1 = TextBox2.Text.ToString Then txt3 = 1 Else txt3 = 0
        chk_ok()
    End Sub

    Private Sub TextBox2_Validated(sender As Object, e As EventArgs) Handles TextBox2.Validated
        Dim txt1 As String = TextBox2.Text.ToString
        If Len(TextBox2.Text) > 0 And txt1 = TextBox1.Text.ToString Then txt4 = 1 Else txt4 = 0
        Chk_ok()
    End Sub

    Private Sub LoginForm2_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        Chk_ok()
    End Sub

    Private Sub LoginForm2_TextChanged(sender As Object, e As EventArgs) Handles MyBase.TextChanged
        Chk_ok()
    End Sub

    Private Sub Chk_ok()
        count_pol = txt1 + txt2 + txt3 + txt4
        If count_pol = 4 Then OK.Enabled = True Else OK.Enabled = False
        Me.Validate()
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim qw As New NpgsqlConnectionStringBuilder("Host = 10.0.1.29; Port = 5432; ApplicationName = CLIENT; Username = " & UsernameTextBox.Text & "; Password = " & PasswordTextBox.Text & "; Database = zakupy")
        Dim NpA As String = qw.ToString
        Dim allowed As Boolean = True
        Try
            Using conB As New NpgsqlConnection(NpA)
                conB.Open()
                Using con As New NpgsqlCommand("select last_modify from public.datatbles where table_name='data'", conB)
                    Dim last_dta = con.ExecuteScalar()
                End Using
            End Using
        Catch ex1 As Exception
            MsgBox("Problem z logowaniem:" & ex1.Message)
            allowed = False
        End Try
        If allowed Then
            Using conB As New NpgsqlConnection(NpA)
                conB.Open()
                Using con As New NpgsqlCommand("ALTER USER @username WITH PASSWORD @psw", conB)
                    con.Parameters.Add("username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = UsernameTextBox.Text.ToString
                    con.Parameters.Add("psw", NpgsqlTypes.NpgsqlDbType.Varchar).Value = TextBox1.Text.ToString
                    con.Prepare()
                    con.ExecuteNonQuery()
                End Using
            End Using
            SaveSetting("ZakupyNEW", "Main", "UID", UsernameTextBox.Text)
            SaveSetting("ZakupyNEW", "Main", "PSW", TextBox1.Text)
            Dim qws As New NpgsqlConnectionStringBuilder("Host = 10.0.1.29; Port = 5432; ApplicationName = CLIENT; Username = " & GetSetting("ZakupyNEW", "Main", "UID", "NoPostgresql") & "; Password = " & GetSetting("ZakupyNEW", "Main", "PSW") & "; Database = zakupy")
            NpA = qws.ToString
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
