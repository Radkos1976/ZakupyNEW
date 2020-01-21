Imports Npgsql
Public Class LoginForm1

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim allowed As Boolean = True
        Try
            SaveSetting("ZakupyNEW", "Main", "UID", UsernameTextBox.Text)
            SaveSetting("ZakupyNEW", "Main", "PSW", PasswordTextBox.Text)
            Dim qw As New NpgsqlConnectionStringBuilder("Host = 10.0.1.29; Port = 5432; ApplicationName = CLIENT; Username = " & GetSetting("ZakupyNEW", "Main", "UID", "NoPostgresql") & "; Password = " & GetSetting("ZakupyNEW", "Main", "PSW") & "; Database = zakupy")
            Dim NpA As String = qw.ToString
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
            signals.Show()
            ' Me.Hide()
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        UsernameTextBox.Text = GetSetting("ZakupyNEW", "Main", "UID", "NoPostgresql")
        PasswordTextBox.Text = GetSetting("ZakupyNEW", "Main", "PSW", "none")
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SaveSetting("ZakupyNEW", "Main", "UID", "NoPostgresql")
        SaveSetting("ZakupyNEW", "Main", "PSW", "none")
        signals.Show()
        'Me.Hide()
    End Sub
End Class
