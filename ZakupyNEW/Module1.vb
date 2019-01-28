Imports System.IO
Imports Npgsql
Module Module1
    Public PROC As String
    Public logon_complete = False
    Public start_cal As Date = Now
    Public driveD As Boolean
    Public Const SHOW_err As Boolean = False
    Dim qw As New NpgsqlConnectionStringBuilder("Host = 10.0.1.29; Port = 5432; ApplicationName = CLIENT; Username = " & GetSetting("ZakupyNEW", "Main", "UID", "NoPostgresql") & "; Password = " & GetSetting("ZakupyNEW", "Main", "PSW") & "; Database = zakupy")
    Public NpA As String = qw.ToString
    Public Refr_date As Date
    Public Function CRE_Soures(Optional ByVal run_now As Boolean = False) As String
        Dim txt As String = "WAIT"
        Try
            txt = CHCK_RKSETS()
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message.ToString & vbNewLine.ToString & ex.Source.ToString & vbTab & ex.Data.ToString & vbNewLine & ex.TargetSite.ToString & " CRE_Soures ex")
        End Try
        Return txt
    End Function
    Public Function Save_ADO(Optional ByVal run_now As Boolean = False) As DataTable
        Using source As New DataTable()
            Using conB As New NpgsqlConnection(NpA)
                conB.Open()
                Using con As New NpgsqlCommand("select * from public.zak_dat", conB)
                    Using rs As NpgsqlDataReader = con.ExecuteReader
                        source.Load(rs)
                    End Using
                End Using
            End Using
            Return source
        End Using
    End Function
    Private Function CHCK_RKSETS() As String
        Dim txt As String = ""
        Dim last_dta As Date
        Try
            Using conB As New NpgsqlConnection(NpA)
                conB.Open()
                Using con As New NpgsqlCommand("select last_modify from public.datatbles where table_name='data'", conB)
                    last_dta = con.ExecuteScalar()
                End Using

                If last_dta <> Refr_date Then
                    txt = "GETDATA"
                Else
                    txt = "WAIT"
                End If
            End Using
            Refr_date = last_dta
            CHCK_RKSETS = txt '& "_" & last_dta.ToString
        Catch
            CHCK_RKSETS = "WAIT"
        End Try
    End Function
    Public Sub QWERTY()
        Try
            If Directory.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES") Then
                ' If File.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\RUN.vbs") Then
                ' Dim startInfo As New ProcessStartInfo(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\RUN.vbs")
                'startInfo.WindowStyle = ProcessWindowStyle.Hidden
                'Process.Start(startInfo)
                'End If
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " " & "QWERTY ex")
        End Try
    End Sub
    Public Sub YTREWQ()
        Try
            If Directory.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES") Then
                ' If File.Exists(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\RUN1.vbs") Then
                ' Dim startInfo As New ProcessStartInfo(IIf(driveD, "D", "C") & ":\ADOXML_SOURCES\RUN.vbs")
                'startInfo.WindowStyle = ProcessWindowStyle.Hidden
                'Process.Start(startInfo)
                'End If
            End If
        Catch ex As Exception
            If SHOW_err Then MsgBox(ex.Message & " " & "QWERTY ex")
        End Try
    End Sub
End Module
