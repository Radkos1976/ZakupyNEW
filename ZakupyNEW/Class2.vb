Option Explicit On

Public Class Progress_Bar
    Implements IDisposable

    Private pb As New Progress_Bar_Window
    Private updateTicker As New Timer

    ''' <summary>
    ''' The progress bar's window title.
    ''' </summary>
    Public Property WindowTitle As String
        Set(ByVal value As String)
            pb.windowTitle = value
        End Set
        Get
            Return pb.windowTitle
        End Get
    End Property

    ''' <summary>
    ''' The percentage value for the progress of overall the task.
    ''' </summary>
    Public Property OverallProgressValue As Int32
        Set(ByVal value As Int32)
            pb.overallProgressValue = value
        End Set
        Get
            Return pb.overallProgressValue
        End Get
    End Property

    ''' <summary>
    ''' The textual description for the progress of overall the task.
    ''' </summary>
    Public Property OverallProgressText As String
        Set(ByVal value As String)
            pb.overallProgressText = value
        End Set
        Get
            Return pb.overallProgressText
        End Get
    End Property

    ''' <summary>
    ''' The percentage value for the progress of partial the task. If omitted, the partial progress bar isn't shown.
    ''' </summary>
    Public Property PartialProgressValue As Int32
        Set(ByVal value As Int32)
            pb.partialProgressValue = value
        End Set
        Get
            Return pb.partialProgressValue
        End Get
    End Property

    ''' <summary>
    ''' The textual description for the progress of the partial task.
    ''' </summary>
    Public Property PartialProgressText As String
        Set(ByVal value As String)
            pb.partialProgressText = value
        End Set
        Get
            Return pb.partialProgressText
        End Get
    End Property

    ''' <summary>
    ''' Optional. The thread that called this class/thread. Use this if you want the progress bar window to close itself when the caller thread ends.
    ''' </summary>
    Public WriteOnly Property CallerThreadSet As Threading.Thread
        'Assigns only once
        Set(ByVal value As Threading.Thread)
            If pb.callerThread Is Nothing Then pb.callerThread = value
        End Set
    End Property

    ''' <summary>
    ''' The thread that called this class/thread.
    ''' </summary>
    Public ReadOnly Property CallerThreadGet As Threading.Thread
        Get
            Return pb.callerThread
        End Get
    End Property

    ''' <summary>
    ''' Optional, default = 60. Time out in seconds without receiving any progress updates. User is asked whether or not to close the window.
    ''' </summary>
    Public Property TimeOut As Int32
        Set(ByVal value As Int32)
            pb.timeOut = value
        End Set
        Get
            Return pb.timeOut
        End Get
    End Property


    Public Sub New()
        StartProgressBar()
    End Sub

    Public Sub New(ByVal windowText As String, _
                   Optional ByVal overallValue As Int32 = 0, _
                   Optional ByVal overallText As String = "", _
                   Optional ByVal partialValue As Int32 = 0, _
                   Optional ByVal partialText As String = "", _
                   Optional ByVal callerThread As Threading.Thread = Nothing, _
                   Optional ByVal idleTimeOut As Int32 = 60)

        Me.WindowTitle = windowText
        Me.OverallProgressValue = overallValue
        Me.OverallProgressText = overallText
        Me.PartialProgressValue = partialValue
        Me.PartialProgressText = partialText
        Me.CallerThreadSet = callerThread
        Me.TimeOut = idleTimeOut
        StartProgressBar()

    End Sub

    Private Sub StartProgressBar()
        Dim th As New Threading.Thread(Sub()
                                           Using pb
                                               Application.Run(pb)
                                           End Using
                                       End Sub)
        th.Start()

        AddHandler updateTicker.Tick, AddressOf UpdateProgressWindow
        updateTicker.Interval = 100
        updateTicker.Start()
    End Sub


    Private Sub UpdateProgressWindow()

        If pb.IsDisposed Then 'If form disposed itself
            updateTicker.Stop()
            Me.Dispose()
            'Exit Sub
        End If

        pb.windowTitle = Me.WindowTitle
        pb.overallProgressValue = Me.OverallProgressValue
        pb.overallProgressText = Me.OverallProgressText
        pb.partialProgressValue = Me.PartialProgressValue
        pb.partialProgressText = Me.PartialProgressText

    End Sub


    Private NotInheritable Class Progress_Bar_Window
        Inherits System.Windows.Forms.Form

        '--- Controls
        Friend pbOverall As New ProgressBar
        Friend pbPartial As New ProgressBar
        Friend lblOverall As New Label
        Friend lblPartial As New Label

        '---

        '--- Updated by progress
        Friend windowTitle As String
        Friend overallProgressValue As Int32
        Friend overallProgressText As String
        Friend partialProgressValue As Int32
        Friend partialProgressText As String
        Friend callerThread As Threading.Thread
        Friend timeOut As Int32 = 60
        '---


        Private progressTicker As New Timer
        Private autoCloseTicker As New Timer
        Private forceClose As Boolean

        Private Const defHeight As Int32 = 120
        Private Const defWidth As Int32 = 400

        Private Const closeAttemptMSG1 As String = "Please wait..."
        Private Const closeAttemptMSG2 As String = "Please wait!"
        Private closeAttempts As Int32

        Private lastUpdate As Date = Now
        Private dialogIsShowing As Boolean
        Private closeDialogCount As Int32


        Private Sub Progress_Bar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            Me.Controls.Add(pbPartial)
            Me.Controls.Add(pbOverall)
            Me.Controls.Add(lblPartial)
            Me.Controls.Add(lblOverall)

            lblOverall.AutoSize = True
            lblOverall.Font = New Font("Tahoma", 8, FontStyle.Regular)
            lblPartial.AutoSize = True
            lblPartial.Font = New Font("Tahoma", 8, FontStyle.Regular)

            '==============
            'Activate the block below if you added progBarSize setting to your project
            'If My.Settings.progBarSize.Height > 0 AndAlso _
            'My.Settings.progBarSize.Width > 0 Then
            'Me.Size = My.Settings.progBarSize
            'Else
            Me.Height = defHeight
            Me.Width = defWidth
            'End If
            '==============

            AddHandler progressTicker.Tick, AddressOf UpdateProgressBarsOnTick
            progressTicker.Interval = 99
            progressTicker.Start()

            Controls_Size_And_Position()

            AddHandler Me.SizeChanged, AddressOf WindowSizeChanged

        End Sub

        Private Sub UpdateProgressBarsOnTick()

            If DateDiff(DateInterval.Second, lastUpdate, Now) > timeOut AndAlso Me.Visible = True AndAlso dialogIsShowing = False Then
                dialogIsShowing = True
                If MessageBox.Show("This window hasn't received any updates for " & timeOut & " seconds." & vbNewLine & _
                                   "Do you want to close it ? ", _
                                   "Time Out - " & windowTitle, _
                                   MessageBoxButtons.YesNo, _
                                   MessageBoxIcon.Question, _
                                   MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    Me.Visible = False
                End If

                closeDialogCount += 1
                If closeDialogCount >= 2 Then
                    Dim objTimeOut As Object = InputBox("Please input a new interval in MINUTES to see this question again.", _
                                                        "Time Out", _
                                                        (timeOut / 60).ToString)
                    If IsNumeric(objTimeOut) AndAlso objTimeOut > 0 Then
                        timeOut = (objTimeOut * 60)
                    End If

                End If

                lastUpdate = Now
                dialogIsShowing = False
            End If

            Dim currentRun As Date = Now 'Get the time it enters here
            Threading.Thread.Sleep(1) 'Make sure is higher than lastUpdate

            If callerThread IsNot Nothing AndAlso callerThread.ThreadState = Threading.ThreadState.Stopped Then
                'Caller thread stopped, so stop this one too.
                Dispose()
            End If

            If overallProgressValue > 100 Then
                overallProgressValue = 100
            End If

            If partialProgressValue > 100 Then
                partialProgressValue = 100
            End If


            If Me.Text <> windowTitle AndAlso Not Me.Text Like windowTitle & "*" Then
                Me.Text = windowTitle
                lastUpdate = Now
            End If

            If Me.pbOverall.Value <> Me.overallProgressValue Then
                Me.pbOverall.Value = Me.overallProgressValue
                lastUpdate = Now
            End If

            If Me.lblOverall.Text <> Me.overallProgressText Then
                Me.lblOverall.Text = Me.overallProgressText
                lastUpdate = Now
            End If

            If Me.pbPartial.Value <> Me.partialProgressValue Then
                If Me.pbPartial.Value = 0 Or Me.partialProgressValue = 0 Then
                    Controls_Size_And_Position()
                End If

                Me.pbPartial.Value = Me.partialProgressValue
                lastUpdate = Now
            End If

            If Me.lblPartial.Text <> Me.partialProgressText Then
                Me.lblPartial.Text = Me.partialProgressText
                lastUpdate = Now
            End If


            If Me.Visible = False AndAlso lastUpdate > currentRun AndAlso Me.IsDisposed = False Then 'There was an update, make form visible
                Me.Visible = True
            End If


            If overallProgressValue >= 100 Then 'If Overall Progress reaches 100, closes after 1,5 secs.
                progressTicker.Stop()

                'Activate auto close timer
                AddHandler autoCloseTicker.Tick, AddressOf CloseWindow
                autoCloseTicker.Interval = 1500
                autoCloseTicker.Start()
            End If


        End Sub

        Private Sub Controls_Size_And_Position()

            If partialProgressValue > 0 Then 'Control positions when a partial progress bar is shown
                '--- Partial progress bar
                'Top is 8.5% of the form's height
                'Left is 3.75% of the form's width
                'Height is 8.5% of the form's height
                'Width is 87.5% of the form's width

                pbPartial.Top = ((Me.Height / 100) * 8.5)
                pbPartial.Left = ((Me.Width / 100) * 3.75)
                pbPartial.Height = ((Me.Height / 100) * 8.5)
                pbPartial.Width = ((Me.Width / 100) * 87.5)
                If pbPartial.Visible = False Then pbPartial.Visible = True
                '---

                '--- Partial progress label
                lblPartial.Top = (pbPartial.Top + pbPartial.Height + 2)
                lblPartial.Left = (pbPartial.Left + 5)
                If lblPartial.Visible = False Then lblPartial.Visible = True
                '---


                '--- Overall progress bar
                'Top is 43.5% of the form's height
                'Left is 3.75% of the form's width
                'Height is 8.5% of the form's height
                'Width is 87.5% of the form's width

                pbOverall.Top = ((Me.Height / 100) * 43.5)
                pbOverall.Left = ((Me.Width / 100) * 3.75)
                pbOverall.Height = ((Me.Height / 100) * 8.5)
                pbOverall.Width = ((Me.Width / 100) * 87.5)
                '---

                '--- Overall progress label
                lblOverall.Top = (pbOverall.Top + pbOverall.Height + 2)
                lblOverall.Left = (pbOverall.Left + 5)
                '---


            Else 'No partial bar

                '--- Partial progress bar
                pbPartial.Visible = False
                '---

                '--- Partial progress label
                lblPartial.Visible = False
                '---


                '--- Overall progress bar
                'Top is 20% of the form's height
                'Left is 3.75% of the form's width
                'Height is 25% of the form's height
                'Width is 87.5% of the form's width

                pbOverall.Top = ((Me.Height / 100) * 20)
                pbOverall.Left = ((Me.Width / 100) * 3.75)
                pbOverall.Height = ((Me.Height / 100) * 25)
                pbOverall.Width = ((Me.Width / 100) * 87.5)
                '---

                '--- Overall progress label
                lblOverall.Top = (pbOverall.Top + pbOverall.Height + 2)
                lblOverall.Left = (pbOverall.Left + 5)
                '---

            End If

        End Sub

        Private Sub WindowSizeChanged() '(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged

            If Me.Height < 120 Then Me.Height = 120 'Minimum size
            If Me.Width < 280 Then Me.Width = 280

            Controls_Size_And_Position()

            If Me.Height <> defHeight And Me.Width <> defWidth Then
                '==============
                'Activate the line below if you added progBarSize setting to your project
                'My.Settings.progBarSize = Me.Size
                '==============
            End If

        End Sub

        Private Sub CloseWindow()
            forceClose = True
            Me.Close()
        End Sub

        Private Sub WindowClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

            If forceClose = False Then 'Prevents user from closing window
                e.Cancel = True
                closeAttempts += 1
                If closeAttempts = 1 Then 'Put message on title
                    Me.Text = Me.Text & " - " & Chr(9) & closeAttemptMSG1

                ElseIf closeAttempts >= 2 Then
                    Me.Text = Strings.Left(Me.Text, InStr(Me.Text, Chr(9))) & closeAttemptMSG2

                End If

            End If

        End Sub


    End Class



#Region "IDisposable"
    '---Code entirely copied from MSDN

    Private managedResource As System.ComponentModel.Component
    Private unmanagedResource As IntPtr
    Protected disposed As Boolean = False

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2213:DisposableFieldsShouldBeDisposed", MessageId:="updateTicker")> <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2213:DisposableFieldsShouldBeDisposed", MessageId:="pb")> Protected Overridable Overloads Sub Dispose( _
        ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                Try
                    managedResource.Dispose()
                Catch
                End Try
            End If
            ' Add code here to release the unmanaged resource.
            unmanagedResource = IntPtr.Zero
            ' Note that this is not thread safe. 
        End If
        Me.disposed = True
    End Sub

    'Private Sub AnyOtherMethods()
    'If Me.disposed Then
    'Throw New ObjectDisposedException(Me.GetType().ToString, "This object has been disposed.")
    'End If
    'End Sub

    'Do not change or add Overridable to these methods. 
    'Put cleanup code in Dispose(ByVal disposing As Boolean). 
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub
#End Region

End Class