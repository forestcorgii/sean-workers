Imports System.ComponentModel
Imports System.Threading
Imports System.Windows.Forms

Public Class Workers
    Private workers As List(Of BackgroundWorker)
    Public Event Worker_DoWork(sender As Object, e As DoWorkEventArgs)
    Public Event Worker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
    Public Event RunWorkersCompleted(sender As Object)
    Public Event WorkersWorkBegin(sender As Object)

    Public QueueArgs As List(Of Object)
    Public UseBuffer As Boolean
#Region "Properties"
    Public ReadOnly Property IsBusy As Boolean
        Get
            Return busyWorkersCount > 0
        End Get
    End Property

    Private ReadOnly Property busyWorkersCount As Integer
        Get
            Return (From res As BackgroundWorker In workers Where res.IsBusy Select res).ToList.Count
        End Get
    End Property
#End Region

    Sub New(Optional queueInterval As Integer = 100, Optional _useBuffer As Boolean = False)
        QueueArgs = New List(Of Object)
        init_Timer()
        SetTimer(queueInterval)
        SetWorker(3)
        UseBuffer = _useBuffer
    End Sub

    Public Sub SetWorker(workerCount As Integer)
        workers = New List(Of BackgroundWorker)
        For i As Integer = 0 To workerCount - 1
            Dim newWorker As New BackgroundWorker
            With newWorker
                AddHandler .DoWork, AddressOf Item_DoWork
                AddHandler .RunWorkerCompleted, AddressOf Item_RunWorkerCompleted
            End With
            workers.Add(newWorker)
        Next
    End Sub

#Region "Event Handlers"
    Private Sub Item_DoWork(sender As Object, e As DoWorkEventArgs)
        RaiseEvent Worker_DoWork(sender, e)
    End Sub

    Private Sub Item_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        RaiseEvent Worker_RunWorkerCompleted(sender, e)
    End Sub
#End Region

    Public Sub AddtoQueue(arg As Object)
        QueueArgs.Add(arg)
        If Not ticking Then
            StartWorking()
        End If
    End Sub

    Public Sub AddRangetoQueue(args As IEnumerable(Of Object))
        If args.Count = 0 Then Exit Sub

        QueueArgs.AddRange(args)
        If Not ticking Then
            StartWorking()
        End If
    End Sub

    Public Function AttemptExecute(Optional args As Object = Nothing) As Boolean
        For i As Integer = 0 To workers.Count - 1
            If workers(i).IsBusy = False Then
                workers(i).RunWorkerAsync(args)
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub UpdateQueue()
        If QueueArgs.Count > 0 Then
            If AttemptExecute(QueueArgs(0)) Then
                QueueArgs.RemoveAt(0)
            End If
        End If
    End Sub

#Region "Timer"
    Private queueTimer As Threading.Timer
    Private interval As Integer = -1
    Private ticking As Boolean

    Private Sub init_Timer()
        queueTimer = New Threading.Timer(AddressOf queueTimer_Tick, Nothing, Timeout.Infinite, Timeout.Infinite)
    End Sub

    Private Sub queueTimer_Tick(sender As Object)
        UpdateQueue()
        If Not IsBusy Then
            StartWorking()
            RaiseEvent RunWorkersCompleted(Me)
        End If
    End Sub

    Public Sub StartWorking()
        queueTimer.Change(0, interval)
        ticking = True
    End Sub

    Public Sub StopWorking()
        queueTimer.Change(-1, -1)
        ticking = False
    End Sub

    Private Sub SetTimer(_interval As Integer)
        interval = _interval
    End Sub
#End Region
End Class
