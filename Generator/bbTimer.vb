Imports System
Imports System.Diagnostics

Public Class bbTimer
    Private ElapsedTicks As Double = 0
    Private x As New Stopwatch
    Private internalName As String
    Private LapCount As Integer = 0

    Public Sub New(ByVal Name As String)
        internalName = Name
    End Sub

    Public Sub StartClock()
        x.Reset()
        LapCount += 1
        x.Start()
    End Sub

    Public Sub StopClock()
        x.Stop()
        ElapsedTicks += x.ElapsedTicks()
    End Sub

    Public Sub ResetClock()
        If x.IsRunning() Then
            StopClock()
        End If
        ElapsedTicks = 0
    End Sub

    Public Sub DebugOut(ByVal TotalTime As Double)
        Debug.Print(internalName & ": " & Format(ElapsedTime() / TotalTime * 100.0, "0.0000") & "% (" & LapCount & " X " & ElapsedTime() / CDbl(LapCount) & " s each)")
    End Sub

    Public Function ElapsedTime() As Double
        Return ElapsedTicks / CDbl(Stopwatch.Frequency)
    End Function
End Class
