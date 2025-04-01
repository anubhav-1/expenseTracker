Imports System

Public Class DateRange
    ' Properties
    Public Property StartDate As DateTime
    Public Property EndDate As DateTime

    ' Constructor
    Public Sub New(startDate As DateTime, endDate As DateTime)
        Me.StartDate = startDate
        Me.EndDate = endDate
    End Sub

    ' Helper methods

    ' Format dates for SQL queries
    Public Function GetFormattedStartDate() As String
        Return StartDate.ToString("yyyy-MM-dd")
    End Function

    Public Function GetFormattedEndDate() As String
        Return EndDate.ToString("yyyy-MM-dd")
    End Function

    ' Check if a specific date is within this range
    Public Function Contains(dateToCheck As DateTime) As Boolean
        Return dateToCheck >= StartDate AndAlso dateToCheck <= EndDate
    End Function

    ' Get number of months in the range
    Public Function GetMonthCount() As Integer
        Dim months As Integer = (EndDate.Year - StartDate.Year) * 12
        months += EndDate.Month - StartDate.Month

        ' Add 1 to include both start and end months
        Return months + 1
    End Function

    ' Get an array of all month/year combinations in the range
    Public Function GetMonthsInRange() As DateTime()
        Dim result As New List(Of DateTime)
        Dim current As DateTime = New DateTime(StartDate.Year, StartDate.Month, 1)

        While current <= New DateTime(EndDate.Year, EndDate.Month, 1)
            result.Add(current)
            current = current.AddMonths(1)
        End While

        Return result.ToArray()
    End Function

    ' Get array of month names in range
    Public Function GetMonthNamesInRange() As String()
        Return GetMonthsInRange().Select(Function(d) d.ToString("MMM yyyy")).ToArray()
    End Function
End Class