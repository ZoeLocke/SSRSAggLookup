
'Allows users to adjust the aggregation type of lookupsets in a cell
Function AggLookup(ByVal choice As String, ByVal items As Object)

    'Ensure passed array is not null or empty
    'Return a zero so you don't have to allow for Nothing
    If items Is Nothing OrElse Not items.length > 0 Then
        Return 0

    End If

    'Call the method specified
    Select Case choice.ToLower 'Using ToLower allows the user to pass in the values without regard to case sensitivity
        Case "sum"
            Return GetSumFromObjectArray(items)

        Case "count"
            Return items.length ' As an array, we can just call its Length property and return that

        Case "countdistinct"
            Return GetDistinctCountFromObjectArray(items)

        Case "avg"
            Dim runningTotal As Decimal = GetSumFromObjectArray(items)
            Return runningTotal / items.length

        Case "min"
            Return GetMinValueFromObjectArray(items)

        Case "max"
            Return GetMaxValueFromObjectArray(items)

        Case "first"
            Return items(0)

        Case "last"
            Return items(items.length - 1)

        Case Else
            Return 0 'If option provided was not valid, return 0 (could change this to return an appropriate error message for the Dev)

    End Select    
    return 0 'If nothing was returned yet, something went wrong; return 0 (again, this could be changed to be more informative for the Dev)

End Function

Private Function GetSumFromObjectArray(items As Object()) As Decimal
    Dim runningTotal As Decimal = 0
    For Each item As Object In items
        Dim thisItemAsDecimal As Decimal
        If Decimal.TryParse(item, thisItemAsDecimal) Then
            runningTotal += thisItemAsDecimal
        Else
            Return 0 'Short circuit; If ANY value is not a number, the entire SUM is invalid, stop evaluating and return 0
        End If
    Next
    Return runningTotal

End Function

Private Function GetMinValueFromObjectArray(items As Object()) As Decimal
    Dim currentMin As Decimal = Decimal.MaxValue 'Default to the max value allowed for a Decimal (Would default to a random value in the array, but it is possible they are not numeric)
    For Each item As Object In items
        Dim thisItemAsDecimal As Decimal
        If Decimal.TryParse(item.ToString, thisItemAsDecimal) Then
            If thisItemAsDecimal < currentMin Then currentMin = thisItemAsDecimal
        Else
            Return 0 'Short circuit
        End If
    Next
    Return currentMin

End Function

Private Function GetMaxValueFromObjectArray(items As Object()) As Decimal
    Dim currentMax As Decimal = Decimal.MinValue 'Default to the min value allowed for a Decimal (Would default to a random value in the array, but it is possible they are not numeric)
    For Each item As Object In items
        Dim thisItemAsDecimal As Decimal
        If Decimal.TryParse(item.ToString, thisItemAsDecimal) Then
            If thisItemAsDecimal > currentMax Then currentMax = thisItemAsDecimal
        Else
            Return 0 'Short circuit
        End If
    Next
    Return currentMax

End Function


''' <summary>
''' Returns the count of distinct items in the array; modified from code retrieved from https://stackoverflow.com/questions/38187819/ssrs-code-to-determine-distinct-count-of-a-lookupset on 5 July 2017.
''' </summary>
''' <param name="items"></param>
''' <returns>Integer</returns>
Public Function GetDistinctCountFromObjectArray(items As Object()) As Integer
    System.Array.Sort(items)
    Dim k As Integer = 0
    If items.Length > 0 Then
        k = 1
        For i As Integer = 1 To items.Length - 1
            If items(i).Equals(items(i - 1)) Then
                Continue For
            End If
            k += 1
        Next
    End If
    Return k
    
End Function