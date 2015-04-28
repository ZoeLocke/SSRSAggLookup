'Allows users to adjust the aggregation type of lookupsets in a cell
Function AggLookup(ByVal choice As String, ByVal items As Object)

    'Ensure passed array is not empty
	'Return a zero so you don't have to allow for Nothing
    If items Is Nothing Then
        Return 0
    End If


    'Define names and data types for all variables
    Dim current As Decimal
    Dim sum As Decimal
    Dim count As Integer
    Dim min As Decimal
    Dim max As Decimal
    Dim err As String

    'Define values for variables where required
    current = 0
    sum = 0
    count = 0
    err = ""

    'Calculate and set variable values
    For Each item As Object In items

        'Calculate Count
        count += 1

        'Check value is a number
        If IsNumeric(item) Then

            'Set current
            current = Convert.ToDecimal(item)

            'Calculate Sum
            sum += current

            'Calculate Min
            If min = Nothing Then
                min = current
            End If
            If min > current Then
                min = current
            End If

            'Calculate the Max
            If max = Nothing Then
                max = current
            End If
            If max < current Then
                max = current
            End If

            'If value is not a number return "NaN"
        Else
            err = "NaN"
        End If

    Next

    'Select and set output based on user choice or parameter one
    If err = "NaN" Then
        If choice = "count" Then
            Return count
        Else
            Return 0
        End If
    Else
        Select Case choice
            Case "sum"
                Return sum
            Case "count"
                Return count
            Case "min"
                Return min
            Case "max"
                Return max
            Case "avg"
                'Calculate the average avoiding divide by zero errors
                If count > 0 Then
                    Return sum / count
                Else
                    Return 0
                End If
        End Select
    End If

    'End
End Function
