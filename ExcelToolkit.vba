Option Explicit

Function AVECELLS(InputRange As Range) As Double
    Dim SUMCELLS As Double, CELLCOUNT As Long
    Dim Arr() As Variant
    Dim R As Long
    Dim C As Long

    ' Transfer the InputRange to the Arr Single or MultiDim Array Holder
    Arr = InputRange

    ' Monitors the Sum of the Cells within the selected range.
    SUMCELLS = 0

    ' Counts the cells that have values
    CELLCOUNT = 0
    
    ' Loops through first dimension (rows)
    For R = 1 To UBound(Arr, 1)

        ' Loops through second dimension (columns)
        For C = 1 To UBound(Arr, 2)
            
            ' Converst the cell contents into trimmed string.
            Dim StringValue As String
            StringValue = Trim(CStr(Arr(R, C)))

            ' Check if the cell is not empty.
            If (Not (StringValue = "")) Then
                ' Gets the double within the cell.
                Dim NumericValue As Double

                ' Converts to numeric value.
                NumericValue =  CDbl(ONLYDIGITS(StringValue))

                ' Increments the sum
                SUMCELLS = SUMCELLS + NumericValue

                ' Increment cell count
                CELLCOUNT = CELLCOUNT + 1
            End If
        Next C
    Next R

    AVECELLS = SUMCELLS / CELLCOUNT
End Function

Function SUMCELLS(InputRange As Range) As Double
    Dim Arr() As Variant
    Arr = InputRange
    Dim R As Long
    Dim C As Long
    SUMCELLS = 0
    
    For R = 1 To UBound(Arr, 1) ' First array dimension is rows.
        For C = 1 To UBound(Arr, 2) ' Second array dimension is columns.
            Debug.Print Arr(R, C)
            Dim NumericValue As Double
            Dim StringValue As String
            
            StringValue = Trim(CStr(Arr(R, C)))
            If (Not (StringValue = "")) Then
                NumericValue = ONLYDIGITS(StringValue)
                SUMCELLS = SUMCELLS + NumericValue
            End If
        Next C
    Next R
End Function

Function ONLYDIGITS(s As String) As String
    Dim retval As String
    Dim i As Integer
    
    Dim periods As Integer
    periods = 0
    
    retval = ""
                                            '
    For i = 1 To Len(s)
        Dim currentCharacter As String
        currentCharacter = Mid(s, i, 1)
        If (currentCharacter >= "0" And currentCharacter <= "9") Then
            retval = retval + currentCharacter
        End If
        
        If (currentCharacter = "." And periods < 1) Then
            periods = periods + 1
            retval = retval + currentCharacter
        End If
    Next
    '
    ONLYDIGITS = retval
End Function

