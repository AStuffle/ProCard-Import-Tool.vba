Sub Convert()
    ' defining initial variables
    Dim myFile As String, rng As Range, cellValue As Variant, i As Integer, j As Integer
    
    'Initially I had it pointing at the user's download folder
    'myFile = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\Downloads\fleet_pcard_upload.txt"
    
    'Now it points directly to the network share
    myFile = "\\wins.lcra.org\maximo_fileprocess\fileprocess\pcard_export.txt"
    
    'Loops to iterate through the grid
    Set rng = Selection
    Open myFile For Output As #1
    For i = 1 To rng.Rows.Count
            For j = 1 To rng.Columns.Count
            cellValue = rng.Cells(i, j).Value
                If i = rng.Rows.Count And j = rng.Columns.Count Then
                    Print #1, cellValue;
                Else
                    If j = rng.Columns.Count Then
                        Print #1, cellValue
                    Else
                        Print #1, cellValue & ",";
                    End If
                End If
            Next j
    Next i
Close #1

End Sub
Sub Button()
Call Convert
End Sub
