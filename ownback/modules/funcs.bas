Attribute VB_Name = "funcs"
Option Compare Database

' Returns the length of a dimension of an array
Function get_len(arr As Variant, dimension As Integer)
    get_len = UBound(arr, dimension) - LBound(arr, dimension) + 1
End Function
    
' Returns the value if not null, otherwise empty string
Function v_or_emp(val As Variant)
    If Not IsNull(val) Then
        v_or_emp = val
    Else
        v_or_emp = ""
    End If
        
End Function


' Gets the values for all fields in specific tables on a specific condition
' Returns 2 Dimensional array
Function get_all(Values As String, Tables As String, Conditions As String)
    Dim RS As DAO.Recordset
    Dim sql As String
    Dim ergebnis() As Variant
    Dim c_item As Integer: c_item = 0
    Dim c_row As Integer: c_row = 0
    
    Dim val_arr() As String
    Dim col_len As Integer
    Dim row_len As Integer
    
    val_arr = Split(Values, ",")
    
    ' Getting the amount of rows and columns to populate the array
    col_len = get_len(val_arr, 1)
    row_len = DCount(val_arr(0), Tables, Conditions)
    ' Keine eintraege
    If CInt(row_len) = 0 Then
        Exit Function
        
    End If
        
    ' Ergbenis array auf datensaetze zuschneiden
    ReDim Preserve ergebnis(0 To row_len - 1, 0 To col_len - 1)
    
'    ReDim row(0 To (val_len - 1))
    
    sql = "select " & Values & " from " & Tables & " where " & Conditions
    
    Set RS = CurrentDb.OpenRecordset(sql)
    
    Dim val As Variant
    Do While Not RS.EOF
    
        For Each val In val_arr
            val = TrimAllWhitespace(val)
            Debug.Print (RS(val))
            ergebnis(c_row, c_item) = "" & RS(val)
            c_item = c_item + 1
        Next val
        
        c_item = 0
        c_row = c_row + 1
        
        RS.MoveNext
    Loop
    
    RS.Close
    Set RS = Nothing
    
    
    get_all = ergebnis
End Function







