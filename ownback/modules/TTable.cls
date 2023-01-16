VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim m_str_data As String
Dim m_arr_data() As String
Dim m_max_w As Integer
Dim m_max_h As Integer
Dim m_del_row As String, m_del_col As String

Public Function init(data As String, max_w As Integer, max_h As Integer, del_row As String, del_col As String)
    
    m_max_w = max_w
    m_max_h = max_h
    m_del_row = del_row     ' vbCr
    m_del_col = del_col     ' ,
    
    update_str_data (data) ' Updates the m_str_data and m_arr_data
End Function

Function get_max_w()
    get_max_w = m_max_w
End Function

Function get_max_h()
    get_max_h = m_max_h
End Function

Function get_del_row()
    get_del_row = m_del_row
End Function

Function get_del_col()
    get_del_col = m_del_col
End Function

Function get_str_data()
    get_str_data = m_str_data
End Function

Function get_arr_data()
    get_arr_data = m_arr_data
End Function

Function update_str_data(new_data As String)
    m_str_data = new_data
    m_arr_data = data_to_arr(data:=m_str_data)
End Function

Function update_arr_data(new_arr() As String)
    m_arr_data = new_arr
    m_str_data = arr_to_data(arr:=m_arr_data)
End Function

Function data_to_arr(data As String)
    Dim ergebnis() As String
    ReDim ergebnis(m_max_h, m_max_w)
    Dim rc As Integer, cc As Integer
    rc = 0
    cc = 0
    
    Dim row, col
    
    For Each row In Split(data, m_del_row)
        For Each col In Split(row, m_del_col)
        
            If rc <= m_max_h And cc <= m_max_w Then
                ergebnis(rc, cc) = col
            Else
                MsgBox "Warning, the data is too large for the table", vbCritical
                Exit For
            End If
            
            
            cc = cc + 1
            Next col
        
        rc = rc + 1
        cc = 0
    Next row
    
    data_to_arr = ergebnis
    
End Function


Function arr_to_data(arr() As String)

    Dim temp_arr_data() As String
    temp_arr_data = arr
    
    arr_len_rows = get_len(temp_arr_data, 1)
    arr_len_cols = get_len(temp_arr_data, 2)
    
    Dim temp_data As String
    temp_data = "" ' string der dann gleich neu befuellt wird
    

    ' Holt sich die Werte
    For rc = 0 To m_max_h
                
        ' Iterating through each element in array
        For cc = 0 To m_max_w
            
            ' iterating through max height and with to get the correct amount of commas,
            ' But there migth not be enough data in array, thats why checking
            If arr_len_rows <= rc And arr_len_cols <= cc Then
                temp_data = temp_data & temp_arr_data(rc, cc)
            End If
            
            If cc <> m_max_w Then
                temp_data = temp_data & m_del_col   ' m_del_col um eintraege zu trennen, achtung im letzten durchlauf nicht
            End If
            
        
        Next cc
        
        ' Delimiter appenden damit daten spaeter richtig ausgelesen werden koennen
        temp_data = temp_data & m_del_row
        
    Next rc
    
    arr_to_data = temp_data
    
End Function

