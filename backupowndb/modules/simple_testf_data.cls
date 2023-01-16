VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "simple_testf_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim str_data As String
Dim arr_data() As String
Dim max_w As Integer
Dim max_h As Integer

Function get_str_data()
    get_str_data = str_data
End Function

Function get_arr_data()
    get_arr_data = arr_data
End Function

Function update_str_data(data As String)
    str_data = data
    data_to_arr
End Function

Function update_arr_data(arr() As String)
    arr_data = arr
    arr_to_data
End Function

Function data_to_arr()
    
End Function

Function arr_to_data()

End Function
