Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit
Function test1()
    Dim m_table As TTable
    Set m_table = New TTable
    Dim data As String
    data = DLookup("csv_data", "simple_testf", "id=" & 1)
    
    m_table.init data:=data, max_w:=4, max_h:=3, del_row:=vbCr, del_col:=","
    
    Debug.Print m_table.get_str_data
    

End Function

Function set_ctrlsources()
    Dim rep
    DoCmd.OpenReport "rpt_testf", acViewDesign
    rep = Reports("rpt_testf")
    Dim max_h, max_w
    max_h = 13
    max_w = 13
    Dim rc, cc

    ' Filling table fields from data
    For rc = 0 To max_h

        ' Iterating through each element in array
        For cc = 0 To max_w

            rep("txt_" & rc & "_" & cc & "").ControlSource = "str_" & rc & "_" & cc & ""

        Next cc
    Next rc
End Function
