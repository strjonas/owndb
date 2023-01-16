Attribute VB_Name = "Module2"
Option Compare Database
Option Explicit


Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' If element is resized, the rest needs to go down
    Dim offset As Integer: offset = 50
    
    Dim max_w, max_h
    max_w = 13
    max_h = 13
    
    Dim rc, cc
    Dim minheight: minheight = 250
    Dim maxheight
    Dim height
    
    Dim padding_rl: padding_side = 100
    Dim padding_tb: padding_tb = 20
    
    Dim umbruch_ab: umbruch_ab = 5000
    
    
    ' Get the height of highest element in a row, resize all the others
    For rc = 0 To max_h
        maxheight = minheight
        
        ' Get the max height
        For cc = 0 To max_w
            
            ' Textheight gibt immer nur wie hoch die schrift ist, auch wenn tatsaechlich es 3 reihen gibt
            ' Da bei umbruch_ab als weite umgebrochen wird kann man damit die reihen bestimmen
            ' Plus eins weil int abrundet aber ich immer aufrunden muss
            height = TextHeight(Me("txt_" & rc & "_" & cc & ""))
            height = height * (Int(TextWidth(Me("txt_" & rc & "_" & cc & "")) / umbruch_ab) + 1)
            If height > maxheight Then
                maxheight = height
            End If
        Next cc
        cc = 0
        

            
        ' Resize all and put the correct position
        For cc = 0 To max_w
            
            Me("txt_" & rc & "_" & cc & "").Top = offset
            Me("txt_" & rc & "_" & cc & "").height = maxheight + padding_tb
        Next cc
        cc = 0
        
        ' offset plus the height the row will be
        offset = offset + maxheight + padding_tb
    Next rc
    
    
    
    ' Same thing for width
    offset = 50
    Dim minwidth: minwidth = 500
    Dim maxwidth
    Dim width
    
    ' gleicher loop aber cc und rc getauscht
    For cc = 0 To max_w
        maxwidth = minwidth
        
        ' Get the max width
        For rc = 0 To max_h
        
            width = TextWidth(Me("txt_" & rc & "_" & cc & ""))
            If width > maxwidth Then
                If width > umbruch_ab Then
                    width = umbruch_ab
                End If
                maxwidth = width
                
            End If
        Next rc
        rc = 0
       
            
        ' Resize all and put the correct position
        For rc = 0 To max_h
            Me("txt_" & rc & "_" & cc & "").Left = offset
            Me("txt_" & rc & "_" & cc & "").width = maxwidth + padding_side
        Next rc
        rc = 0
    
        ' offset plus the width the col will have
        offset = offset + maxwidth + padding_side
    Next cc
    
    ' Hide empty rows and columns
    Debug.Print "done"

End Sub

