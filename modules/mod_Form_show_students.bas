Attribute VB_Name = "mod_Form_show_students"
Option Compare Database

Function TrimAllWhitespace(ByVal str As String)

    str = Trim(str)

    Do Until Not Left(str, 1) = Chr(9)
        str = Trim(Mid(str, 2, Len(str) - 1))
    Loop

    Do Until Not Right(str, 1) = Chr(9)
        str = Trim(Left(str, Len(str) - 1))
    Loop

    TrimAllWhitespace = str

End Function
