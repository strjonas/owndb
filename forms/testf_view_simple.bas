Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =6
    Right =21250
    Bottom =11980
    RecSrcDt = Begin
        0xe34b947d54f0e540
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =5952
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1697
                    Top =907
                    Width =1591
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_name"
                    GridlineColor =10921638

                    LayoutCachedLeft =1697
                    LayoutCachedTop =907
                    LayoutCachedWidth =3288
                    LayoutCachedHeight =1207
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =910
                            Width =570
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label1"
                            Caption ="name"
                            GridlineColor =10921638
                            LayoutCachedLeft =170
                            LayoutCachedTop =910
                            LayoutCachedWidth =740
                            LayoutCachedHeight =1210
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1697
                    Top =1478
                    Width =3921
                    Height =1880
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_tbl"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =1697
                    LayoutCachedTop =1478
                    LayoutCachedWidth =5618
                    LayoutCachedHeight =3358
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =1474
                            Width =670
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label3"
                            Caption ="tabelle"
                            GridlineColor =10921638
                            LayoutCachedLeft =170
                            LayoutCachedTop =1474
                            LayoutCachedWidth =840
                            LayoutCachedHeight =1774
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4081
                    Top =4818
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_id"
                    DefaultValue ="-1"
                    GridlineColor =10921638

                    LayoutCachedLeft =4081
                    LayoutCachedTop =4818
                    LayoutCachedWidth =5782
                    LayoutCachedHeight =5118
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_load()
    Dim ergebnis As String
    
    ' Get ID from custom function
    ID = get_id
    
    ergebnis = load_data(ID)
    
    On Error Resume Next
    txt_tbl.Value = ergebnis
End Sub


    
Private Function load_data(ID As Variant)

    
    Dim sql As String
    Dim name As String
    Dim records() As Variant
    Dim ergebnis As String
    
    On Error GoTo errHandler
func:
    ' Set the name to the name of the testf
    Me.txt_name.Value = DLookup("s_name", "tbl_testf", "id=" & ID)

    
    ' hier nicht ID sonder fk_testf, weil das ja die eintraege sind die zu der id gehoeren
    records = get_all("ID,s_stoff", "tbl_testf_rows", "fk_testf=" & ID)
    
    rec_len_rows = get_len(records, 1)
    rec_len_cols = get_len(records, 2)
    
    ergebnis = ""
    ' Array to String
    For i = 0 To rec_len_rows - 1
        For j = 0 To rec_len_cols - 1
            ergebnis = ergebnis & records(i, j) & " "
        Next j
        ergebnis = ergebnis & "<br>"
    Next i
    
    load_data = ergebnis
    
    Exit Function
    
errHandler:
    Debug.Print Err
    Select Case Err.Number
    Case 13:
        Debut.Print "13"
    Case 94:
        txt_id.Value = -1
        If MsgBox("Id not in database", vbRetryCancel) = vbRetry Then
            get_id
            Resume func
            
        Else
            DoCmd.Close acForm, Me.name
        End If
    Case Default:
        Debut.Print Err
    End Select


End Function

' Prompts the user until there is a valid ID
Private Function get_id()
    Do While txt_id.Value < 0
       txt_id.Value = InputBox("Enter new id")
    Loop
    
    get_id = txt_id.Value
    
End Function
