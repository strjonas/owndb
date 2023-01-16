Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9978
    DatasheetFontHeight =11
    ItemSuffix =28
    Right =21250
    Bottom =12570
    RecSrcDt = Begin
        0x1819f7310bf0e540
    End
    RecordSource ="students"
    Caption ="show_students"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1026
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =57
                    Top =57
                    Width =3036
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label8"
                    Caption ="show_students"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3093
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            NewRowOrCol =1
            Height =8617
            Name ="Detail"
            AlternateBackColor =15658734
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2672
                    Top =741
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_name"
                    ControlSource ="s_name"
                    GridlineColor =10921638

                    LayoutCachedLeft =2672
                    LayoutCachedTop =741
                    LayoutCachedWidth =9312
                    LayoutCachedHeight =1321
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =741
                            Width =2240
                            Height =310
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_name_Label"
                            Caption ="name"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =741
                            LayoutCachedWidth =2582
                            LayoutCachedHeight =1051
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2672
                    Top =1425
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_lname"
                    ControlSource ="s_lname"
                    GridlineColor =10921638

                    LayoutCachedLeft =2672
                    LayoutCachedTop =1425
                    LayoutCachedWidth =9312
                    LayoutCachedHeight =2005
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1425
                            Width =2240
                            Height =310
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_lname_Label"
                            Caption ="lastname"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1425
                            LayoutCachedWidth =2582
                            LayoutCachedHeight =1735
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2672
                    Top =2109
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_year"
                    ControlSource ="s_year"
                    GridlineColor =10921638

                    LayoutCachedLeft =2672
                    LayoutCachedTop =2109
                    LayoutCachedWidth =4082
                    LayoutCachedHeight =2419
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2109
                            Width =2240
                            Height =310
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_year_Label"
                            Caption ="year"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2109
                            LayoutCachedWidth =2582
                            LayoutCachedHeight =2419
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2664
                    Top =2621
                    Height =290
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text10"
                    ConditionalFormat = Begin
                        0x01000000b0000000020000000100000000000000000000000b00000001000000 ,
                        0x0000000000b7ef0001000000000000000c000000270000000100000000000000 ,
                        0x2f36990000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0073005f0079006500610072005d003c003500000000005b0073005f007900 ,
                        0x6500610072005d003c0031003000200041006e00640020005b0073005f007900 ,
                        0x6500610072005d003e00340000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2664
                    LayoutCachedTop =2621
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =2911
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010000000000000000b7ef000a0000005b00 ,
                        0x73005f0079006500610072005d003c0035000000000000000000000000000000 ,
                        0x00000000000000010000000000000001000000000000002f3699001a0000005b ,
                        0x0073005f0079006500610072005d003c0031003000200041006e00640020005b ,
                        0x0073005f0079006500610072005d003e00340000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =963
                            Top =2607
                            Width =1520
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label11"
                            Caption ="graduation level"
                            GridlineColor =10921638
                            LayoutCachedLeft =963
                            LayoutCachedTop =2607
                            LayoutCachedWidth =2483
                            LayoutCachedHeight =2907
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7880
                    Top =3118
                    Height =300
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_id"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =7880
                    LayoutCachedTop =3118
                    LayoutCachedWidth =9581
                    LayoutCachedHeight =3418
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =226
                    Top =7937
                    Height =403
                    TabIndex =5
                    ForeColor =4210752
                    Name ="Command13"
                    Caption ="vorschau"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =226
                    LayoutCachedTop =7937
                    LayoutCachedWidth =1927
                    LayoutCachedHeight =8340
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5045
                    Top =2834
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btn_view"
                    Caption ="View"
                    GridlineColor =10921638

                    LayoutCachedLeft =5045
                    LayoutCachedTop =2834
                    LayoutCachedWidth =6746
                    LayoutCachedHeight =3117
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7993
                    Top =3628
                    TabIndex =7
                    ForeColor =4210752
                    Name ="autofiller"
                    Caption ="autofiller"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7993
                    LayoutCachedTop =3628
                    LayoutCachedWidth =9694
                    LayoutCachedHeight =3911
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =3398
                    Width =7085
                    Height =2700
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text24"
                    ControlSource ="s_desc"
                    GridlineColor =10921638

                    LayoutCachedLeft =283
                    LayoutCachedTop =3398
                    LayoutCachedWidth =7368
                    LayoutCachedHeight =6098
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =340
                    Top =6292
                    Width =7031
                    Height =1530
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text26"
                    ControlSource ="s_long"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =340
                    LayoutCachedTop =6292
                    LayoutCachedWidth =7371
                    LayoutCachedHeight =7822
                End
            End
        End
        Begin FormFooter
            Height =1587
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =283
                    Top =680
                    Width =810
                    Height =360
                    ForeColor =4210752
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =283
                    LayoutCachedTop =680
                    LayoutCachedWidth =1093
                    LayoutCachedHeight =1040
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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

Private Sub autofiller_Click()
    If s_year < 5 Then
        btn_view.BackColor = RGB(0, 0, 0)
    End If
    
End Sub

Private Sub Command13_Click()
      MsgBox ("Die folgende Ansicht der Spezifikationkann mit ESC verlassen werden.")

    DoCmd.OpenReport "Report2", acViewPreview, , " " & txt_id.Value & " = id"
    
End Sub

Private Sub Refresh_Click()
Me.Requery

    Me.Refresh
    
End Sub
