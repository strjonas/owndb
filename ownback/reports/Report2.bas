Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =7
    RecSrcDt = Begin
        0xb1f663430bf0e540
    End
    RecordSource ="students"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin PageHeader
            Height =1134
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Left =1300
                    Top =170
                    Width =750
                    Height =280
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label0"
                    Caption ="student"
                    GridlineColor =10921638
                    LayoutCachedLeft =1300
                    LayoutCachedTop =170
                    LayoutCachedWidth =2050
                    LayoutCachedHeight =450
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =5952
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =2494
                    Top =1190
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text1"
                    ControlSource ="s_name"
                    GridlineColor =10921638

                    LayoutCachedLeft =2494
                    LayoutCachedTop =1190
                    LayoutCachedWidth =4195
                    LayoutCachedHeight =1490
                    Begin
                        Begin Label
                            Left =793
                            Top =1190
                            Width =550
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label2"
                            Caption ="Text1"
                            GridlineColor =10921638
                            LayoutCachedLeft =793
                            LayoutCachedTop =1190
                            LayoutCachedWidth =1343
                            LayoutCachedHeight =1490
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2494
                    Top =1610
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text3"
                    ControlSource ="s_lname"
                    GridlineColor =10921638

                    LayoutCachedLeft =2494
                    LayoutCachedTop =1610
                    LayoutCachedWidth =4195
                    LayoutCachedHeight =1910
                    Begin
                        Begin Label
                            Left =793
                            Top =1610
                            Width =550
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label4"
                            Caption ="Text1"
                            GridlineColor =10921638
                            LayoutCachedLeft =793
                            LayoutCachedTop =1610
                            LayoutCachedWidth =1343
                            LayoutCachedHeight =1910
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2494
                    Top =2040
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text5"
                    ControlSource ="s_year"
                    GridlineColor =10921638

                    LayoutCachedLeft =2494
                    LayoutCachedTop =2040
                    LayoutCachedWidth =4195
                    LayoutCachedHeight =2340
                    Begin
                        Begin Label
                            Left =793
                            Top =2040
                            Width =550
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label6"
                            Caption ="Text1"
                            GridlineColor =10921638
                            LayoutCachedLeft =793
                            LayoutCachedTop =2040
                            LayoutCachedWidth =1343
                            LayoutCachedHeight =2340
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =1134
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
