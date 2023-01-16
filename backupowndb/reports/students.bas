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
    Width =11184
    DatasheetFontHeight =11
    ItemSuffix =9
    RecSrcDt = Begin
        0x42efa10c0bf0e540
    End
    RecordSource ="students"
    Caption ="students"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="s_name"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="s_lname"
        End
        Begin BreakLevel
            ControlSource ="s_year"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =919
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =57
                    Top =57
                    Width =1460
                    Height =520
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label6"
                    Caption ="students"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =1517
                    LayoutCachedHeight =577
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1077
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1098
            BreakLevel =1
            Name ="GroupHeader1"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =4308
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =7540
                    Top =1020
                    Width =1410
                    Height =310
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_year"
                    ControlSource ="s_year"
                    GridlineColor =10921638

                    LayoutCachedLeft =7540
                    LayoutCachedTop =1020
                    LayoutCachedWidth =8950
                    LayoutCachedHeight =1330
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2953
                    Top =850
                    Width =3340
                    Height =310
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_lname"
                    ControlSource ="s_lname"
                    GridlineColor =10921638

                    LayoutCachedLeft =2953
                    LayoutCachedTop =850
                    LayoutCachedWidth =6293
                    LayoutCachedHeight =1160
                    Begin
                        Begin Label
                            Left =623
                            Top =850
                            Width =2240
                            Height =310
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_lname_Label"
                            Caption ="s_lname"
                            GridlineColor =10921638
                            LayoutCachedLeft =623
                            LayoutCachedTop =850
                            LayoutCachedWidth =2863
                            LayoutCachedHeight =1160
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2896
                    Top =1303
                    Width =3340
                    Height =310
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_name"
                    ControlSource ="s_name"
                    GridlineColor =10921638

                    LayoutCachedLeft =2896
                    LayoutCachedTop =1303
                    LayoutCachedWidth =6236
                    LayoutCachedHeight =1613
                    Begin
                        Begin Label
                            Left =566
                            Top =1303
                            Width =2240
                            Height =310
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_name_Label"
                            Caption ="s_name"
                            GridlineColor =10921638
                            LayoutCachedLeft =566
                            LayoutCachedTop =1303
                            LayoutCachedWidth =2806
                            LayoutCachedHeight =1613
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =2097
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =57
                    Top =228
                    Width =5040
                    Height =310
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text7"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =57
                    LayoutCachedTop =228
                    LayoutCachedWidth =5097
                    LayoutCachedHeight =538
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6087
                    Top =228
                    Width =5040
                    Height =310
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text8"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6087
                    LayoutCachedTop =228
                    LayoutCachedWidth =11127
                    LayoutCachedHeight =538
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1644
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
