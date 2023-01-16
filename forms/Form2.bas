Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =17177
    DatasheetFontHeight =11
    ItemSuffix =18
    Right =25440
    Bottom =12550
    RecSrcDt = Begin
        0x63bf4f2455f0e540
    End
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
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
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
        Begin Tab
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin WebBrowser
            OldBorderStyle =1
            Width =4536
            Height =2835
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationControl
            BorderWidth =1
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationButton
            Width =283
            Height =283
            ForeColor =-2
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            HoverColor =-2
            HoverThemeColorIndex =2
            HoverTint =20.0
            PressedColor =-2
            PressedThemeColorIndex =2
            PressedTint =60.0
            HoverForeColor =-2
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =-2
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            BackThemeColorIndex =1
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            FontName ="Calibri"
            FontWeight =400
            FontSize =11
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Section
            Height =12860
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =5760
                    Left =7483
                    Top =6009
                    Height =300
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Combo10"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [students].[ID], [students].[s_name], [students].[s_lname], [students].[s"
                        "_year], [students].[s_number] FROM students ORDER BY [s_number], [s_name], [s_ye"
                        "ar]; "
                    ColumnWidths ="0;1440;1440;1440;1440"
                    GridlineColor =10921638

                    LayoutCachedLeft =7483
                    LayoutCachedTop =6009
                    LayoutCachedWidth =9184
                    LayoutCachedHeight =6309
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5782
                            Top =6009
                            Width =860
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_name_Label"
                            Caption ="s_name"
                            GridlineColor =10921638
                            LayoutCachedLeft =5782
                            LayoutCachedTop =6009
                            LayoutCachedWidth =6642
                            LayoutCachedHeight =6329
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    Left =2948
                    Top =960
                    Width =10204
                    Height =4880
                    TabIndex =1
                    Name ="TabCtl13"
                    FontName ="Calibri Light"
                    GridlineColor =10921638

                    LayoutCachedLeft =2948
                    LayoutCachedTop =960
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =5840
                    BackColor =14277081
                    BorderColor =11573124
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    ForeColor =4210752
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =3000
                            Top =1370
                            Width =10100
                            Height =4420
                            BorderColor =10921638
                            Name ="Page14"
                            GridlineColor =10921638
                            LayoutCachedLeft =3000
                            LayoutCachedTop =1370
                            LayoutCachedWidth =13100
                            LayoutCachedHeight =5790
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =5160
                                    Top =2780
                                    Width =380
                                    Height =280
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Label17"
                                    Caption ="bye"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =5160
                                    LayoutCachedTop =2780
                                    LayoutCachedWidth =5540
                                    LayoutCachedHeight =3060
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =3000
                            Top =1370
                            Width =10100
                            Height =4420
                            BorderColor =10921638
                            Name ="Page15"
                            GridlineColor =10921638
                            LayoutCachedLeft =3000
                            LayoutCachedTop =1370
                            LayoutCachedWidth =13100
                            LayoutCachedHeight =5790
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =5330
                                    Top =2040
                                    Width =500
                                    Height =280
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Label16"
                                    Caption ="hallo"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =5330
                                    LayoutCachedTop =2040
                                    LayoutCachedWidth =5830
                                    LayoutCachedHeight =2320
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
