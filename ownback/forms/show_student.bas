Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11457
    DatasheetFontHeight =11
    ItemSuffix =18
    Right =21510
    Bottom =12550
    RecSrcDt = Begin
        0x3dd1eb2b52f0e540
    End
    RecordSource ="students"
    Caption ="show_student"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
                    Width =2844
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label14"
                    Caption ="show_student"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =2901
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =8787
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =114
                    Top =414
                    Width =11286
                    Height =285
                    ColumnWidth =1701
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =114
                    LayoutCachedTop =414
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =699
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =114
                            Top =114
                            Width =11286
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ID_Label"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =114
                            LayoutCachedTop =114
                            LayoutCachedWidth =11400
                            LayoutCachedHeight =414
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =114
                    Top =999
                    Width =11286
                    Height =850
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_name"
                    ControlSource ="s_name"
                    GridlineColor =10921638

                    LayoutCachedLeft =114
                    LayoutCachedTop =999
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =1849
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =114
                            Top =699
                            Width =11286
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_name_Label"
                            Caption ="s_name"
                            GridlineColor =10921638
                            LayoutCachedLeft =114
                            LayoutCachedTop =699
                            LayoutCachedWidth =11400
                            LayoutCachedHeight =999
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =114
                    Top =2149
                    Width =11286
                    Height =850
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_lname"
                    ControlSource ="s_lname"
                    GridlineColor =10921638

                    LayoutCachedLeft =114
                    LayoutCachedTop =2149
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =2999
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =114
                            Top =1849
                            Width =11286
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_lname_Label"
                            Caption ="s_lname"
                            GridlineColor =10921638
                            LayoutCachedLeft =114
                            LayoutCachedTop =1849
                            LayoutCachedWidth =11400
                            LayoutCachedHeight =2149
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =114
                    Top =3299
                    Width =11286
                    Height =285
                    ColumnWidth =1410
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_year"
                    ControlSource ="s_year"
                    GridlineColor =10921638

                    LayoutCachedLeft =114
                    LayoutCachedTop =3299
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =3584
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =114
                            Top =2999
                            Width =11286
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_year_Label"
                            Caption ="s_year"
                            GridlineColor =10921638
                            LayoutCachedLeft =114
                            LayoutCachedTop =2999
                            LayoutCachedWidth =11400
                            LayoutCachedHeight =3299
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =114
                    Top =3884
                    Width =11286
                    Height =1450
                    ColumnWidth =3860
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_desc"
                    ControlSource ="s_desc"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =114
                    LayoutCachedTop =3884
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =5334
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =114
                            Top =3584
                            Width =11286
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_desc_Label"
                            Caption ="s_desc"
                            GridlineColor =10921638
                            LayoutCachedLeft =114
                            LayoutCachedTop =3584
                            LayoutCachedWidth =11400
                            LayoutCachedHeight =3884
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =113
                    Top =5742
                    Width =11286
                    Height =1740
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_long"
                    ControlSource ="s_long"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =113
                    LayoutCachedTop =5742
                    LayoutCachedWidth =11399
                    LayoutCachedHeight =7482
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =113
                            Top =5442
                            Width =11286
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_long_Label"
                            Caption ="s_long"
                            GridlineColor =10921638
                            LayoutCachedLeft =113
                            LayoutCachedTop =5442
                            LayoutCachedWidth =11399
                            LayoutCachedHeight =5742
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =226
                    Top =7896
                    Width =1410
                    Height =285
                    ColumnWidth =1410
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="s_number"
                    ControlSource ="s_number"
                    GridlineColor =10921638

                    LayoutCachedLeft =226
                    LayoutCachedTop =7896
                    LayoutCachedWidth =1636
                    LayoutCachedHeight =8181
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =226
                            Top =7596
                            Width =1410
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="s_number_Label"
                            Caption ="s_number"
                            GridlineColor =10921638
                            LayoutCachedLeft =226
                            LayoutCachedTop =7596
                            LayoutCachedWidth =1636
                            LayoutCachedHeight =7896
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9693
                    Top =7834
                    Height =300
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_id"
                    DefaultValue ="-1"
                    GridlineColor =10921638

                    LayoutCachedLeft =9693
                    LayoutCachedTop =7834
                    LayoutCachedWidth =11394
                    LayoutCachedHeight =8134
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4592
                    Top =8050
                    TabIndex =8
                    ForeColor =4210752
                    Name ="btn_different"
                    Caption ="show other student"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4592
                    LayoutCachedTop =8050
                    LayoutCachedWidth =6293
                    LayoutCachedHeight =8333
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
            End
        End
        Begin FormFooter
            Height =1587
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_different_Click()
txt_id.Value = -1
load_data
End Sub

Private Sub Form_load()
load_data
    
End Sub
Private Sub load_data()
    If txt_id.Value = -1 Then
        txt_id.Value = InputBox("Enter new id")
    End If
    'Check if ID is existant
    If DCount("id", "students", "s_number=" & txt_id.Value) = 0 Then
        DoCmd.OpenForm "add_student"
        DoCmd.Close acForm, Me.name
    
    End If
    
    Me.RecordSource = "select * from students where id=" & txt_id.Value
End Sub
