Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10771
    DatasheetFontHeight =11
    ItemSuffix =14
    Right =24150
    Bottom =11820
    RecSrcDt = Begin
        0xb5727edc0cf1e540
    End
    DatasheetFontName ="Calibri"
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
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
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
        Begin PageHeader
            DisplayWhen =1
            Height =1360
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =1133
                    Top =226
                    Width =2325
                    Height =735
                    FontSize =16
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label2"
                    Caption ="Menu"
                    GridlineColor =10921638
                    LayoutCachedLeft =1133
                    LayoutCachedTop =226
                    LayoutCachedWidth =3458
                    LayoutCachedHeight =961
                End
            End
        End
        Begin Section
            Height =7256
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =1133
                    Top =3968
                    Height =568
                    ForeColor =4210752
                    Name ="btn_edit"
                    Caption ="Bearbeiten"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1133
                    LayoutCachedTop =3968
                    LayoutCachedWidth =2834
                    LayoutCachedHeight =4536
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
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1133
                    Top =5669
                    Height =568
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Command1"
                    Caption ="Neu"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1133
                    LayoutCachedTop =5669
                    LayoutCachedWidth =2834
                    LayoutCachedHeight =6237
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
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =1133
                    Top =2551
                    Width =3120
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label5"
                    Caption ="Id zum bearbeiten auswaehlen:"
                    GridlineColor =10921638
                    LayoutCachedLeft =1133
                    LayoutCachedTop =2551
                    LayoutCachedWidth =4253
                    LayoutCachedHeight =2836
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1133
                    Top =3061
                    Width =3126
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="combo_id"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT id, str_name FROM simple_testf; "
                    ColumnWidths ="284"
                    GridlineColor =10921638

                    LayoutCachedLeft =1133
                    LayoutCachedTop =3061
                    LayoutCachedWidth =4259
                    LayoutCachedHeight =3376
                End
                Begin Label
                    OverlapFlags =93
                    Left =1133
                    Top =5102
                    Width =3120
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label8"
                    Caption ="Oder neu:"
                    GridlineColor =10921638
                    LayoutCachedLeft =1133
                    LayoutCachedTop =5102
                    LayoutCachedWidth =4253
                    LayoutCachedHeight =5387
                End
                Begin OptionGroup
                    OverlapFlags =247
                    Left =566
                    Top =1700
                    Width =4530
                    Height =5114
                    TabIndex =3
                    BorderColor =10921638
                    Name ="Frame9"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =1700
                    LayoutCachedWidth =5096
                    LayoutCachedHeight =6814
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =5669
                    Top =1700
                    Width =4530
                    Height =5114
                    TabIndex =4
                    BorderColor =10921638
                    Name ="Frame11"
                    GridlineColor =10921638

                    LayoutCachedLeft =5669
                    LayoutCachedTop =1700
                    LayoutCachedWidth =10199
                    LayoutCachedHeight =6814
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6236
                    Top =2834
                    Height =568
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btn_show_all"
                    Caption ="Alle Anzeigen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6236
                    LayoutCachedTop =2834
                    LayoutCachedWidth =7937
                    LayoutCachedHeight =3402
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
                    Overlaps =1
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =1134
            Name ="PageFooterSection"
            AutoHeight =1
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

Private Sub btn_edit_Click()
    Dim ID As String
    If IsNull(Me.combo_id.Value) Then
        MsgBox "Bitte id auswaehlen zum editieren"
        Exit Sub
    End If
    
    ID = Me.combo_id.Value
    
    DoCmd.OpenForm "edit_simple_testf"
    Forms!edit_simple_testf.txt_id.Value = ID
    Call Form_edit_simple_testf.load_data
    ' DoCmd.Close acForm, Me.name
    
    
    
End Sub

Private Sub btn_show_all_Click()
    DoCmd.OpenForm "show_all_testf"
    DoCmd.Close acForm, Me.name
End Sub

Private Sub Command1_Click()
    DoCmd.OpenForm "add_simple_testf", acNormal
    DoCmd.Close acForm, Me.name
End Sub
