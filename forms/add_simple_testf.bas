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
    ItemSuffix =11
    Right =24150
    Bottom =11820
    RecSrcDt = Begin
        0x97b04a010cf1e540
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
            Height =1134
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =226
                    Top =283
                    Width =3450
                    Height =465
                    FontSize =18
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label9"
                    Caption ="Testf hinzufuegen"
                    GridlineColor =10921638
                    LayoutCachedLeft =226
                    LayoutCachedTop =283
                    LayoutCachedWidth =3676
                    LayoutCachedHeight =748
                End
            End
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
                    Left =1474
                    Top =1077
                    Width =2556
                    Height =525
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_name"
                    GridlineColor =10921638

                    LayoutCachedLeft =1474
                    LayoutCachedTop =1077
                    LayoutCachedWidth =4030
                    LayoutCachedHeight =1602
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1474
                    Top =1722
                    Width =5106
                    Height =2475
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_data"
                    GridlineColor =10921638

                    LayoutCachedLeft =1474
                    LayoutCachedTop =1722
                    LayoutCachedWidth =6580
                    LayoutCachedHeight =4197
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1474
                    Top =4308
                    Width =2556
                    Height =525
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_bem"
                    GridlineColor =10921638

                    LayoutCachedLeft =1474
                    LayoutCachedTop =4308
                    LayoutCachedWidth =4030
                    LayoutCachedHeight =4833
                End
                Begin Label
                    OverlapFlags =85
                    Left =113
                    Top =1133
                    Width =630
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label6"
                    Caption ="Name"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =1133
                    LayoutCachedWidth =743
                    LayoutCachedHeight =1418
                End
                Begin Label
                    OverlapFlags =85
                    Left =113
                    Top =1757
                    Width =765
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label7"
                    Caption ="Tabelle"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =1757
                    LayoutCachedWidth =878
                    LayoutCachedHeight =2042
                End
                Begin Label
                    OverlapFlags =85
                    Left =56
                    Top =4308
                    Width =1140
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label8"
                    Caption ="Bemerkung"
                    GridlineColor =10921638
                    LayoutCachedLeft =56
                    LayoutCachedTop =4308
                    LayoutCachedWidth =1196
                    LayoutCachedHeight =4593
                End
            End
        End
        Begin FormFooter
            Height =1134
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =396
                    ForeColor =4210752
                    Name ="btn_add"
                    Caption ="add"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =396
                    LayoutCachedWidth =2267
                    LayoutCachedHeight =679
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
                    Left =2834
                    Top =396
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btn_menu"
                    Caption ="menu"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2834
                    LayoutCachedTop =396
                    LayoutCachedWidth =4535
                    LayoutCachedHeight =679
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

Private Sub btn_add_Click()
    Dim name
    Dim data
    Dim bem
    
    name = txt_name.Value
    bem = txt_bem.Value
    
    If IsNull(txt_data.Value) Then
        MsgBox "Enter all fields"
        Exit Sub
        
    End If
    
    data = Replace(txt_data.Value, ";", "")
    

    DoCmd.RunSQL "Insert into simple_testf (str_name, str_bem, csv_data) values ('" & name & "','" & bem & "','" & data & "');"
    
    MsgBox "sucessfully added"
    
    DoCmd.OpenForm "menu_testf"
    DoCmd.Close acForm, Me.name
    
    
End Sub

Private Sub btn_menu_Click()
    DoCmd.OpenForm "menu_testf"
    DoCmd.Close acForm, Me.name
End Sub
