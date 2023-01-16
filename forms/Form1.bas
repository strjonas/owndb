Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13152
    DatasheetFontHeight =11
    ItemSuffix =10
    Right =25700
    Bottom =11960
    RecSrcDt = Begin
        0x797729db52f0e540
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7362
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =737
                    Top =737
                    Width =10434
                    Height =2950
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSource"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =737
                    LayoutCachedTop =737
                    LayoutCachedWidth =11171
                    LayoutCachedHeight =3687
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =800
                    Top =3911
                    Width =10366
                    Height =2950
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtHTML"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =800
                    LayoutCachedTop =3911
                    LayoutCachedWidth =11166
                    LayoutCachedHeight =6861
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

Private Sub txtHTML_Change()
    txtSource = txtHTML.Text
End Sub

Private Sub txtSource_Change()
    txtHTML = txtSource.Text
End Sub
