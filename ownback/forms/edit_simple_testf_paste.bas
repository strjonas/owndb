Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9127
    DatasheetFontHeight =11
    Left =4650
    Top =3165
    Right =22260
    Bottom =15015
    RecSrcDt = Begin
        0xa7f718252ef1e540
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
        Begin Section
            Height =5952
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1417
                    Top =1077
                    Width =6526
                    Height =3465
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_data"
                    GridlineColor =10921638

                    LayoutCachedLeft =1417
                    LayoutCachedTop =1077
                    LayoutCachedWidth =7943
                    LayoutCachedHeight =4542
                End
                Begin Label
                    OverlapFlags =85
                    Left =56
                    Top =1112
                    Width =765
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label7"
                    Caption ="Tabelle"
                    GridlineColor =10921638
                    LayoutCachedLeft =56
                    LayoutCachedTop =1112
                    LayoutCachedWidth =821
                    LayoutCachedHeight =1397
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1417
                    Top =4662
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btn_add"
                    Caption ="save"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1417
                    LayoutCachedTop =4662
                    LayoutCachedWidth =3118
                    LayoutCachedHeight =4945
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
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7200
                    Top =283
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_id"
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =283
                    LayoutCachedWidth =8901
                    LayoutCachedHeight =598
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

' REQUIRES PARENT FROM (OHNE PASTE) OPEN, BECAUSE IT ACESSES ITS VARIABLES


Private Sub btn_add_Click()
    
    Call Form_edit_simple_testf.paste_form_save
    
    MsgBox "sucessfully saved"
    DoCmd.Close acForm, Me.name
End Sub
