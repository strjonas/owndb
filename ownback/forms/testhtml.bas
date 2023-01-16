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
    ItemSuffix =1
    Right =17610
    Bottom =11850
    RecSrcDt = Begin
        0x24c986f8f3f0e540
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
        Begin Section
            Height =5952
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =3741
                    Top =2777
                    ForeColor =4210752
                    Name ="Command0"
                    Caption ="Command0"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3741
                    LayoutCachedTop =2777
                    LayoutCachedWidth =5442
                    LayoutCachedHeight =3060
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

Private Sub Command0_Click()
    DoCmd.OutputTo acOutputReport, "testhtml", _
       acFormatTXT, "G:\Meine Ablage\Praxis1\dev\playground\rep.html"
       
       Dim fileHTML As String
    Dim fileRTF As String
    Dim pathApp As String
    
    pathApp = CurrentProject.Path
    fileHTML = pathApp & "\sample.html"
    fileRTF = pathApp & "\sample.rtf"
    
    Dim theWord As Object
    
    Set theWord = CreateObject("Word.Application")
    theWord.Documents.Open FileName:=fileHTML
    
    theWord.Documents(fileHTML).SaveAs _
      FileName:=fileRTF, FileFormat:=wdFormatRTF
      
      theWord.Documents(fileRTF).Close
    theWord.Quit
    Set theWord = Nothing
End Sub
