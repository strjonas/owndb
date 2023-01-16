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
    ItemSuffix =3
    Right =17610
    Bottom =11850
    RecSrcDt = Begin
        0x79650489f6f0e540
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
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =566
                    Top =1984
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txt_id"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =1984
                    LayoutCachedWidth =2267
                    LayoutCachedHeight =2299
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =2891
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btn_show"
                    Caption ="Command2"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =2891
                    LayoutCachedWidth =2267
                    LayoutCachedHeight =3174
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

Private Sub btn_show_Click()
    Dim ID As Integer
    ID = Me!txt_id.Value

    Dim name As String
    Dim data As String
    Dim bem As String
    
    Dim sql As String
    
    ' Get Data
    Dim temp()
    temp = get_all("csv_data,str_name,str_bem", "simple_testf", "id=" & ID)
    name = temp(0, 1)
    data = Replace(Replace(temp(0, 0), """", ""), ";", "")
    bem = temp(0, 2)
    Erase temp
    
    ' Clear the temp db, otherwise there will be old values if theyre not overwritten
    DoCmd.RunSQL "delete from temp_testf where key='1';"
    DoCmd.RunSQL "insert into temp_testf (key) values ('1');"
    
    ' Inserting name und bem into the temp table
    sql = "update temp_testf set s_bem='" & bem & "' where key='1';" ' key=1 ist einfach die default row im temp table die immer genutzt wird
    DoCmd.RunSQL sql
    sql = "update temp_testf set s_name='" & name & "' where key='1';"
    DoCmd.RunSQL sql
    
    ' Prepare Data
    Dim rows() As String
    Dim col() As String
    Dim colc
    
    
    ' Splitting by new line
    rows = Split(data, vbCr)
    
    
    ' Der loop fuellte die Tabelle
    For rc = 1 To get_len(rows, 1)
        col = Split(rows(rc - 1), ",")
        colc = 0
        ' Iterating through each element in row for each row and updating all values in the temp table
        For Each el In col
            If el = "" Then
                el = " "
            End If
            
            ' Updating each column, names have the form str_rownum_columnnum, temp_tesf hat nur eine row mit id 1, die immer ueberschrieben wird
            sql = "update temp_testf set str_" & rc & "_" & colc & "='" & el & "' where key='1';"
            DoCmd.RunSQL sql
            colc = colc + 1
        
        Next el
    
    Next rc
    
    
    
    
    '  Now open the report:
    DoCmd.OpenReport "rpt_testf", acViewPreview
    
    

End Sub
