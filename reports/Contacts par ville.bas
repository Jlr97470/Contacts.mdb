Version =19
VersionRequired =19
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9070
    DatasheetFontHeight =10
    ItemSuffix =25
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x07b623aaefb6e140
    End
    RecordSource ="Contacts"
    Caption ="Contacts par ville"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Ville"
        End
        Begin BreakLevel
            ControlSource ="NomFamille"
        End
        Begin BreakLevel
            ControlSource ="Prenom"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =938
            Name ="EntêteÉtat"
            Begin
                Begin Label
                    BackStyle =1
                    Left =57
                    Top =57
                    Width =2970
                    Height =510
                    FontSize =20
                    Name ="Étiquette12"
                    Caption ="Contacts par ville"
                End
                Begin Line
                    BorderWidth =3
                    Top =57
                    Width =9070
                    BorderColor =12632256
                    Name ="Trait15"
                End
                Begin Line
                    BorderWidth =3
                    Top =86
                    Width =9070
                    BorderColor =12632256
                    Name ="Trait16"
                End
                Begin Line
                    BorderWidth =3
                    Top =909
                    Width =9070
                    BorderColor =12632256
                    Name ="Trait17"
                End
                Begin Line
                    BorderWidth =3
                    Top =938
                    Width =9070
                    BorderColor =12632256
                    Name ="Trait18"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="ZoneEntêtePage"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =870
            Name ="EntêteGroupe0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =2037
                    Top =57
                    Width =2760
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="Ville"
                    ControlSource ="Ville"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            Left =57
                            Top =57
                            Width =1860
                            Height =345
                            Name ="Ville_Étiquette"
                            Caption ="Ville"
                        End
                    End
                End
                Begin Label
                    Left =1083
                    Top =513
                    Width =1368
                    Height =300
                    Name ="NomFamille_Étiquette"
                    Caption ="Nom"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2508
                    Top =513
                    Width =1368
                    Height =300
                    Name ="Prénom_Étiquette"
                    Caption ="Prénom"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3933
                    Top =513
                    Width =1368
                    Height =300
                    Name ="NomSociété_Étiquette"
                    Caption ="Nom de la société"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =5358
                    Top =513
                    Width =1368
                    Height =300
                    Name ="Titre_Étiquette"
                    Caption ="Titre"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6783
                    Top =513
                    Width =2230
                    Height =300
                    Name ="TélProfessionnel_Étiquette"
                    Caption ="Tél. Prof."
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1083
                    Top =513
                    Width =7930
                    Name ="Trait19"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1083
                    Top =484
                    Width =7930
                    Name ="Trait20"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1083
                    Top =813
                    Width =7930
                    Name ="Trait21"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1083
                    Top =842
                    Width =7930
                    Name ="Trait22"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =384
            Name ="Détail"
            Begin
                Begin TextBox
                    Left =1083
                    Top =57
                    Width =1368
                    Height =270
                    Name ="NomFamille"
                    ControlSource ="NomFamille"
                End
                Begin TextBox
                    Left =2508
                    Top =57
                    Width =1368
                    Height =270
                    TabIndex =1
                    Name ="Prénom"
                    ControlSource ="Prenom"
                End
                Begin TextBox
                    Left =3933
                    Top =57
                    Width =1368
                    Height =270
                    TabIndex =2
                    Name ="NomSociété"
                    ControlSource ="NomSociete"
                End
                Begin TextBox
                    Left =5358
                    Top =57
                    Width =1368
                    Height =270
                    TabIndex =3
                    Name ="Titre"
                    ControlSource ="Titre"
                End
                Begin TextBox
                    Left =6783
                    Top =57
                    Width =2230
                    Height =270
                    TabIndex =4
                    Name ="TélProfessionnel"
                    ControlSource ="TelProfessionnel"
                    InputMask ="00\\ 00\\ 00\\ 00\\ 00"
                End
            End
        End
        Begin PageFooter
            Height =528
            Name ="ZonePiedPage"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =57
                    Top =228
                    Width =4421
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Texte13"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =4592
                    Top =228
                    Width =4421
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Texte14"
                    ControlSource ="=\"Page \" & [Page] & \" sur \" & [Pages]"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Width =9070
                    BorderColor =12632256
                    Name ="Trait23"
                End
                Begin Line
                    BorderWidth =3
                    Top =28
                    Width =9070
                    BorderColor =12632256
                    Name ="Trait24"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="PiedÉtat"
        End
    End
End
