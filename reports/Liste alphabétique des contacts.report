Version =19
VersionRequired =19
Begin Report
    LayoutForPrint = NotDefault
    MaxButton = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13943
    DatasheetFontHeight =10
    ItemSuffix =18
    Top =600
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xf49c42ebf1b6e140
    End
    RecordSource ="SELECT DISTINCTROW Contacts.* FROM Contacts; "
    Caption ="Liste alphabétique des contacts"
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
            GroupOn =1
            ControlSource ="NomFamille"
        End
        Begin BreakLevel
            ControlSource ="NomFamille"
        End
        Begin BreakLevel
            ControlSource ="=[Prenom]"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =909
            Name ="EntêteÉtat"
            Begin
                Begin Label
                    BackStyle =1
                    Left =57
                    Top =57
                    Width =5220
                    Height =510
                    FontSize =20
                    Name ="Étiquette15"
                    Caption ="Liste alphabétique des contacts"
                End
            End
        End
        Begin PageHeader
            Height =456
            Name ="ZoneEntêtePage"
            Begin
                Begin Label
                    Left =1026
                    Top =57
                    Width =2700
                    Height =300
                    Name ="Nom contact_Étiquette"
                    Caption ="Nom contact"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Nom_contact_Étiquette"
                End
                Begin Label
                    Left =3783
                    Top =57
                    Width =2700
                    Height =300
                    Name ="NomSociété_Étiquette"
                    Caption ="Nom de la société"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6540
                    Top =57
                    Width =2000
                    Height =300
                    Name ="Titre_Étiquette"
                    Caption ="Titre"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8597
                    Top =57
                    Width =1700
                    Height =300
                    Name ="TélProfessionnel_Étiquette"
                    Caption ="Tél. Prof."
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =10354
                    Top =57
                    Width =700
                    Height =300
                    Name ="PosteTravail_Étiquette"
                    Caption ="Ext."
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =11111
                    Top =57
                    Width =1700
                    Height =300
                    Name ="NumTélécopie_Étiquette"
                    Caption ="N° de télécopie"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =57
                    Width =13774
                    Name ="Ligne1"
                End
                Begin Line
                    BorderWidth =2
                    Left =57
                    Top =399
                    Width =13774
                    Name ="Ligne2"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =387
            Name ="EntêteGroupe"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =57
                    Width =1000
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="PremièreLettre"
                    ControlSource ="=left([NomFamille],1)"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            Name ="Détail"
            Begin
                Begin TextBox
                    Left =1026
                    Width =2700
                    Height =270
                    Name ="Nom contact"
                    ControlSource ="=[NomFamille] & \" ,\" & [Prenom]"
                    EventProcPrefix ="Nom_contact"
                End
                Begin TextBox
                    Left =3783
                    Width =2700
                    Height =270
                    TabIndex =1
                    Name ="NomSociété"
                    ControlSource ="NomSociete"
                End
                Begin TextBox
                    Left =6540
                    Width =2000
                    Height =270
                    TabIndex =2
                    Name ="Titre"
                    ControlSource ="Titre"
                End
                Begin TextBox
                    Left =8597
                    Width =1700
                    Height =270
                    TabIndex =3
                    Name ="TélProfessionnel"
                    ControlSource ="TelProfessionnel"
                    InputMask ="00\\ 00\\ 00\\ 00\\ 00"
                End
                Begin TextBox
                    Left =10354
                    Width =700
                    Height =270
                    TabIndex =4
                    Name ="PosteTravail"
                    ControlSource ="PosteTravail"
                End
                Begin TextBox
                    Left =11111
                    Width =1700
                    Height =270
                    TabIndex =5
                    Name ="NumTélécopie"
                    ControlSource ="=[NumTelecopie]"
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
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Texte16"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =8903
                    Top =228
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Texte17"
                    ControlSource ="=\"Page \" & [Page] & \" sur \" & [Pages]"
                    FontName ="Times New Roman"
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
CodeBehindForm
' See "Liste alphabétique des contacts.cls"
