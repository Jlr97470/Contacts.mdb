Version =19
VersionRequired =19
Begin Form
    MaxButton = NotDefault
    AutoCenter = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =2
    GridY =10
    Width =10374
    DatasheetFontHeight =10
    ItemSuffix =66
    Left =2100
    Top =1560
    Right =12765
    Bottom =5895
    DatasheetGridlinesColor =12632256
    OnFilter ="[Event Procedure]"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xaf7bb53fefb6e140
    End
    RecordSource ="Contacts"
    Caption ="Contacts"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin PageBreak
            Width =283
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
            Width =5103
            Height =3402
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="EntêteFormulaire"
        End
        Begin Section
            Height =7638
            BackColor =-2147483633
            Name ="Détail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1797
                    Top =171
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    Name ="Prenom"
                    ControlSource ="Prenom"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =171
                            Width =1482
                            Height =255
                            Name ="Prénom_Étiquette"
                            Caption ="Prénom"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1797
                    Top =456
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="NomFamille"
                    ControlSource ="NomFamille"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =456
                            Width =1482
                            Height =255
                            Name ="NomFamille_Étiquette"
                            Caption ="Nom"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1797
                    Top =741
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="NomSociété"
                    ControlSource ="NomSociete"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =741
                            Width =1482
                            Height =255
                            Name ="NomSociété_Étiquette"
                            Caption ="Société"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1797
                    Top =1026
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="Nomfamilier"
                    ControlSource ="Nomfamilier"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =1026
                            Width =1482
                            Height =255
                            Name ="Nom familier_Étiquette"
                            Caption ="Nom familier"
                            EventProcPrefix ="Nom_familier_Étiquette"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =1797
                    Top =1311
                    Width =2763
                    Height =450
                    TabIndex =4
                    Name ="Adresse"
                    ControlSource ="Adresse"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =1311
                            Width =1482
                            Height =255
                            Name ="Adresse_Étiquette"
                            Caption ="Adresse"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    Left =1797
                    Top =1767
                    Width =1560
                    Height =255
                    ColumnWidth =1560
                    TabIndex =5
                    Name ="CodePostal"
                    ControlSource ="CodePostal"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="00000;;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =1767
                            Width =1482
                            Height =255
                            Name ="CodePostal_Étiquette"
                            Caption ="Code postal"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1797
                    Top =2052
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Ville"
                    ControlSource ="Ville"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =2052
                            Width =1482
                            Height =255
                            Name ="Ville_Étiquette"
                            Caption ="Ville"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1797
                    Top =2337
                    Width =1560
                    Height =255
                    ColumnWidth =1560
                    TabIndex =7
                    Name ="DepartementOuRegion"
                    ControlSource ="DepartementOuRegion"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =2337
                            Width =1482
                            Height =255
                            Name ="DépartementOuRégion_Étiquette"
                            Caption ="Département"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1797
                    Top =2622
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="Pays"
                    ControlSource ="Pays"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =2622
                            Width =1482
                            Height =255
                            Name ="Pays_Étiquette"
                            Caption ="Pays"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6243
                    Top =171
                    Width =1566
                    Height =255
                    ColumnWidth =1701
                    TabIndex =9
                    Name ="RéfContact"
                    ControlSource ="RefContact"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4617
                            Top =171
                            Width =1482
                            Height =255
                            Name ="RéfContact_Étiquette"
                            Caption ="Réf contact"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6243
                    Top =456
                    Width =1566
                    Height =255
                    TabIndex =10
                    Name ="Titre"
                    ControlSource ="Titre"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4617
                            Top =456
                            Width =1482
                            Height =255
                            Name ="Titre_Étiquette"
                            Caption ="Titre"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6243
                    Top =741
                    Width =1566
                    Height =255
                    TabIndex =11
                    Name ="TélProfessionnel"
                    ControlSource ="TelProfessionnel"
                    InputMask ="00\\ 00\\ 00\\ 00\\ 00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4617
                            Top =741
                            Width =1482
                            Height =255
                            Name ="TélProfessionnel_Étiquette"
                            Caption ="Tél. Prof."
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6243
                    Top =1026
                    Width =1560
                    Height =255
                    ColumnWidth =1560
                    TabIndex =12
                    Name ="PosteTravail"
                    ControlSource ="PosteTravail"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4617
                            Top =1026
                            Width =1482
                            Height =255
                            Name ="PosteTravail_Étiquette"
                            Caption ="Poste travail"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6243
                    Top =1311
                    Width =1566
                    Height =255
                    TabIndex =13
                    Name ="TéléphoneMobile"
                    ControlSource ="TelMobile"
                    InputMask ="00\\ 00\\ 00\\ 00\\ 00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4617
                            Top =1311
                            Width =1482
                            Height =255
                            Name ="TéléphoneMobile_Étiquette"
                            Caption ="Tél. mobile"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6243
                    Top =1596
                    Width =1566
                    Height =255
                    TabIndex =14
                    Name ="NumTélécopie"
                    ControlSource ="NumTelecopie"
                    InputMask ="00\\ 00\\ 00\\ 00\\ 00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4617
                            Top =1596
                            Width =1482
                            Height =255
                            Name ="NumTélécopie_Étiquette"
                            Caption ="N° de télécopie"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1875
                    Top =3990
                    Width =1600
                    Height =255
                    TabIndex =15
                    Name ="Nom contact"
                    ControlSource ="=NomComplet()"
                    EventProcPrefix ="Nom_contact"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =3990
                            Width =1560
                            Height =255
                            Name ="Nom contact_Étiquette"
                            Caption ="Nom contact"
                            EventProcPrefix ="Nom_contact_Étiquette"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =2835
                    Left =1875
                    Top =4275
                    Width =1600
                    Height =255
                    TabIndex =16
                    Name ="RéfTypeContact"
                    ControlSource ="RéfTypeContact"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW [Types de contacts].* FROM [Types de contacts] ORDER BY [Type"
                        "s de contacts].[TypeContact]; "
                    ColumnWidths ="0;2880"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =4275
                            Width =1560
                            Height =255
                            Name ="RéfTypeContact_Étiquette"
                            Caption ="Type contact"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1875
                    Top =4560
                    Width =1600
                    Height =255
                    TabIndex =17
                    Name ="NomCourrierÉlec"
                    ControlSource ="NomCourrierÉlec"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =4560
                            Width =1560
                            Height =255
                            Name ="NomCourrierÉlec_Étiquette"
                            Caption ="Email"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1875
                    Top =4845
                    Width =1600
                    Height =255
                    TabIndex =18
                    Name ="RecommandéPar"
                    ControlSource ="RecommandéPar"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =4845
                            Width =1560
                            Height =255
                            Name ="RecommandéPar_Étiquette"
                            Caption ="Recommandé par"
                        End
                    End
                End
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1875
                    Top =5130
                    Width =2800
                    Height =840
                    TabIndex =19
                    Name ="Remarques"
                    ControlSource ="Remarques"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =5130
                            Width =1560
                            Height =255
                            Name ="Remarques_Étiquette"
                            Caption ="Remarques"
                        End
                    End
                End
                Begin PageBreak
                    OverlapFlags =85
                    Top =3819
                    Name ="Saut de page45"
                    EventProcPrefix ="Saut_de_page45"
                End
                Begin ComboBox
                    OverlapFlags =85
                    AccessKey =84
                    ColumnCount =2
                    ListWidth =2325
                    Left =7370
                    Top =2324
                    Width =1927
                    Height =397
                    TabIndex =20
                    BackColor =12632256
                    ForeColor =0
                    Name ="Modifiable50"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contacts.[RefContact], [NomFamille] & \", \" & [Prenom] AS Expr1 FROM Con"
                        "tacts ORDER BY [NomFamille] & \", \" & [Prenom]; "
                    ColumnWidths ="0;2325"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =5039
                            Top =2324
                            Width =2160
                            Height =240
                            FontWeight =700
                            Name ="&Trouver enregistrement :_Étiquette"
                            Caption ="&Trouver enregistrement :"
                            EventProcPrefix ="Ctl_Trouver_enregistrement___Étiquette"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =6009
                    Top =3174
                    Width =1937
                    Height =943
                    TabIndex =21
                    Name ="OptionsFiltre"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =0
                            Left =6129
                            Top =3054
                            Width =1350
                            Height =240
                            Name ="Étiquette60"
                            Caption ="Options de filtrage"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            AccessKey =84
                            Left =6195
                            Top =3412
                            OptionValue =1
                            Name ="Option62"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6425
                                    Top =3384
                                    Width =1335
                                    Height =240
                                    Name ="Étiquette63"
                                    Caption ="&Tous les contacts"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            AccessKey =67
                            Left =6195
                            Top =3742
                            OptionValue =2
                            Name ="Option64"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6425
                                    Top =3714
                                    Width =1095
                                    Height =240
                                    Name ="Étiquette65"
                                    Caption ="&Contacts Paris"
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =519
            BackColor =-2147483633
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =171
                    Top =57
                    Width =855
                    Height =405
                    Name ="Appels"
                    Caption ="&Appels..."
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =77
                    Left =1083
                    Top =57
                    Width =1080
                    Height =405
                    TabIndex =1
                    Name ="Composer"
                    Caption ="Co&mposer..."
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2494
                    Top =56
                    Width =495
                    Height =240
                    Name ="Page"
                    Caption ="Page:"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =49
                    Left =3061
                    Top =56
                    Width =380
                    Height =400
                    TabIndex =2
                    Name ="Page1"
                    Caption ="&1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =50
                    Left =3514
                    Top =56
                    Width =380
                    Height =400
                    TabIndex =3
                    Name ="Page2"
                    Caption ="&2"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "Contacts.cls"
