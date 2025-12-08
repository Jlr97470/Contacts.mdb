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
    Width =14000
    DatasheetFontHeight =10
    ItemSuffix =20
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    RecordSource ="SELECT DISTINCTROW Contacts.Prénom, Contacts.NomFamille, Appels.DateAppel, Appel"
        "s.Sujet, Appels.HeureAppel, Contacts.TélProfessionnel, Contacts.PosteTravail, Co"
        "ntacts.TéléphoneMobile, Contacts.NumTélécopie, Contacts.NomCourrierÉlec FROM Con"
        "tacts INNER JOIN Appels ON [Contacts].[RéfContact]=[Appels].[RéfContact] WHERE ("
        "([Appels].[DateAppel]>=[forms]![Période de l'état]![Date premier appel] And [App"
        "els].[DateAppel]<=[forms]![Période de l'état]![Date dernier appel])); "
    Caption ="Statistiques appels hebdo"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            Width =1701
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ListBox
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="DateAppel"
        End
        Begin BreakLevel
            ControlSource ="HeureAppel"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1866
            Name ="EntêteÉtat"
            Begin
                Begin TextBox
                    OverlapFlags =4
                    Left =57
                    Top =1197
                    Width =5760
                    Height =270
                    Name ="Pér de l'état"
                    ControlSource ="=\"Pour les appels effectués entre : \" & [forms]![Période de l'état]![Date prem"
                        "ier appel] & \" et \" & [forms]![Période de l'état]![Date dernier appel]"
                    EventProcPrefix ="Pér_de_l_état"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =4
                    Left =57
                    Top =57
                    Width =4230
                    Height =510
                    FontSize =20
                    Name ="Étiquette17"
                    Caption ="Statistiques appels hebdo"
                End
            End
        End
        Begin PageHeader
            Height =456
            Name ="ZoneEntêtePage"
            Begin
                Begin Label
                    OverlapFlags =4
                    TextAlign =3
                    Left =57
                    Top =57
                    Width =900
                    Height =300
                    Name ="DateAppel_Étiquette"
                    Caption ="Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =3
                    Left =912
                    Top =57
                    Width =900
                    Height =300
                    Name ="HeureAppel_Étiquette"
                    Caption ="Heure"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =4
                    Left =1869
                    Top =57
                    Width =2600
                    Height =300
                    Name ="Nom contact_Étiquette"
                    Caption ="Nom contact"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Nom_contact_Étiquette"
                End
                Begin Label
                    OverlapFlags =4
                    Left =4526
                    Top =57
                    Width =3400
                    Height =300
                    Name ="Sujet_Étiquette"
                    Caption ="Sujet"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =4
                    Left =7983
                    Top =57
                    Width =1700
                    Height =300
                    Name ="TélProfessionnel_Étiquette"
                    Caption ="Tél. Prof."
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =4
                    Left =9740
                    Top =57
                    Width =700
                    Height =300
                    Name ="PosteTravail_Étiquette"
                    Caption ="Ext."
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =4
                    Left =10497
                    Top =57
                    Width =1700
                    Height =300
                    Name ="NumTélécopie_Étiquette"
                    Caption ="N° de télécopie"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =4
                    Left =57
                    Width =13774
                    Name ="Ligne1"
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =4
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
            Name ="EntêteGroupe0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OverlapFlags =4
                    TextFontFamily =18
                    Left =57
                    Width =900
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="DateAppel"
                    ControlSource ="DateAppel"
                    Format ="Short Date"
                    FontName ="Times New Roman"
                    InputMask ="99/99/00;0"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =327
            Name ="Détail"
            Begin
                Begin TextBox
                    Left =912
                    Width =900
                    Height =270
                    Name ="HeureAppel"
                    ControlSource ="HeureAppel"
                    Format ="Medium Time"
                    InputMask ="09:00\\ >LL;0"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    Left =1869
                    Width =2600
                    Height =270
                    TabIndex =1
                    Name ="Nom contact"
                    ControlSource ="=[NomFamille] & \", \" & [Prénom]"
                    EventProcPrefix ="Nom_contact"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    Left =4526
                    Width =3400
                    Height =270
                    TabIndex =2
                    Name ="Sujet"
                    ControlSource ="Sujet"
                End
                Begin TextBox
                    Left =7983
                    Width =1700
                    Height =270
                    TabIndex =3
                    Name ="TélProfessionnel"
                    ControlSource ="TélProfessionnel"
                    InputMask ="00\\ 00\\ 00\\ 00\\ 00"
                End
                Begin TextBox
                    Left =9740
                    Width =700
                    Height =270
                    TabIndex =4
                    Name ="PosteTravail"
                    ControlSource ="PosteTravail"
                End
                Begin TextBox
                    Left =10497
                    Width =1700
                    Height =270
                    TabIndex =5
                    Name ="NumTélécopie"
                    ControlSource ="NumTélécopie"
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
                    OverlapFlags =4
                    TextAlign =1
                    TextFontFamily =18
                    Left =57
                    Top =228
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Texte18"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OverlapFlags =4
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
                    Name ="Texte19"
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
' See "Statistiques appels hebdo.cls"
