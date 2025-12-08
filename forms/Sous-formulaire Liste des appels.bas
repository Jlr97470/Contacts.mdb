Version =19
VersionRequired =19
Begin Form
    MaxButton = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6321
    ItemSuffix =6
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf8a8fa3fefb6e140
    End
    RecordSource ="Appels"
    Caption ="Appels"
    DatasheetFontName ="Tahoma"
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
            Width =850
            Height =850
        End
        Begin Line
            SpecialEffect =3
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
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
            Height =1191
            BackColor =-2147483633
            Name ="Détail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1704
                    Width =825
                    Height =255
                    ColumnWidth =825
                    Name ="DateAppel"
                    ControlSource ="DateAppel"
                    Format ="Short Date"
                    DefaultValue ="=Date()"
                    InputMask ="99/99/00;0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1560
                            Height =255
                            Name ="DateAppel_Étiquette"
                            Caption ="Date de l'appel"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1704
                    Top =342
                    Width =750
                    Height =255
                    ColumnWidth =750
                    TabIndex =1
                    Name ="HeureAppel"
                    ControlSource ="HeureAppel"
                    Format ="Medium Time"
                    DefaultValue ="=Time()"
                    InputMask ="09:00\\ >LL;0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =342
                            Width =1560
                            Height =255
                            Name ="HeureAppel_Étiquette"
                            Caption ="Heure de l'appel"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1704
                    Top =684
                    Width =4560
                    Height =450
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Sujet"
                    ControlSource ="Sujet"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =684
                            Width =1560
                            Height =255
                            Name ="Sujet_Étiquette"
                            Caption ="Sujet"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="PiedFormulaire"
        End
    End
End
