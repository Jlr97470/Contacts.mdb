Version =19
VersionRequired =19
Begin Form
    MaxButton = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4242
    DatasheetFontHeight =10
    ItemSuffix =4
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4b0f3040efb6e140
    End
    RecordSource ="Types de contacts"
    Caption ="Types de contacts"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
            Height =1281
            BackColor =-2147483633
            Name ="Détail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1875
                    Top =342
                    Height =255
                    ColumnWidth =1701
                    Name ="RéfTypeContact"
                    ControlSource ="RéfTypeContact"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =342
                            Width =1560
                            Height =255
                            Name ="RéfTypeContact_Étiquette"
                            Caption ="Réf type de contact"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1875
                    Top =684
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="TypeContact"
                    ControlSource ="TypeContact"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =171
                            Top =684
                            Width =1560
                            Height =255
                            Name ="TypeContact_Étiquette"
                            Caption ="Type de contact"
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
CodeBehindForm
' See "Types de contacts.cls"
