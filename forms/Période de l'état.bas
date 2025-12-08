Version =19
VersionRequired =19
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3849
    DatasheetFontHeight =10
    ItemSuffix =5
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa04d4940efb6e140
    End
    Caption ="Période de l'état"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
            Height =1659
            BackColor =-2147483633
            Name ="Détail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =2046
                    Top =171
                    Width =825
                    Height =255
                    ColumnWidth =825
                    Name ="Date premier appel"
                    Format ="Short Date"
                    InputMask ="99/99/00"
                    EventProcPrefix ="Date_premier_appel"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =171
                            Width =1560
                            Height =255
                            Name ="Date premier appel_Étiquette"
                            Caption ="Date 1er appel"
                            EventProcPrefix ="Date_premier_appel_Étiquette"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2046
                    Top =513
                    Width =825
                    Height =255
                    ColumnWidth =825
                    TabIndex =1
                    Name ="Date dernier appel"
                    Format ="Short Date"
                    InputMask ="99/99/00"
                    EventProcPrefix ="Date_dernier_appel"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =513
                            Width =1560
                            Height =255
                            Name ="Date dernier appel_Étiquette"
                            Caption ="Date dernier appel"
                            EventProcPrefix ="Date_dernier_appel_Étiquette"
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    Left =2907
                    Top =1083
                    Width =885
                    Height =405
                    TabIndex =2
                    Name ="Aperçu"
                    Caption ="&Aperçu..."
                    OnClick ="[Event Procedure]"
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
' See "Période de l'état.cls"
