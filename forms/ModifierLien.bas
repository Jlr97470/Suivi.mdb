Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5340
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =1485
    Top =2775
    Right =7560
    Bottom =6195
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa27716451346e140
    End
    Caption ="Modifier le lien hypertexte"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin Section
            Height =1800
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1253
                    Top =120
                    Width =3913
                    Height =251
                    Name ="DisplayText"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =123
                            Top =127
                            Width =1073
                            Height =249
                            Name ="Label1"
                            Caption ="Texte du lien"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1263
                    Top =480
                    Width =3901
                    TabIndex =1
                    Name ="Address"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =483
                            Width =772
                            Height =249
                            Name ="Label3"
                            Caption ="Adresse"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1263
                    Top =840
                    Width =3901
                    TabIndex =2
                    Name ="Subaddress"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =837
                            Width =1073
                            Height =248
                            Name ="Label5"
                            Caption ="Sous-adresse"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =2760
                    Top =1320
                    Width =1323
                    TabIndex =4
                    Name ="Annuler"
                    Caption ="Annuler"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =1320
                    Top =1320
                    Width =1323
                    TabIndex =3
                    Name ="OK"
                    Caption ="OK"
                End
            End
        End
    End
End
CodeBehindForm
' See "ModifierLien.cls"
