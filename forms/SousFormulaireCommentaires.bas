Version =19
VersionRequired =19
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7797
    RowHeight =432
    ItemSuffix =12
    Left =1200
    Top =3945
    Right =8445
    Bottom =5535
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7ffcf4914ebde140
    End
    RecordSource ="RequêteCommentaires"
    Caption ="Commentaires"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin Section
            Height =2542
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =2220
                    Top =120
                    ColumnOrder =0
                    Name ="RéfCommentaire"
                    ControlSource ="RéfCommentaire"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =1335
                            Height =240
                            Name ="Label1"
                            Caption ="Réf. Commentaire"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =2220
                    Top =480
                    ColumnOrder =1
                    TabIndex =1
                    Name ="RéfProblème"
                    ControlSource ="RéfProblème"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =1080
                            Height =240
                            Name ="Label3"
                            Caption ="Réf. Problème"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =2220
                    Top =1200
                    ColumnWidth =780
                    ColumnOrder =3
                    TabIndex =3
                    Name ="DateSaisie"
                    ControlSource ="DateSaisie"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1200
                            Width =1100
                            Height =249
                            Name ="Label7"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =2220
                    Top =1560
                    Width =5433
                    Height =862
                    ColumnWidth =3940
                    ColumnOrder =5
                    TabIndex =4
                    Name ="Commentaires"
                    ControlSource ="Commentaires"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1560
                            Width =990
                            Height =240
                            Name ="Label9"
                            Caption ="Commentaire"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =2220
                    Top =840
                    Width =2814
                    ColumnWidth =1244
                    ColumnOrder =2
                    TabIndex =2
                    Name ="NomEmployé"
                    ControlSource ="NomEmployé"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =840
                            Width =877
                            Height =249
                            Name ="Label5"
                            Caption ="Employé"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5700
                    Top =1200
                    ColumnWidth =877
                    ColumnOrder =4
                    TabIndex =5
                    Name ="Etat"
                    ControlSource ="Commentaires.Etat"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4260
                            Top =1200
                            Width =1047
                            Height =249
                            Name ="Label11"
                            Caption ="État"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SousFormulaireCommentaires.cls"
