Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
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
    Width =7020
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =2445
    Top =1320
    Right =9465
    Bottom =5505
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5705e01377a4e140
    End
    Caption ="Nouvelle assignation"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin ListBox
            SpecialEffect =2
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin Section
            Height =4200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    ColumnCount =2
                    ListWidth =2160
                    Left =2622
                    Top =236
                    Width =2280
                    Height =262
                    Name ="AssignéA"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW Employés.RéfEmployé, Employés.NomEmployé FROM Employés ORDER "
                        "BY Employés.NomEmployé;"
                    ColumnWidths ="0;2160"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =118
                            Top =236
                            Width =2415
                            Height =240
                            Name ="Label13"
                            Caption ="&Afficher les problèmes assignés à"
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =87
                    AccessKey =83
                    MultiSelect =1
                    ColumnCount =3
                    Left =150
                    Top =1320
                    Width =4752
                    Height =2576
                    TabIndex =2
                    Name ="ListeProblème"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW Suivi.RéfProblème, Suivi.Titre, RecherchePriorité.Priorité, S"
                        "uivi.AssignéA FROM RecherchePriorité RIGHT JOIN Suivi ON RecherchePriorité.RéfPr"
                        "iorité = Suivi.Priorité WHERE (((Suivi.AssignéA)=[Forms]![Réassigner]![AssignéA]"
                        ")) ORDER BY Suivi.RéfProblème;"
                    ColumnWidths ="432;3240;720"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =118
                            Top =1087
                            Width =1905
                            Height =240
                            Name ="lbProblèmes"
                            Caption ="&Sélectionnez un problème"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    AccessKey =80
                    ColumnCount =2
                    ListWidth =2160
                    Left =2622
                    Top =708
                    Width =2280
                    Height =262
                    TabIndex =1
                    Name ="NouvelleAssignation"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW Employés.RéfEmployé, Employés.NomEmployé FROM Employés ORDER "
                        "BY Employés.NomEmployé;"
                    ColumnWidths ="0;2160"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =118
                            Top =614
                            Width =2094
                            Height =249
                            Name ="Label3"
                            Caption ="Réassigner le &problème à"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =84
                    Left =5102
                    Top =874
                    Width =1746
                    Height =420
                    TabIndex =4
                    Name ="ToutRéassigner"
                    Caption ="&Tout réassigner"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =5100
                    Top =240
                    Width =1747
                    Height =510
                    TabIndex =3
                    Name ="NouvelleAssignationSélectionnée"
                    Caption ="&Nouvelle assignation sélectionnée"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =70
                    Left =5102
                    Top =1417
                    Width =1741
                    Height =393
                    TabIndex =5
                    Name ="FermerFormulaire"
                    Caption ="&Fermer"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "Réassigner.cls"
