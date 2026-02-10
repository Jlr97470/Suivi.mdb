Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ShortcutMenu = NotDefault
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
    Width =5940
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =2985
    Top =1800
    Right =8925
    Bottom =4545
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x922c12263142e140
    End
    Caption ="Saisie du nom de l'employé"
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin Section
            Height =2760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =732
                    Top =188
                    Width =4320
                    Height =390
                    FontSize =14
                    Name ="Label0"
                    Caption ="Bienvenue dans l'application Suivi"
                    FontName ="Arial"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =2880
                    Left =1380
                    Top =1320
                    Width =2828
                    Name ="cbxRéfUtilisateur"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW Employés.RéfEmployé, Employés.NomEmployé FROM Employés ORDER "
                        "BY Employés.NomEmployé;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1380
                            Top =780
                            Width =2409
                            Height =472
                            Name ="Label4"
                            Caption ="Veuillez saisir votre nom ou le sélectionner dans la liste :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =661
                    Top =2149
                    Width =2130
                    Height =405
                    TabIndex =1
                    Name ="AfficherTout"
                    Caption ="&Afficher tous les problèmes"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =77
                    Left =3180
                    Top =2160
                    Width =1875
                    Height =405
                    TabIndex =2
                    Name ="FiltrerProblèmes"
                    Caption ="Afficher &mes problèmes"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "DialogueNomEmployé.cls"
