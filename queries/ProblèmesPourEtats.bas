Operation =1
Option =8
Where ="(((Suivi.Etat)<3))"
Begin InputTables
    Name ="Employés"
    Name ="RechercheEtat"
    Name ="RecherchePriorité"
    Name ="RechercheRésolution"
    Name ="Employés"
    Alias ="Employés_1"
    Name ="Employés"
    Alias ="Employés_2"
    Name ="Suivi"
End
Begin OutputColumns
    Expression ="Suivi.RéfProblème"
    Expression ="Suivi.Titre"
    Expression ="Suivi.Description"
    Expression ="RechercheEtat.Etat"
    Expression ="RecherchePriorité.Priorité"
    Expression ="Suivi.DateSaisie"
    Alias ="Saisi par"
    Expression ="Employés_1.NomEmployé"
    Alias ="AssignéA"
    Expression ="Employés_2.NomEmployé"
    Expression ="Suivi.DateRésolution"
    Expression ="RechercheRésolution.Résolution"
    Alias ="Résolu par"
    Expression ="Employés.NomEmployé"
End
Begin Joins
    LeftTable ="RecherchePriorité"
    RightTable ="Suivi"
    Expression ="RecherchePriorité.RéfPriorité = Suivi.Priorité"
    Flag =3
    LeftTable ="Suivi"
    RightTable ="Employés_2"
    Expression ="Suivi.AssignéA = Employés_2.RéfEmployé"
    Flag =2
    LeftTable ="RechercheRésolution"
    RightTable ="Suivi"
    Expression ="RechercheRésolution.RéfRésolution = Suivi.Résolution"
    Flag =3
    LeftTable ="Employés"
    RightTable ="Suivi"
    Expression ="Employés.RéfEmployé = Suivi.RésoluPar"
    Flag =3
    LeftTable ="RechercheEtat"
    RightTable ="Suivi"
    Expression ="RechercheEtat.RéfEtat = Suivi.Etat"
    Flag =3
    LeftTable ="Employés_1"
    RightTable ="Suivi"
    Expression ="Employés_1.RéfEmployé = Suivi.SaisiPar"
    Flag =3
End
Begin OrderBy
    Expression ="Suivi.RéfProblème"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Saisi par "
        dbMemo "Caption" ="Entered By"
    End
    Begin
        dbText "Name" ="Résolu par "
        dbMemo "Caption" ="Resolved By"
    End
End
Begin
    State =0
    Left =4
    Top =4
    Right =790
    Bottom =439
    Left =-1
    Top =-1
    Right =779
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =171
        Top =3
        Right =280
        Bottom =110
        Top =0
        Name ="Employés"
        Name =""
    End
    Begin
        Left =4
        Top =137
        Right =137
        Bottom =214
        Top =0
        Name ="RechercheEtat"
        Name =""
    End
    Begin
        Left =179
        Top =140
        Right =314
        Bottom =217
        Top =0
        Name ="RecherchePriorité"
        Name =""
    End
    Begin
        Left =349
        Top =136
        Right =493
        Bottom =213
        Top =0
        Name ="RechercheRésolution"
        Name =""
    End
    Begin
        Left =289
        Top =6
        Right =400
        Bottom =113
        Top =0
        Name ="Employés_1"
        Name =""
    End
    Begin
        Left =422
        Top =5
        Right =536
        Bottom =112
        Top =0
        Name ="Employés_2"
        Name =""
    End
    Begin
        Left =16
        Top =7
        Right =112
        Bottom =114
        Top =7
        Name ="Suivi"
        Name =""
    End
End
