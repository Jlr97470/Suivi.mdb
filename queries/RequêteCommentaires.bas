Operation =1
Option =8
Begin InputTables
    Name ="Commentaires"
    Name ="RechercheEtat"
    Name ="Employés"
End
Begin OutputColumns
    Expression ="Commentaires.RéfCommentaire"
    Expression ="Commentaires.RéfProblème"
    Expression ="Commentaires.RéfEmployé"
    Expression ="Employés.NomEmployé"
    Expression ="Commentaires.DateSaisie"
    Expression ="Commentaires.Etat"
    Expression ="RechercheEtat.Etat"
    Expression ="Commentaires.Commentaires"
End
Begin Joins
    LeftTable ="Commentaires"
    RightTable ="RechercheEtat"
    Expression ="Commentaires.Etat = RechercheEtat.RéfEtat"
    Flag =1
    LeftTable ="Commentaires"
    RightTable ="Employés"
    Expression ="Commentaires.RéfEmployé = Employés.RéfEmployé"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =2
    Left =-4
    Top =-23
    Right =800
    Bottom =485
    Left =-1
    Top =-1
    Right =797
    Bottom =165
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-1
        Top =2
        Right =95
        Bottom =137
        Top =0
        Name ="Commentaires"
        Name =""
    End
    Begin
        Left =129
        Top =2
        Right =225
        Bottom =136
        Top =0
        Name ="RechercheEtat"
        Name =""
    End
    Begin
        Left =254
        Top =4
        Right =350
        Bottom =141
        Top =0
        Name ="Employés"
        Name =""
    End
End
