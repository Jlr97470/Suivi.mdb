Operation =1
Option =8
Begin InputTables
    Name ="Employés"
    Name ="Employés"
    Alias ="Employés_1"
    Name ="Suivi"
End
Begin OutputColumns
    Expression ="Suivi.RéfProblème"
    Expression ="Suivi.Titre"
    Expression ="Suivi.Description"
    Expression ="Suivi.Etat"
    Expression ="Suivi.Priorité"
    Expression ="Suivi.DateSaisie"
    Expression ="Suivi.SaisiPar"
    Expression ="Suivi.AssignéA"
    Expression ="Suivi.DateRésolution"
    Expression ="Suivi.Résolution"
    Expression ="Suivi.RésoluPar"
    Alias ="Résolu par"
    Expression ="Employés.NomEmployé"
End
Begin Joins
    LeftTable ="Employés_1"
    RightTable ="Suivi"
    Expression ="Employés_1.RéfEmployé = Suivi.SaisiPar"
    Flag =3
    LeftTable ="Employés"
    RightTable ="Suivi"
    Expression ="Employés.RéfEmployé = Suivi.RésoluPar"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =0
    Left =40
    Top =22
    Right =778
    Bottom =348
    Left =-1
    Top =-1
    Right =731
    Bottom =165
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =157
        Top =6
        Right =253
        Bottom =113
        Top =0
        Name ="Employés"
        Name =""
    End
    Begin
        Left =271
        Top =4
        Right =367
        Bottom =111
        Top =0
        Name ="Employés_1"
        Name =""
    End
    Begin
        Left =27
        Top =4
        Right =123
        Bottom =111
        Top =7
        Name ="Suivi"
        Name =""
    End
End
