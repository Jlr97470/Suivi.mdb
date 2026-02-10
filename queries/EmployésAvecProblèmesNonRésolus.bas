Operation =1
Option =8
Begin InputTables
    Name ="Employés"
    Name ="Suivi"
End
Begin OutputColumns
    Expression ="Employés.NomEmployé"
End
Begin Joins
    LeftTable ="Employés"
    RightTable ="Suivi"
    Expression ="Employés.RéfEmployé = Suivi.AssignéA"
    Flag =1
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
        Left =17
        Top =8
        Right =121
        Bottom =115
        Top =0
        Name ="Employés"
        Name =""
    End
    Begin
        Left =158
        Top =6
        Right =254
        Bottom =140
        Top =4
        Name ="Suivi"
        Name =""
    End
End
