SELECT DISTINCTROW
  Employés.NomEmployé
FROM
  Employés
  INNER JOIN Suivi ON Employés.RéfEmployé = Suivi.AssignéA;
