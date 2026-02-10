SELECT
  DISTINCTROW Suivi.RéfProblème,
  Suivi.Titre,
  Suivi.Description,
  Suivi.Etat,
  Suivi.Priorité,
  Suivi.DateSaisie,
  Suivi.SaisiPar,
  Suivi.AssignéA,
  Suivi.DateRésolution,
  Suivi.Résolution,
  Suivi.RésoluPar,
  Employés.NomEmployé AS [Résolu par]
FROM
  Employés
  RIGHT JOIN (
    Employés AS Employés_1
    RIGHT JOIN Suivi ON Employés_1.RéfEmployé = Suivi.SaisiPar
  ) ON Employés.RéfEmployé = Suivi.RésoluPar;
