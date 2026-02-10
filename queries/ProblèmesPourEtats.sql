SELECT
  DISTINCTROW Suivi.RéfProblème,
  Suivi.Titre,
  Suivi.Description,
  RechercheEtat.Etat,
  RecherchePriorité.Priorité,
  Suivi.DateSaisie,
  Employés_1.NomEmployé AS [Saisi par],
  Employés_2.NomEmployé AS AssignéA,
  Suivi.DateRésolution,
  RechercheRésolution.Résolution,
  Employés.NomEmployé AS [Résolu par]
FROM
  Employés AS Employés_1
  RIGHT JOIN (
    RechercheEtat
    RIGHT JOIN (
      Employés
      RIGHT JOIN (
        RechercheRésolution
        RIGHT JOIN (
          (
            RecherchePriorité
            RIGHT JOIN Suivi ON RecherchePriorité.RéfPriorité = Suivi.Priorité
          )
          LEFT JOIN Employés AS Employés_2 ON Suivi.AssignéA = Employés_2.RéfEmployé
        ) ON RechercheRésolution.RéfRésolution = Suivi.Résolution
      ) ON Employés.RéfEmployé = Suivi.RésoluPar
    ) ON RechercheEtat.RéfEtat = Suivi.Etat
  ) ON Employés_1.RéfEmployé = Suivi.SaisiPar
WHERE
  (
    (
      (Suivi.Etat)< 3
    )
  )
ORDER BY
  Suivi.RéfProblème;
