SELECT
  DISTINCTROW Commentaires.RéfCommentaire,
  Commentaires.RéfProblème,
  Commentaires.RéfEmployé,
  Employés.NomEmployé,
  Commentaires.DateSaisie,
  Commentaires.Etat,
  RechercheEtat.Etat,
  Commentaires.Commentaires
FROM
  (
    Commentaires
    INNER JOIN RechercheEtat ON Commentaires.Etat = RechercheEtat.RéfEtat
  )
  INNER JOIN Employés ON Commentaires.RéfEmployé = Employés.RéfEmployé;
