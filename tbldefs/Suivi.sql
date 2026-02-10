CREATE TABLE [Suivi] (
  [RéfProblème] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Titre] VARCHAR (50),
  [Description] LONGTEXT,
  [Etat] LONG,
  [Priorité] LONG,
  [DateSaisie] DATETIME,
  [SaisiPar] LONG,
  [AssignéA] LONG,
  [DateRésolution] DATETIME,
  [Résolution] LONG,
  [RésoluPar] LONG,
  [SiteWeb] LONGTEXT
)
