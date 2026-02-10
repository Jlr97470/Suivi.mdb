CREATE TABLE [Commentaires] (
  [RéfCommentaire] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [RéfProblème] LONG,
  [RéfEmployé] LONG,
  [DateSaisie] DATETIME,
  [Etat] LONG,
  [Commentaires] LONGTEXT,
   CONSTRAINT 
)
