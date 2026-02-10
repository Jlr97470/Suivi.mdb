CREATE TABLE [Employés] (
  [RéfEmployé] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [NomEmployé] VARCHAR (50),
  [Téléphone] VARCHAR (30),
  [RéfDirecteur] LONG
)
