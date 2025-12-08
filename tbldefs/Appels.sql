CREATE TABLE [Appels] (
  [RéfAppel] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [RéfContact] LONG,
  [DateAppel] DATETIME,
  [HeureAppel] DATETIME,
  [Sujet] VARCHAR (255),
  [Remarques] LONGTEXT
)
