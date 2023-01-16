CREATE TABLE [students] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [s_name] VARCHAR (255),
  [s_lname] VARCHAR (255),
  [s_year] LONG ,
  [s_desc] LONGTEXT ,
  [s_long] LONGTEXT ,
  [s_number] LONG 
)
