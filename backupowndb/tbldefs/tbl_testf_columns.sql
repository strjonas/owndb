CREATE TABLE [tbl_testf_columns] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [fk_rows] LONG ,
  [col_num] LONG ,
  [s_val] VARCHAR (255)
)
