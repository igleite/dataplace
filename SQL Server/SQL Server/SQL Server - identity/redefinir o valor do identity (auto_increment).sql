/*  Ver se na tabela te auto_inclement */
SELECT COLUMN_NAME
     , TABLE_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE COLUMNPROPERTY(object_id(TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 1
  AND table_name = 'orcamento'
ORDER BY TABLE_NAME

/* Redefine o auto_increment para o próximo número no caso do pendenciavenda1 o próximo número será para 454 */
DBCC CHECKIDENT (pendenciavenda1, RESEED, 453)
DBCC CHECKIDENT (orcamento, RESEED, 844)