DECLARE @Tab TABLE
             (
                 Codigo  INT,
                 Produto VARCHAR(30),
                 Estoque INT
             )
INSERT INTO @Tab
VALUES (1, 'Motor BMW', 5)
INSERT INTO @Tab
VALUES (2, 'Trator Agrale', 3);

WITH Seq
         AS (SELECT 1 AS ID
             UNION ALL
             SELECT ID + 1
             FROM Seq
             WHERE ID < 100)
SELECT t.*
FROM @Tab t
         INNER JOIN Seq
                    ON Seq.ID BETWEEN 0 AND t.Estoque
ORDER BY Codigo
OPTION (MAXRECURSION 0);