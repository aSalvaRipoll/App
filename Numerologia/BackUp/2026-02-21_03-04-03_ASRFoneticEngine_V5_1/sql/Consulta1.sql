SELECT Prefijo
FROM qryPrefijos
WHERE Activo = 1 
            AND Tipo Like 'auténtico' 
            AND [ca-va] = true
ORDER BY Len(Prefijo) DESC , Prefijo;

