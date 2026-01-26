SELECT tbuFonetica.IDFonetica, tbuFonetica.IDPersona, Choose([ModoFonetico],'Tradicional Clásico','Fonético','Tradicional Moderno') AS Modo, tbuFonetica.FonNombre, tbuFonetica.FonApe1, tbuFonetica.FonApe2, tbuFonetica.FechaCalculo, Exists (SELECT 1
        FROM tbuResultados
        WHERE tbuResultados.IDPersona = tbuFonetica.IDPersona
          AND tbuResultados.IDFonetica = tbuFonetica.IDFonetica
    ) AS YaCalculado
FROM tbuFonetica
WHERE (((tbuFonetica.IDFonetica)=1) AND ((tbuFonetica.IDPersona)=18) AND ((tbuFonetica.Activo)=True));

