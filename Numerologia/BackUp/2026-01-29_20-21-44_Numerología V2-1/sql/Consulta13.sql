SELECT qryUnionNombres.NombreOriginal, qryUnionNombres.IdiomaOriginal, qryUnionNombres.NombreEquivalente, qryUnionNombres.IdiomaEquivalente, qryUnionNombres.Genero, qryUnionNombres.Activo INTO tbmEquivNombre_Base
FROM qryUnionNombres;

