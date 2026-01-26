SELECT DISTINCTROW qryUnionTotal.NombreOriginal, qryUnionTotal.IdiomaOriginal, qryUnionTotal.NombreEquivalente, qryUnionTotal.IdiomaEquivalente, qryUnionTotal.Tipo, qryUnionTotal.Notas INTO tbmEquivNombre
FROM qryUnionTotal
WHERE (((qryUnionTotal.NombreOriginal) Is Not Null And (qryUnionTotal.NombreOriginal)<>"—"))
ORDER BY qryUnionTotal.NombreOriginal, qryUnionTotal.IdiomaOriginal, qryUnionTotal.IdiomaEquivalente;

