SELECT DISTINCT tbmEquivNombre.NombreOriginal, tbmEquivNombre.Tipo, tbmEquivNombre.Raiz
FROM tbmEquivNombre
WHERE (((tbmEquivNombre.Raiz) Is Not Null));

