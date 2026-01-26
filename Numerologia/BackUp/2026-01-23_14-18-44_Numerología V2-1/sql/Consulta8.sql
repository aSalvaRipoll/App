SELECT tbmFoneticaCompleta.idFonema, tbmFoneticaCompleta.Fonema, tbmFoneticaCompleta.ASCII, tbmFoneticaCompleta.EsVocal, ValoresFonéticos.ValorExtendido, ValoresFonéticos.ValorAlba, tbmFoneticaCompleta.ValorPitagórico, tbmFoneticaCompleta.ValorFonético
FROM tbmFoneticaCompleta INNER JOIN ValoresFonéticos ON tbmFoneticaCompleta.idFonema = ValoresFonéticos.idFonema
ORDER BY tbmFoneticaCompleta.idFonema;

