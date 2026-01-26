INSERT INTO tbmDicExcepciones ( Idioma, Palabra, Grafemas, idFonemas, EsVocal, Valor, Notas, Activo )
SELECT tbmDicExcepciones1.Idioma, tbmDicExcepciones1.Palabra, tbmDicExcepciones1.Grafemas, tbmDicExcepciones1.idFonemas, tbmDicExcepciones1.EsVocal, tbmDicExcepciones1.Valor, tbmDicExcepciones1.Notas, tbmDicExcepciones1.Activo
FROM tbmDicExcepciones1
ORDER BY tbmDicExcepciones1.Idioma, tbmDicExcepciones1.Palabra;

