INSERT INTO tbmPrefijos ( Prefijo, Vocal, Tipo, Origen, Notas, Activo, es, ca, [ca-va], [ca-ib], gl, eu, [pt-eu], [pt-br], fr, [en-gb] )
SELECT DISTINCT tbmPrefijos_Temp.Prefijo, tbmPrefijos_Temp.Vocal, tbmPrefijos_Temp.Tipo, tbmPrefijos_Temp.Origen, tbmPrefijos_Temp.Notas, tbmPrefijos_Temp.Activo, tbmPrefijos_Temp.[es], tbmPrefijos_Temp.ca, tbmPrefijos_Temp.[ca-va], tbmPrefijos_Temp.[ca-ib], tbmPrefijos_Temp.gl, tbmPrefijos_Temp.eu, tbmPrefijos_Temp.[pt-eu], tbmPrefijos_Temp.[pt-br], tbmPrefijos_Temp.fr, tbmPrefijos_Temp.[en-gb]
FROM tbmPrefijos_Temp;

