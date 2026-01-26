SELECT tbuResultados.IDPersona, tbuResultados.IDFonetica, tbuResultados.IDResultado, tbmModos.Calculo, [NumCiclos] & " ciclos" AS Expr2, tbmModos_1.Ciclos, tbmModos_2.Tarot, Format([FechaCalculo],"dd/mm/yyyy") AS Expr1
FROM ((tbuResultados LEFT JOIN tbmModos ON tbuResultados.SistemaCalculo = tbmModos.ID) LEFT JOIN tbmModos AS tbmModos_1 ON tbuResultados.MetodoCiclos = tbmModos_1.ID) LEFT JOIN tbmModos AS tbmModos_2 ON tbuResultados.SistemaTarot = tbmModos_2.ID
ORDER BY tbuResultados.FechaCalculo DESC;

