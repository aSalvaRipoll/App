cabecera = Replace(cabecera, "{{NOMBRE}}", Persona.Nombre)
cabecera = Replace(cabecera, "{{FECHA_NAC}}", Persona.FechaNacimiento)
cabecera = Replace(cabecera, "{{EDAD}}", r.Edad)

cabecera = Replace(cabecera, "{{SISTEMA_FONETICO}}", NombreModoFonetico(r.IDFonetica))
cabecera = Replace(cabecera, "{{SISTEMA_CALCULO}}", NombreModoCalculo(r.SistemaCalculo))
cabecera = Replace(cabecera, "{{NUM_CICLOS}}", r.NumCiclos)
cabecera = Replace(cabecera, "{{METODO_CICLOS}}", NombreModoCiclos(r.MetodoCiclos))
cabecera = Replace(cabecera, "{{SISTEMA_TAROT}}", NombreModoTarot(r.SistemaTarot))

cabecera = Replace(cabecera, "{{FECHA_CALCULO}}", Format(r.FechaCalculo, "dd/mm/yyyy"))
cabecera = Replace(cabecera, "{{VERSION_MOTOR}}", r.VersionMotor)
