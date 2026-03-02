<div style="text-align:center; margin-top:120px;">

# **Sistema Pitagórico de Valores Fonéticos**  
### **Documento Técnico de Especificación**  
### **Versión reconstruida y documentada**

<br><br>

## **Alba**  
### Barcelona, Catalunya  
### 2026  

<br><br><br>

<img src="https://upload.wikimedia.org/wikipedia/commons/thumb/5/5a/Pythagoras_of_Samos.jpg/640px-Pythagoras_of_Samos.jpg" width="260" style="border-radius:8px; opacity:0.92;">

<br><br>

### *“La armonía del universo está hecha de números.”*  
#### — Tradición pitagórica

</div>

---

<div style="page-break-after: always;"></div>


# Sistema Pitagórico de Valores Fonéticos
## Documento técnico de especificación

Versión reconstruida y documentada

1. Principios generales del sistema

El sistema asigna a cada fonema un Valor Pitagórico, obtenido mediante la suma de valores positivos asociados a sus características articulatorias.
El orden conceptual seguido es:
Altura (vocales) o Modo (consonantes)
Posterioridad (vocales) o Punto de articulación (consonantes)
Redondez (solo vocales) o Sonoridad (consonantes)
Nasalidad (vocales, consonantes y modificadores)
Modificadores prosódicos (acento, alargamiento, ligadura, etc.)
El sistema es estrictamente aditivo:
no se restan valores en ningún caso.

Los valores se inspiran en la idea pitagórica de escalas numéricas armónicas, donde cada rasgo añade un “peso” que sitúa al fonema en un nivel relativo dentro de su serie.

2. Valores pitagóricos para vocales

2.1. Valores asignados a cada característica

. Altura (V\_Altura)

. Altura	Valor
. abierta	20
. media‑cerrada	18
. media‑abierta	18
. media	16
. cerrada	15

. Posterioridad (V\_Posterioridad)

. Posterioridad	Valor
. anterior	4
. central	3
. posterior	2

. Redondez (V\_Redondez)

. Redondez	Valor
. no	2
. sí	1

. Nasalidad (V\_Nasal)

. Nasal	Valor
. no	0
. sí	1

2.2. Fórmula general para vocales

Valor=Altura+Posterioridad+Redondez+Nasal

2.3. Ejemplos documentados

Vocal |                Rasgos                  |    Cálculo       | Valor |
------|----------------------------------------|------------------|-------|
/a/ | abierta, central, no redondeada          |  20 + 3 + 2      |  25 | 
/e/ | media‑cerrada, anterior, no redondeada   |  18 + 4 + 2      |  24 | 
/i/ | cerrada, anterior, no redondeada         |  15 + 4 + 2      |  21 | 
/o/ | media‑cerrada, posterior, redondeada     |  18 + 2 + 1      |  21 | 
/u/ | cerrada, posterior, redondeada           |  15 + 2 + 1      |  18 | 
/ã/ | abierta, central, no redondeada, nasal   |  20 + 3 + 2 + 1  |  26 | 

3. Valores pitagóricos para consonantes

3.1. Valores asignados a cada característica

. Modo (V\_Modo)

Modo	Valor
oclusiva	8
fricativa	7
africada	8
nasal	9
lateral	7
vibrante	7
aproximante	6

. Punto de articulación (V\_Punto)

Punto	Valor
bilabial	6
labiodental	5
dental	4
alveolar	4
postalveolar	3
palatal	3
retrofleja	3
velar	2
uvular	2
glotal	1

.Sonoridad (V\_Sonoridad)

Sonoridad	Valor
sorda	0
sonora	1

.Nasalidad (V\_Nasal)

Nasal	Valor
no	0
sí	1

3.2. Fórmula general para consonantes

Valor = Modo + Punto + Sonoridad + Nasal

3.3. Ejemplos documentados

Consonante	Rasgos	Cálculo	Valor
/p/	oclusiva, bilabial, sorda	8 + 6 + 0	14
/b/	oclusiva, bilabial, sonora	8 + 6 + 1	15
/t/	oclusiva, alveolar, sorda	8 + 4 + 0	12
/d/	oclusiva, alveolar, sonora	8 + 4 + 1	13
/k/	oclusiva, velar, sorda	8 + 2 + 0	10
/g/	oclusiva, velar, sonora	8 + 2 + 1	11
/m/	nasal, bilabial, sonora, nasal	9 + 6 + 1 + 1	17
/s/	fricativa, alveolar, sorda	7 + 4 + 0	11
/z/	fricativa, alveolar, sonora	7 + 4 + 1	12
/ʃ/	fricativa, palatal, sorda	7 + 3 + 0	10
/ʒ/	fricativa, palatal, sonora	7 + 3 + 1	11
/ɾ/	vibrante, alveolar, sonora	7 + 4 + 1	12
4. Modificadores prosódicos
Los modificadores se aplican después del valor fonético base.

4.1. Valores asignados
ID	ASCII	IPA	Descripción	Tipo	Nasal	Valor
80	´	ˈ	acento primario	prosódico	0	5
81	`	ˌ	acento secundario	prosódico	0	3
82	0	0	sílaba átona	prosódico	0	0
83	:	ː	alargamiento	prosódico	0	2
84	_	‿	enlace	prosódico	0	2
85	^	̃͡	nasalización prosódica	prosódico	1	4
86	::	ːː	geminación	prosódico	0	5
4.2. Fórmula general con modificadores
ValorFinal
=
ValorFonema
+
∑
ValorModificador
Ejemplo:

/m/ nasalizada con acento primario:

17
+
5
=
22
5. Resumen conceptual del sistema
El sistema es armónico, aditivo y jerárquico.

Cada rasgo añade un valor que sitúa al fonema en una escala pitagórica.

Las vocales se estructuran por altura → posterioridad → redondez → nasalidad.

Las consonantes se estructuran por modo → punto → sonoridad → nasalidad.

Los modificadores prosódicos se aplican al final, sumando su propio peso.

No se realizan restas en ningún caso.

El sistema es completamente determinista y reproducible.