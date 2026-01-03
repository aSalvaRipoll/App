# 📋 GUÍA PRÁCTICA: RECOPILACIÓN DE NOMBRES CON DÍGRAFOS
## **Base de Datos para el Motor Fonético de Universal Magic**

---

## 🎯 OBJETIVO

Crear bases de datos completas de nombres y apellidos que contienen los dígrafos españoles (CH, LL, RR) para:

1. **Validar** el motor fonético
2. **Probar** casos reales
3. **Documentar** ejemplos
4. **Crear** interpretaciones específicas
5. **Educar** usuarios sobre diferencias

---

## 📊 ESTRUCTURA DE LAS LISTAS

### **Formato de Archivo CSV Propuesto**

```csv
Nombre,Tipo,Idioma,Digrafos,Genero,Frecuencia,Notas
CHARO,Nombre,Español,CH,F,Media,Diminutivo de Rosario
LLUC,Nombre,Catalán,LL,M,Alta,Muy común en Catalunya
TORRE,Apellido,Español,RR,U,Alta,Apellido muy frecuente
CHILLÓN,Apellido,Español,"CH,LL",U,Baja,Dos dígrafos
```

**Campos:**
- **Nombre:** El nombre/apellido completo en MAYÚSCULAS
- **Tipo:** Nombre / Apellido
- **Idioma:** Español / Catalán / Euskera / Gallego / etc.
- **Dígrafos:** CH / LL / RR / CH,LL / etc. (si tiene múltiples)
- **Género:** M (masculino) / F (femenino) / U (unisex/apellido)
- **Frecuencia:** Alta / Media / Baja
- **Notas:** Información adicional relevante

---

## 🇪🇸 ESPAÑOL (CASTELLANO) - PRIORIDAD MÁXIMA

### **NOMBRES CON CH**

**Femeninos:**
```
CHARO (diminutivo de Rosario) - Frecuencia: Media
CHELO (diminutivo de Consuelo) - Frecuencia: Media
CHUS (diminutivo de Jesús/María Jesús) - Frecuencia: Media
CHABELI (diminutivo de Isabel) - Frecuencia: Baja
CHON (diminutivo de Concepción) - Frecuencia: Baja
CHONI (diminutivo de Concepción) - Frecuencia: Baja
CONCHA (diminutivo de Concepción) - Frecuencia: Alta
CONCHITA (diminutivo de Concepción) - Frecuencia: Media
CHARITO (diminutivo de Rosario) - Frecuencia: Baja
CHELITO (diminutivo de Consuelo) - Frecuencia: Muy baja
```

**Masculinos:**
```
CHUCHO (diminutivo de Jesús) - Frecuencia: Baja
NACHO (diminutivo de Ignacio) - Frecuencia: Alta
PANCHO (diminutivo de Francisco) - Frecuencia: Media
LUCHO (diminutivo de Luis) - América Latina
CHEMA (diminutivo de José María) - Frecuencia: Media
CHENTE (diminutivo de Vicente) - América Latina
CHECO (gentilicio, usado como nombre) - Frecuencia: Baja
CHENCHO (diminutivo de Inocencio) - Frecuencia: Muy baja
```

**Nombres compuestos:**
```
MARÍA CONCHA
JOSÉ NACHO
FRANCISCA CONCHITA
```

**NOTA IMPORTANTE:** En España, muchos nombres con CH son diminutivos cariñosos que se usan como nombres propios.

---

### **NOMBRES CON LL (en Español)**

**NOTA:** En español peninsular estándar, LL al inicio de nombre es MUY raro. La mayoría son:
- De origen catalán
- Préstamos de otras lenguas
- Apellidos convertidos en nombres

**Ejemplos raros:**
```
LLANOS (advocación mariana, "Virgen de los Llanos")
LLOYD (préstamo inglés, usado en España)
```

**Para nombres con LL, ver sección de CATALÁN más abajo.**

---

### **APELLIDOS CON RR - ALTA PRIORIDAD**

**Apellidos muy frecuentes (Top 50 España):**

```
HERRERA - Frecuencia: MUY ALTA (#20 aprox.)
  Origen: Lugar donde se trabaja el hierro
  Distribución: Nacional
  Variantes: Ferreiro (gallego), Ferrer (catalán)

GUERRA - Frecuencia: MUY ALTA (#40 aprox.)
  Origen: Apodo o profesión (guerrero)
  Distribución: Nacional

SERRANO - Frecuencia: ALTA (#35 aprox.)
  Origen: De la sierra, montañés
  Distribución: Nacional

NAVARRO - Frecuencia: ALTA (#45 aprox.)
  Origen: De Navarra
  Distribución: Nacional

FERRER - Frecuencia: ALTA
  Origen: Herrero (catalán)
  Distribución: Catalunya, Valencia, Baleares

GUERRERO - Frecuencia: ALTA
  Origen: Profesión (guerrero)
  Distribución: Nacional

PARRA - Frecuencia: MEDIA-ALTA
  Origen: Planta de la vid
  Distribución: Nacional

BECERRA - Frecuencia: MEDIA
  Origen: Vaca joven
  Distribución: Norte de España

BARRERA - Frecuencia: MEDIA
  Origen: Obstáculo, barrera
  Distribución: Nacional

SIERRA - Frecuencia: MEDIA
  Origen: Montaña, herramienta
  Distribución: Nacional

CORREA - Frecuencia: MEDIA
  Origen: Tira de cuero
  Distribución: Nacional

BARRA - Frecuencia: MEDIA
  Origen: Pieza alargada
  Distribución: Nacional
```

**Apellidos con frecuencia media:**

```
TORRENTE
TORREGROSA
TORRALBA
TORRE (y derivados: Torres, Torrejón, etc.)
BERROCAL
CARRERA
CERRADA
CORRALES
FERREIRA
FIGUEROA
PERROTE
TORRERO
YERRO
ZORRERO
BORREGO
CARRASQUILLA
HERRANZ
HERREROS
PARRILLA
PERALTA
SERRALTA
TERRAZAS
TERRÓN
TORRADO
VERDUGO (tiene RR en algunas pronunciaciones)
```

**Apellidos con RR doble o múltiple:**

```
HERRERO - Frecuencia: ALTA
  RR simple pero apellido muy común
  
FERREIRO - Frecuencia: MEDIA (Galicia)
  Variante gallega de Ferrer/Herrero

CARRASCO - Frecuencia: MEDIA
  RR + posible doblete fonético

BARRANCO - Frecuencia: MEDIA
  RR en medio

SERRADOR - Frecuencia: BAJA
  RR doble (dos RR separadas)
```

---

### **APELLIDOS CON CH**

```
CHACÓN - Frecuencia: MEDIA
CHAVES - Frecuencia: MEDIA
CHAVES - Frecuencia: MEDIA  
CHAMORRO - Frecuencia: MEDIA
CHECA - Frecuencia: BAJA
CHICO - Frecuencia: MEDIA
CHINCHILLA - Frecuencia: BAJA (¡tiene CH dos veces!)
CHUECA - Frecuencia: BAJA
MONTESDEOCA - Frecuencia: BAJA (contiene CH en "deoCA")
SANTAMARÍA (algunos pronuncian el CH en variantes)
```

---

### **APELLIDOS CON LL**

**NOTA:** En apellidos españoles, LL es relativamente frecuente:

```
LLAMAS - Frecuencia: MEDIA
LLORENTE - Frecuencia: MEDIA-ALTA
LLORET - Frecuencia: MEDIA
LLORENS - Frecuencia: MEDIA (más catalán)
LLOPIS - Frecuencia: MEDIA (más valenciano)
LLOBREGAT - Frecuencia: BAJA
CASTELLANOS - Frecuencia: ALTA (LL interna)
CASTILLO - Frecuencia: MUY ALTA (LL interna)
PORTILLO - Frecuencia: MEDIA (LL interna)
CARRILLO - Frecuencia: MEDIA (LL interna + RR)
MURILLO - Frecuencia: MEDIA (LL interna)
MEDINA-SIDONIA (tiene LL en algunas variantes)
CABELLO - Frecuencia: MEDIA
SELLO - Frecuencia: BAJA
BELLIDO - Frecuencia: MEDIA
BELLVER - Frecuencia: BAJA
GALLARDO - Frecuencia: MEDIA (LL interna)
GALLEGO - Frecuencia: ALTA (LL interna)
VALLE - Frecuencia: MEDIA-ALTA
VALLEJO - Frecuencia: MEDIA
VILLAR - Frecuencia: ALTA (LL interna)
```

---

### **APELLIDOS CON MÚLTIPLES DÍGRAFOS**

**Estos son casos especiales MUY interesantes:**

```
CARRILLO - Contiene: RR + LL
  Ejemplo completo: PEDRO CARRILLO
  Cálculo fonético: múltiples dígrafos

CHILLÓN - Contiene: CH + LL
  Ejemplo: CARMEN CHILLÓN
  Dos dígrafos maestros/especiales

TORRALBA - Contiene: RR + LL (potencialmente)
  Depende de pronunciación regional

BORRULL - Contiene: RR + LL
  Apellido catalán/valenciano
  Ejemplo: JOAN BORRULL

CARRASCO - Contiene: RR + SC
  (SC no es dígrafo, pero RR sí)
```

---

## 🏴 CATALÁN - PRIORIDAD ALTA

### **Contexto Lingüístico**

El catalán tiene:
- **LL** /ʎ/ - Lateral palatal (como LL español tradicional)
- **L·L** /l.l/ - Dos eles separadas (con punto volado)
- **NY** /ɲ/ - Equivalente a Ñ española
- **IG** /ʧ/ - Al final de palabra = CH española
- **TX** /ʧ/ - Equivalente a CH española
- **RR** /r/ - Vibrante múltiple

### **NOMBRES CON LL**

**Masculinos muy comunes:**
```
LLUC - Frecuencia: MUY ALTA en Catalunya
  Variante catalana de Lucas
  Pronunciación: /ʎuk/

LLUÍS - Frecuencia: ALTA
  Variante catalana de Luis
  Pronunciación: /ʎu'is/

LLORENÇ - Frecuencia: MEDIA
  Variante catalana de Lorenzo
  Pronunciación: /ʎu'rɛns/
```

**Femeninos comunes:**
```
LLÚCIA - Frecuencia: MEDIA
  Variante catalana de Lucía
  Pronunciación: /'ʎusiə/

LLUNA - Frecuencia: MEDIA
  Significa "Luna"
  Pronunciación: /'ʎunə/
```

**Apellidos catalanes con LL:**
```
LLORENS - Frecuencia: MUY ALTA
LLORET - Frecuencia: ALTA
LLOMBART - Frecuencia: MEDIA
LLOBERA - Frecuencia: MEDIA
LLOBREGAT - Frecuencia: BAJA
LLULL - Frecuencia: MEDIA (Ramon Llull, famoso)
```

---

### **NOMBRES CON NY (equivalente a Ñ)**

```
MUNTANYA - No es nombre propio, pero aparece en topónimos
ESPANYA - Igual, topónimo
CAÑELLAS - Versión española: Cañellas
```

**En catalán, NY = /ɲ/ tiene el mismo valor que Ñ española = 5**

---

### **NOMBRES CON IG FINAL**

**El grupo IG al final suena /ʧ/ (como CH):**

```
PUIG - Apellido muy común
  Pronunciación: /puʧ/
  Significa "colina, montaña"
  Frecuencia: MUY ALTA en Catalunya

ROIG - Apellido común
  Pronunciación: /roʧ/
  Significa "rojo"
  Frecuencia: ALTA

DESIG - Menos común como apellido
  Pronunciación: /də'ziʧ/

VIG - Poco común
```

**En sistema fonético:** IG final = 11 (mismo valor que CH)

---

### **NOMBRES CON TX (equivalente a CH)**

```
TXELL - Nombre femenino
  Pronunciación: /ʧeʎ/
  Frecuencia: MEDIA en Euskadi

TXEMA - Nombre masculino (diminutivo vasco de José María)
  Pronunciación: /'ʧema/
  Frecuencia: ALTA en País Vasco
```

---

## 🟢 EUSKERA (VASCO) - PRIORIDAD MEDIA

### **Contexto Lingüístico**

El euskera tiene varios dígrafos especiales:

- **TX** /ʧ/ - Africada postalveolar (= CH español)
- **TS** /ts̻/ - Africada alveolar
- **TZ** /ts̺/ - Africada apicoalveolar  
- **TT** /c/ - Oclusiva
- **DD** /ɟ/ - Oclusiva palatal
- **RR** /r/ - Vibrante múltiple

### **NOMBRES CON TX**

```
TXOMIN - Masculino
  Equivalente vasco de Domingo
  Pronunciación: /'ʧomin/
  Frecuencia: ALTA en Euskadi

TXEMA - Masculino
  Diminutivo de José María
  Pronunciación: /'ʧema/
  Frecuencia: MUY ALTA

TXELL - Femenino
  Pronunciación: /ʧeʎ/
  Frecuencia: MEDIA

ITXASO - Femenino
  Significa "mar"
  Pronunciación: i'ʧaso
  Frecuencia: ALTA

TXARO - Femenino
  Variante vasca
  Pronunciación: /'ʧaro/
  Frecuencia: MEDIA
```

**En sistema fonético:** TX = 11 (mismo valor que CH español)

---

### **APELLIDOS VASCOS CON TX**

```
ETXEBERRIA - Muy común
  Significa "casa nueva"
  Contiene TX

ETXEBARRIA - Variante
  También contiene TX

OTXOA - Común
  Contiene TX
```

---

### **NOMBRES CON RR**

```
GORKA - Común
  (No tiene RR pero es muy vasco)

GARRIDO - Apellido común en zona vasca
  Contiene RR
```

---

## 🌊 GALLEGO - PRIORIDAD BAJA

### **Contexto Lingüístico**

El gallego tiene:
- **LL** /ʎ/ - Lateral palatal
- **NH** /ɲ/ - Equivalente a Ñ española
- **CH** /ʧ/ - Como español
- **RR** /r/ - Vibrante múltiple

### **NOMBRES Y APELLIDOS GALLEGOS**

```
FERREIRO - Apellido muy común
  Significa "herrero"
  Contiene RR
  Frecuencia: ALTA en Galicia

CARREIRA - Apellido común
  Contiene RR
  Frecuencia: MEDIA

BARRAL - Apellido
  Contiene RR
  Frecuencia: MEDIA
```

**Nombres con NH:**
```
MINHO - Topónimo (río)
CUNHA - Apellido
```

---

## 🏝️ BALEAR (MALLORQUÍN, MENORQUÍN, IBICENCO)

### **Contexto Lingüístico**

El catalán balear tiene características propias pero usa los mismos dígrafos que el catalán estándar:

- **LL** /ʎ/
- **NY** /ɲ/  
- **IG** final /ʧ/
- **RR** /r/

### **NOMBRES ESPECÍFICOS DE BALEARES**

```
LLUC - Muy popular en Mallorca
  Patrón de la isla
  Frecuencia: MUY ALTA

CATALINA - No tiene dígrafos pero muy balear
TOMEU - No tiene dígrafos pero típico mallorquín
BIEL - No tiene dígrafos pero típico mallorquín
```

**Apellidos baleares con dígrafos:**
```
FERRER - Muy común
  Contiene RR
  
OLIVER - Común
  Contiene LL interna

LLABRÉS - Común
  Contiene LL inicial

FERRAGUT - Común
  Contiene RR
```

---

## 🍊 VALENCIANO - PRIORIDAD MEDIA

### **Contexto Lingüístico**

El valenciano es una variante del catalán con los mismos dígrafos:

- **LL** /ʎ/
- **NY** /ɲ/
- **RR** /r/

### **NOMBRES Y APELLIDOS VALENCIANOS**

```
LLORENS - Apellido común
  Contiene LL
  Frecuencia: ALTA

FERRER - Muy común
  Contiene RR
  Frecuencia: MUY ALTA

BORRELL - Apellido
  Contiene RR + LL
  Frecuencia: MEDIA

BORRULL - Apellido
  Contiene RR + LL
  Caso especial: dos dígrafos
```

---

## 📈 ESTADÍSTICAS Y PRIORIDADES

### **Resumen de Frecuencias por Dígrafo**

```
╔══════════════════════════════════════════════════════╗
║  DÍGRAFO  │  NOMBRES  │  APELLIDOS  │  IMPACTO      ║
╠══════════════════════════════════════════════════════╣
║  RR       │  Muy bajo │  MUY ALTO   │  ⭐⭐⭐⭐⭐    ║
║  LL       │  Medio    │  Alto       │  ⭐⭐⭐⭐      ║
║  CH       │  Medio    │  Medio      │  ⭐⭐⭐        ║
╚══════════════════════════════════════════════════════╝
```

**Conclusión estadística:**
- **RR en apellidos** es el caso más importante (afecta ~15-20% de población española)
- **LL en nombres** es relevante sobre todo en Catalunya
- **CH en nombres** es moderadamente común (diminutivos)

---

## 🗂️ ORGANIZACIÓN DE ARCHIVOS SUGERIDA

```
/Nombres_Digrafos/
├── Español/
│   ├── nombres_con_CH.csv
│   ├── nombres_con_LL.csv
│   ├── apellidos_con_RR.csv (⭐ PRIORIDAD)
│   ├── apellidos_con_CH.csv
│   ├── apellidos_con_LL.csv
│   └── nombres_multiples_digrafos.csv
├── Catalan/
│   ├── nombres_con_LL.csv (⭐ PRIORIDAD)
│   ├── nombres_con_NY.csv
│   ├── apellidos_con_IG.csv
│   └── apellidos_con_TX.csv
├── Euskera/
│   ├── nombres_con_TX.csv
│   ├── nombres_con_TS.csv
│   └── apellidos_vascos.csv
├── Gallego/
│   └── apellidos_con_RR.csv
├── Balear/
│   └── nombres_apellidos_baleares.csv
└── Valenciano/
    └── nombres_apellidos_valencianos.csv
```

---

## 🔍 METODOLOGÍA DE RECOPILACIÓN

### **Fuentes Oficiales Recomendadas**

**ESPAÑA:**
1. **INE (Instituto Nacional de Estadística)**
   - URL: https://www.ine.es/
   - Sección: Nombres y apellidos más frecuentes
   - Filtrar por: Comunidades autónomas

2. **Páginas de registros civiles**
   - Listados oficiales por región

**CATALUNYA:**
1. **Idescat (Institut d'Estadística de Catalunya)**
   - URL: https://www.idescat.cat/
   - Nombres catalanes más populares por año

**EUSKADI:**
1. **Eustat (Instituto Vasco de Estadística)**
   - URL: https://www.eustat.eus/
   - Nombres vascos registrados

**GALICIA:**
1. **IGE (Instituto Galego de Estatística)**
   - URL: https://www.ige.eu/
   - Nombres gallegos

---

### **Herramientas de Extracción**

```python
# Script Python ejemplo para procesar datos del INE

import pandas as pd

# Cargar datos
df = pd.read_csv('nombres_ine.csv', encoding='utf-8')

# Filtrar nombres con dígrafos
nombres_con_CH = df[df['Nombre'].str.contains('CH', na=False)]
nombres_con_LL = df[df['Nombre'].str.contains('LL', na=False)]
nombres_con_RR = df[df['Nombre'].str.contains('RR', na=False)]

# Exportar
nombres_con_CH.to_csv('nombres_con_CH.csv', index=False)
# ... etc
```

---

### **Criterios de Inclusión**

**NOMBRES:**
- ✅ Incluir si tiene al menos 100 registros en España
- ✅ Incluir diminutivos usados como nombres propios
- ✅ Incluir variantes regionales
- ❌ Excluir nombres extranjeros no adaptados

**APELLIDOS:**
- ✅ Incluir todos los del Top 500 España
- ✅ Incluir apellidos regionales comunes (Top 100 por región)
- ✅ Incluir apellidos con múltiples dígrafos (prioridad)
- ❌ Excluir apellidos con <10 portadores

---

## ✅ LISTA DE VERIFICACIÓN DE TAREAS

### **Fase 1: Recopilación Básica (PRIORIDAD MÁXIMA)**

- [ ] Descargar datos del INE (apellidos españoles)
- [ ] Filtrar apellidos con RR (estimar: 500-1000 apellidos)
- [ ] Crear CSV con top 200 apellidos con RR
- [ ] Documentar 50 ejemplos completos (nombre + apellido con RR)
- [ ] Verificar cálculos manuales de 20 casos

### **Fase 2: Nombres Catalanes (PRIORIDAD ALTA)**

- [ ] Descargar datos Idescat
- [ ] Listar nombres con LL inicial (estimar: 50-100 nombres)
- [ ] Documentar LLUC, LLUÍS, LLORENÇ con ejemplos completos
- [ ] Verificar cálculos de 10 casos catalanes

### **Fase 3: Casos Especiales (PRIORIDAD MEDIA)**

- [ ] Buscar apellidos con múltiples dígrafos (CARRILLO, CHILLÓN, etc.)
- [ ] Documentar 20 casos especiales
- [ ] Crear tabla comparativa (tradicional vs fonético)
- [ ] Verificar que generan resultados diferentes

### **Fase 4: Otros Idiomas (PRIORIDAD BAJA)**

- [ ] Nombres vascos con TX (10-20 ejemplos)
- [ ] Apellidos gallegos con RR (20-30 ejemplos)
- [ ] Nombres baleares únicos (10 ejemplos)
- [ ] Documentación mínima de cada región

---

## 🎯 OBJETIVOS MÍNIMOS VIABLES

**Para lanzar la Versión 1.0 de Universal Magic necesitas:**

1. ✅ **100 apellidos con RR** documentados y probados
2. ✅ **30 nombres con CH** documentados
3. ✅ **20 nombres catalanes con LL** documentados
4. ✅ **10 casos con múltiples dígrafos** documentados
5. ✅ **50 ejemplos completos** (nombre+apellido) calculados con ambos sistemas

**Total estimado de registros:** ~200 entradas en base de datos

**Tiempo estimado:** 2-3 días de trabajo (si usas fuentes oficiales)

---

## 📝 PLANTILLA DE DOCUMENTACIÓN POR NOMBRE

```markdown
### EJEMPLO: CHARO TORRE

**ANÁLISIS FONÉTICO:**

Nombre: CHARO
- Fonemas: /ʧ/ + /a/ + /ɾ/ + /o/ = 4 elementos
- Sistema Tradicional: C(3)+H(8)+A(1)+R(9)+O(6) = 27 → 9
- Sistema Fonético: CH(11)+A(1)+R(9)+O(6) = 27 → 9
- Piedra Angular: C(3) vs CH(11) ⭐ DIFERENTE

Apellido: TORRE
- Fonemas: /t/ + /o/ + /r̄/ + /e/ = 4 elementos
- Sistema Tradicional: T(2)+O(6)+R(9)+R(9)+E(5) = 31 → 4
- Sistema Fonético: T(2)+O(6)+RR(9)+E(5) = 22 ⭐ MAESTRO
- Resultado final: 4 vs 22 ⭐⭐⭐ MUY DIFERENTE

**SIGNIFICADO DEL CAMBIO:**
El sistema fonético detecta que TORRE tiene una vibración maestra (22)
debido a la intensidad única del fonema /r̄/ (RR). El sistema tradicional
trata las dos R como elementos separados, perdiendo esa intensidad.

**INTERPRETACIÓN:**
- Tradicional: Constructor práctico, trabajador estable (4)
- Fonético: Maestro constructor, edificador visionario (22)

La diferencia es filosóficamente profunda y afecta toda la carta.
```

---

## 🌟 CASOS DE ESTUDIO PRIORITARIOS

### **Top 10 Casos Más Importantes para Documentar**

1. **TORRE** (apellido) - Genera maestro 22
2. **HERRERA** (apellido) - Muy frecuente con RR
3. **GUERRA** (apellido) - Muy frecuente con RR
4. **CHARO** (nombre) - Piedra Angular maestra
5. **LLUC** (nombre catalán) - Piedra Angular especial
6. **CARRILLO** (apellido) - Dos dígrafos (RR+LL)
7. **CHILLÓN** (apellido) - Dos dígrafos (CH+LL)
8. **NACHO** (nombre) - CH común en España
9. **CONCHITA** (nombre) - CH + diminutivo
10. **LLORENS** (apellido catalán) - LL muy común

---

## 💻 SCRIPT DE AYUDA PARA PROCESAMIENTO

```python
import csv
import re

def detectar_digrafos(texto):
    """
    Detecta dígrafos en un texto.
    Retorna lista de dígrafos encontrados.
    """
    texto_upper = texto.upper()
    digrafos = []
    
    if 'CH' in texto_upper:
        digrafos.append('CH')
    if 'LL' in texto_upper:
        digrafos.append('LL')
    if 'RR' in texto_upper:
        digrafos.append('RR')
    
    return digrafos

def procesar_lista_nombres(archivo_entrada, archivo_salida):
    """
    Procesa una lista de nombres y filtra los que tienen dígrafos.
    """
    with open(archivo_entrada, 'r', encoding='utf-8') as f_in:
        with open(archivo_salida, 'w', encoding='utf-8', newline='') as f_out:
            reader = csv.DictReader(f_in)
            fieldnames = ['Nombre', 'Tipo', 'Digrafos', 'Frecuencia']
            writer = csv.DictWriter(f_out, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for row in reader:
                nombre = row['Nombre']
                digrafos = detectar_digrafos(nombre)
                
                if digrafos:
                    writer.writerow({
                        'Nombre': nombre,
                        'Tipo': row.get('Tipo', 'Desconocido'),
                        'Digrafos': ','.join(digrafos),
                        'Frecuencia': row.get('Frecuencia', 'Media')
                    })

# Uso
procesar_lista_nombres('nombres_todos.csv', 'nombres_con_digrafos.csv')
```

---

## 📚 RECURSOS ADICIONALES

### **Libros de Referencia**

1. **"Nombres y Apellidos Españoles"** - Roberto Faure et al.
2. **"Diccionario de Apellidos Españoles"** - Instituto de Genealogía
3. **"Onomástica Catalana"** - Institut d'Estudis Catalans

### **Sitios Web Útiles**

1. **Forebears.io** - Distribución geográfica de apellidos
2. **Behind the Name** - Etimología de nombres
3. **Apellidosespañoles.com** - Frecuencias aproximadas

---

## ✨ CONCLUSIÓN

Esta guía te proporciona:

1. ✅ **Estructura clara** de qué datos recopilar
2. ✅ **Prioridades** por frecuencia e impacto
3. ✅ **Metodología** de recopilación
4. ✅ **Herramientas** para procesar datos
5. ✅ **Objetivos mínimos** para V1.0

**Siguiente paso:** Comenzar con la recopilación de apellidos con RR del INE (máxima prioridad).

---

**Documento creado para:** Sistema Universal Magic  
**Versión:** 1.0 - Guía de Recopilación  
**Autor:** Alba - Proyecto de Numerología Fonética  
**Fecha:** Enero 2025

---

🎯 **Con esta guía tienes un plan claro y ejecutable para crear tu base de datos de nombres con dígrafos** 🎯
