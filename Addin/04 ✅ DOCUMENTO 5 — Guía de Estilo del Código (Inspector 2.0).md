# Inspector VBA — Guía de Estilo del Código (Versión 2.0)

Esta guía define las convenciones oficiales para escribir, organizar y mantener el código del Inspector.  
Su objetivo es garantizar:

- Claridad
- Consistencia
- Mantenibilidad
- Escalabilidad
- Profesionalidad

Estas reglas aplican a todos los módulos, clases y formularios del proyecto.

---

# 1. Convenciones de Nombres

## 1.1 Módulos estándar

Formato:

NN_modNombre

Ejemplos:

- `04_modFunciones`
- `20_modAnalisis`
- `21_modReglas`

Reglas:

- **NN** = número de orden (00–99)
- `mod` siempre presente
- Nombre en PascalCase
- Un módulo = una responsabilidad

---

## 1.2 Clases

Formato:

clsNombre


Ejemplos:

- `clsResultadoAnalisis`
- `clsResultados`
- `clsAnalizador`

Reglas:

- Nombre en PascalCase
- Deben representar entidades o comportamientos claros
- No deben contener lógica de UI

---

## 1.3 Formularios

Formato:

frmNombre


Ejemplos:

- `frmInspector`

Reglas:

- Solo UI
- Sin lógica de análisis
- Sin duplicación de funciones de formateo

---

## 1.4 Controles de formulario

Formato:

```<tipo><Nombre>```


Ejemplos:
- `lstResultados`
- `lblSeveridad`
- `cmdAnalizar`

Tipos comunes:
- `lbl` = Label
- `cmd` = CommandButton
- `lst` = ListBox
- `txt` = TextBox
- `chk` = CheckBox

---

# 2. Estructura de Módulos

Cada módulo debe seguir esta estructura:

```vba
Option Compare Database
Option Explicit

'===============================================================
' Nombre del módulo
' Descripción breve
'---------------------------------------------------------------
' Responsabilidades:
'   - Punto 1
'   - Punto 2
'   - Punto 3
'===============================================================
```

**Reglas:**

- Siempre usar Option Explicit
- Siempre incluir cabecera descriptiva
- Agrupar funciones por secciones
- Separar secciones con comentarios largos

# 3. Estructura de Clases

Cada clase debe seguir esta estructura:

```vba

Option Compare Database
Option Explicit

'===============================================================
' clsNombre
' Descripción breve
'===============================================================

'--- Variables privadas ----------------------------------------
Private mNombre As String
Private mTipo As String

'--- Propiedades ------------------------------------------------
Public Property Get Nombre() As String
End Property

Public Property Let Nombre(value As String)
End Property

'--- Métodos públicos ------------------------------------------
Public Sub Analizar()
End Sub

'--- Métodos privados ------------------------------------------
Private Function Interno()
End Function
```

**Reglas:**

- Variables privadas con prefijo m
- Propiedades siempre con Get/Let/Set
- Métodos públicos arriba, privados abajo
- Nada de lógica de UI

# 4. Comentarios

## 4.1 Comentarios de cabecera

**Obligatorios en:**

- Módulos
- Clases
- Funciones públicas
  
## 4.2 Comentarios en línea

**Usar solo cuando:**

- La intención no sea obvia
- Se documente una decisión arquitectónica
- Se explique un comportamiento no trivial

**Evitar:**

- Comentarios redundantes
- Comentarios que repiten el código

### 1. Estilo de Código

## 5.1 Indentación

- 4 espacios
- Nunca usar tabuladores

## 5.2 Líneas largas

- Máximo recomendado: 120 caracteres
- Usar _ para dividir líneas complejas

## 5.3 Espacios

- Espacio antes y después de operadores
- Espacio después de coma
- Sin espacios antes de paréntesis en llamadas
- Ejemplo correcto:

```vba
If valor > 10 Then
    Procesar valor, True
End If
```
# 6. Manejo de Errores

## 6.1 Reglas generales

- No usar On Error Resume Next salvo casos muy controlados
- Siempre capturar errores en puntos críticos
- Registrar errores en el futuro módulo de logs (versión 3.0)

## 6.2 Patrón recomendado

```vba
On Error GoTo ErrHandler

' Código principal

Exit Sub

ErrHandler:
    MsgBox "Error en X: " & Err.Description, vbExclamation
```

# 7. Funciones y Procedimientos

## 7.1 Nombres

- Verbos para procedimientos (AnalizarModulo)
- Sustantivos para funciones (IconoSeveridad)
- Nombres descriptivos, nunca abreviaturas crípticas

## 7.2 Parámetros

- Usar nombres claros
- Evitar parámetros booleanos ambiguos
- Preferir enumeraciones cuando sea posible

# 8. Reglas de Formateo Visual (Inspector)

Toda la lógica visual debe estar en:

```04_modFunciones```

**Incluye:**

- Iconos Unicode
- Formato de severidad
- Formato de elemento
- Formato de miembro
- Truncado de texto
- Normalización

El formulario no debe contener lógica de formateo.

# 9. Ordenación y Navegación

## 9.1 Ordenación

- Implementada en clsResultados
- Encabezados gestionan asc/desc
- Indicadores visuales desde tblUnicode

## 9.2 Navegación

Implementada en modNavegacion
El formulario solo llama a funciones públicas

10. Código Limpio
Reglas esenciales:

No duplicar lógica

No mezclar UI con lógica

No mezclar análisis con formateo

No usar variables globales salvo casos muy justificados

Mantener módulos pequeños y enfocados

## 11. Estado actual (Versión 2.0)

- ✅ Estilo consistente
- ✅ Arquitectura limpia
- ✅ Módulos bien definidos
- ✅ Clases bien estructuradas
- ✅ Formateo centralizado
- ✅ Preparado para crecer

# 12. Próximas ampliaciones (2.1 / 3.0)

- Guía de estilo para reglas
- Guía de estilo para exportación
- Guía de estilo para paneles adicionales
- Plantillas de módulos y clases

---

# ✅ Documento 5 completado.
