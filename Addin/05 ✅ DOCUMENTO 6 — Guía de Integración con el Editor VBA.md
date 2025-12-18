# Inspector VBA — Guía de Integración con el Editor VBA (Versión 2.0)

Esta guía describe cómo el Inspector interactúa con el Editor de VBA para:
- Navegar a módulos
- Seleccionar procedimientos
- Ir a líneas específicas
- Sincronizar resultados con el código real

El objetivo es mantener una integración estable, clara y extensible.

---

# 1. Objetivo de la integración

El Inspector no solo analiza el proyecto: también permite **navegar directamente al código** desde los resultados.

Esto incluye:
- Abrir el módulo correspondiente
- Seleccionar el miembro (Sub/Function/Property)
- Posicionar el cursor en la línea exacta
- Preparar el editor para inspección manual

---

# 2. Componentes implicados

La integración se realiza a través de:

### ✅ `modNavegacion`  
Módulo encargado de:
- Abrir módulos
- Seleccionar miembros
- Ir a líneas
- Controlar el editor

### ✅ `clsResultadoAnalisis`  
Cada resultado contiene:
- Nombre del módulo
- Nombre del miembro
- Línea
- Tipo de elemento

### ✅ `frmResultados`  
El formulario llama a `modNavegacion` cuando el usuario hace clic en un resultado.

---

# 3. Flujo de navegación

```text
Usuario hace clic en un resultado
        ↓
frmInspector obtiene la clave del resultado
        ↓
clsResultados devuelve el objeto clsResultadoAnalisis
        ↓
frmInspector llama a modNavegacion
        ↓
modNavegacion abre el módulo en el Editor VBA
        ↓
modNavegacion selecciona el miembro
        ↓
modNavegacion posiciona el cursor en la línea
```

# 4. Funciones principales de modNavegacion

## 4.1 Abrir un módulo

``vba
Public Sub IrAModulo(nombreModulo As String)
    Application.VBE.ActiveVBProject.VBComponents(nombreModulo).Activate
End Sub```

**Reglas:**

- El nombre debe coincidir exactamente con el módulo real
- No se debe usar lógica de búsqueda difusa

## 4.2 Seleccionar un miembro

```vba
Public Sub IrAMiembro(nombreModulo As String, nombreMiembro As String)

    Dim comp As VBIDE.VBComponent
    Dim modCode As VBIDE.CodeModule
    Dim linea As Long

    Set comp = Application.VBE.ActiveVBProject.VBComponents(nombreModulo)
    Set modCode = comp.CodeModule

    linea = modCode.ProcStartLine(nombreMiembro, vbext_pk_Proc)

    comp.Activate
    modCode.CodePane.SetSelection linea, 1, linea, 1

End Sub
```

**Reglas:**

- Usar ProcStartLine para localizar el inicio del procedimiento
- Activar el módulo antes de seleccionar

## 4.3 Ir a una línea específica

```vba

Public Sub IrALinea(nombreModulo As String, numeroLinea As Long)

    Dim comp As VBIDE.VBComponent
    Dim modCode As VBIDE.CodeModule

    Set comp = Application.VBE.ActiveVBProject.VBComponents(nombreModulo)
    Set modCode = comp.CodeModule

    comp.Activate
    modCode.CodePane.SetSelection numeroLinea, 1, numeroLinea, 1

End Sub
```

**Reglas:**

- Solo se usa si numeroLinea > 0
- Si no hay línea, se selecciona el miembro
  
# 5. Integración desde el formulario

En `frmResultados`, el evento lstResultados_Click debe:

- Obtener la clave del resultado
- Recuperar el objeto `clsResultadoAnalisis`
- Llamar a `modNavegacion`

Ejemplo:

```vba
Private Sub lstResultados_Click()

    Dim clave As String
    Dim res As clsResultadoAnalisis

    clave = lstResultados.Column(0)
    Set res = gResultadosInspector.ObtenerPorClave(clave)

    If res.linea > 0 Then
        IrALinea res.nombreElemento, res.linea
    Else
        IrAMiembro res.nombreElemento, res.nombreMiembro
    End If

End Sub
```

# 6. Reglas de diseño para la integración

1. La UI nunca debe contener lógica de navegación. Solo debe llamar a funciones públicas de modNavegacion.
2. La navegación nunca debe fallar silenciosamente. Si un módulo o miembro no existe, debe mostrarse un mensaje claro.
3. La navegación debe ser rápida y directa. Sin animaciones, sin esperas, sin pasos intermedios.
4. La navegación debe ser determinista. El mismo resultado siempre lleva al mismo punto del código.
5. La navegación debe ser extensible. Preparada para:
- resaltar líneas
- seleccionar bloques
- mostrar tooltips (futuro)
- abrir panel de detalles (versión 3.0)
  
# 7. Problemas comunes y soluciones

❌ El módulo no se abre
✅ Verificar que el nombre coincide exactamente con el de VBComponents.

❌ El miembro no se selecciona
✅ Asegurarse de que ProcStartLine devuelve un valor > 0.

❌ La línea no se selecciona
✅ Confirmar que el número de línea existe en el módulo.

❌ Unicode aparece como ? en el editor
✅ Nunca insertar Unicode directamente en el código.
✅ Usar siempre tblUnicode + IconoUnicode.

# 8. Estado actual (Versión 2.0)

✅ Navegación a módulos
✅ Navegación a miembros
✅ Navegación a líneas
✅ Integración estable con el Editor VBA
✅ Arquitectura limpia y extensible

# 9. Próximas ampliaciones (2.1 / 3.0)

- Resaltar línea analizada
- Resaltar bloque del procedimiento
- Navegación inversa (desde el editor al Inspector)
- Panel de detalles sincronizado
- Modo “seguir análisis” paso a paso
  
---

# ✅ Documento 6 completado.
