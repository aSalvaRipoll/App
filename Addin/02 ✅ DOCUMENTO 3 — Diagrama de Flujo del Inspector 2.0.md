# Inspector VBA — Diagrama de Flujo (Versión 2.0)

Este documento describe el flujo completo del Inspector, desde la inicialización hasta la visualización de resultados.  
El objetivo es proporcionar una visión clara, estructurada y mantenible del funcionamiento interno del sistema.

---

## 1. Visión General del Flujo

El Inspector sigue un flujo en cinco fases:

1. **Inicialización**
2. **Análisis del proyecto**
3. **Ejecución de reglas**
4. **Agregación de resultados**
5. **Visualización e interacción**

---

## 2. Diagrama General del Proceso

```text
┌────────────────────┐
│   Inicialización    │
└──────────┬─────────┘
           │
           ▼
┌────────────────────┐
│   Cargar Proyecto   │
│  (clsProyectoVBA)   │
└──────────┬─────────┘
           │
           ▼
┌────────────────────┐
│   Analizar Módulos  │
│   (clsAnalizador)   │
└──────────┬─────────┘
           │
           ▼
┌────────────────────┐
│  Ejecutar Reglas    │
│   (modReglas)       │
└──────────┬─────────┘
           │
           ▼
┌────────────────────┐
│  Generar Resultados │
│ (clsResultadoAnal.) │
└──────────┬─────────┘
           │
           ▼
┌────────────────────┐
│ Agregar y Ordenar   │
│   (clsResultados)   │
└──────────┬─────────┘
           │
           ▼
┌────────────────────┐
│   Mostrar en UI     │
│   (frmInspector)    │
└────────────────────┘

## 3. Fase 1 — Inicialización

frmInspector → cmdAnalizar_Click

**Acciones:**

- Se inicializa el analizador (clsAnalizador)
- Se cargan iconos desde tblUnicode (vía 04_modFunciones)
- Se preparan reglas (modReglas)
- Se limpia el ListBox de resultados

Objetivo: dejar el sistema listo para analizar el proyecto.

## 4. Fase 2 — Carga del Proyecto

clsAnalizador → clsProyectoVBA

**Acciones:**

- Se obtiene el VBProject activo
- Se enumeran módulos estándar
- Se enumeran módulos de clase
- Se crean objetos clsModuloVBA y clsMiembroVBA

Objetivo: construir un modelo interno del proyecto.

## 5. Fase 3 — Análisis de Módulos

clsAnalizador → AnalizarModulo(mod)

**Para cada módulo:**
- Se analiza su código
- Se detectan miembros (Sub, Function, Property)
- Se identifican líneas relevantes
- Se prepara el contexto para las reglas

Objetivo: extraer información estructural.

## 6. Fase 4 — Ejecución de Reglas

modReglas → EjecutarReglas(mod, miembro)

**Cada regla:**

- Evalúa una condición
- Determina severidad
- Genera un clsResultadoAnalisis si aplica

**Ejemplos de reglas:**

- Nombres incorrectos
- Procedimientos demasiado largos
- Falta de Option Explicit
- Código muerto (futuro)
- Complejidad ciclomática (futuro)

Objetivo: detectar problemas y oportunidades de mejora.

## 7. Fase 5 — Agregación de Resultados

clsResultados → Add(resultado)

**Acciones:**

- Se agregan todos los resultados
- Se ordenan según criterio por defecto
- Se preparan para visualización

Objetivo: consolidar la información en un formato uniforme.

## 8. Fase 6 — Visualización en la UI

**Acciones:**

- Se formatea cada resultado:
- - Icono de severidad
- - Icono de elemento
- - Icono de miembro
- - Línea
- Se añade al ListBox
- Se aplican truncados si es necesario
- Se actualizan indicadores de ordenación

Objetivo: presentar los resultados de forma clara y profesional.

## 9. Interacción del Usuario

### 9.1 Ordenación por columnas

lblSeveridad_Click → AlternarOrden → OrdenarPor → MostrarResultados

- Orden asc/desc alternada
- Indicadores visuales (flechas Unicode)
- Estado persistente entre clics

### 9.2 Navegación al código

lstResultados_Click → modNavegacion → Editor VBA

- Seleccionar módulo
- Seleccionar miembro
- Ir a línea específica

## 10. Flujo Completo (Resumen)

[Usuario pulsa Analizar]
        │
        ▼
Inicializar → Cargar Proyecto → Analizar Módulos → Ejecutar Reglas
        │
        ▼
Generar Resultados → Agregar → Ordenar → Mostrar en UI
        │
        ├── Clic en encabezado → Reordenar → Mostrar
        └── Clic en resultado → Navegar al código

## 11. Estado actual (Versión 2.0)

- ✅ Flujo completo implementado 
- ✅ Ordenación avanzada 
- ✅ Iconografía centralizada 
- ✅ Navegación al editor 
- ✅ Arquitectura modular 
- ✅ Preparado para paneles adicionales

## 12. Próximos pasos (2.1 / 3.0)

- Panel de detalles
- Filtros avanzados
- Exportación
- Reglas configurables
- Métricas de rendimiento
- Integración con Git

---

# ✅ Documento 3 completado.
