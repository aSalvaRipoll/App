📘 Guía de Extensión – InspectorVBA 2.2

# 🧩 Guía de Extensión – InspectorVBA 2.2
Guía oficial para ampliar el InspectorVBA manteniendo su arquitectura, estilo y estándares.

---

## 🎯 Objetivo de esta guía

El InspectorVBA está diseñado para ser modular, extensible y seguro.  
Esta guía explica cómo añadir nuevas funcionalidades sin romper la arquitectura existente.

Incluye:

- Dónde colocar nuevo código  
- Cómo crear nuevas reglas de análisis  
- Cómo extender exportaciones  
- Cómo añadir nuevas entidades (clases)  
- Cómo integrar nuevas opciones en la interfaz  
- Buenas prácticas y patrones recomendados  

---

# 2. Estructura base que debes respetar

El InspectorVBA se organiza en capas y subsistemas bien definidos.  
Respetar esta estructura garantiza que cualquier extensión sea estable, mantenible y coherente.

---

## 2.1 Capa pública (fachada)

- 00_modMain  
- Expone funciones públicas limpias  
- No contiene lógica interna  
- Es el punto de entrada para la interfaz (Ribbon, menús, botones)

La regla fundamental es que esta capa solo delega, nunca implementa.

---

## 2.2 Capa Core

- 02_modCore  
- Orquesta el flujo completo del Inspector:
  - Inicialización  
  - Análisis  
  - Reparación  
  - Exportación  
  - Reset  

El Core coordina, pero no implementa detalles.

---

## 2.3 Subsistemas principales

- 10–19 → Análisis  
- 30–39 → Reparación  
- 40–49 → Exportación  
- 50–59 → Navegación y utilidades internas  
- 60–69 → Entorno, configuración y preferencias  
- 70–79 → Interfaz (Ribbon, menús, callbacks)  
- 90–99 → Prototipos y extensiones experimentales  

---

## 2.4 Modelo de datos (clases ds*)

Las clases representan entidades del análisis:

- dsSimbolo  
- dsMiembro  
- dsModulo  
- dsClase  
- dsCatalogoInspector  
- dsResultadoAnalisis  
- dsResultados  
- dsEstadisticas  
- dsInformeSimbolos  

Estas clases contienen datos y utilidades, no lógica de análisis ni de UI.

---

## 2.5 Regla de oro de la arquitectura

Cada módulo hace una sola cosa.  
Cada clase representa una sola entidad.  
Cada subsistema tiene un propósito único.

---

# 3. Crear nuevas reglas de análisis

El sistema de análisis del InspectorVBA está diseñado para ser modular y ampliable.

---

## 3.1 Ubicación de las reglas

Las reglas deben implementarse exclusivamente en:

- 13_modReglas  
- dsResultadoAnalisis  
- dsResultados  
- dsCatalogoInspector  

---

## 3.2 Flujo general de una regla

1. Recorrer los elementos del catálogo.  
2. Evaluar la condición de la regla.  
3. Crear un resultado.  
4. Añadirlo a la colección global.  
5. Registrar la actividad en los logs.  

(Ejemplo 1) — Crear un resultado de análisis

---

## 3.3 Tipos de elementos que puede analizar una regla

- Símbolos  
- Módulos estándar  
- Módulos de clase  
- Formularios  
- Miembros  
- Referencias  
- Estructuras del proyecto  

---

## 3.4 Buenas prácticas al crear reglas

- Mantén cada regla en un procedimiento independiente.  
- Usa nombres descriptivos.  
- Evita duplicar lógica.  
- No mezcles análisis con reparación o exportación.  
- Documenta cada regla.  

---

## 3.5 Pruebas recomendadas

1. Proyecto pequeño  
2. Proyecto grande  
3. Proyecto vacío  
4. Proyecto con referencias rotas  
5. Falsos positivos  
6. Rendimiento  

---

## 3.6 Checklist final

- Está en 13_modReglas  
- Usa dsResultadoAnalisis y dsResultados  
- No rompe la arquitectura  
- Está documentada  
- Está probada  

---

## 3.7 Ejemplo conceptual

(Ejemplo 2) — Regla que detecta funciones públicas sin comentario

---

# 4. Extender el sistema de reparación

---

## 4.1 Dónde implementar reparaciones

- 30_modReparar → manuales  
- 31_modAutoRepair → automáticas  

---

## 4.2 Flujo general de una reparación

1. Recibir un resultado.  
2. Identificar el elemento afectado.  
3. Aplicar la modificación.  
4. Registrar la acción.  
5. Actualizar el estado.  

(Ejemplo 3) — Reparación automática

---

## 4.3 Tipos de reparaciones habituales

- Cambiar visibilidad  
- Eliminar símbolos no usados  
- Renombrar duplicados  
- Corregir referencias  
- Insertar comentarios  
- Normalizar nombres  

---

## 4.4 Buenas prácticas

- Reparaciones pequeñas y claras  
- Registrar siempre  
- Evitar cambios masivos  
- Reparaciones arriesgadas → manuales  

---

## 4.5 Pruebas recomendadas

1. Proyecto pequeño  
2. Proyecto grande  
3. Proyecto con errores reales  
4. Sin inconsistencias  
5. Sin efectos colaterales  
6. Logs correctos  

---

## 4.6 Checklist final

- En módulo correcto  
- Asociada a resultados  
- Documentada  
- Probada  

---

# 5. Extender el sistema de exportación

---

## 5.1 Dónde viven las exportaciones

- 41_modExportAux  
- 42_modExportTXT  
- 43_modExportExcel  
- 44_modExportHTML  

---

## 5.2 Flujo general

1. Validar parámetros  
2. Preparar datos  
3. Generar archivo  
4. Registrar  
5. Devolver información  

(Ejemplo 4) — Exportación a TXT

---

## 5.3 Añadir un nuevo formato

1. Crear módulo 40–49  
2. Implementar formato  
3. Registrar en 41_modExportAux  
4. Añadir opción en interfaz  
5. Documentar  

---

## 5.4 Buenas prácticas

- Un módulo por formato  
- Reutilizar lógica  
- Registrar siempre  
- Mantener coherencia  

---

## 5.5 Pruebas recomendadas

1. Proyecto pequeño  
2. Proyecto grande  
3. Archivo válido  
4. Sin sobrescritura accidental  
5. Datos correctos  
6. Logs correctos  

---

# 6. Extender el modelo de datos (clases ds*)

---

## 6.1 Cuándo crear una nueva clase

- Nueva entidad del análisis  
- Datos estructurados  
- Acceso compartido  
- Evitar variables globales  

---

## 6.2 Buenas prácticas

- Prefijo ds  
- Propiedades públicas, campos privados  
- Sin lógica compleja  
- Documentación clara  

(Ejemplo 5) — Nueva clase ds*

---

## 6.3 Extender clases existentes

- Añadir propiedades  
- Añadir utilidades  
- Mantener compatibilidad  

---

# 7. Interfaz: Ribbon, menús y navegación

---

## 7.1 Dónde vive la interfaz

- 70_modRibbon  
- 71_modMenus  
- 50_modNavegacion  

---

## 7.2 Extender el Ribbon

1. Añadir control en XML  
2. Crear callback  
3. Delegar en 00_modMain  
4. Mantener coherencia  

(Ejemplo 6) — Callback de Ribbon

---

## 7.3 Extender menús

1. Registrar comando  
2. Asociarlo a función pública  
3. Mantener estructura  

---

## 7.4 Navegación

- Abrir módulos  
- Seleccionar miembros  
- Posicionar cursor  
- Resaltar elementos  

(Ejemplo 7) — Navegar a un miembro

---

# 8. Entorno, configuración y preferencias

---

## 8.1 Dónde vive

- 60_modEntorno  
- 61_modPreferencias  
- 62_modConfig  

---

## 8.2 Entorno

- Idioma  
- Rutas  
- Versión  
- Estado del editor  

---

## 8.3 Preferencias

- Análisis  
- Exportación  
- Reparación  
- Comportamiento  

(Ejemplo 8) — Nueva preferencia

---

## 8.4 Configuración interna

- Constantes  
- Parámetros  
- Ajustes internos  

---

# 9. Buenas prácticas generales

---

## 9.1 Principios fundamentales

- Una responsabilidad por módulo  
- Una entidad por clase  
- UI sin lógica  
- Análisis sin modificaciones  
- Reparación sin análisis  
- Exportación sin UI  

---

## 9.2 Organización del código

- Respetar numeración  
- Nombres descriptivos  
- Evitar duplicación  
- Documentar todo  

---

## 9.3 Logs y trazabilidad

(Ejemplo 9) — Registrar en logs

---

## 9.4 Rendimiento

- Evitar bucles innecesarios  
- Minimizar accesos al editor  
- Probar con proyectos grandes  

---

# 10. Ejemplo completo de extensión

---

## 10.1 Objetivo

Detectar funciones públicas sin comentario e insertar encabezado.

---

## 10.2 Análisis

(Ejemplo 10) — Regla completa

---

## 10.3 Reparación

(Ejemplo 11) — Reparación automática

---

## 10.4 Exportación

(Ejemplo 12) — Exportación del resultado

---

## 10.5 Interfaz

(Ejemplo 13) — Botón en Ribbon

---

## 10.6 Pruebas finales

- Análisis  
- Reparación  
- Exportación  
- Interfaz  
- Logs  

---

# 11. Cierre de la guía

Esta guía proporciona la estructura oficial para extender el InspectorVBA 2.2 de forma segura, modular y profesional.  
Siguiendo estos principios, cualquier extensión será coherente, mantenible y totalmente integrada.
