# Inspector VBA — Diccionario de Iconos (tblUnicode)

Este documento define la estructura, categorías, reglas y catálogo de iconos utilizados por el Inspector.  
Todos los iconos se almacenan en la tabla `tblUnicode` y se consumen a través del módulo `04_modFunciones`.

El objetivo es centralizar la iconografía para:

- Evitar problemas con Unicode en el editor VBA
- Permitir cambios visuales sin modificar código
- Facilitar la expansión en futuras versiones (2.x, 3.x, 4.x)
- Mantener una estética coherente y profesional

---

# 1. Estructura de la tabla `tblUnicode`

La tabla debe contener al menos estas columnas:

| Campo        | Tipo        | Descripción |
|--------------|-------------|-------------|
| **ID**       | Número      | Identificador interno |
| **Nombre**   | Texto corto | Clave única usada en código |
| **Texto**    | Texto corto | Icono Unicode |
| **Categoria**| Texto corto | Grupo funcional |
| **Descripcion** | Texto largo | Explicación del uso |

Ejemplo:

| ID | Nombre | Texto | Categoria | Descripcion |
|----|--------|--------|-----------|-------------|
| 1 | Info | ℹ️ | Severidad | Mensaje informativo |

---

# 2. Categorías oficiales

Los iconos se agrupan en categorías para facilitar su uso y mantenimiento:

- **Severidad**  
  Iconos usados para INFO, AVISO, ERROR, CRÍTICO…

- **Estado**  
  Iconos para elementos nuevos, modificados, eliminados, bloqueados…

- **Acción**  
  Iconos para botones, comandos, navegación, exportación…

- **Elemento**  
  Iconos para módulos, clases, formularios, informes…

- **Inspector**  
  Iconos internos del Inspector (reglas, resultados, análisis…)

- **Rendimiento**  
  Iconos para tiempos, velocidad, optimización…

- **Orden**  
  Iconos para orden ascendente/descendente en encabezados.

---

# 3. Catálogo de iconos actuales (Inspector 2.0)

## ✅ 3.1 Iconos de Severidad

| Nombre | Icono | Descripción |
|--------|--------|-------------|
| Info | ℹ️ | Información general |
| Aviso | ⚠️ | Advertencia |
| Error | ❗ | Error |
| Critico | ❌ | Error crítico |
| Info2 | 🛈 | Alternativa a ℹ️ |
| Aviso2 | ❕ | Advertencia leve |
| Error2 | ❗❗ | Error doble |
| AdvertenciaSuave | ⚠ | Advertencia suave |
| AdvertenciaFuerte | ⚠️⚠️ | Advertencia fuerte |

---

## ✅ 3.2 Iconos de Estado

| Nombre | Icono | Descripción |
|--------|--------|-------------|
| Ok | ✅ | Correcto |
| Nuevo | ✨ | Nuevo elemento |
| Editado | ✏️ | Modificado |
| Eliminado | 🗑️ | Eliminado |
| Bloqueado | 🔒 | Bloqueado |
| Desbloqueado | 🔓 | Desbloqueado |
| Experimental | 🧪 | Función experimental |

---

## ✅ 3.3 Iconos de Acción

| Nombre | Icono | Descripción |
|--------|--------|-------------|
| Buscar | 🔍 | Buscar / localizar |
| Depurar | 🐞 | Depuración |
| Config | ⚙️ | Configuración |
| Exportar | 📤 | Exportar datos |
| Importar | 📥 | Importar datos |
| Filtrar | 🔽 | Filtro |
| Ordenar | ↕️ | Ordenación |
| Pregunta | ❓ | Ayuda |

---

## ✅ 3.4 Iconos de Elemento

| Nombre | Icono | Descripción |
|--------|--------|-------------|
| Archivo | 📄 | Archivo genérico |
| Carpeta | 📁 | Carpeta / contenedor |
| Clase | 🧩 | Módulo de clase |
| Modulo | 📘 | Módulo estándar |
| Funcion | 🔧 | Procedimiento o función |
| Evento | 🎯 | Evento |

---

## ✅ 3.5 Iconos del Inspector

| Nombre | Icono | Descripción |
|--------|--------|-------------|
| Regla | 📏 | Regla del Inspector |
| Resultado | 📊 | Resultado del análisis |
| InfoDetallada | 📝 | Detalles del resultado |
| Analisis | 🔎 | Análisis del proyecto |

---

## ✅ 3.6 Iconos de Rendimiento

| Nombre | Icono | Descripción |
|--------|--------|-------------|
| Tiempo | ⏱️ | Operación lenta |
| RendimientoAlto | 🚀 | Muy rápido |
| RendimientoBajo | 🐢 | Muy lento |

---

## ✅ 3.7 Iconos de Orden

| Nombre | Icono | Descripción |
|--------|--------|-------------|
| FlechaArriba | ▲ | Orden ascendente |
| FlechaAbajo | ▼ | Orden descendente |

---

# 4. Reglas de uso

1. **Nunca insertar Unicode directamente en el código VBA.**  
   Siempre usar `IconoUnicode("Nombre")`.
2. **Cada icono debe tener un nombre único.**
3. **Las categorías deben mantenerse coherentes.**
4. **Los iconos deben ser simples y legibles.**
5. **Los iconos de severidad deben ser visualmente distintos.**
6. **Los iconos de ordenación deben ser monocromáticos y discretos.**
7. **Los iconos nuevos deben añadirse siempre al final de la tabla.**

---

# 5. Ejemplos de uso en código

```vba
severidad = IconoSeveridad(item.Severidad)
elemento = IconoElemento(item.tipoElemento)
miembro = IconoMiembro(item.tipoMiembro)
flecha = IconoUnicode("FlechaArriba")```

# 6. Estado actual (Versión 2.0)

- ✅ Iconografía completa 
- ✅ Tabla preparada para crecer 
- ✅ Integración total con 04_modFunciones 
- ✅ Indicadores de ordenación incluidos 
- ✅ Lista ampliable sin tocar código

# 7. Próximas ampliaciones (2.1 / 3.0)

- Iconos para categorías de reglas
- Iconos para paneles adicionales
- Iconos para exportación avanzada
- Iconos para estados de análisis
- Iconos para métricas de complejidad


---

# ✅ Documento 4 completado.
