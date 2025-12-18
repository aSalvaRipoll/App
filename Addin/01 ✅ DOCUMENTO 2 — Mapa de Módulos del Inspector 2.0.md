# Inspector VBA — Mapa de Módulos (Versión 2.0)

Este documento describe todos los módulos del Inspector, su numeración, responsabilidades y relaciones.  
La numeración sigue un esquema profesional que facilita la lectura, el mantenimiento y la escalabilidad del proyecto.

---

# 1. Convenciones de numeración

Los módulos estándar siguen el patrón:

- NN_modNombre
  
Donde:

- **NN** = número de orden (00–99)
- **mod** = módulo estándar
- **Nombre** = responsabilidad principal

Las clases siguen el patrón:

- clsNombre

Los formularios:

- frmNombre

---

# 2. Estructura general por rangos

| Rango | Propósito |
|-------|-----------|
| **00–09** | Infraestructura, utilidades, constantes |
| **10–19** | Clases y modelo del Inspector |
| **20–29** | Lógica del análisis y reglas |
| **30–39** | Interfaz de usuario |
| **40–49** | Exportación y reporting (futuro) |
| **50–99** | Reservado para futuras versiones |

---

# 3. Módulos actuales (Inspector 2.0)

## ✅ **00–09 — Infraestructura**

### **04_modFunciones**

Módulo central de utilidades generales:

- Iconos Unicode (desde `tblUnicode`)
- Formato visual (elemento, miembro, severidad)
- Normalización de texto
- Truncado de cadenas
- Conversión de colecciones
- Funciones auxiliares para columnas
- Preparado para futuras utilidades transversales

> Este módulo es la base de toda la interfaz y formateo del Inspector.

---

## ✅ **10–19 — Clases y Modelo**

### **clsResultadoAnalisis**

Representa un resultado individual:

- Severidad  
- Elemento  
- Miembro  
- Línea  
- Mensaje  
- Clave única  

### **clsResultados**

Colección de resultados:

- Agregar resultados  
- Ordenar por columna  
- Filtrar (futuro)  
- Exportar (futuro)  

### **clsRegla**

Representa una regla del Inspector:

- Nombre  
- Descripción  
- Severidad  
- Método de evaluación  

### **clsAnalizador**

Motor principal del análisis:

- Recorre módulos y miembros  
- Ejecuta reglas  
- Genera resultados  

### **clsProyectoVBA**

Representa el proyecto completo:

- Lista de módulos  
- Lista de clases  
- Acceso al VBProject  

### **clsModuloVBA**

Representa un módulo estándar:

- Nombre  
- Tipo  
- Miembros  
- Código  

### **clsMiembroVBA**

Representa un procedimiento:

- Nombre  
- Tipo (Sub, Function, Property)  
- Línea de inicio  
- Código asociado  

---

## ✅ **20–29 — Lógica del Inspector**

### **20_modAnalisis**

Coordinación del análisis:

- Inicializa el analizador  
- Ejecuta reglas  
- Devuelve `clsResultados`  

### **21_modReglas**

Contiene las reglas del Inspector:

- Reglas de estilo  
- Reglas de limpieza  
- Reglas de estructura  
- Reglas de rendimiento (futuro)  
- Reglas de seguridad (futuro)  

### **22_modNavegacion**

Integración con el Editor VBA:

- Ir a módulo  
- Ir a miembro  
- Seleccionar línea  
- Resaltar código (futuro)  

---

## ✅ **30–39 — Interfaz de Usuario**

### **frmInspector**

Formulario principal:

- ListBox de resultados  
- Encabezados clicables  
- Orden asc/desc con indicadores  
- Botón de análisis  
- Panel de estado (futuro)  
- Panel de detalles (futuro)  

---

# 4. Módulos futuros (reservados)

## ✅ **40–49 — Exportación y Reporting**

### **40_modExportacion** *(futuro)*

- Exportación a texto  
- Exportación a Excel  
- Exportación a Markdown  
- Exportación a JSON (opcional)  

### **41_modLogs** *(futuro)*

- Registro de análisis  
- Auditoría  
- Historial de ejecuciones  

---

# 5. Módulos reservados para versiones 3.x y 4.x

| Número | Nombre sugerido | Propósito |
|--------|------------------|-----------|
| 50 | modFiltros | Filtros avanzados en UI |
| 51 | modCategorias | Categorías de reglas |
| 52 | modConfig | Configuración del usuario |
| 53 | modRendimiento | Métricas y tiempos |
| 54 | modInspectorUI | Paneles adicionales |
| 60–69 | modReglasAvanzadas | Reglas de seguridad, rendimiento, arquitectura |
| 70–79 | modIntegraciones | Git, exportación externa |
| 80–99 | Reservado | Expansión futura |

---

# 6. Relación entre módulos

```text
frmInspector
    ↓ muestra
clsResultados
    ↓ contiene
clsResultadoAnalisis
    ↑ generado por
clsAnalizador
    ↑ usa reglas de
21_modReglas
    ↑ coordinado por
20_modAnalisis
    ↑ usa utilidades de
04_modFunciones

# ✅ Documento 2 completado.
