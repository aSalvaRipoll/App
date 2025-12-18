## 🧩 Diagrama de arquitectura de módulos – InspectorVBA 2.2

Este diagrama representa la estructura modular del InspectorVBA, agrupando los módulos por función y responsabilidad. Cada bloque muestra los grupos funcionales principales del sistema, con sus relaciones jerárquicas y dependencias.

![Diagrama de arquitectura](sandbox:/mnt/data/graphic_art/InspectorVBA_Modular_Architecture.png)

### 🔹 Grupos funcionales

- **00–09**: Núcleo y utilidades generales  
- **10–19**: Análisis del proyecto  
- **30–39**: Reparación y autoreparación  
- **40–49**: Exportación  
- **50–59**: Navegación, logs y depuración  
- **60–69**: Entorno y preferencias  
- **70–79**: Interfaz (Ribbon y menús)  
- **90–99**: Stub y extensiones futuras  

### 🔹 Propósito

Este diagrama sirve como referencia visual para entender la arquitectura del InspectorVBA, facilitar la documentación técnica y guiar futuras extensiones o colaboraciones.

---
## 📑 Tabla de módulos y responsabilidades – InspectorVBA 2.2

| Grupo | Módulo | Descripción |
|-------|--------|-------------|
| **00–09 Núcleo y utilidades** |||
| 00 | modMain | Punto de entrada público del Inspector. Delegación hacia el Core. |
| 01 | modConstantes | Enumeraciones, constantes y valores globales. |
| 02 | modCore | Lógica central: inicialización, análisis, reparación, reset. |
| 03 | modVBIDE | Acceso al editor VBA, inspección del proyecto, navegación interna. |
| 04 | modFunciones | Funciones generales reutilizables en todo el Inspector. |
| 05 | modMensajes | Textos, mensajes y utilidades de comunicación con el usuario. |
| **10–19 Análisis del proyecto** |||
| 10 | modAnalisisAux | Funciones auxiliares para el análisis del proyecto. |
| 11 | modBuscarObjetos | Localización de módulos, formularios, clases y componentes. |
| 12 | modBuscarReferencias | Detección de referencias, dependencias y vínculos. |
| 13 | modReglas | Reglas de inspección, validación y análisis estático. |
| 14 | modSimbolos | Catálogo de símbolos, miembros y elementos analizados. |
| **30–39 Reparación** |||
| 30 | modReparar | Aplicación de reparaciones manuales sobre el proyecto. |
| 31 | modAutoRepair | Reparaciones automáticas y sugeridas por el Inspector. |
| **40–49 Exportación** |||
| 40 | mod_ControlRutasExportacion | Validación, normalización y preparación de rutas de exportación. |
| 41 | modExportAux | Coordinador de exportación, resumen, extensiones y utilidades. |
| 42 | modExportTXT | Exportación a formato TXT. |
| 43 | modExportExcel | Exportación a Excel (XLSX). |
| 44 | modExportHTML | Exportación a HTML con estilos. |
| **50–59 Navegación, logs y depuración** |||
| 50 | modNavegacion | Navegación entre elementos del proyecto desde la interfaz. |
| 51 | modLogs | Registro de acciones, errores y eventos del Inspector. |
| 52 | modDebug | Herramientas internas de depuración. |
| **60–69 Entorno y preferencias** |||
| 60 | modEntorno | Detección del entorno de ejecución y configuración base. |
| 61 | modEntornoInspector | Configuración específica del InspectorVBA. |
| 62 | modInicioUsuario | Inicialización personalizada según el usuario. |
| 63 | modInicioFin | Flujo de arranque y cierre del Inspector. |
| 64 | modPreferencias | Carga y guardado de preferencias del usuario. |
| **70–79 Interfaz** |||
| 70 | modRibbon | Definición y callbacks del Ribbon personalizado. |
| 71 | modMenus | Menús contextuales y comandos asociados. |
| **90–99 Extensiones y stub** |||
| 90 | ModStub | Módulo de pruebas, prototipos y extensiones futuras. |

## 🧩 Diagrama extendido de arquitectura – InspectorVBA 2.2

Este diagrama representa la arquitectura completa del InspectorVBA, incluyendo:

- Módulos estándar agrupados por función (00–99)
- Módulos de clase que encapsulan lógica, datos y entidades del análisis
- Relaciones jerárquicas y funcionales entre componentes

![Diagrama extendido](sandbox:/mnt/data/graphic_art/InspectorVBA_Modular_Architecture.png)

### 🔹 Módulos de clase incluidos

- `dsAddin`  
- `dsCatalogoInspector`  
- `dsCatalogoSimbolos`  
- `dsClase`  
- `dsEstadisticas`  
- `dsInformeSimbolos`  
- `dsMiembro`  
- `dsModulo`  
- `dsResultadoAnalisis`  
- `dsResultados`  
- `dsSimbolo`  

Estos módulos encapsulan entidades clave del análisis y permiten una arquitectura orientada a objetos dentro del entorno VBA.

---
## 🧩 Diagrama extendido de arquitectura – InspectorVBA 2.2

Este diagrama representa la arquitectura completa del InspectorVBA, incluyendo:

- Módulos estándar agrupados por función (00–99)
- Módulos de clase que encapsulan lógica, datos y entidades del análisis
- Relaciones jerárquicas y funcionales entre componentes

![Diagrama extendido](sandbox:/mnt/data/graphic_art/InspectorVBA_Modular_Architecture.png)

### 🔹 Módulos de clase incluidos

| Clase | Propósito |
|-------|-----------|
| `dsAddin` | Representa el Add-In y su integración con Access |
| `dsCatalogoInspector` | Catálogo principal de símbolos inspeccionados |
| `dsCatalogoSimbolos` | Catálogo auxiliar de símbolos individuales |
| `dsClase` | Representación de clases VBA |
| `dsEstadisticas` | Cálculo y almacenamiento de estadísticas del análisis |
| `dsInformeSimbolos` | Generación de informes sobre símbolos no usados |
| `dsMiembro` | Propiedades, métodos y eventos de clases o módulos |
| `dsModulo` | Representación de módulos estándar y de clase |
| `dsResultadoAnalisis` | Resultado individual de una regla aplicada |
| `dsResultados` | Colección de resultados del análisis completo |
| `dsSimbolo` | Entidad básica del análisis: variable, función, propiedad, etc. |

---
## 📑 Tabla de clases y propósito – InspectorVBA 2.2

| Clase | Propósito |
|-------|-----------|
| `dsAddin` | Representa el Add-In y su integración con Access. |
| `dsCatalogoInspector` | Catálogo principal de símbolos inspeccionados. |
| `dsCatalogoSimbolos` | Catálogo auxiliar de símbolos individuales. |
| `dsClase` | Representación de clases VBA. |
| `dsEstadisticas` | Cálculo y almacenamiento de estadísticas del análisis. |
| `dsInformeSimbolos` | Generación de informes sobre símbolos no usados. |
| `dsMiembro` | Propiedades, métodos y eventos de clases o módulos. |
| `dsModulo` | Representación de módulos estándar y de clase. |
| `dsResultadoAnalisis` | Resultado individual de una regla aplicada. |
| `dsResultados` | Colección de resultados del análisis completo. |
| `dsSimbolo` | Entidad básica del análisis: variable, función, propiedad, etc. |

## 🔗 Tabla de dependencias entre clases – InspectorVBA 2.2

| Clase | Depende de | Relación |
|-------|------------|----------|
| `dsCatalogoInspector` | `dsSimbolo`, `dsModulo`, `dsClase`, `dsMiembro` | Contiene todos los elementos inspeccionados. |
| `dsCatalogoSimbolos` | `dsSimbolo` | Subconjunto filtrado de símbolos. |
| `dsResultados` | `dsResultadoAnalisis` | Colección de resultados generados por reglas. |
| `dsResultadoAnalisis` | `dsSimbolo`, `dsModulo` | Resultado vinculado a un símbolo o módulo. |
| `dsInformeSimbolos` | `dsCatalogoSimbolos`, `dsEstadisticas` | Genera informes a partir de símbolos y estadísticas. |
| `dsEstadisticas` | `dsResultados`, `dsCatalogoInspector` | Calcula métricas a partir del análisis completo. |
| `dsMiembro` | `dsClase`, `dsModulo` | Pertenece a una clase o módulo. |
| `dsModulo` | `dsMiembro`, `dsSimbolo` | Contiene miembros y símbolos. |
| `dsClase` | `dsMiembro` | Contiene miembros propios. |
| `dsAddin` | `dsCatalogoInspector`, `dsResultados` | Orquesta la ejecución y exportación del análisis. |

## 🧱 Inventario completo: formularios, módulos y clases – InspectorVBA 2.2

| Tipo        | Nombre                       | Grupo / Nº | Descripción |
|------------|------------------------------|-----------|-------------|
| **Formulario** | Form_frmInicio              | –         | Pantalla inicial del Inspector, punto de entrada visual. |
| **Formulario** | Form_frmResultados          | –         | Visualización de resultados del análisis. |
| **Formulario** | Form_subExportarInspector   | –         | Subformulario para opciones y acciones de exportación. |
| **Módulo**  | 00_modMain                   | 00–09    | Punto de entrada público del Inspector. Delegación hacia el Core. |
| **Módulo**  | 01_modConstantes             | 00–09    | Enumeraciones, constantes y valores globales. |
| **Módulo**  | 02_modCore                   | 00–09    | Lógica central: inicialización, análisis, reparación, reset. |
| **Módulo**  | 03_modVBIDE                  | 00–09    | Acceso al editor VBA, inspección del proyecto, navegación interna. |
| **Módulo**  | 04_modFunciones              | 00–09    | Funciones generales reutilizables. |
| **Módulo**  | 05_modMensajes               | 00–09    | Textos, mensajes y utilidades de comunicación con el usuario. |
| **Módulo**  | 10_modAnalisisAux            | 10–19    | Funciones auxiliares para el análisis del proyecto. |
| **Módulo**  | 11_modBuscarObjetos          | 10–19    | Localización de módulos, formularios, clases y componentes. |
| **Módulo**  | 12_modBuscarReferencias      | 10–19    | Detección de referencias, dependencias y vínculos. |
| **Módulo**  | 13_modReglas                 | 10–19    | Reglas de inspección, validación y análisis estático. |
| **Módulo**  | 14_modSimbolos               | 10–19    | Catálogo y gestión de símbolos y elementos analizados. |
| **Módulo**  | 30_modReparar                | 30–39    | Aplicación de reparaciones manuales sobre el proyecto. |
| **Módulo**  | 31_modAutoRepair             | 30–39    | Reparaciones automáticas y sugeridas por el Inspector. |
| **Módulo**  | 40_mod_ControlRutasExportacion | 40–49 | Validación, normalización y preparación de rutas de exportación. |
| **Módulo**  | 41_modExportAux              | 40–49    | Coordinador de exportación, resumen, extensiones y utilidades. |
| **Módulo**  | 42_modExportTXT              | 40–49    | Exportación a formato TXT. |
| **Módulo**  | 43_modExportExcel            | 40–49    | Exportación a Excel (XLSX). |
| **Módulo**  | 44_modExportHTML             | 40–49    | Exportación a HTML con estilos. |
| **Módulo**  | 50_modNavegacion             | 50–59    | Navegación entre elementos del proyecto desde la interfaz. |
| **Módulo**  | 51_modLogs                   | 50–59    | Registro de acciones, errores y eventos del Inspector. |
| **Módulo**  | 52_modDebug                  | 50–59    | Herramientas internas de depuración. |
| **Módulo**  | 60_modEntorno                | 60–69    | Detección del entorno de ejecución y configuración base. |
| **Módulo**  | 61_modEntornoInspector       | 60–69    | Configuración específica del InspectorVBA. |
| **Módulo**  | 62_modInicioUsuario          | 60–69    | Inicialización personalizada según el usuario. |
| **Módulo**  | 63_modInicioFin              | 60–69    | Flujo de arranque y cierre del Inspector. |
| **Módulo**  | 64_modPreferencias           | 60–69    | Carga y guardado de preferencias del usuario. |
| **Módulo**  | 70_modRibbon                 | 70–79    | Definición y callbacks del Ribbon personalizado. |
| **Módulo**  | 71_modMenus                  | 70–79    | Menús contextuales y comandos asociados. |
| **Módulo**  | 90_ModStub                   | 90–99    | Módulo de pruebas, prototipos y extensiones futuras. |
| **Clase**   | dsAddin                      | Clases   | Representa el Add-In y su integración con Access. |
| **Clase**   | dsCatalogoInspector          | Clases   | Catálogo principal de símbolos inspeccionados. |
| **Clase**   | dsCatalogoSimbolos           | Clases   | Catálogo auxiliar de símbolos individuales. |
| **Clase**   | dsClase                      | Clases   | Representación de clases VBA. |
| **Clase**   | dsEstadisticas               | Clases   | Cálculo y almacenamiento de estadísticas del análisis. |
| **Clase**   | dsInformeSimbolos            | Clases   | Generación de informes sobre símbolos no usados. |
| **Clase**   | dsMiembro                    | Clases   | Propiedades, métodos y eventos de clases o módulos. |
| **Clase**   | dsModulo                     | Clases   | Representación de módulos estándar y de clase. |
| **Clase**   | dsResultadoAnalisis          | Clases   | Resultado individual de una regla aplicada. |
| **Clase**   | dsResultados                 | Clases   | Colección de resultados del análisis completo. |
| **Clase**   | dsSimbolo                    | Clases   | Entidad básica del análisis: variable, función, propiedad, etc. |

## 🗺️ Mapa de lectura – InspectorVBA 2.2

El InspectorVBA es un sistema modular y extensible. Esta guía te orienta sobre qué módulos y clases leer primero según el área que quieras comprender o extender.

---

### 🔵 1. Si quieres entender el funcionamiento general del Inspector
Empieza por:
- **00_modMain** → Punto de entrada público.
- **02_modCore** → Lógica central: inicialización, análisis, reparación, reset.
- **05_modMensajes** → Mensajes y textos clave.

---

### 🟢 2. Si quieres entender cómo se analiza un proyecto
Lee en este orden:
1. **10_modAnalisisAux** → Funciones auxiliares del análisis.  
2. **11_modBuscarObjetos** → Localización de módulos, formularios, clases.  
3. **12_modBuscarReferencias** → Dependencias y referencias.  
4. **13_modReglas** → Reglas de inspección.  
5. **14_modSimbolos** → Catálogo de símbolos.

Clases relevantes:
- `dsCatalogoInspector`
- `dsCatalogoSimbolos`
- `dsSimbolo`
- `dsModulo`
- `dsClase`
- `dsMiembro`

---

### 🟡 3. Si quieres entender cómo se generan los resultados
Revisa:
- **13_modReglas** → Cada regla produce un resultado.
- **14_modSimbolos** → Estructura de símbolos.
- **02_modCore** → Ensamblado final de resultados.

Clases clave:
- `dsResultadoAnalisis`
- `dsResultados`
- `dsEstadisticas`

---

### 🟠 4. Si quieres entender la reparación del proyecto
Orden recomendado:
1. **30_modReparar** → Reparaciones manuales.  
2. **31_modAutoRepair** → Reparaciones automáticas.  

Clases relacionadas:
- `dsResultadoAnalisis`  
- `dsResultados`  

---

### 🔴 5. Si quieres entender la exportación
Orden recomendado:
1. **40_mod_ControlRutasExportacion** → Validación y normalización de rutas.  
2. **41_modExportAux** → Coordinador de exportación.  
3. **42_modExportTXT**  
4. **43_modExportExcel**  
5. **44_modExportHTML**

Clases relacionadas:
- `dsInformeSimbolos`
- `dsResultados`
- `dsCatalogoInspector`

---

### 🟣 6. Si quieres entender la interfaz (Ribbon, menús, navegación)
Lee:
- **70_modRibbon** → Callbacks del Ribbon.  
- **71_modMenus** → Menús contextuales.  
- **50_modNavegacion** → Navegación entre elementos.  

---

### ⚙️ 7. Si quieres entender el entorno, arranque y preferencias
Orden recomendado:
1. **60_modEntorno**  
2. **61_modEntornoInspector**  
3. **62_modInicioUsuario**  
4. **63_modInicioFin**  
5. **64_modPreferencias**

---

### 🧪 8. Si quieres experimentar o extender el Inspector
Módulo pensado para pruebas:
- **90_ModStub**

---

### 🧱 9. Si quieres entender las entidades del modelo (clases)
Empieza por:
- `dsSimbolo` → La unidad básica del análisis.  
- `dsMiembro` → Propiedades, métodos, eventos.  
- `dsModulo` → Representación de módulos.  
- `dsClase` → Representación de clases.  
- `dsCatalogoInspector` → El “árbol” completo del proyecto.  
- `dsResultadoAnalisis` y `dsResultados` → Resultados del análisis.  
- `dsEstadisticas` → Métricas.  
- `dsInformeSimbolos` → Informes de símbolos no usados.

---

Este mapa te permite navegar el InspectorVBA de forma rápida y eficiente, entendiendo qué partes leer según tu objetivo.
## 🗺️ Mapa de lectura – InspectorVBA 2.2

El InspectorVBA es un sistema modular y extensible. Esta guía te orienta sobre qué módulos y clases leer primero según el área que quieras comprender o extender.

---

### 🔵 1. Si quieres entender el funcionamiento general del Inspector
Empieza por:
- **00_modMain** → Punto de entrada público.
- **02_modCore** → Lógica central: inicialización, análisis, reparación, reset.
- **05_modMensajes** → Mensajes y textos clave.

---

### 🟢 2. Si quieres entender cómo se analiza un proyecto
Lee en este orden:
1. **10_modAnalisisAux** → Funciones auxiliares del análisis.  
2. **11_modBuscarObjetos** → Localización de módulos, formularios, clases.  
3. **12_modBuscarReferencias** → Dependencias y referencias.  
4. **13_modReglas** → Reglas de inspección.  
5. **14_modSimbolos** → Catálogo de símbolos.

Clases relevantes:
- `dsCatalogoInspector`
- `dsCatalogoSimbolos`
- `dsSimbolo`
- `dsModulo`
- `dsClase`
- `dsMiembro`

---

### 🟡 3. Si quieres entender cómo se generan los resultados
Revisa:
- **13_modReglas** → Cada regla produce un resultado.
- **14_modSimbolos** → Estructura de símbolos.
- **02_modCore** → Ensamblado final de resultados.

Clases clave:
- `dsResultadoAnalisis`
- `dsResultados`
- `dsEstadisticas`

---

### 🟠 4. Si quieres entender la reparación del proyecto
Orden recomendado:
1. **30_modReparar** → Reparaciones manuales.  
2. **31_modAutoRepair** → Reparaciones automáticas.  

Clases relacionadas:
- `dsResultadoAnalisis`  
- `dsResultados`  

---

### 🔴 5. Si quieres entender la exportación
Orden recomendado:
1. **40_mod_ControlRutasExportacion** → Validación y normalización de rutas.  
2. **41_modExportAux** → Coordinador de exportación.  
3. **42_modExportTXT**  
4. **43_modExportExcel**  
5. **44_modExportHTML**

Clases relacionadas:
- `dsInformeSimbolos`
- `dsResultados`
- `dsCatalogoInspector`

---

### 🟣 6. Si quieres entender la interfaz (Ribbon, menús, navegación)
Lee:
- **70_modRibbon** → Callbacks del Ribbon.  
- **71_modMenus** → Menús contextuales.  
- **50_modNavegacion** → Navegación entre elementos.  

---

### ⚙️ 7. Si quieres entender el entorno, arranque y preferencias
Orden recomendado:
1. **60_modEntorno**  
2. **61_modEntornoInspector**  
3. **62_modInicioUsuario**  
4. **63_modInicioFin**  
5. **64_modPreferencias**

---

### 🧪 8. Si quieres experimentar o extender el Inspector
Módulo pensado para pruebas:
- **90_ModStub**

---

### 🧱 9. Si quieres entender las entidades del modelo (clases)
Empieza por:
- `dsSimbolo` → La unidad básica del análisis.  
- `dsMiembro` → Propiedades, métodos, eventos.  
- `dsModulo` → Representación de módulos.  
- `dsClase` → Representación de clases.  
- `dsCatalogoInspector` → El “árbol” completo del proyecto.  
- `dsResultadoAnalisis` y `dsResultados` → Resultados del análisis.  
- `dsEstadisticas` → Métricas.  
- `dsInformeSimbolos` → Informes de símbolos no usados.

---

Este mapa te permite navegar el InspectorVBA de forma rápida y eficiente, entendiendo qué partes leer según tu objetivo.

## 🗺️ Mapa de lectura visual – InspectorVBA 2.2

Este diagrama muestra de forma gráfica cómo recorrer el código fuente del InspectorVBA según el área funcional que quieras comprender o extender.

Cada bloque representa un grupo temático, con los módulos y clases más relevantes conectados por orden de lectura recomendado.

![Mapa de lectura visual](sandbox:/mnt/data/graphic_art/InspectorVBA_Modular_Architecture.png)

### 🔹 Categorías incluidas

1. **Funcionamiento general**  
   - modMain  
   - modCore  
   - modMensajes  

2. **Análisis del proyecto**  
   - modAnalisisAux  
   - modBuscarObjetos  
   - modBuscarReferencias  
   - modReglas  
   - modSimbolos  
   - dsCatalogoInspector, dsCatalogoSimbolos, dsSimbolo, dsModulo, dsClase, dsMiembro  

3. **Resultados del análisis**  
   - modCore  
   - modReglas  
   - modSimbolos  
   - dsResultadoAnalisis, dsResultados, dsEstadisticas  

4. **Reparación**  
   - modReparar  
   - modAutoRepair  
   - dsResultadoAnalisis, dsResultados  

5. **Exportación**  
   - mod_ControlRutasExportacion  
   - modExportAux  
   - modExportTXT  
   - modExportExcel  
   - modExportHTML  
   - dsInformeSimbolos, dsResultados, dsCatalogoInspector  

6. **Interfaz (Ribbon y navegación)**  
   - modRibbon  
   - modMenus  
   - modNavegacion  

7. **Entorno y preferencias**  
   - modEntorno  
   - modEntornoInspector  
   - modInicioUsuario  
   - modInicioFin  
   - modPreferencias  

8. **Extensiones y pruebas**  
   - ModStub  

9. **Entidades del modelo**  
   - dsSimbolo, dsMiembro, dsModulo, dsClase, dsCatalogoInspector  
   - dsResultadoAnalisis, dsResultados, dsEstadisticas, dsInformeSimbolos  

---

Este mapa te permite navegar el InspectorVBA de forma rápida y eficiente, entendiendo qué partes leer según tu objetivo técnico o funcional.
