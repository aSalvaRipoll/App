🌌 KOSMOS 2.1 — Arquitectura General del Motor Fonético Multilingüe
Documento de arquitectura técnica y editorial
1. Visión general del sistema
KOSMOS 2.1 es un motor fonético multilingüe diseñado para:

tokenizar texto en grafemas (monógrafos, dígrafos, trigrafemas)

silabear según reglas específicas de cada idioma

detectar la sílaba tónica

normalizar vocales según el sistema fonético de cada lengua

asignar fonemas universales mediante un conjunto de reglas por idioma

generar una secuencia ordenada de objetos fonéticos (clsFonema)

El sistema está construido con una arquitectura modular, extensible y editorialmente impecable, donde cada idioma es autónomo y el dispatcher central coordina el flujo.

2. Mapa de módulos
Código
KOSMOS 2.1
│
├── modMotorFonetico_V2_1_Idiomas.bas   ← Dispatcher principal
│
├── modMotor_Idioma_ES.bas              ← Castellano
├── modMotor_Idioma_CA.bas              ← Catalán central
├── modMotor_Idioma_CA_VA.bas           ← Valenciano
├── modMotor_Idioma_CA_IB.bas           ← Mallorquín
├── modMotor_Idioma_EU.bas              ← Euskera
├── modMotor_Idioma_GL.bas              ← Gallego
├── modMotor_Idioma_PT_EU.bas           ← Portugués europeo
├── modMotor_Idioma_PT_BR.bas           ← Portugués brasileño
├── modMotor_Idioma_EN_GB.bas           ← Inglés británico
├── modMotor_Idioma_FR.bas              ← Francés
│
├── modMotor_Idioma_AUX_CA.bas          ← Auxiliares catalán
├── modMotor_Idioma_AUX_PT.bas          ← Auxiliares portugués
│
└── clsFonema.cls                       ← Objeto fonético universal
3. Flujo de ejecución (alto nivel)
Código
Entrada: NombreOriginal + AbreviadoIdioma
↓
1. Normalización previa (espacios, mayúsculas)
↓
2. Silabeo por idioma (motor puro + revisión opcional)
↓
3. Detección de sílaba tónica por idioma
↓
4. Normalización vocálica por idioma
↓
5. Tokenización grafema por grafema
    - trigrafema
    - dígrafo
    - monógrafo
↓
6. Inserción del acento universal (fonema 61)
↓
7. Aplicación de reglas fonéticas por idioma
↓
8. Construcción de objetos clsFonema
↓
Salida: Collection de clsFonema (grafema + idFonema + orden)
4. Componentes principales
4.1 Dispatcher: MF21_ConvertirNombreAParGrafemaIDFonema
Es el núcleo del sistema.
Coordina:

normalización

marcado de tónica

normalización vocálica

tokenización

aplicación de reglas fonéticas

creación de objetos fonéticos

Es totalmente agnóstico al idioma: delega todo en módulos especializados.

4.2 Motores por idioma
Cada idioma implementa:

✔ Silabear (motor puro)
Reglas propias:

VV

VCV

CCV

grupos inseparables

hiatos

diptongos

ela geminada (CA)

nasales (PT/FR)

etc.

✔ Silabear con revisión
Interfaz universal:

convierte colección → string con “-”

abre formulario

valida

reconstruye colección

✔ Marcar tónica
Cada idioma define su propia prosodia:

ES → reglas ortográficas

CA → penúltima si no hay tilde

CA‑VA → reglas valencianas

CA‑IB → penúltima (mallorquín)

EU → penúltima fija

GL → paroxítona salvo excepciones

PT‑EU → reglas portuguesas

PT‑BR → paroxítona por defecto

FR → última

EN‑GB → heurísticas

✔ Normalización vocálica
Cada idioma define:

qué vocales se preservan

cuáles se convierten

cómo se representan hiatos

cómo se marcan tensiones

✔ Reglas fonéticas
Cada idioma asigna:

dígrafos

trigrafemas

vocales

consonantes

nasales

diptongos

casos especiales

4.3 Auxiliares por idioma
Catalán
EsVocal_CA

EsConsonant_CA

EsGrupInseparable_CA

EsDiptong_CA

EsHiat_CA

Portugués
EsVocal_PT

EsConsonant_PT

EsGrupInseparable_PT

4.4 Objeto fonético universal: clsFonema
Cada fonema contiene:

grafema original

idFonema

orden

esVocal

valor fonético

5. Tokenizador universal: SiguienteGrafema
El tokenizador es multilingüe y jerárquico:

trigrafemas

dígrafos

monógrafos

Incluye:

nasales

ela geminada

dígrafos universales

dígrafos catalanes

dígrafos portugueses

dígrafos franceses

diptongos acentuados

normalización Unicode → precompuestos

Es uno de los componentes más potentes del sistema.

6. Acento universal (fonema 61)
El sistema inserta un fonema especial:

Código
Grafema: "#"
idFonema: 61
Justo antes de la sílaba tónica.
Esto permite:

análisis prosódico

síntesis

alineación fonética

exportación a otros sistemas

7. Extensibilidad
Añadir un idioma nuevo requiere:

MF_MarcarTonica_<IDIOMA>

Silabear_<IDIOMA>

Silabear_<IDIOMA>_ConRevision

Reglas<IDIOMA>

MF_NormalizarVocales_<IDIOMA>

Auxiliares opcionales

Añadir al dispatcher

La arquitectura está diseñada para crecer sin romper nada.

8. Resumen ejecutivo
KOSMOS 2.1 es un motor fonético:

multilingüe

modular

editorialmente impecable

extensible

robusto ante Unicode

preciso en prosodia

coherente en tokenización

perfectamente organizado por idioma