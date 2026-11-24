# 💼 Procesador y Unificador de Comisiones NL (National Life Group)

Aplicación web minimalista y ejecutiva desarrollada en **Python con Streamlit** diseñada para cargar, unificar y procesar reportes de comisiones en formato Excel (`.xlsx`), generando:
1. **Un archivo PDF unificado** de alta resolución que replica fielmente el diseño, colores, fuentes y estructura del reporte oficial de referencia.
2. **Un bloque de texto preconstruido** listo para copiar con un solo clic, clasificando a los clientes en comisiones del agente principal y de su equipo con sus respectivos estados (`Activa` o `Chargeback`).

---

## 🚀 Inicio Rápido (Ejecución Local)

### 1. Prerrequisitos
- Python 3.9 o superior instalado.
- Gestor de paquetes `pip`.

### 2. Clonar o descargar el repositorio
Navega a la carpeta del proyecto en tu terminal:
```bash
cd c:\Users\User\Documents\Projects\Reader_Comission_NL
```

### 3. (Opcional pero recomendado) Crear un entorno virtual
```bash
# En Windows (PowerShell / CMD)
python -m venv venv
.\venv\Scripts\activate

# En macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

### 4. Instalar las dependencias
```bash
pip install -r requirements.txt
```

### 5. Iniciar la aplicación Streamlit
```bash
streamlit run app.py
```
La aplicación se abrirá automáticamente en tu navegador web en `http://localhost:8501`.

---

## 📁 Estructura Modular del Proyecto

```text
Reader_Comission_NL/
│
├── app.py              # Aplicación web Streamlit (UI/UX, visualización interactiva y descargas)
├── processor.py        # Módulo de lógica de negocio (extracción, unificación, clasificación y estados)
├── pdf_generator.py    # Módulo de generación de PDF con ReportLab (diseño, colores y tipografía idéntica)
├── requirements.txt    # Dependencias exactas del proyecto
├── test_app.py         # Suite de pruebas unitarias automatizadas (unittest)
├── sample_data/        # Archivos de prueba reales y PDF de referencia
│   ├── NLG_PayableGrossCommission_11232026.xlsx
│   ├── NLG_PayableGrossCommission_11172026.xlsx
│   └── NLG_PayableGrossCommission_11172026_Referencia.pdf
└── README.md           # Documentación completa y manual de uso
```

---

## ⚙️ Lógica de Procesamiento de Datos

### 1. Extracción del Agente Principal
- El sistema escanea las celdas del archivo Excel en busca de la etiqueta `Exported By:`.
- Extrae el nombre del agente (por ejemplo: `Gutierrez, Manuel Hernandez` o `AlexanderArellano`).
- **Separación de mayúsculas pegadas (CamelCase / PascalCase):** Si el nombre en el Excel viene sin espacios entre palabras (ej. `AlexanderArellano`, `JuanLopez Bravo` o `Arellano,Alexander`), el sistema inserta automáticamente los espacios correspondientes resultando en `Alexander Arellano`, `Juan Lopez Bravo` o `Arellano, Alexander`.
- **Normalización tokenizada:** Permite comparar sin fallos nombres escritos en formato `Apellidos, Nombres` con registros escritos en formato `NOMBRES APELLIDOS` o con variaciones de tildes y mayúsculas/minúsculas.

### 2. Unificación de Archivos (1 o 2 archivos `.xlsx`)
- Si se carga 1 archivo, procesa y genera las salidas para dicho reporte.
- Si se cargan 2 archivos, combina automáticamente sus registros en una tabla unificada preservando las 10 columnas oficiales:
  1. `Agent #`
  2. `Agent Name`
  3. `Policy #`
  4. `Insured/Annuitant`
  5. `Process Date`
  6. `Premium`
  7. `Gross Com. Paid`
  8. `Com. Rate`
  9. `Part. %`
  10. `Payment Date`
- Limpia filas en blanco, metadatos y filas de totales.

### 3. Clasificación en dos listados
Cada registro se compara contra el Agente Principal detectado:
- **Comision y Chargeback NL:** Registros donde `Agent Name` **coincide** con el nombre del agente principal en `Exported By:`.
- **Comision y Chargeback de Equipo NL:** Registros donde `Agent Name` **no coincide** (pertenecen a miembros del equipo).

### 4. Regla de Duplicados y Determinación de Estado
- Se agrupan los registros por `Insured/Annuitant` dentro de cada listado.
- Se limpian y convierten a valores numéricos los importes de `Gross Com. Paid` (manejando formatos monetarios `$`, comas y números negativos entre paréntesis como `($202.50)`).
- Se calcula la suma neta para cada persona:
  - Si la suma es **mayor a 0** (`> 0`) → `Activa`
  - Si la suma es **menor o igual a 0** (`<= 0`) → `Chargeback`
- El listado del equipo se presenta ordenado alfabéticamente por agente y por cliente, garantizando consistencia y orden ejecutivo.

---

## 📄 Especificaciones de las Salidas

### 1. Archivo PDF Unificado
- **Fidelidad al PDF de referencia:**
  - Encabezado verde corporativo suave (`#C6E0B4`) con texto en negrita centrado.
  - Filas de datos con fondo alternado en azul suave (`#83CCEB`) y azul hielo (`#C0E6F5`).
  - Cuadrícula completa de líneas nítidas (`0.68 pt`).
  - Formato contable estándar para números y monedas: valores negativos representados entre paréntesis `($102.65)` y porcentajes formateados a dos decimales.
  - Pie de página oficial con `Exported On:` y `Exported By:`.
- **Nombre de archivo automático:**
  `NLG_[nombre del agente]_PayableGrossComission.pdf` (limpio de caracteres inválidos de sistema operativo).

### 2. Bloque de Texto Preconstruido
Generado exactamente con el formato requerido, listo para copiar:

```text
Comision y Chargeback NL
* Rafael Aguilar Perez: Activa
* Emilio O Olivera: Activa
* Gabino Ruiz: Chargeback

Comision y Chargeback de Equipo NL
* Sofia Moreno Aguilera: Activa
* Guerreros Anaya: Activa
* Felix Rodriguez: Activa
* Erick Jr Lucero: Chargeback
* Elly Maria: Chargeback
* Ana Sanchez Durazo: Activa
* Katering Gonzalez Guzman: Activa
* Maria Consuelo Ramirez: Activa
* Franki J Leon Calzadilla: Chargeback
* Yulissa Rodriguez Solis: Activa
* Gerardo Martinez Pinon: Activa
* Hugo A Plascencia Lope: Activa
* Francisco Javier Lopez: Activa
```

---

## 🧪 Pruebas Automatizadas

El proyecto incluye un conjunto de pruebas unitarias para garantizar la fiabilidad del procesamiento:
```bash
python test_app.py
```
Las pruebas validan:
- Extracción de datos y metadatos desde archivos Excel reales.
- Normalización y comparación inteligente de nombres.
- Limpieza contable de monedas y valores negativos.
- Agrupación por cliente y asignación de estados (`Activa` / `Chargeback`).
- Generación de PDF con dimensiones y nombres de archivo correctos.
- Unificación de dos archivos simultáneos.
