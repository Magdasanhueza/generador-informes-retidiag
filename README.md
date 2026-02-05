# Generador de Informes Retinográficos PDF

Programa en Python que genera informes PDF de fondo de ojo para pacientes diabéticos a partir de un archivo Excel.

## Requisitos

- Python 3.9 o superior
- Librerías: `pandas`, `openpyxl`, `reportlab`

### Instalación de dependencias

```bash
pip3 install pandas openpyxl reportlab
```

## Uso

### Ejecución básica (usa el archivo Excel por defecto)

```bash
python3 /Users/magda/Documents/Retidiag/pdfPatientsCreator/generar_informes.py
```

### Con archivo Excel específico

```bash
python3 /Users/magda/Documents/Retidiag/pdfPatientsCreator/generar_informes.py /ruta/al/archivo.xlsx
```

### Con la plantilla

```bash
python3 /Users/magda/Documents/Retidiag/pdfPatientsCreator/generar_informes.py /Users/magda/Documents/Retidiag/pdfPatientsCreator/Plantilla_Informes_Retidiag.xlsx
```

## Estructura del archivo Excel

El archivo Excel debe tener una hoja llamada `INPUT` con las siguientes columnas:

| Columna | Descripción |
|---------|-------------|
| COMUNA | Comuna del establecimiento |
| FECHA | Fecha del examen |
| ESTABLECIMIENTO | Nombre del centro de salud |
| NOMBRE PACIENTE | Nombre completo del paciente |
| RUT | RUT del paciente |
| EDAD | Edad del paciente |
| EVALUACION TMO | Evaluación del tecnólogo médico |
| OBSERVACIONES | Observaciones adicionales |
| RESULTADO FINAL | Diagnóstico: NORMAL, DG NORMAL, CATARATA, RD, OTROS |
| DETALLE OD | Detalle ojo derecho |
| DETALLE OI | Detalle ojo izquierdo |
| Derivacion | Indicación de derivación |
| OFTALMOLOGO | Nombre del oftalmólogo (para DG NORMAL, RD, OTROS) |

## Estructura de salida

Los PDFs se generan en la siguiente estructura:

```
/Users/magda/Documents/informes retidiag/
└── [COMUNA]/
    └── [Establecimiento]_[Fecha]/
        ├── NORMAL/
        ├── DG_NORMAL/
        ├── CATARATA/
        ├── RD/
        └── OTROS/
```

### Ejemplo:

```
/Users/magda/Documents/informes retidiag/
└── PEÑALOLÉN/
    └── Centro de Salud Familiar Lo Hermida_2026-01-16/
        ├── CATARATA/
        │   ├── ERNESTO ACEVEDO CERDA.pdf
        │   └── ...
        ├── DG_NORMAL/
        │   ├── GISELLA SANCHEZ VILLAGRA.pdf
        │   └── ...
        ├── NORMAL/
        │   ├── ADOLFO DIAZ BALLESTEROS.pdf
        │   └── ...
        └── RD/
            ├── JUAN SOLAR.pdf
            └── ...
```

## Tipos de informes

| Tipo | Firma | Descripción |
|------|-------|-------------|
| NORMAL | Tecnólogo Médico | Sin signos de retinopatía diabética |
| DG NORMAL | Médico Oftalmólogo | Normal, evaluado por oftalmólogo |
| CATARATA | Tecnólogo Médico | Sospecha de cataratas |
| RD | Médico Oftalmólogo | Retinopatía diabética |
| OTROS | Médico Oftalmólogo | Otras alteraciones |

## Imágenes

Las imágenes (logos y firmas) están en la carpeta `imagenes/`:

```
imagenes/
├── logos/
│   ├── logo_retidiag.jpg
│   ├── logo_penalolen.jpg
│   ├── logo_las_condes.png
│   └── logo_el_monte.png
└── firmas/
    ├── firma_tmo_felipe_rojas.jpg
    ├── firma_felipe_contreras.jpg
    ├── firma_yasmine_eltit.png
    └── ... (otras firmas)
```

## Agregar nuevos establecimientos

Para agregar un nuevo establecimiento, editar el diccionario `LOGOS_ESTABLECIMIENTO` en `generar_informes.py`:

```python
LOGOS_ESTABLECIMIENTO = {
    "PEÑALOLÉN": "logo_penalolen.jpg",
    "LAS CONDES": "logo_las_condes.png",
    "EL MONTE": "logo_el_monte.png",
    "NUEVA COMUNA": "logo_nueva_comuna.png",  # Agregar aquí
}
```

## Agregar nuevos oftalmólogos

Para agregar un nuevo oftalmólogo, editar el diccionario `FIRMAS_OFTALMOLOGO`:

```python
FIRMAS_OFTALMOLOGO = {
    "DR. CONTRERAS": "firma_felipe_contreras.jpg",
    "DRA. ELTIT": "firma_yasmine_eltit.png",
    "DR. NUEVO": "firma_nuevo_doctor.png",  # Agregar aquí
}
```

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `generar_informes.py` | Script principal para generar PDFs |
| `Plantilla_Informes_Retidiag.xlsx` | Plantilla para ingresar datos |
| `README.md` | Esta documentación |
| `imagenes/logos/` | Logos de Retidiag y establecimientos |
| `imagenes/firmas/` | Firmas de TMO y oftalmólogos |

## Autor

Retidiag - 2026
