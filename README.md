# ETL GCA (Excel + PDF) — Tkinter

Proyecto modularizado para extraer datos desde Excel/PDF, transformarlos y cargarlos a Excel o PostgreSQL,
con interfaz gráfica en Tkinter.

## Estructura
```
etl_gca/
├─ main.py
├─ requirements.txt
├─ README.md
├─ config/
│  └─ settings.py
├─ etl/
│  ├─ __init__.py
│  ├─ utils.py
│  ├─ extract.py
│  ├─ transform.py
│  └─ load.py
├─ ui/
│  ├─ __init__.py
│  ├─ inicio.py
│  └─ principal.py
└─ data/
   ├─ entrada/
   └─ salida/
```

## Requisitos
- Python 3.9+
- Windows/Linux/Mac (Tkinter viene incluido con la mayoría de instalaciones de Python en Windows/Mac).
- Para PostgreSQL: servidor accesible y credenciales válidas.

Instala dependencias (excepto Tkinter que ya viene con Python):
```bash
pip install -r requirements.txt
```

## Ejecutar
```bash
python main.py
```

## Notas
- El guardado a Excel se hace por defecto en `data/salida/Datos_GCA_Limpios.xlsx` (se fusiona si existe).
- Ajusta `config/settings.py` para la conexión a PostgreSQL.
