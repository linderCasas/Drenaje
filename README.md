# Proyecto Drenaje

Proyecto para construir un proceso ETL sobre los datos del proyecto de drenaje y
publicar sus resultados en reportes web.

## Estructura

```text
data/
  raw/                 Archivos fuente originales
  processed/           Datos limpios listos para consumo
reports/               Reportes exportados
src/proyecto_drenaje/
  etl/
    extract/           Lectura de fuentes
    transform/         Limpieza y reglas de negocio
    load/              Persistencia de resultados
  web/                 Aplicación del reporte web
tests/                 Pruebas automatizadas
```

El archivo `Maquinaria.xlsx` y el entorno virtual `.venv` existentes se
mantienen en la raíz. La aplicación web de maquinaria lee el Excel en cada
solicitud y presenta el corte financiero y la disponibilidad operativa.

Para iniciar el sitio en Windows:

```powershell
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe -m proyecto_drenaje.web.app
```

Abre <http://127.0.0.1:5000> en el navegador. Si cambian fórmulas en Excel,
guarda el libro recalculado para que la aplicación pueda leer sus valores
actualizados.

## Entorno

La carpeta ya cuenta con un entorno virtual en `.venv`, configurado con Python
3.13.15. En PowerShell, se puede activar con:

```powershell
.\.venv\Scripts\Activate.ps1
```

Las dependencias de la aplicación web y del lector de Excel se declaran en
`pyproject.toml`.

## Próximas etapas

1. Inspeccionar y especificar los datos de entrada.
2. Implementar extracción, transformación y validaciones.
3. Persistir datos procesados en un formato adecuado para el reporte.
4. Completar el proceso ETL y ampliar las validaciones del reporte web.
