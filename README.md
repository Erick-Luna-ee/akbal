# AKBAL — Aplicación Streamlit para Yacimientos I

Prototipo de consulta y análisis basado en `Excel AKBAL.xlsx`.

## Garantías de diseño

- El selector de etapas es de consulta; no hay formularios para editar el historial.
- Se exige el inventario fijo de etapas `[0, 2, 5, 6, 7, 8]`. Si el libro no coincide, la app se detiene y muestra un error; no corrige ni reordena etapas automáticamente.
- El archivo Excel se incluye como fuente de datos de solo lectura para la aplicación.
- Las ecuaciones trasladadas se identifican en `app.py` y los resultados se calculan en Python.
- El módulo de balance de materia no estima N ni RF hasta definir la ecuación apropiada y validarla.

## Ejecutar localmente

1. Instala Python 3.11 o 3.12 y Visual Studio Code.
2. Abre esta carpeta en VS Code.
3. En la terminal, crea un entorno virtual:

   Windows:
   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

4. Instala dependencias:
   ```powershell
   python -m pip install -r requirements.txt
   ```
5. Ejecuta:
   ```powershell
   streamlit run app.py
   ```
6. Streamlit mostrará la dirección local en la terminal, normalmente `http://localhost:8501`.

## Publicar en Streamlit Community Cloud

1. Crea un repositorio de GitHub y sube `app.py`, `requirements.txt`, `README.md` y `Excel AKBAL.xlsx`.
2. En Streamlit Community Cloud, selecciona **Create app**, el repositorio, la rama y `app.py`.
3. Pulsa **Deploy**. Al finalizar, usa el enlace público que genera la plataforma.

No publiques datos sensibles en el repositorio. Este proyecto contiene datos académicos del libro adjunto.

## Notas técnicas pendientes

- Revisar con el docente o responsable del proyecto la convención/unidad física de `Bg` y su tratamiento cuando el Excel usa cero como marcador no aplicable.
- Definir y validar la ecuación de balance de materia para estimar N y RF.
- Ejecutar comparación celda por celda entre Excel y Python antes de considerar el prototipo validado.
