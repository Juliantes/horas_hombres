# Horas Hombre vs OT

Aplicación Streamlit para importar confirmaciones SAP, registrar horas disponibles por grupo técnico y comparar cumplimiento (gráficos, Excel maestro, exportación PDF).

---

## Requisitos generales

| Contexto | Requisitos |
|----------|------------|
| **Desarrollo (Mac/Linux)** | Python 3.12+, navegador |
| **Windows en planta** | Windows 10/11 **64 bits** (Intel o AMD: el paquete es `amd64` = cualquier PC 64-bit) |
| **Permisos** | No hace falta administrador si la app vive en carpeta del usuario (escritorio, USB, `Documentos`) |
| **Red** | No hace falta Internet **para usar** la app; sí puede hacer falta **una vez** para armar el entorno (opción liviana) |

Los datos se guardan en la carpeta `data\` (`horas_hombre.db`, `confirmaciones.xlsx`).

---

## 1. Desarrollo en Mac (o Linux)

### Instalación

```bash
cd "Horas Hombre"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Ejecutar

```bash
streamlit run horas_hombre.py
```

La base y el Excel maestro quedan en la **raíz del proyecto** (misma carpeta que `horas_hombre.py`), salvo que definas `HH_DATA_DIR`.

### Tests

```bash
PYTHONPATH=. .venv/bin/pytest tests/ -q
```

---

## 2. Uso en Windows — elegir una opción

Hay **tres** caminos. Elija según Internet en planta y cuánto puede copiar por USB.

| Opción | Qué llevas | Internet en planta | Tamaño aprox. al transportar |
|--------|------------|--------------------|------------------------------|
| **A — Código + build** | `HorasHombre-codigo-win.zip` | Sí, **una vez** (`build.ps1`) | Pocos KB |
| **B — Portable embebido** | `HorasHombre-portable-win64.zip` | No | ~150 MB |
| **C — Carpeta ya armada en Windows** | Carpeta completa con `venv\` | No (si ya corriste build en esa PC u otra igual) | ~150–250 MB |

### Arranque en Windows (todas las opciones)

1. Descomprimir o copiar la carpeta a un lugar **escribible** (no `C:\Program Files`).
2. Doble clic en **`HorasHombre.bat`**.
3. No cerrar la ventana negra mientras usás la app.
4. URL: **http://127.0.0.1:8501** (el `.bat` intenta abrir Firefox; si no está, abrir el navegador a mano).

---

## Opción A — Paquete liviano (recomendado con Internet una vez en planta)

### Qué llevar desde la Mac

Archivo generado:

```text
dist/HorasHombre-codigo-win.zip
```

Para regenerarlo:

```bash
./scripts/pack_codigo_win.sh
```

Contenido del zip:

- `HorasHombre.bat`
- `build.ps1`
- `horas_hombre.py`
- `sap_import.py`
- `requirements-portable.txt`
- `PAQUETE-WINDOWS.txt`

**No llevar:** `.venv` de Mac, `dist/HorasHombre` completo, carpeta `build/`, `tests/`, ni el zip `build/python-3.12.8-embed-amd64.zip` (eso es solo caché del build en Mac).

### Instalación en Windows

1. Descomprimir en una carpeta (ej. `C:\Users\tu_usuario\HorasHombre`).
2. Abrir PowerShell en esa carpeta y ejecutar:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build.ps1
   ```

   También: clic derecho en `build.ps1` → ejecutar con PowerShell (si la política de la empresa lo permite).

3. Necesitás **Python 3.12+** en PATH (`python` o `py -3.12`) solo para este paso. No requiere admin en la mayoría de entornos si instalás Python “solo para mi usuario”.

4. Cuando termine (~5–15 minutos), doble clic en **`HorasHombre.bat`**.

### Copiar a otra PC sin Internet

Después del `build.ps1`, copiá la carpeta **entera** incluyendo `venv\`. Si en otra máquina no arranca, volvé a ejecutar `build.ps1` allí.

---

## Opción B — Portable completo (sin Python ni build en planta)

Para plantas **sin Internet** y sin Python instalado.

### Generar el portable en la Mac

```bash
.venv/bin/python scripts/build_portable_win.py
```

Salida: carpeta `dist/HorasHombre/` con `python\` embebido, `app\`, `HorasHombre.bat`, `LEEME.txt`, `data\`.

### Crear el zip para llevar

```bash
cd dist
zip -r HorasHombre-portable-win64.zip HorasHombre
```

### Instalación en Windows

1. Copiar **`HorasHombre-portable-win64.zip`** (~150 MB).
2. Extraer todo (no abrir el zip como carpeta).
3. Entrar a la carpeta **`HorasHombre`** (debe verse `HorasHombre.bat` junto a `python\`, `app\`, `data\`).
4. Doble clic en **`HorasHombre.bat`**.

**Importante:** la carpeta `build/` del proyecto Mac **no** es el instalable. Ahí solo está la descarga intermedia de Python (`python-3.12.8-embed-amd64.zip`).

---

## Opción C — Misma máquina / entorno ya armado con `build.ps1`

Si ya ejecutaste `build.ps1` en una PC Windows, el paquete de trabajo es la carpeta del proyecto con:

```text
HorasHombre.bat
horas_hombre.py
sap_import.py
venv/
data/          (opcional; se crea al usar)
```

Llevá esa carpeta completa por USB. Arranque: `HorasHombre.bat`.

---

## Datos entre Mac, Windows y otras PCs

| Archivo | Ubicación en Windows (modo `.bat` raíz) | Ubicación portable Mac (`dist`) |
|---------|----------------------------------------|-----------------------------------|
| Base SQLite | `data\horas_hombre.db` | `data\horas_hombre.db` |
| Excel maestro | `data\confirmaciones.xlsx` | `data\confirmaciones.xlsx` |

En desarrollo Mac sin `HH_DATA_DIR`, los archivos están en la raíz del repo. Para llevar datos a Windows, copialos a `data\` antes o después del primer uso.

Variables de entorno (las define `HorasHombre.bat`):

- `HH_DATA_DIR` — carpeta de la base y maestro
- `MPLCONFIGDIR` — caché de matplotlib en `data\.matplotlib`

---

## Estructura del repositorio (referencia)

```text
horas_hombre.py          # aplicación Streamlit
sap_import.py            # importación SAP
HorasHombre.bat          # lanzador Windows
build.ps1                # crea venv en Windows
requirements.txt         # desarrollo (+ pytest)
requirements-portable.txt
scripts/
  build_portable_win.py  # arma dist/HorasHombre en Mac
  pack_codigo_win.sh     # arma dist/HorasHombre-codigo-win.zip
dist/
  HorasHombre/                    # portable embebido
  HorasHombre-portable-win64.zip
  HorasHombre-codigo-win.zip
build/                   # caché; NO copiar a Windows
tests/
```

---

## Problemas frecuentes (Windows)

**SmartScreen bloquea el `.bat`**

Propiedades del archivo → Desbloquear, o “Más información” → Ejecutar de todas formas.

**Antivirus y `python.exe`**

Pedir excepción para la carpeta completa `HorasHombre`, no solo el `.bat`.

**No encuentra Python al ejecutar el `.bat`**

- Opción A: falta correr `build.ps1` o el `venv\` no se copió.
- Opción B: descomprimiste mal (falta la carpeta `python\`).

**Puerto 8501 ocupado**

Cerrar otra instancia de Horas Hombre o de otra app Streamlit en el mismo puerto.

**“No está la carpeta horas_hombre”**

No existe esa carpeta. El paquete se llama **`HorasHombre`**; la app es el archivo `horas_hombre.py`.

**Firefox no abre**

Ir manualmente a http://127.0.0.1:8501 (Edge/Chrome también sirven).

---

## Cierre de la aplicación

Cerrar la pestaña del navegador y luego la ventana de consola, o `Ctrl+C` en la consola.

---

## Resumen rápido

| Objetivo | Acción |
|----------|--------|
| Programar en Mac | `venv` + `streamlit run horas_hombre.py` |
| Llevar poco por USB y hay Internet en planta | `HorasHombre-codigo-win.zip` + `build.ps1` + `HorasHombre.bat` |
| Plant a sin red ni Python | `HorasHombre-portable-win64.zip` generado en Mac |
| Regenerar portable en Mac | `python scripts/build_portable_win.py` |
| Regenerar zip liviano en Mac | `./scripts/pack_codigo_win.sh` |