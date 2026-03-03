# modECUFixFile — Corrector de Escenarios de Caso de Uso

Este módulo VBA permite migrar el contenido de archivos `.docx` de escenarios de caso de uso a una plantilla actualizada, de forma automática y en lote.

---

## ¿Qué hace?

Para cada archivo `.docx` en la carpeta seleccionada, el macro:

1. Abre el archivo original de escenario de caso de uso.
2. Crea una nueva copia de la plantilla seleccionada.
3. Transfiere automáticamente el contenido del original a la plantilla:
   - **Título** del documento.
   - **Encabezado** de la tabla (filas de metadatos del caso de uso).
   - **Secuencia normal** (ajustando la cantidad de filas según el original).
   - **Excepciones** (ídem).
   - **Postcondición** y pie de tabla.
4. Guarda el archivo resultante en una subcarpeta llamada `FIXED FILES`, dentro de la misma carpeta de los archivos originales.

Los archivos originales **no son modificados**.

---

## Requisitos del sistema

| Requisito | Detalle |
|---|---|
| Sistema operativo | Windows |
| Aplicación | Microsoft Word (cualquier versión con soporte de macros VBA) |
| Formato de archivos | `.docx` (plantilla y escenarios) |
| Macros habilitadas | El archivo `.docm` debe ejecutarse con macros habilitadas |

---

## Estructura del repositorio

```
escenario-caso-uso/
├── modECUFixFile.bas   ← Módulo VBA con la lógica del macro
└── ...
```

---

## Configuración inicial (primera vez)

El módulo `.bas` debe importarse en un archivo Word habilitado para macros (`.docm`). Seguir los siguientes pasos:

1. Abrir Microsoft Word y crear un documento nuevo.
2. Guardar el archivo como tipo **Documento de Word habilitado para macros** (`.docm`).
3. Abrir el editor de VBA con el atajo `Alt + F11`.
4. En el panel izquierdo, hacer clic derecho sobre el proyecto del documento → **Importar archivo...**.
5. Seleccionar el archivo `modECUFixFile.bas`.
6. Cerrar el editor de VBA con `Alt + Q`.
7. Guardar el archivo `.docm`.

> Este paso solo es necesario la primera vez. Una vez configurado, el archivo `.docm` puede reutilizarse directamente.

---

## Uso

1. Abrir el archivo `.docm` que contiene el macro.
2. Ir a la pestaña **Programador** → **Macros**, o usar el atajo `Alt + F8`.
3. Seleccionar el macro `main` y hacer clic en **Ejecutar**.
4. Se abrirán dos diálogos en secuencia:

   **Paso 1 — Seleccionar la plantilla**
   > Elegir el archivo `.docx` que se usará como plantilla base. Este archivo define la estructura y el formato al que se migrarán los escenarios.

   **Paso 2 — Seleccionar la carpeta**
   > Elegir la carpeta que contiene los archivos `.docx` de escenarios de caso de uso a corregir. El macro procesará **todos** los `.docx` que encuentre en esa carpeta.

5. El macro se ejecuta automáticamente. Al finalizar, los archivos corregidos estarán disponibles en:

```
<carpeta seleccionada>/FIXED FILES/
```

---

## Notas importantes

- La carpeta `FIXED FILES` se crea automáticamente si no existe.
- Si un archivo de la carpeta no sigue la estructura esperada de la tabla (no contiene las palabras clave `Secuencia normal`, `Excepción` y `Postcondición:`), el macro mostrará un mensaje de error e indicará qué archivo falló.
- La plantilla debe tener al menos las filas base esperadas para secuencia normal (3), excepción (3) y postcondición.
- No se procesan subcarpetas; solo los `.docx` directamente dentro de la carpeta seleccionada.
