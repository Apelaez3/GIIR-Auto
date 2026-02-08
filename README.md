# SAP GR/IR F.13 Automation (Excel VBA)

Macro de Excel que automatiza el proceso F.13 en SAP para el mayor `20206100` (GR/IR) y registra el resultado en una columna con la fecha de hoy.

## Qué hace
- Busca la hoja `Clearing GRIR 20206100`.
- Encuentra la fila de encabezados donde exista la columna `Company`.
- Crea (si no existe) una columna con el encabezado de hoy en formato `mm.dd.yyyy`.
- Por cada compañía sin resultado en la columna de hoy:
  - Abre `F.13` en SAP.
  - Ejecuta el proceso para el año fiscal actual y la cuenta `20206100`.
  - Guarda el texto de la barra de estado de SAP en la columna de hoy.

## Requisitos
- Excel con macros habilitadas.
- SAP GUI con scripting habilitado.
- Sesión activa de SAP (log in antes de ejecutar la macro).
- Hoja con:
  - Nombre: `Clearing GRIR 20206100`
  - Columna de encabezado: `Company`
  - Lista de compañías debajo de ese encabezado.

## Uso
1. Abrir el archivo de Excel que contiene la hoja `Clearing GRIR 20206100`.
2. Verificar que la columna `Company` tenga las compañías a procesar.
3. Iniciar sesión en SAP GUI.
4. Ejecutar la macro `Run_GRIR_F13_ToTodayColumn`.

## Configuración
Las constantes están al inicio del módulo:
```
Private Const SHEET_NAME As String = "Clearing GRIR 20206100"
Private Const HEADER_TEXT_COMPANY As String = "Company"
Private Const GL_ACCOUNT As String = "20206100"
```
Si cambias el nombre de hoja, el texto del encabezado o la cuenta GL, actualiza esos valores.

## Archivos
- `sapauto.txt`: código VBA de la macro.

## Notas
- La macro evita reprocesar filas que ya tengan un valor en la columna de hoy.
- Si SAP devuelve un mensaje de confirmación de “production run”, la macro lo confirma automáticamente.
- Los errores se registran en la celda correspondiente con el formato `ERROR: <número> - <descripción>`.
