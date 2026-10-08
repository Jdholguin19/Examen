# UiPath: especificación técnica de conexiones (Integration Service) y vinculación de XAML

> **Audiencia:** agente de IA que leerá o modificará este proyecto.
> **Alcance:** proyecto `MiniProyect` (Studio 26.0, `schemaVersion 4.1`, `targetFramework Windows`, `expressionLanguage VisualBasic`).
> **Estado del documento:** describe el estado del repositorio al 2026-10-07. Los números de línea son orientativos; verificar siempre con `grep` antes de editar.

---

## 0. Invariantes (leer primero)

1. **Las credenciales nunca están en el repositorio.** Una actividad con conexión solo guarda el GUID de la conexión (`ConnectionId`). El token OAuth vive en Integration Service (Orchestrator) y se resuelve en tiempo de ejecución.
2. **La conexión no se pasa como argumento entre XAML.** Cada actividad con conexión tiene su propio `ConnectionId` literal. Para que dos XAML usen la misma cuenta, ambas actividades deben tener el mismo GUID.
3. **`ConnectionId` y `ConnectionKey` deben ser iguales** dentro de una misma actividad (la actividad y sus sub-objetos `DriveItemArgument` / `MailboxArgument`).
4. **En `InvokeWorkflowFile`, el `x:Key` de cada argumento debe coincidir exactamente** con un `x:Property Name` del XAML destino: mismo nombre, misma dirección (`InArgument` / `OutArgument` / `InOutArgument`) y mismo `x:TypeArguments`.
5. **`WorkflowFileName` es una ruta relativa a la raíz del proyecto.** Si se renombra o mueve un XAML, actualizar todas las referencias.
6. **`.local/` está en `.gitignore`.** `bindings_v2.json` es un artefacto que Studio regenera. No editarlo a mano como fuente de verdad: la fuente de verdad es el XAML.
7. **Según `CLAUDE.md`, las ediciones de `.xaml` deben hacerse con el skill `uipath-rpa`.** El diagnóstico se hace con `uipath-troubleshoot` y la auditoría con `uipath-review`. Este documento es referencia, no sustituye la validación de esos skills.

---

## 1. Inventario

### 1.1 Dependencias relevantes ([project.json](../project.json))

```json
"UiPath.Excel.Activities": "[3.6.1]",
"UiPath.IntegrationService.Activities": "[1.33.0]",
"UiPath.MicrosoftOffice365.Activities": "[3.11.11]",
"UiPath.System.Activities": "[26.8.2]"
```

`UiPath.MicrosoftOffice365.Activities` contiene las actividades `*Connections` usadas. Si se elimina, los tres XAML con conexión dejan de compilar.

### 1.2 Conexiones referenciadas

| ConnectionId | Conector (`metadata.Connector`) | Archivo | Actividad (clase XAML) | `IdRef` | DisplayName |
|---|---|---|---|---|---|
| `4dd5d2ee-b922-4d45-b468-b34b9cdb191a` | `uipath-microsoft-365` | [GetSharePointFiles.xaml](../GetSharePointFiles.xaml) (~L204) | `umaf:GetFileListConnections` | `GetFileListConnections_2` | Get File or Folder List |
| `4dd5d2ee-b922-4d45-b468-b34b9cdb191a` | `uipath-microsoft-365` | [AnalyzeFileVersions.xaml](../AnalyzeFileVersions.xaml) (~L125) | `umag:HttpRequestConnections` | `HttpRequestConnections_1` | Get File Versions (Graph) |
| `4dd5d2ee-b922-4d45-b468-b34b9cdb191a` | `uipath-microsoft-365` | [AnalyzeFileVersions.xaml](../AnalyzeFileVersions.xaml) (~L145) | `umag:HttpRequestConnections` | `HttpRequestConnections_2` | — |
| `d516c359-7e1c-4be0-b596-30c3fd2d53cd` | `uipath-microsoft-outlook365` | [SendReportEmail.xaml](../SendReportEmail.xaml) (~L138) | `umam:SendMailConnections` | `SendMailConnections_2` | Send Email |

Los demás XAML no tienen actividades con conexión: `Main`, `AnalyzeDuplicates`, `AnalyzeLargeOldFiles`, `CategorizeFiles`, `BuildCleanupProposal`, `GenerateExcelReport` y `PublishToPortal`. `PublishToPortal` usa una API key (`in_ApiKey`), no Integration Service.

### 1.3 Discrepancia conocida

`.local/content/bindings_v2.json` solo contiene `HttpRequestConnections_1` y `SendMailConnections_2`. **Faltan `GetFileListConnections_2` y `HttpRequestConnections_2`.** Hipótesis: el archivo no se regeneró después de la última edición. Acción: abrir y guardar el proyecto en Studio (o volver a empaquetar) y comprobar que aparecen las entradas. Si una actividad no tiene binding, Orchestrator no permitirá reasignar su conexión en el proceso publicado.

---

## 2. Modelo de ejecución de una conexión

```
Robot ──(job)──► Activity *Connections
                    │  ConnectionId = GUID (o el valor sobrescrito por el binding)
                    ▼
            Integration Service (Orchestrator, carpeta F)
                    │  valida acceso del robot a la carpeta F
                    │  devuelve un access token vigente (refresh automático)
                    ▼
            API destino (Microsoft Graph / Outlook)
```

- **Resolución del GUID:** el valor efectivo es `binding override (proceso en Orchestrator)` si existe; si no, `ConnectionId` del XAML.
- **Ámbito:** la conexión pertenece a una carpeta de Orchestrator. La identidad del robot que ejecuta debe tener acceso a esa carpeta.
- **Estados de la conexión:** `Enabled`, o `Failed`/`Disconnected` (token revocado o caducado, MFA, cambio de contraseña). La reautenticación se hace en Orchestrator y **no requiere cambios en el XAML** porque el GUID se mantiene.
- **Scopes:** si la operación necesita permisos que la conexión no tiene, Studio serializa `AuthScopesInvalid="True"`. En tiempo de ejecución el resultado esperado es un 401 o 403.

---

## 3. Esquema XAML de actividades con conexión

### 3.1 Namespaces y ensamblados requeridos

Prefijos usados en el proyecto (cabecera `<Activity ...>`):

| Prefijo | CLR namespace (`assembly=UiPath.MicrosoftOffice365.Activities` salvo indicación) | Uso |
|---|---|---|
| `umaf` | `UiPath.MicrosoftOffice365.Activities.Files` | `GetFileListConnections` |
| `umafm` | `UiPath.MicrosoftOffice365.Activities.Files.Models` | `DriveItemArgument` |
| `umafe` | `UiPath.MicrosoftOffice365.Activities.Files.Enums` | `EDriveItemMode` |
| `umag` | `UiPath.MicrosoftOffice365.Activities.General` | `HttpRequestConnections` |
| `umam` | `UiPath.MicrosoftOffice365.Activities.Mail` | `SendMailConnections` |
| `umame` / `umamm` | `...Mail.Enums` / `...Mail.Models` | enums y modelos de correo |
| `umsm` | `UiPath.MicrosoftOffice365.Sharepoint.Models` (`assembly=UiPath.MicrosoftOffice365`) | `Office365SharepointListColumnSlim` |
| `usau` | `UiPath.Shared.Activities.Utils` | `BackupSlot` |
| `upr` | `UiPath.Platform.ResourceHandling` (`assembly=UiPath.Platform`) | `IResource` / `LocalResource` (adjuntos) |

`TextExpression.ReferencesForImplementation` debe incluir `UiPath.MicrosoftOffice365.Activities` y `UiPath.MicrosoftOffice365`. `NamespacesForImplementation` debe incluir los namespaces usados en expresiones VB (por ejemplo `UiPath.MicrosoftOffice365.Enums` para `FindFilesAndFoldersResultType`).

### 3.2 Atributos comunes a toda actividad `*Connections`

| Atributo | Tipo / valor | Semántica |
|---|---|---|
| `ConnectionId` | GUID literal | Identificador de la conexión. **Es el vínculo.** |
| `UseConnectionService` | `"True"` | Modo Integration Service. `False` implica autenticación clásica dentro de un scope de Office 365 (no se usa en este proyecto). |
| `AuthScopesInvalid` | `"True"` / `"False"` | Calculado por Studio en diseño. No ponerlo a `False` para ocultar un problema de permisos. |
| `ConnectionAccountName` | `{x:Null}` | Informativo; se deja nulo. |
| `sap2010:WorkflowViewState.IdRef` | `<Clase>_<n>` | Debe ser único por archivo. Lo usa `bindings_v2.json` como `activityId`. |

### 3.3 `GetFileListConnections` (GetSharePointFiles.xaml)

Atributos no nulos actuales:

```xml
<umaf:GetFileListConnections
  ConnectionId="4dd5d2ee-b922-4d45-b468-b34b9cdb191a"
  UseConnectionService="True" AuthScopesInvalid="False"
  DisplayName="Get File or Folder List"
  IncludeSubfolders="True" MaxResults="100000"
  SearchMode="UseSimple" SortOptions="Ascending" TrimDuplicates="False"
  WhatToReturn="[UiPath.MicrosoftOffice365.Enums.FindFilesAndFoldersResultType.FilesAndFolders]"
  WorkingWithDocumentLibrary="False"
  Result="[resultadoArchivos]"
  ItemIdBackup="{x:Reference __ReferenceID0}"
  AllColumns="{x:Array Type=umsm:Office365SharepointListColumnSlim}"
  sap2010:WorkflowViewState.IdRef="GetFileListConnections_2">
  <umaf:GetFileListConnections.DriveItemArgument>
    <umafm:DriveItemArgument
      ConnectionKey="4dd5d2ee-b922-4d45-b468-b34b9cdb191a"
      ItemSelectionMode="ItemUrl"
      ManualEntryItemUrl="[currentFolderUrl]" ParentFullPathHint="">
      <umafm:DriveItemArgument.ItemIdBackup>
        <usau:BackupSlot x:TypeArguments="umafe:EDriveItemMode" x:Name="__ReferenceID0" StoredValue="ItemUrl">
          <!-- BackupValues: valores del modo "Browse" (driveId, siteUrl, etc.) -->
        </usau:BackupSlot>
      </umafm:DriveItemArgument.ItemIdBackup>
    </umafm:DriveItemArgument>
  </umaf:GetFileListConnections.DriveItemArgument>
  <umaf:GetFileListConnections.FilterSelectionBackup> ... </umaf:GetFileListConnections.FilterSelectionBackup>
</umaf:GetFileListConnections>
```

Notas:

- `ItemSelectionMode="ItemUrl"` con `ManualEntryItemUrl` como expresión. La carpeta se resuelve en tiempo de ejecución a partir de `in_SiteUrl` + `in_LibraryPath` (variable `currentFolderUrl`). Por eso el sitio de SharePoint **no está fijado** en la actividad.
- `BackupSlot` / `__ReferenceID0` guarda los valores de modos alternativos (por ejemplo `Browse`). Contiene IDs de drive y URLs de OneDrive del autor. Son **datos de diseño, no se usan en ejecución** mientras `StoredValue="ItemUrl"`. El `x:Name` y el `x:Reference` deben seguir emparejados.
- El `x:Reference` hacia `__ReferenceID0` va **antes** de la definición `x:Name`. XAML lo admite y no hay que reordenarlo.

### 3.4 `HttpRequestConnections` (AnalyzeFileVersions.xaml)

```xml
<umag:HttpRequestConnections
  ConnectionId="4dd5d2ee-b922-4d45-b468-b34b9cdb191a" UseConnectionService="True"
  ContinueOnError="True" DisplayName="Get File Versions (Graph)"
  RequestMethod="[UiPath.MicrosoftOffice365.Activities.General.Enums.RequestMethod.GET]"
  RequestUrl="[currentUrl]"
  ResponseContent="[responseContent]" ResponseStatus="[responseStatus]"
  Body="{x:Null}" HeadersVariable="{x:Null}" ParametersVariable="{x:Null}" ResponseHeaders="{x:Null}"
  sap2010:WorkflowViewState.IdRef="HttpRequestConnections_1" />
```

- Petición HTTP autenticada contra Microsoft Graph reutilizando el token de la conexión. `RequestUrl` es una URL de Graph construida en un `InvokeCode` previo.
- `ContinueOnError="True"`: los errores HTTP no lanzan excepción. El `InvokeCode` siguiente evalúa `responseStatus` (rango 2xx) y parsea `responseContent` con `Newtonsoft.Json.Linq`.
- Las variables `responseContent`/`responseStatus` se reinician entre iteraciones para no arrastrar la respuesta del archivo anterior.

### 3.5 `SendMailConnections` (SendReportEmail.xaml)

```xml
<umam:SendMailConnections
  ConnectionId="d516c359-7e1c-4be0-b596-30c3fd2d53cd" UseConnectionService="True" AuthScopesInvalid="False"
  DisplayName="Send Email" InputType="HTML" Importance="Normal" SaveAsDraft="False"
  To="[in_Recipients.Split(New Char(){&quot;;&quot;c, &quot;,&quot;c}, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) x.Trim()).ToArray()]"
  Subject="[emailSubject]" Body="[emailBody]"
  AttachmentInputMode="Builder"
  MailboxBackup="{x:Reference __ReferenceID1}"
  sap2010:WorkflowViewState.IdRef="SendMailConnections_2">
  <umam:SendMailConnections.AttachmentList>
    <scg:List x:TypeArguments="InArgument(upr:IResource)" Capacity="4">
      <InArgument x:TypeArguments="upr:IResource">[UiPath.Platform.ResourceHandling.LocalResource.FromPath(in_ReportPath)]</InArgument>
    </scg:List>
  </umam:SendMailConnections.AttachmentList>
  ...
</umam:SendMailConnections>
```

- `To` es de tipo `String[]`. El argumento `in_Recipients` (String) se separa por `;` o `,`.
- Los adjuntos se pasan como `IResource` (`LocalResource.FromPath`), no como rutas de texto, porque `AttachmentInputMode="Builder"`.
- El buzón remitente es la cuenta de la conexión `d516c359-…`.

---

## 4. Bindings (`.local/content/bindings_v2.json`)

Esquema de cada recurso de tipo conexión:

```json
{
  "activityDisplayName": "<DisplayName>",
  "activityId": "<IdRef>",
  "resource": "Connection",
  "key": "<ConnectionId>",
  "id": "Connection<ConnectionId><ConnectionId>False",
  "value": {
    "ConnectionId": {
      "defaultValue": "<ConnectionId>",
      "isExpression": false,
      "displayName": "Connection",
      "description": "The connection to be used"
    }
  },
  "metadata": {
    "BindingsVersion": "2.2",
    "Connector": "<connector-key>",
    "UseConnectionService": "True"
  }
}
```

- Generado por Studio a partir de las actividades con `UseConnectionService="True"`. Las reglas de extracción por tipo de actividad están en `.project/PackageBindingsMetadata.json`.
- Se empaqueta en el `.nupkg`. En Orchestrator, en Process → Requirements/Connections, cada `key` se puede reasignar a otra conexión **del mismo conector**.
- `isExpression: false`: el `ConnectionId` es literal. Si se cambiara a una expresión (variable o asset), el binding pierde la capacidad de reasignarse desde Orchestrator. No hacerlo.

---

## 5. Vinculación entre XAML

### 5.1 Entry point

- `project.json` tiene `"main": "Main.xaml"` y `entryPoints[0].filePath = "Main.xaml"` (`uniqueId 1bd0d046-24e7-4ccc-a418-6cb9ec770132`).
- `entry-points.json` expone el input `in_RunId: string` (opcional).
- `runtimeOptions.excludedLoggedData = ["Private:*", "*password*"]`.

### 5.2 Firmas de los workflows (`x:Property`)

```
Main                  in_RunId:String
GetSharePointFiles    in_SiteUrl:String, in_LibraryPath:String
                      → out_Inventory:DataTable, out_FailedFolders:DataTable
AnalyzeDuplicates     in_FileInventory:DataTable, in_CaseSensitive:Boolean
                      → out_Duplicates:DataTable
AnalyzeLargeOldFiles  in_FileInventory:DataTable, in_LargeFileThresholdBytes:Double, in_OldFileThresholdYears:Int32
                      → out_LargeOldFiles:DataTable
AnalyzeFileVersions   in_FileInventory:DataTable, in_MinSizeBytes:Double, in_MaxFiles:Int32
                      → out_VersionInfo:DataTable, out_VersionDetail:DataTable
CategorizeFiles       in_FileInventory:DataTable, in_CategoryMapping:DataTable
                      → out_EnrichedInventory:DataTable, out_CategorySummary:DataTable, out_DateSummary:DataTable
BuildCleanupProposal  in_Duplicates:DataTable, in_LargeOldFiles:DataTable
                      → out_CleanupProposal:DataTable, out_TotalFilesFlagged:Int32, out_TotalEstimatedRecoverableMB:Double
GenerateExcelReport   in_Inventory, in_Duplicates, in_LargeOldFiles, in_Versions, in_CategorySummary,
                      in_DateSummary, in_CleanupProposal, in_ResumenEjecutivo, in_FailedFolders : DataTable
                      in_OutputFolderPath:String
                      → out_ReportFilePath:String
PublishToPortal       in_RunId, in_PortalBaseUrl, in_ApiKey, in_SiteUrl, in_LibraryPath, in_ReportPath : String
                      in_FilesScanned:Int32, in_VersionInfo:DataTable, in_VersionDetail:DataTable
                      → out_PortalUrl:String
SendReportEmail       in_Recipients, in_ReportPath, in_SiteUrl, in_LibraryPath, in_PortalUrl : String
                      in_TopN:Int32, in_VersionInfo:DataTable, in_FailedFolders:DataTable
```

Tipos en XAML: `x:String`, `x:Int32`, `x:Double`, `x:Boolean`, `sd:DataTable` (`xmlns:sd="clr-namespace:System.Data;assembly=System.Data.Common"`).

### 5.3 Secuencia de Main.xaml y mapeo variable → argumento

| # | DisplayName | Tipo | Mapeo (variable de Main ⇄ argumento del hijo) | Conexión |
|---|---|---|---|---|
| 1 | Registrar hora de inicio | Assign/InvokeCode | → `startTimeText` | — |
| 2 | Leer Config - Settings | `ui:ReadRange` `Config.xlsx`!`Settings` `A4:C40` | → `dtSettings` | — |
| 3 | Leer Config - CategoryMapping | `ui:ReadRange` `Config.xlsx`!`CategoryMapping` `A4:B300` | → `dtCategoryMapping` | — |
| 4 | Interpretar configuracion | `ui:InvokeCode` | `dtSettings`, `in_RunId` → `siteUrl, libraryPath, outputFolder, largeThresholdBytes, oldYears, versionMinBytes, versionMaxFiles, emailRecipients, sendEmail, emailTopN, caseSensitive, portalBaseUrl, portalApiKey, publishToPortal, runId, portalUrl` | — |
| 5 | Invoke GetSharePointFiles | InvokeWorkflowFile | `siteUrl→in_SiteUrl`, `libraryPath→in_LibraryPath`, `out_Inventory→dtInventory`, `out_FailedFolders→dtFailedFolders` | M365 |
| 6 | Invoke AnalyzeDuplicates | InvokeWorkflowFile | `dtInventory→in_FileInventory`, `caseSensitive→in_CaseSensitive`, `out_Duplicates→dtDuplicates` | — |
| 7 | Invoke AnalyzeLargeOldFiles | InvokeWorkflowFile | `dtInventory`, `largeThresholdBytes`, `oldYears` → `dtLargeOld` | — |
| 8 | Invoke AnalyzeFileVersions | InvokeWorkflowFile | `dtInventory`, `versionMinBytes`, `versionMaxFiles` → `dtVersions`, `dtVersionDetail` | M365 |
| 9 | Invoke CategorizeFiles | InvokeWorkflowFile | `dtInventory`, `dtCategoryMapping` → `dtEnriched`, `dtCategorySummary`, `dtDateSummary` | — |
| 10 | Invoke BuildCleanupProposal | InvokeWorkflowFile | `dtDuplicates`, `dtLargeOld` → `dtCleanupProposal`, `totalFlagged`, `totalRecoverableMB` | — |
| 11 | Construir resumen ejecutivo | `ui:InvokeCode` | varias tablas y totales → `dtResumen` | — |
| 12 | Invoke GenerateExcelReport | InvokeWorkflowFile | `dtEnriched→in_Inventory`, …, `dtResumen→in_ResumenEjecutivo`, `outputFolder→in_OutputFolderPath`, `out_ReportFilePath→reportPath` | — |
| 13 | If `[publishToPortal]` → Invoke PublishToPortal | InvokeWorkflowFile | `runId, portalBaseUrl, portalApiKey, siteUrl, libraryPath, reportPath, dtInventory.Rows.Count, dtVersions, dtVersionDetail` → `portalUrl` | — (API key) |
| 14 | If `[sendEmail]` → Invoke SendReportEmail | InvokeWorkflowFile | `emailRecipients, dtVersions, reportPath, siteUrl, libraryPath, emailTopN, portalUrl, dtFailedFolders` | Outlook |
| 15 | Log fin | LogMessage | — | — |

Dependencias de orden: 5 → {6, 7, 8, 9} → 10 → 11 → 12 → 13 → 14. El paso 14 consume `portalUrl`, que produce el paso 13. Si `publishToPortal=False`, `portalUrl` conserva el valor del paso 4.

### 5.4 Forma canónica de `InvokeWorkflowFile`

```xml
<ui:InvokeWorkflowFile ContinueOnError="{x:Null}" DisplayName="Invoke <Nombre>" UnSafe="False"
                       WorkflowFileName="<Nombre>.xaml"
                       sap2010:WorkflowViewState.IdRef="InvokeWorkflowFile_<n>">
  <ui:InvokeWorkflowFile.Arguments>
    <InArgument  x:TypeArguments="<tipo>" x:Key="in_<X>">[<expresión VB>]</InArgument>
    <OutArgument x:TypeArguments="<tipo>" x:Key="out_<Y>">[<variable l-value>]</OutArgument>
  </ui:InvokeWorkflowFile.Arguments>
</ui:InvokeWorkflowFile>
```

Reglas de validación:

- `OutArgument` / `InOutArgument` requieren un l-value (una variable), no una expresión.
- Un `x:Key` sin `x:Property` correspondiente en el destino produce un error de validación en Studio.
- Un `x:Property` del destino que no está en `Arguments` recibe su valor por defecto (`Nothing` en tipos por referencia).
- `UnSafe="False"` hace que el XAML se ejecute en el mismo proceso. Las referencias a `DataTable` se comparten (mutar `in_FileInventory` en el hijo muta `dtInventory` en Main).
- Toda variable referenciada debe estar declarada en `Sequence.Variables` de un ámbito que contenga el Invoke.

---

## 6. Procedimientos

### 6.1 Cambiar la conexión de una actividad (mismo conector)

1. Obtener el GUID nuevo (Orchestrator → Integration Service → Connections, o `uipath-platform`).
2. En el XAML, reemplazar **todas** las apariciones del GUID antiguo dentro de la actividad (`ConnectionId` y cualquier `ConnectionKey` hijo).
3. No tocar `BackupSlot.BackupValues` salvo que se cambie `ItemSelectionMode`.
4. Regenerar los bindings (abrir/guardar en Studio o empaquetar) y comprobar el `key` en `bindings_v2.json`.

> Alternativa preferida para cambios por entorno: **no editar el XAML** y reasignar el binding en el proceso publicado.

### 6.2 Añadir una actividad con conexión en un XAML nuevo o existente

1. Asegurar que la dependencia del paquete está en `project.json`.
2. Declarar en la cabecera los prefijos `xmlns` necesarios (§3.1) y añadir los ensamblados/namespaces a `TextExpression.*ForImplementation`.
3. Serializar con `ConnectionId`, `UseConnectionService="True"` e `IdRef` único.
4. Encapsular la integración en su propio XAML con argumentos `in_*`/`out_*`. Main solo debe mapear datos.

### 6.3 Añadir un sub-workflow e invocarlo

1. Crear `<Nombre>.xaml` con `x:Class="<Nombre>"` y `x:Members` / `x:Property` por cada argumento.
2. En Main, insertar `InvokeWorkflowFile` (§5.4) en el punto correcto del orden de dependencias (§5.3).
3. Declarar en Main las variables de salida.
4. Validar con `uipath-rpa` (o *Validate* / *Analyze* en Studio).

### 6.4 Renombrar un argumento o un XAML

- Argumento: actualizar el `x:Property Name`, todos los usos internos y **cada** `x:Key` en los Invoke que lo llaman.
- XAML: actualizar `x:Class`, el nombre de archivo y todos los `WorkflowFileName`. Si es entry point, actualizar también `project.json` y `entry-points.json`.

---

## 7. Comandos de verificación (Git Bash, desde la raíz del proyecto)

```bash
# Actividades con conexión y su GUID
grep -n -o -E 'ConnectionId="[^"]*"|ConnectionKey="[^"]*"|UseConnectionService="[^"]*"|AuthScopesInvalid="True"' *.xaml

# GUIDs distintos en uso
grep -h -o -E 'Connection(Id|Key)="[0-9a-f-]{36}"' *.xaml | sort | uniq -c

# Actividades con conexión que faltan en bindings
for id in $(grep -h -o -E 'IdRef="[A-Za-z]+Connections_[0-9]+"' *.xaml | cut -d'"' -f2); do
  grep -q "\"activityId\": \"$id\"" .local/content/bindings_v2.json || echo "SIN BINDING: $id"
done

# Firma de cada XAML
grep -o -E 'x:Property Name="[^"]*" Type="[^"]*"' <archivo>.xaml

# Claves pasadas por los Invoke de Main
grep -n -E 'WorkflowFileName=|x:Key="(in|out|io)_' Main.xaml
```

Comprobación de coherencia entre Invoke y firma: para cada `WorkflowFileName=X`, el conjunto de `x:Key` con prefijo `in_`/`out_`/`io_` debe estar incluido en el conjunto de `x:Property Name` de `X`, con la misma dirección y tipo.

---

## 8. Diagnóstico

| Síntoma | Causa | Verificación / acción |
|---|---|---|
| `Connection not found` / actividad inválida en Studio | El GUID no existe en las carpetas visibles para el usuario o robot | Listar conexiones de la carpeta. Reasignar el binding o corregir `ConnectionId` |
| HTTP 401 / token expirado | Conexión `Disconnected` | Reconectar en Integration Service. El XAML no cambia |
| HTTP 403 en Graph | La cuenta no tiene permisos sobre el sitio o la biblioteca, o faltan scopes | Revisar `AuthScopesInvalid` y los permisos en SharePoint |
| `responseStatus` ≠ 2xx sin excepción en AnalyzeFileVersions | Comportamiento esperado por `ContinueOnError="True"` | Revisar logs del `InvokeCode` posterior |
| Funciona en Studio, falla en Orchestrator | El binding del proceso apunta a otra conexión (o ninguna), o el robot no tiene acceso a la carpeta | Revisar Process → Connections y los permisos de la carpeta |
| Argumento `Nothing` en el hijo | El `x:Key` no coincide con el `x:Property` | Comando de coherencia del §7 |
| Error de compilación por un tipo no encontrado (`umaf:*`) | Falta `xmlns` o falta el ensamblado en `ReferencesForImplementation` | §3.1 |
| El email sale de una cuenta inesperada | El remitente es la cuenta de la conexión `d516c359-…` | Reasignar la conexión de Outlook |
