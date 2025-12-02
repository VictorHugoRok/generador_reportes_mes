# Revisión de Prima Pagada — Documentación

La revisión de prima pagada es una actividad que requiere múltiples pasos para verificar que la póliza ha sido pagada en el mes en curso y no ha sido cancelada. Validar que la póliza está pagada requiere la descarga de reportes generados en la página de la compañía **AXA**, ya sea por número de agente o por centro de costo.  
En este caso se utilizan los reportes mensuales por centro de costo; por lo que el archivo que se procesa para la comparación es el del reporte mensual por cada uno de los centros de costo en **Grupo Ponce A.S.C.**

El presente documento describe el uso y funcionamiento de los códigos en **Visual Basic for Applications (VBA)** para la revisión de primas pagadas con reportes mensuales generados en la página de AXA. Además, incluye los códigos que generan un reporte mensual usando estos mismos reportes.

Para entender y mantener el uso de las macros que utilizan código VBA, se describen las secciones correspondientes, permitiendo su adaptación ante cambios o actualizaciones futuras.

---

## Abreviaturas

- **VBA:** Visual Basic for Applications  
- **AXA:** Compañía de seguros AXA  
- **GPA:** Grupo Ponce y Asociados  

---

## Centros de costo

La empresa **GPA** cuenta con **5 centros de costo**, los cuales agrupan a los diferentes agentes. Para el uso de las macros aplicadas en el *Reporte de Pólizas*, deben descargarse los reportes mensuales de cada centro de costo.

- Para saber cómo descargar cada archivo, consulte a un colaborador del área.  
- Para utilizar las macros, deben tenerse los archivos de los 5 centros de costo, para incluir a todos los agentes.

---

## Formato del archivo de AXA

El archivo descargado de AXA es mensual, y de él se obtiene la prima pagada correspondiente al mes específico y al mes en curso.  
La fecha de corte abarca desde el inicio del mes hasta su fin, o hasta el día en que se solicite.

El archivo tiene dos características importantes:

1. **La estructura en que se descarga**  
2. **El formato europeo de cantidades (importante para el ranking)**

### Estructura

Las pólizas se encuentran en una columna específica. Actualmente es la **columna E**, por lo que en VBA se trabaja sobre esta ubicación.

En los módulos:

- `REVISAR_PRIMA`
- `REVISAR_PRIMA_CNMENSUALIDADES`
- `REVISAR_PRIMA_VIDA`
- `REVISAR_PRIMA_VIDA_MENSUALES`

Existe una constante llamada **`COL_POLIZA_PAG_EXT`**, que indica el número de columna donde se encuentran las pólizas.  
Esta constante puede modificarse si la estructura del reporte cambia.

---

### Formato europeo (aplica para ranking)

Los reportes mensuales de AXA usan:

- Coma (`,`) como separador decimal  
- Punto o espacio como separador de miles  

Ejemplos:

- `€300,10`  
- `100.100,61`

Para convertir estos valores al formato americano, el formulario `.frm` del proyecto incluye la función:

```vb
Function ConvertirFormatoEuropeo(texto As String) As String 
```

Esta función asegura que los cálculos sean correctos durante la generación del ranking.

Hojas asignadas

Para que las macros funcionen, debe estar correctamente definido el nombre de la hoja de cálculo usada en el proceso.

En la línea 62 de los códigos VBA se define el nombre de la hoja:
```vb
Set wsRegistro = wb.Sheets("Nombre_Hoja_De_Calculo")
' Ingresa el nombre de la hoja del reporte que se usa
```

Si el nombre no coincide, la macro mostrará un error indicando que la hoja no se encontró.

Sufijos de pólizas

Los sufijos de póliza pueden variar cada año, por lo que las macros permiten ingresarlos manualmente.

Para ingresar varios sufijos, se separan por comas:
V00,U00
(sin espacios)

La macro procesa los sufijos con:
```vb
arrSufijos = Split(Replace(sufijosIngresados, " ", ""), ",")
```

El prefijo de póliza está definido por defecto como 1, usando la expresión regular:

"1*sufijo"


Para cambiar el prefijo, se puede modificar en el código:
```vb
patronLike = "1*" & sufijo
```
## Enviar actualización

La función envía una actualización de los reportes de **GMM** y **VIDA**, creando un nuevo libro donde solo se incluyen las hojas que contienen dichos reportes.  
El archivo generado se guarda en el **mismo directorio** en el que se encuentra el documento con la macro.

---

### Cambiar el nombre del archivo generado

Para modificar el nombre del documento que se generará, se puede editar la siguiente línea:

```vb
newFileName = "Actualizacion_reporte_" & Format(Date, "yyyy-mm-dd") & ".xlsx"
```

##Hojas que se copian al nuevo libro

En la siguiente parte del código se definen las hojas que serán copiadas al nuevo libro.
Si se desean agregar más hojas, deberán añadirse de la misma forma o, en su caso, únicamente cambiar los nombres:
```vb
Set wsVida = ThisWorkbook.Sheets("Polizas de VIDA en 2025")
Set wsGMM = ThisWorkbook.Sheets("Polizas de GMM en 2025")
```

# Gestión de la configuración

En la carpeta del proyecto se encuentran cada uno de los códigos aplicados en las macros del archivo de reporte de pólizas de GPA. Estos se pueden abrir directamente en un IDE para editarse en caso de no querer usar directamente el editor de Excel.

El orden de las carpetas se encuentra de la siguiente manera:

- **Documento_Proyecto_Revision_Polizas**
- **reportes**
  - `FILTRACION_NUEVOS_AXA.bas`
  - `NUEVO_GENERADOR_REPORTE_MENSUAL.bas`
  - `RANKING_MES_GMM.bas`
- **revision_primas**
  - `Formulario_botones.frm`
  - `Formulario_botones.frx`
  - `LIMPIAR_REPORTE_AXA.bas`
  - `REVISAR_PRIMA.bas`
  - `REVISAR_PRIMAS_CON_RECLICADAS.bas`
  - `REVISAR_PRIMA_CC_GMM.bas`
  - `REVISAR_PRIMA_CNMENSUALIDADES.bas`
- `ENVIAR ACTUALIZACION.bas`

