# Revisión de Prima Pagada — Documentación

Este pequeño proyecto es resultado y forma parte de mi tiempo de practicas profesionales que lleve a cabo en la promotoría Grupo Ponce y Asociados alrededor de 3 meses cumpliendo 320 horas de prácticas profesionales.
Agradezco a Grupo Ponce y Asociados por la oportunidad que me brindaron de haber sido parte de su equipo durante este tiempo, el cual fue de gran aprendizaje y emoción día con día durante esta etapa.

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

Finalmente el nombre del nuevo libro con el reporte guardado será: 
```vb
Actualizacion_reporte(fecha).xlsx
```
# Proceso de revision de pólizas

# Revisión de Prima Pagada

Para la revisión de prima pagada se recomienda seguir los siguientes pasos para el correcto uso de las macros en los libros asignados.

1. **Descargar los archivos de reportes mensuales** de cada centro de costo desde la página de AXA Seguros.  
   - Para conocer el proceso de descarga de los archivos, consulte a un encargado del área.

2. **Ingrese a la hoja _REVISION Y REPORTES_.**

3. Seleccione la opción **REVISAR PRIMA PAGADA POR CENTRO DE COSTO**.

4. Se abrirá una pestaña en la que debe **ingresar la terminación de las pólizas a revisar**.

5. Una vez ingresado, dé clic en **Aceptar** o presione **Enter**.

6. Se abrirá una ventana del explorador de archivos donde debe **seleccionar uno de los archivos del centro de costo previamente descargados**.

7. Se abrirá una pestaña donde debe **ingresar el mes a revisar del reporte de GPA**.  
   - Una vez ingresado, presione **Enter**.

8. **Repita los pasos anteriores** para cada uno de los archivos de los respectivos centros de costo.

9. Diríjase a su **hoja de reporte de pólizas (GMM o VIDA)**.

10. En la hoja verá la columna de las pólizas del mes ingresado, en dos colores:
    - **Verde** – Indica que la póliza aparece en el reporte mensual de Axa, por lo que se puede marcar como pagada.  
      - Marque la celda de **“PAGADA”** como **SI**.
    - **Rojo** – Indica que la póliza no aparece en el reporte mensual de Axa; esto indica que **la póliza aún no está pagada**.

11. Marque las celdas de **prima pagada** correspondientes con **SI**, según el color de la celda de la póliza.

### Diagrama del proceso

<img width="1354" height="344" alt="DiagramaPrima drawio" src="https://github.com/user-attachments/assets/07a48d68-0fd9-4b70-a339-b18e1c7a2f00" />

# Proceso de Revisión de Prima con Mensualidades

Para la revisión de pólizas con frecuencia de pago mensual, el proceso es bastante similar al de revisión por mes explicado en la sección anterior.

1. **Descargar los archivos de reportes mensuales** de cada centro de costo desde la página de AXA Seguros.  
   - Para conocer el proceso de descarga de los archivos, consulte a un encargado del área.

2. **Ingrese a la hoja _REVISION Y REPORTES_.**

3. Seleccione la opción **REVISAR PRIMA PAGADA CON MENSUALIDADES INCLUIDAS POR CENTRO DE COSTO**.

4. Ingrese la **terminación o las terminaciones de las pólizas** que se revisarán durante todo el año y que se pagan mensualmente.

5. Se abrirá el explorador de archivos donde debe **seleccionar el archivo de reporte mensual de AXA** de cada centro de costo.

6. **Repita los pasos anteriores** para cada centro de costo.

7. Diríjase a su **hoja de reporte de pólizas (GMM o VIDA)**.

8. En la hoja verá que **todas las columnas de pólizas pagadas** (es decir, que aparecen en el reporte de AXA) estarán **señaladas en color Verde**.  
   - Esto indica que la póliza de pago mensual se ha pagado nuevamente en el mes del reporte descargado y, por lo tanto, **sigue vigente**.

9. **Señale el mes o la fecha del último pago** de la póliza mensual, según lo indicado por el encargado del área.

---

### Diagrama del proceso
<img width="1384" height="274" alt="Mensualidades drawio" src="https://github.com/user-attachments/assets/12d5c88f-4947-484d-9bd7-cc55f399f22e" />


# Generar Reportes

Para el proceso que genera un reporte mensual con ranking de agentes, se deben seguir los siguientes pasos:

1. **Descargar todos los archivos de reporte mensual de AXA** de todos los centros de costo.  
   - Para conocer el proceso de descarga de los archivos, consulte a un encargado del área.

2. **Ingrese a la hoja _REVISION Y REPORTES_.**

3. Seleccione la opción **GENERACIÓN DE REPORTES**.

4. Se abrirá una ventana con los pasos a seguir:

   a. **Paso 1:**  
      Seleccione la primera opción e **ingrese los archivos de los cinco centros de costo** previamente descargados.  
      - Se genera una nueva hoja de cálculo en el libro, con el nombre **“Reporte consolidado”**.

   b. **Paso 2:**  
      Seleccione la segunda opción, que **limpia los datos para “nuevos en AXA” mayores a 0**.  
      - Se genera una nueva hoja de cálculo llamada **“Comparativo pólizas”**.

   c. **Paso 3:**  
      Seleccione la tercera opción, que **genera el ranking**.  
      - Se genera una hoja de cálculo llamada **“Ranking Agentes”**.

De esta manera, se crea un **nuevo ranking mensual**, que **no toma en cuenta las pólizas con valor 0 en “nuevos en AXA”**.

<img width="1173" height="354" alt="Reportes drawio" src="https://github.com/user-attachments/assets/8a3d4550-baa4-488a-8bd8-bcf1312ad7c2" />


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

