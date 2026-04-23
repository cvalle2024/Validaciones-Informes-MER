# 📖 Documentación – Validaciones de indicadores MER

## 1. Introducción
Este portal describe la estructura y funcionalidad del sistema de validaciones automáticas para los indicadores del proyecto VIHCA, conforme a los lineamientos establecidos por la **Guía MER de PEPFAR**.

Su propósito es asegurar la integridad, consistencia y calidad de los datos reportados, mediante la ejecución de validaciones automatizadas y la presentación de alertas tempranas para la oportuna corrección de errores antes del envío final de la información.

---

## 2. Instrucciones de uso del Portal de Validaciones
1. Ingresar al portal mediante el enlace compartido por el equipo de M&E regional.
2. **Cargar** uno o varios archivos `.xlsx` o un archivo `.zip` con subcarpetas.
3. Pulsar el botón **Procesar**.
4. Utilizar los **segmentadores** para filtrar por:
   - **País**
   - **Departamento**
   - **Sitio**
5. Revisar las secciones de:
   - **Resumen**
   - **% de error**
   - **Detalle**
   - **Métricas**
6. **Descargar** el archivo Excel completo o filtrado.
7. Aplicar las correcciones necesarias antes del envío final a la jefatura inmediata o a la plataforma correspondiente.

> **Importante:** para la validación de conciliación TX_CURR trimestral, se recomienda cargar archivos de más de un trimestre para que el sistema pueda comparar la cohorte base con el trimestre siguiente.

---

## 3. Objetivos del Portal de Validaciones
- Detectar errores comunes de forma anticipada en las bases de datos locales de cada país, antes de cargar información en DATIM.
- Generar visualizaciones y tablas resumen de los errores encontrados en los archivos cargados.
- Fortalecer la calidad, consistencia y confiabilidad de los datos reportados por los equipos nacionales.
- Facilitar el análisis de brechas entre cohortes trimestrales en indicadores de tratamiento.

---

## 4. Indicadores y reglas que se validan

### 4.1 Formato fecha diagnóstico (HTS_TST)
- **Regla:** la fecha del diagnóstico debe estar en un formato de fecha válido, preferiblemente `dd/mm/yyyy`.

### 4.2 ID duplicado (HTS_TST)
- **Regla:** se verifica que el mismo ID de expediente no se repita dentro del mismo trimestre.

### 4.3 Fecha de inicio de TARV < Fecha del diagnóstico (HTS_TST)
- **Regla:** la `Fecha de inicio TARV` no debe ser menor que la `Fecha del diagnóstico`.

### 4.4 CD4 vacío en diagnósticos positivos (HTS_TST)
- **Regla:** si el `Resultado de la prueba de VIH` es **Positivo**, el campo `CD4 Basal` no debe estar vacío.

### 4.5 TX_PVLS Numerador > TX_PVLS Denominador
- **Regla:** el `Numerador` no debe ser mayor que el `Denominador`.
- **Variables revisadas:** **Sexo + Tipo de población + Rango de edad**.

### 4.6 TX_PVLS Denominador > TX_CURR
- **Regla:** el `Denominador` de TX_PVLS no debe ser mayor que el `TX_CURR`.
- **Variables revisadas:** **Sexo + Tipo de población + Rango de edad**.

### 4.7 TX_CURR ≠ Dispensación_TARV (cuadros dentro de TX_CURR)
- **Regla:** se verifica que el valor por sexo y rango de edad sea el mismo en ambos cuadros.
- **Variables revisadas:** **Sexo + Rango de edad**.

### 4.8 Verificación de Sexo (HTS_TST)
- **Regla:** en la columna `Sexo` solo deben registrarse los valores:
  - `Femenino`
  - `Masculino`

### 4.9 TX_ML: Última cita esperada vacía
- **Regla:** en la hoja `TX_ML`, la columna `Fecha de su última cita esperada` no debe estar vacía cuando el registro corresponde a un sitio válido.

### 4.10 Conciliación TX_CURR trimestral
- **Regla:** se verifica la coherencia entre la cohorte base del trimestre anterior y el TX_CURR reportado en el trimestre siguiente.

#### Fórmula de conciliación:
`TX_CURR base + TX_NEW + TX_RTT + Traslado recibido - TX_ML total = TX_CURR esperado`

Luego se compara:
`TX_CURR esperado` vs `TX_CURR real reportado`

#### Lógica de comparación:
La comparación se realiza **trimestre a trimestre**, de forma encadenada:
- **Q1 → Q2**
- **Q2 → Q3**
- **Q3 → Q4**
- **Q4 → Q1** del siguiente año fiscal

#### Criterios:
- Si la diferencia es **0**, el sitio **cuadra**.
- Si la diferencia es distinta de **0**, se marca como **error de conciliación**.
- El sistema clasifica el resultado como:
  - **Cuadra**
  - **TX_CURR real mayor al esperado**
  - **TX_CURR real menor al esperado**

#### Componentes analizados:
- **TX_CURR base**
- **TX_NEW**
- **TX_RTT**
- **Traslado recibido**
- **TX_ML total**
- Modalidades de salida en TX_ML, tales como:
  - Fallecido
  - ITT <3 meses en TAR
  - ITT 3 a 5 meses en TAR
  - ITT >6 meses en TAR
  - Paciente rechaza (finaliza) el tratamiento
  - Paciente transferido

---

## 5. Segmentadores (filtros)
En esta sección podrá seleccionar:
- **País**
- **Departamento**
- **Sitio**

Orden recomendado de uso:
**País → Departamento → Sitio**

> En la validación de conciliación TX_CURR trimestral, los filtros permiten revisar los hallazgos por sitio y por país.

---

## 6. Cálculos y % de errores
- **Errores:** cantidad de registros que incumplen la regla validada.
- **Chequeos:** cantidad de registros o combinaciones revisadas por la validación.
- **% Error** = `errores / chequeos * 100`

### En la conciliación TX_CURR trimestral:
- Un **check** corresponde a una comparación válida entre:
  - un **trimestre base**
  - y un **trimestre comparado**
  para un sitio determinado.

Ejemplo:
- **Base:** Q1 FY26
- **Comparado:** Q2 FY26

Si el cálculo no coincide con el TX_CURR reportado, se cuenta como error.

---

## 7. Archivo exportable Excel

El archivo exportable puede incluir las siguientes hojas, según los errores encontrados:

### 7.1 Hojas generales
- **Resumen**
  - Número de errores encontrados por indicador.
- **Resumen de errores por indicador**
  - Cada validación se exporta en hojas separadas.
  - Si no se encuentran errores, la hoja puede no mostrarse.

### 7.2 Hojas de conciliación TX_CURR
Cuando existan datos suficientes para realizar el análisis de cohorte trimestral, el sistema puede generar además:

- **Conciliación TX_CURR trimestral**
  - Muestra el detalle del cálculo por transición trimestral.
- **Auditoria_Sitio TX_CURR**
  - Muestra el detalle limpio por sitio, incluyendo:
    - `Q_Base`
    - `Q_Comparado`
    - `Transición`
    - `País`
    - `Site`
    - `TX_CURR_Base`
    - `TX_NEW`
    - `TX_RTT`
    - `Traslado_Recibido`
    - `TX_ML_Total`
    - modalidades de TX_ML
    - `TX_CURR_Esperado`
    - `TX_CURR_Real`
    - `Brecha`
    - `Cuadra`
    - `Tipo_Error`
- **Resumen_Conciliacion_TX_CURR**
  - Resume por transición trimestral:
    - cantidad de sitios evaluados
    - sitios que cuadran
    - sitios con error
    - porcentaje de error
    - brecha neta
    - brecha absoluta

### 7.3 Resaltado en el Excel
- Se resalta en rojo la **columna con error** en cada hoja.
- En conciliación TX_CURR trimestral se prioriza el resaltado de:
  - **Brecha**
  - **TX_CURR esperado**
  - **TX_CURR real**

---

## 8. Interpretación de resultados de conciliación TX_CURR

### Cuando un sitio cuadra
Significa que:
`TX_CURR base + entradas - salidas = TX_CURR real`

### Cuando un sitio no cuadra
Puede indicar una o varias de las siguientes situaciones:
- omisión de registros en `TX_NEW`
- omisión de retornos o traslados recibidos en `TX_RTT`
- omisión o clasificación incorrecta en `TX_ML`
- diferencias entre la cohorte base y la cohorte reportada
- problemas de calidad del dato en el sitio
- inconsistencias entre archivos de distintos meses del trimestre

---

## 9. Recomendaciones
- Cada error identificado automáticamente permite fortalecer la calidad del dato en campo, en clínicas y durante el procesamiento de bases.
- Con base en la frecuencia de errores encontrados, se pueden reforzar las indicaciones sobre cómo construir correctamente cada indicador según la Guía MER.
- Si no existen **checks** válidos en una selección, revisar:
  - filtros aplicados
  - fechas cargadas
  - cantidad de trimestres disponibles
  - estructura de archivos
- Mantener un registro histórico de errores frecuentes y de las acciones correctivas implementadas facilita el seguimiento y mejora continua.
- Para la conciliación TX_CURR trimestral, es recomendable revisar especialmente los sitios que presentan brechas persistentes entre trimestres consecutivos.

---

## 10. Consideraciones finales
Este portal constituye una herramienta de apoyo para el aseguramiento de calidad del dato. Los resultados deben ser utilizados como insumo para la revisión técnica de los equipos nacionales y regionales, y no sustituyen la validación programática de los indicadores antes de su reporte final.
