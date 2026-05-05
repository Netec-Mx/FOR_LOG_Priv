---LAB_START---
LAB_ID: 01-00-01
---MARKDOWN---
# Práctica 1 — Análisis Exploratorio de Datos Históricos de Ventas en Excel

## 1. Metadatos

| Campo            | Detalle                                      |
|------------------|----------------------------------------------|
| **Duración**     | 12 minutos                                   |
| **Complejidad**  | Básica (Easy)                                |
| **Nivel Bloom**  | Aplicar (Apply)                              |
| **Módulo**       | Módulo 1 — Fundamentos de Forecasting        |
| **Práctica**     | 1 de 10 (inicio de secuencia)                |
| **Archivo base** | `Dataset_Ventas_24M_8SKUs.xlsx` (provisto por el instructor) |

---

## 2. Descripción General

En esta práctica explorarás un dataset real de ventas mensuales de 24 meses para 8 SKUs de una empresa distribuidora ficticia. Aplicarás funciones básicas de Excel para construir una tabla de resumen exploratorio que te permita identificar el comportamiento histórico de la demanda por producto. Los resultados de esta práctica son el punto de partida para todas las prácticas posteriores del curso: el dataset depurado y la tabla de resumen que generes aquí serán insumos directos de las Prácticas 2, 3 y 4. El enfoque está anclado al concepto central de la Lección 1.1: **el forecasting comienza por entender la demanda histórica real** antes de proyectar cualquier escenario futuro.

---

## 3. Objetivos de Aprendizaje

Al completar esta práctica serás capaz de:

- [ ] Identificar la estructura de un dataset de ventas históricas (campos clave: fecha, SKU, descripción, unidades vendidas, canal) como insumo para el proceso de forecasting logístico.
- [ ] Aplicar funciones de Excel (`SUMA`, `PROMEDIO`, `MAX`, `MIN`, `CONTAR`, `SUMAR.SI`) para calcular métricas exploratorias básicas por SKU y por período.
- [ ] Construir una tabla de resumen exploratorio con métricas de demanda histórica (total anual, promedio mensual, máximo y mínimo) que permita distinguir SKUs de comportamiento estable versus variable.
- [ ] Reconocer visualmente en tablas de datos los componentes del comportamiento de la demanda (tendencia, estacionalidad, variabilidad) como precursor de la construcción del pronóstico.
- [ ] Documentar observaciones iniciales sobre la calidad de los datos e identificar la diferencia entre demanda histórica real y demanda proyectada.

---

## 4. Prerequisitos

### Conocimiento previo requerido
- Experiencia operativa básica en procesos de inventario, distribución o logística.
- Manejo básico de Excel: navegación entre hojas, ingreso de datos, aplicación de fórmulas simples en celdas.
- Haber revisado los conceptos teóricos del Módulo 1, Lección 1.1 (Introducción al Forecasting y su Aplicación en Logística): concepto de forecasting, rol en inventarios y abastecimiento, impacto en rotación, nivel de servicio y quiebres de stock.

### Acceso y recursos necesarios
- Computadora con Microsoft Excel 2016 o superior instalado (Microsoft 365 recomendado).
- Archivo `Dataset_Ventas_24M_8SKUs.xlsx` distribuido por el instructor (debe estar descargado y accesible antes de iniciar).
- No se requiere conexión a internet para esta práctica.

---

## 5. Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente       | Mínimo                        | Recomendado                   |
|------------------|-------------------------------|-------------------------------|
| Procesador       | Intel Core i5 / AMD Ryzen 5   | Intel Core i7 / AMD Ryzen 7   |
| RAM              | 4 GB                          | 8 GB                          |
| Almacenamiento   | 500 MB libres                 | 2 GB libres                   |
| Pantalla         | 1280 × 768                    | 1920 × 1080                   |

### Software requerido

| Software              | Versión mínima     | Notas                                      |
|-----------------------|--------------------|--------------------------------------------|
| Microsoft Excel       | 2016               | Microsoft 365 recomendado                  |
| Archivo de datos      | `.xlsx` provisto   | No modificar el archivo original; trabajar en copia |

### Configuración inicial (antes de comenzar)

1. Localiza el archivo `Dataset_Ventas_24M_8SKUs.xlsx` en la carpeta de práctica proporcionada por el instructor.
2. **Crea una copia de trabajo** antes de abrir el archivo original:
   - Haz clic derecho sobre el archivo → **Copiar** → pega en la misma carpeta.
   - Renombra la copia como: `P1_TuNombre_Exploratorio.xlsx`
3. Abre la copia de trabajo en Excel.
4. Verifica que el archivo contenga al menos las siguientes hojas:
   - `Ventas_Historicas` — datos mensuales de ventas (hoja principal de trabajo).
   - `Catalogo_SKU` — descripción y categoría de cada SKU (referencia).
5. Asegúrate de que Excel muestra el archivo en **Vista Normal** (no en modo de compatibilidad).

> ⚠️ **Importante:** Nunca trabajes sobre el archivo original. Si cometes un error que corrompe los datos, la copia te permite reiniciar sin perder el dataset base.

---

## 6. Instrucciones Paso a Paso

---

### Paso 1 — Revisión de la Estructura del Dataset

**Objetivo:** Comprender la organización del dataset antes de calcular cualquier métrica. En forecasting logístico, conocer la estructura de los datos disponibles es el primer paso para determinar qué modelos de pronóstico se pueden aplicar.

#### Instrucciones

1. En Excel, activa la hoja **`Ventas_Historicas`**.

2. Identifica las columnas disponibles en el dataset. El archivo debe contener las siguientes columnas (verifica que todas estén presentes):

   | Columna           | Descripción                                            |
   |-------------------|--------------------------------------------------------|
   | `Fecha`           | Mes y año del registro (formato: `ene-22`, `feb-22`, …) |
   | `SKU`             | Código único del producto (ej. `SKU-001`)              |
   | `Descripcion`     | Nombre del producto                                    |
   | `Categoria`       | Categoría del producto (Alimentos / Limpieza / Cuidado Personal) |
   | `Canal`           | Canal de venta (Retail / Mayorista / E-commerce)       |
   | `Unidades_Vendidas` | Cantidad de unidades vendidas en el período           |

3. Anota en papel o en una celda vacía de la hoja `Ventas_Historicas` (por ejemplo, celda `I1`) cuántas filas de datos contiene el dataset (sin contar el encabezado). Puedes contar manualmente o usar:
   ```
   =CONTAR(A2:A1000)
   ```
   > El resultado esperado es **192 filas** (8 SKUs × 24 meses = 192 registros). Si obtienes un número diferente, notifícalo al instructor antes de continuar.

4. Activa la hoja **`Catalogo_SKU`** y revisa los 8 SKUs disponibles. Anota mentalmente (o en papel) las tres categorías presentes: **Alimentos**, **Limpieza** y **Cuidado Personal**.

5. Regresa a la hoja `Ventas_Historicas`.

6. Aplica **filtros automáticos** al dataset:
   - Haz clic en cualquier celda dentro del rango de datos.
   - Ve a la pestaña **Datos** → grupo **Ordenar y filtrar** → clic en **Filtro**.
   - Aparecerán flechas desplegables en cada encabezado de columna.

7. Usa el filtro de la columna `SKU` para visualizar únicamente los registros de `SKU-001`. Observa cuántos meses de datos aparecen (debe ser 24). Luego **limpia el filtro** (Datos → Borrar) para volver a ver todos los registros.

#### Resultado esperado

- El dataset tiene 192 filas de datos (8 SKUs × 24 meses).
- Todas las columnas descritas están presentes y tienen datos.
- Los filtros automáticos están activos y funcionando correctamente.

#### Verificación

> ✅ Confirma: ¿La celda `I1` (o donde hayas colocado la fórmula `CONTAR`) muestra el valor **192**? ¿Los filtros automáticos están activos en todos los encabezados? Si ambas respuestas son sí, continúa al Paso 2.

---

### Paso 2 — Creación de la Hoja de Resumen Exploratorio

**Objetivo:** Preparar el espacio de trabajo donde construirás la tabla de resumen. Mantener el análisis en una hoja separada es una buena práctica que preserva los datos originales intactos.

#### Instrucciones

1. Crea una nueva hoja en el archivo:
   - Haz clic en el botón **`+`** (nueva hoja) en la barra inferior de pestañas.
   - Haz doble clic sobre la nueva pestaña y renómbrala como: **`Resumen_Exploratorio`**

2. En la hoja `Resumen_Exploratorio`, construye el encabezado de la tabla de resumen. Ingresa los siguientes textos en las celdas indicadas:

   | Celda | Texto a ingresar          |
   |-------|---------------------------|
   | `A1`  | `SKU`                     |
   | `B1`  | `Descripcion`             |
   | `C1`  | `Categoria`               |
   | `D1`  | `Total_Año1`              |
   | `E1`  | `Total_Año2`              |
   | `F1`  | `Total_24M`               |
   | `G1`  | `Promedio_Mensual`        |
   | `H1`  | `Maximo_Mensual`          |
   | `I1`  | `Minimo_Mensual`          |
   | `J1`  | `Meses_con_Datos`         |
   | `K1`  | `Comportamiento`          |

3. Aplica **formato de tabla** al rango `A1:K1` (solo el encabezado por ahora):
   - Selecciona `A1:K1`.
   - Ve a **Inicio** → **Estilos** → **Dar formato como tabla** → elige un estilo con encabezado oscuro (por ejemplo, "Estilo de tabla medio 2").
   - Cuando Excel pregunte si la tabla tiene encabezados, selecciona **Sí**.

4. Ingresa los códigos de los 8 SKUs en la columna `A`, filas 2 a 9:

   | Celda | Valor     |
   |-------|-----------|
   | `A2`  | `SKU-001` |
   | `A3`  | `SKU-002` |
   | `A4`  | `SKU-003` |
   | `A5`  | `SKU-004` |
   | `A6`  | `SKU-005` |
   | `A7`  | `SKU-006` |
   | `A8`  | `SKU-007` |
   | `A9`  | `SKU-008` |

5. Completa las columnas `B` (Descripcion) y `C` (Categoria) consultando la hoja `Catalogo_SKU`. Puedes ingresar los valores manualmente o usar referencias directas a esa hoja. Por ejemplo, en `B2`:
   ```
   =Catalogo_SKU!B2
   ```
   > Si los datos del catálogo están organizados de forma diferente, ingresa los valores manualmente consultando la hoja de referencia.

#### Resultado esperado

- La hoja `Resumen_Exploratorio` existe con el encabezado de 11 columnas correctamente etiquetadas.
- Los 8 SKUs están listados en las filas 2 a 9 con su descripción y categoría.

#### Verificación

> ✅ Confirma: ¿La hoja `Resumen_Exploratorio` tiene 11 columnas encabezadas y 8 filas de SKUs con descripción y categoría? Si es así, continúa al Paso 3.

---

### Paso 3 — Cálculo de Métricas de Agregación con SUMAR.SI

**Objetivo:** Calcular los totales de ventas por SKU para el Año 1 (meses 1–12) y el Año 2 (meses 13–24), y el total acumulado de 24 meses. Esta información permite identificar si la demanda de un SKU crece, decrece o se mantiene estable en el tiempo, lo cual es fundamental para decidir qué tipo de modelo de pronóstico aplicar (concepto central de la Lección 1.1).

#### Instrucciones

1. Activa la hoja `Resumen_Exploratorio`.

2. En la celda **`D2`**, ingresa la fórmula para calcular el total de ventas de `SKU-001` durante el **Año 1** (primeros 12 meses del dataset):
   ```
   =SUMAR.SI(Ventas_Historicas!$B:$B,A2,Ventas_Historicas!$F:$F)
   ```
   > **Nota:** Esta fórmula suma todas las unidades vendidas en la columna `F` de la hoja `Ventas_Historicas` donde el valor de la columna `B` (SKU) coincide con el valor en `A2`. Si la columna de SKU o de unidades vendidas en tu dataset está en columnas diferentes, ajusta las referencias de columna (`$B:$B` y `$F:$F`) según corresponda.

   > ⚠️ **Ajuste para Año 1 vs Año 2:** Para calcular el total del Año 1 (primeros 12 meses) necesitarás filtrar también por rango de fechas. La forma más práctica en este nivel es usar un rango de filas específico. Dado que el dataset tiene 24 filas por SKU ordenadas cronológicamente, los primeros 12 registros de cada SKU corresponden al Año 1 y los siguientes 12 al Año 2. Usa la siguiente versión simplificada:
   ```
   =SUMAR.SI(Ventas_Historicas!$B$2:$B$97,A2,Ventas_Historicas!$F$2:$F$97)
   ```
   *(Filas 2 a 97 = primeros 12 meses × 8 SKUs = 96 registros del Año 1)*

3. En la celda **`E2`**, calcula el total del **Año 2** (meses 13–24):
   ```
   =SUMAR.SI(Ventas_Historicas!$B$98:$B$193,A2,Ventas_Historicas!$F$98:$F$193)
   ```
   *(Filas 98 a 193 = segundos 12 meses × 8 SKUs = 96 registros del Año 2)*

4. En la celda **`F2`**, calcula el **total de 24 meses**:
   ```
   =D2+E2
   ```
   > Alternativamente: `=SUMAR.SI(Ventas_Historicas!$B:$B,A2,Ventas_Historicas!$F:$F)`

5. **Copia las fórmulas de `D2:F2` hacia abajo** hasta la fila 9 (SKU-008):
   - Selecciona el rango `D2:F2`.
   - Posiciona el cursor en la esquina inferior derecha del rango seleccionado hasta que aparezca el cursor de cruz negra (+).
   - Arrastra hacia abajo hasta la fila 9.

6. Revisa los resultados. Los totales deben ser números positivos y razonables (en el rango de cientos a miles de unidades dependiendo del SKU). Si alguna celda muestra `0` o `#¡VALOR!`, verifica que las referencias de columna sean correctas.

#### Resultado esperado

- Las columnas `D`, `E` y `F` tienen valores numéricos para los 8 SKUs.
- Algunos SKUs mostrarán `Total_Año2` mayor que `Total_Año1` (tendencia creciente), otros al revés (tendencia decreciente) y otros similares (demanda estable).

#### Verificación

> ✅ Confirma: ¿Todas las celdas `D2:F9` contienen valores numéricos positivos (no ceros ni errores)? ¿Puedes identificar al menos un SKU cuyo Año 2 es notablemente diferente al Año 1? Anota ese SKU. Continúa al Paso 4.

---

### Paso 4 — Cálculo de Promedio, Máximo y Mínimo Mensual por SKU

**Objetivo:** Calcular las métricas de dispersión básicas que permiten cuantificar la variabilidad de la demanda de cada SKU. El promedio mensual representa la demanda base (baseline), mientras que el máximo y mínimo permiten identificar la amplitud de la variación, un indicador clave de la estacionalidad y la volatilidad que se estudiarán en profundidad en las Prácticas 3 y 4.

#### Instrucciones

1. En la celda **`G2`**, calcula el **promedio mensual** de ventas de `SKU-001` en los 24 meses:
   ```
   =PROMEDIO(SI(Ventas_Historicas!$B$2:$B$193=A2,Ventas_Historicas!$F$2:$F$193))
   ```
   > ⚠️ **Esta es una fórmula matricial.** Después de escribirla, en lugar de presionar solo `Enter`, presiona **`Ctrl + Shift + Enter`** (Excel 2016/2019). En Microsoft 365, puedes presionar `Enter` normalmente ya que las fórmulas matriciales se gestionan automáticamente. Si la fórmula funciona correctamente, verás llaves `{ }` alrededor de ella en la barra de fórmulas (Excel 2016/2019).

   > **Alternativa sin fórmula matricial** (más simple, recomendada si tienes dificultades):
   > Puedes calcular el promedio dividiendo el total entre 24:
   ```
   =F2/24
   ```

2. En la celda **`H2`**, calcula el **máximo mensual** de ventas de `SKU-001`:
   ```
   =MAX(SI(Ventas_Historicas!$B$2:$B$193=A2,Ventas_Historicas!$F$2:$F$193))
   ```
   > ⚠️ También es fórmula matricial: presiona **`Ctrl + Shift + Enter`** en Excel 2016/2019.

3. En la celda **`I2`**, calcula el **mínimo mensual** de ventas de `SKU-001`:
   ```
   =MIN(SI(Ventas_Historicas!$B$2:$B$193=A2,Ventas_Historicas!$F$2:$F$193))
   ```
   > ⚠️ También es fórmula matricial: presiona **`Ctrl + Shift + Enter`** en Excel 2016/2019.

4. En la celda **`J2`**, calcula el **número de meses con datos** para `SKU-001`:
   ```
   =CONTAR.SI(Ventas_Historicas!$B:$B,A2)
   ```
   > El resultado debe ser **24** para todos los SKUs. Si algún SKU muestra un número menor, significa que hay meses sin datos — una observación importante para documentar en el Paso 6.

5. **Copia las fórmulas de `G2:J2` hacia abajo** hasta la fila 9:
   - Selecciona `G2:J2`.
   - Arrastra el controlador de relleno hacia abajo hasta la fila 9.
   - Verifica que los valores cambien para cada SKU (si todos muestran el mismo valor, las referencias de fila no se están ajustando correctamente — revisa que `A2` en las fórmulas no esté bloqueado con `$A$2`).

6. Formatea las columnas `G`, `H` e `I` para mostrar **un decimal**:
   - Selecciona `G2:I9`.
   - Clic derecho → **Formato de celdas** → **Número** → Posiciones decimales: **1** → Aceptar.

#### Resultado esperado

La tabla de resumen parcial debe verse similar a esta estructura (los valores específicos dependen del dataset):

| SKU     | Descripcion         | Categoria | Total_Año1 | Total_Año2 | Total_24M | Promedio_Mensual | Maximo_Mensual | Minimo_Mensual | Meses_con_Datos |
|---------|---------------------|-----------|------------|------------|-----------|-----------------|----------------|----------------|-----------------|
| SKU-001 | Detergente Líquido  | Limpieza  | 4,320      | 4,680      | 9,000     | 375.0           | 520            | 280            | 24              |
| SKU-002 | Arroz 5kg           | Alimentos | 6,100      | 5,900      | 12,000    | 500.0           | 650            | 400            | 24              |
| …       | …                   | …         | …          | …          | …         | …               | …              | …              | …               |

> Los valores de tu tabla serán diferentes; lo importante es que la estructura sea correcta y los números sean coherentes.

#### Verificación

> ✅ Confirma: ¿Las columnas `G`, `H` e `I` muestran valores diferentes para cada SKU? ¿La columna `J` muestra **24** para todos los SKUs (o un número menor para alguno, lo cual deberás anotar)? ¿Los valores de `H` (máximo) son siempre mayores que los de `I` (mínimo)? Si todo es correcto, continúa al Paso 5.

---

### Paso 5 — Clasificación Manual del Comportamiento de la Demanda

**Objetivo:** Clasificar cada SKU como de comportamiento **Estable**, **Variable** o **Muy Variable** basándose en la relación entre el máximo y el mínimo mensual. Esta clasificación es la base para decidir qué nivel de sofisticación requiere el modelo de pronóstico de cada producto — un SKU estable puede pronosticarse con un promedio simple, mientras que uno muy variable requiere modelos que capturen estacionalidad o tendencia (conceptos que se profundizarán en las Prácticas 3 y 4).

#### Instrucciones

1. En la celda **`K2`**, ingresa manualmente la clasificación del comportamiento de `SKU-001` usando el siguiente criterio:

   **Criterio de clasificación (basado en la razón Máximo/Mínimo):**

   | Razón Máximo / Mínimo | Clasificación     |
   |-----------------------|-------------------|
   | Menor a 1.5           | `Estable`         |
   | Entre 1.5 y 2.5       | `Variable`        |
   | Mayor a 2.5           | `Muy Variable`    |

   Para calcular la razón de `SKU-001`, puedes usar una celda auxiliar temporal:
   ```
   =H2/I2
   ```
   Evalúa el resultado y escribe manualmente en `K2` la clasificación correspondiente (`Estable`, `Variable` o `Muy Variable`).

   > **¿Por qué este criterio?** Una razón de 1.5 significa que el mes de mayor demanda vendió 50% más que el mes de menor demanda. Eso indica cierta variabilidad pero dentro de rangos manejables. Una razón mayor a 2.5 sugiere que hay meses que duplican o triplican a otros — señal de estacionalidad marcada o eventos extraordinarios que deben considerarse en el pronóstico.

2. Repite el proceso para los 8 SKUs (filas 3 a 9). Calcula la razón H/I para cada uno y escribe la clasificación en la columna `K`.

   > **Tip de eficiencia:** Puedes usar la columna `L` como columna auxiliar temporal para calcular todas las razones:
   > - En `L2`: `=H2/I2`, copia hasta `L9`.
   > - Evalúa cada valor y escribe la clasificación en `K`.
   > - Una vez terminado, puedes eliminar la columna `L` auxiliar.

3. Aplica **formato de color** a la columna `K` para hacer la clasificación visualmente intuitiva:
   - Selecciona `K2:K9`.
   - Ve a **Inicio** → **Estilos** → **Formato condicional** → **Resaltar reglas de celdas** → **Texto que contiene**.
   - Regla 1: Texto que contiene `Estable` → Formato: **Relleno verde claro**.
   - Regla 2: Texto que contiene `Variable` → Formato: **Relleno amarillo**.
   - Regla 3: Texto que contiene `Muy Variable` → Formato: **Relleno rojo claro**.

4. Observa la distribución de clasificaciones. Deberías tener una mezcla de los tres tipos entre los 8 SKUs.

#### Resultado esperado

- La columna `K` tiene una clasificación (`Estable`, `Variable` o `Muy Variable`) para los 8 SKUs.
- El formato condicional muestra verde, amarillo o rojo según la clasificación.
- Al menos 2 SKUs deben clasificarse como `Estable`, al menos 2 como `Variable` y al menos 1 como `Muy Variable` (dependiendo del dataset provisto).

#### Verificación

> ✅ Confirma: ¿Todos los 8 SKUs tienen una clasificación en la columna `K`? ¿El formato condicional de colores está activo y es coherente con las clasificaciones? ¿Puedes identificar cuál SKU tiene la mayor variabilidad (razón más alta)? Anótalo. Continúa al Paso 6.

---

### Paso 6 — Documentación de Observaciones Iniciales

**Objetivo:** Registrar observaciones sobre la calidad de los datos y el comportamiento de la demanda. Esta documentación es un insumo crítico para la Práctica 2 (depuración de datos) y refleja la práctica profesional real: en logística, el analista de demanda siempre documenta sus hallazgos para que otros miembros del equipo puedan entender el estado de los datos.

#### Instrucciones

1. En la hoja `Resumen_Exploratorio`, desplázate hacia abajo (a partir de la fila 12) y crea una sección de observaciones con el siguiente encabezado en la celda **`A12`**:
   ```
   OBSERVACIONES INICIALES DEL DATASET
   ```
   Aplica **negrita** y **relleno gris claro** a esta celda para destacarla.

2. A partir de la celda **`A13`**, documenta tus observaciones respondiendo las siguientes preguntas. Escribe una observación por fila:

   **Preguntas guía para las observaciones:**

   | # | Pregunta de análisis                                                                 | Celda de inicio |
   |---|--------------------------------------------------------------------------------------|-----------------|
   | 1 | ¿Algún SKU tiene menos de 24 meses de datos? ¿Cuál y cuántos meses tiene?           | `A13`           |
   | 2 | ¿Qué SKU tiene el mayor volumen total de ventas en 24 meses? ¿Y el menor?           | `A14`           |
   | 3 | ¿Qué SKU muestra la mayor variabilidad (razón Máx/Mín más alta)?                    | `A15`           |
   | 4 | ¿Algún SKU tiene un Total_Año2 significativamente diferente al Total_Año1? ¿Qué sugiere esto sobre su tendencia? | `A16` |
   | 5 | ¿Hay algún valor de Mínimo_Mensual igual a 0 o cercano a 0? ¿Qué podría indicar?   | `A17`           |
   | 6 | ¿Los datos históricos disponibles (24 meses) son suficientes para construir un pronóstico confiable? ¿Por qué? | `A18` |

3. Para cada pregunta, escribe tu respuesta en la **columna B** de la misma fila. Sé específico: menciona los SKUs por código, cita los valores de la tabla.

   **Ejemplo de observación bien documentada (fila 14):**
   > *"El SKU con mayor volumen total es SKU-002 (Arroz 5kg) con 12,000 unidades en 24 meses. El de menor volumen es SKU-007 (Crema Facial) con 3,840 unidades. La diferencia de 3x entre ambos sugiere que los modelos de pronóstico y los niveles de inventario deben gestionarse de forma diferenciada por SKU."*

4. Agrega una observación adicional libre (fila `A19`) donde respondas la siguiente pregunta desde la perspectiva logística:
   > *"¿Cómo podría impactar la variabilidad que observaste en la demanda de los SKUs clasificados como 'Muy Variable' en el nivel de servicio y en la posibilidad de quiebres de stock de la empresa?"*

   Esta pregunta conecta directamente con el concepto de la Lección 1.1: el impacto del forecasting en el nivel de servicio y los quiebres de stock.

5. **Guarda el archivo** con `Ctrl + S`.

#### Resultado esperado

- La sección de observaciones tiene al menos 7 filas (preguntas 1–6 más la observación libre).
- Cada observación menciona SKUs específicos y valores numéricos de la tabla.
- La observación libre conecta la variabilidad de la demanda con el impacto operativo en nivel de servicio y quiebres de stock.

#### Verificación

> ✅ Confirma: ¿Tienes al menos 7 observaciones documentadas con valores específicos? ¿La observación libre menciona el impacto en nivel de servicio o quiebres de stock? ¿El archivo está guardado? Si todo es correcto, has completado la práctica.

---

## 7. Validación y Prueba Final

Antes de dar por finalizada la práctica, realiza esta verificación integral de tu trabajo:

### Lista de verificación final

Revisa cada ítem y marca ✅ cuando lo hayas confirmado:

| # | Ítem de verificación                                                                 | Estado |
|---|--------------------------------------------------------------------------------------|--------|
| 1 | La hoja `Ventas_Historicas` tiene filtros automáticos activos y el dataset tiene 192 filas de datos. | ☐ |
| 2 | La hoja `Resumen_Exploratorio` existe con 11 columnas correctamente encabezadas.     | ☐ |
| 3 | Los 8 SKUs están listados en filas 2–9 con descripción y categoría.                  | ☐ |
| 4 | Las columnas `D`, `E` y `F` tienen totales por año y total de 24 meses para todos los SKUs. | ☐ |
| 5 | Las columnas `G`, `H` e `I` tienen promedio, máximo y mínimo mensual para todos los SKUs. | ☐ |
| 6 | La columna `J` muestra el número de meses con datos (idealmente 24 para todos).     | ☐ |
| 7 | La columna `K` tiene clasificación de comportamiento con formato condicional de colores. | ☐ |
| 8 | La sección de observaciones tiene al menos 7 entradas con valores específicos.       | ☐ |
| 9 | La observación libre conecta la variabilidad con el impacto en nivel de servicio.   | ☐ |
| 10 | El archivo está guardado como `P1_TuNombre_Exploratorio.xlsx`.                      | ☐ |

### Prueba de coherencia numérica

Realiza esta verificación rápida para confirmar que tus fórmulas son correctas:

1. Selecciona la celda `F2` (Total_24M de SKU-001).
2. Ve a la hoja `Ventas_Historicas`.
3. Aplica el filtro de columna `SKU` para mostrar solo `SKU-001`.
4. En una celda vacía, calcula: `=SUMA(F2:F25)` (ajusta el rango según las filas visibles).
5. El resultado debe coincidir exactamente con el valor en `Resumen_Exploratorio!F2`.
6. Si coincide: ✅ tus fórmulas son correctas. Si no coincide: revisa las referencias de columna en las fórmulas `SUMAR.SI`.
7. Limpia el filtro antes de continuar.

---

## 8. Solución de Problemas

### Problema 1: Las fórmulas SUMAR.SI devuelven 0 para todos los SKUs

**Síntoma:** Las celdas `D2:F9` muestran el valor `0` a pesar de que el dataset tiene datos visibles en la hoja `Ventas_Historicas`.

**Causa probable:** Las referencias de columna en la fórmula `SUMAR.SI` no corresponden a las columnas reales donde están el SKU y las unidades vendidas en el dataset. Esto ocurre cuando el dataset del instructor tiene las columnas en un orden diferente al descrito en estas instrucciones, o cuando la columna de SKU contiene espacios adicionales (espacios al inicio o al final del texto) que impiden la coincidencia.

**Solución:**
1. Ve a la hoja `Ventas_Historicas` y verifica en qué columna está el código SKU (¿es la columna B, C u otra?) y en qué columna están las unidades vendidas.
2. Ajusta las referencias en tu fórmula. Por ejemplo, si el SKU está en la columna C y las unidades en la columna G:
   ```
   =SUMAR.SI(Ventas_Historicas!$C:$C,A2,Ventas_Historicas!$G:$G)
   ```
3. Si el problema persiste, verifica si hay espacios en los códigos SKU: en una celda vacía escribe `=LARGO(Ventas_Historicas!B2)`. Si el resultado es mayor que la longitud esperada del código (ej. `SKU-001` tiene 7 caracteres, si `LARGO` devuelve 8 o más hay espacios ocultos). Solución: usa `SUMAR.SI` con la función `ESPACIOS()` o solicita al instructor un dataset limpio.

---

### Problema 2: Las fórmulas matriciales (MAX/MIN/PROMEDIO con SI) muestran el mismo valor para todos los SKUs

**Síntoma:** Las celdas `G2:I9` muestran exactamente el mismo número en todas las filas, independientemente del SKU.

**Causa probable:** La fórmula matricial no se confirmó correctamente con `Ctrl + Shift + Enter` en Excel 2016/2019. Sin la confirmación matricial, Excel interpreta la función `SI` de forma incorrecta y devuelve un único valor global en lugar de filtrar por SKU. En Microsoft 365 esto puede ocurrir si la fórmula fue copiada desde otra fuente con formato incorrecto.

**Solución:**
1. Haz clic en la celda `G2`.
2. Verifica en la barra de fórmulas si la fórmula aparece entre llaves `{ }`. Si no aparece con llaves, la fórmula no es matricial.
3. Haz clic dentro de la barra de fórmulas (sin cambiar la fórmula) y presiona **`Ctrl + Shift + Enter`** (en Excel 2016/2019). Las llaves deben aparecer automáticamente.
4. Si estás en Microsoft 365 y el problema persiste, usa la alternativa simplificada sin fórmula matricial:
   - Para el promedio: `=F2/24`
   - Para el máximo y mínimo: crea columnas auxiliares en la hoja `Ventas_Historicas` con los valores por SKU usando `SUMAR.SI` con rangos de fila específicos, o solicita al instructor el archivo de checkpoint de esta práctica.

---

## 9. Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos de limpieza para dejar tu entorno ordenado y listo para la Práctica 2:

1. **Elimina columnas auxiliares temporales** (si creaste la columna `L` con las razones Máx/Mín): selecciona el encabezado de la columna `L` → clic derecho → **Eliminar**.

2. **Verifica que los filtros de la hoja `Ventas_Historicas` estén limpios** (sin ningún filtro activo): ve a la hoja `Ventas_Historicas` → pestaña **Datos** → **Borrar** (en el grupo Ordenar y filtrar). Todos los registros deben ser visibles.

3. **Guarda el archivo** por última vez con `Ctrl + S`. Verifica que el nombre del archivo sea `P1_TuNombre_Exploratorio.xlsx`.

4. **Entrega o comparte el archivo** según las instrucciones del instructor (carpeta compartida, correo electrónico o plataforma del curso). Este archivo será el insumo de la Práctica 2.

5. **No cierres Excel** si vas a continuar inmediatamente con la Práctica 2 — mantén el archivo abierto para agilizar la transición.

> 📁 **Archivo de salida de esta práctica:** `P1_TuNombre_Exploratorio.xlsx` con las hojas `Ventas_Historicas` (con filtros activos), `Catalogo_SKU` y `Resumen_Exploratorio` (con tabla de métricas y observaciones documentadas).

---

## 10. Resumen y Próximos Pasos

### Lo que lograste en esta práctica

En 12 minutos aplicaste el primer paso del ciclo de forecasting logístico: **conocer los datos disponibles antes de construir cualquier pronóstico**. Específicamente:

- **Exploraste la estructura** de un dataset real de 24 meses y 8 SKUs, identificando los campos clave que alimentarán los modelos de pronóstico.
- **Calculaste métricas exploratorias** (totales anuales, promedios, máximos y mínimos) usando `SUMAR.SI`, `PROMEDIO`, `MAX` y `MIN`, obteniendo una visión cuantitativa del comportamiento histórico de la demanda.
- **Clasificaste los SKUs** por estabilidad de demanda, distinguiendo entre productos de comportamiento estable (más fáciles de pronosticar) y productos muy variables (que requieren modelos más sofisticados).
- **Conectaste el análisis estadístico con la operación logística**: identificaste que los SKUs de alta variabilidad son los que mayor riesgo representan para el nivel de servicio y los quiebres de stock — el concepto central de la Lección 1.1.

### Conexión con los conceptos de la Lección 1.1

| Concepto de la Lección 1.1                          | Aplicación en esta práctica                                           |
|-----------------------------------------------------|-----------------------------------------------------------------------|
| El forecasting reduce la incertidumbre              | La tabla de resumen cuantifica la variabilidad real de cada SKU       |
| Rol del pronóstico en la gestión de inventarios     | Los totales y promedios son la base para calcular stock de seguridad  |
| Impacto en nivel de servicio                        | SKUs "Muy Variable" son los de mayor riesgo de quiebre de stock       |
| Diferencia entre demanda histórica y proyectada     | Los datos explorados son históricos; la proyección vendrá en P3–P5    |

### Próximos pasos — Práctica 2

En la **Práctica 2** trabajarás con un dataset que contiene errores intencionados (valores atípicos, registros faltantes, fechas inconsistentes). Utilizarás las métricas que calculaste hoy (promedio, máximo, mínimo) como **referencia de validación** para identificar y corregir esos errores. Por eso era fundamental completar esta práctica primero: sin conocer el comportamiento normal de la demanda de cada SKU, no puedes detectar qué es un error y qué es simplemente una variación legítima.

### Recursos de referencia

- **Chopra & Meindl** — *Supply Chain Management* (Cap. 7: Demand Forecasting): fundamentos teóricos del forecasting en cadenas de suministro.
- **Ballou** — *Business Logistics/Supply Chain Management* (Cap. 9): gestión de inventarios y el rol del pronóstico.
- **Ayuda de Microsoft Excel** — Función `SUMAR.SI`: [support.microsoft.com](https://support.microsoft.com/es-es/office/funci%C3%B3n-sumar-si-169b8c99-c05c-4483-a712-1697a653039b)
- **MIT OpenCourseWare** — Supply Chain Planning (15.762): modelos cuantitativos de pronóstico aplicados a logística.

---

> 💡 **Reflexión final para el instructor:** Antes de pasar a la Práctica 2, dedica 2 minutos a preguntar al grupo: *"¿Cuál de los 8 SKUs les preocuparía más gestionar desde el punto de vista del inventario y por qué?"* Esta pregunta ancla el análisis estadístico al contexto operativo y prepara mentalmente a los participantes para la depuración de datos de la siguiente práctica.

---
LAB_END---

---

# Práctica 2 — Depuración y organización de datos para análisis de pronóstico

## 1. Metadatos

| Campo            | Detalle                                                                 |
|------------------|-------------------------------------------------------------------------|
| **Duración**     | 12 minutos                                                              |
| **Complejidad**  | Fácil                                                                   |
| **Nivel Bloom**  | Aplicar (Apply)                                                         |
| **Módulo**       | 1 — Fundamentos de Forecasting Logístico                                |
| **Práctica**     | 2 de 10                                                                 |
| **Archivo base** | `P02_Dataset_Sucio.xlsx` (proporcionado por el instructor)              |
| **Archivo salida** | `P02_Dataset_Depurado.xlsx` + tabla de log de cambios                |

---

## 2. Descripción general

En la Práctica 1 exploraste la estructura de un dataset limpio de demanda histórica. En esta práctica trabajarás con una versión del mismo dataset que contiene problemas de calidad de datos intencionalmente introducidos: valores nulos, filas duplicadas, outliers y formatos de fecha inconsistentes. Aplicarás técnicas de depuración en Excel —filtros, formato condicional, eliminación de duplicados y columnas auxiliares— para normalizar el dataset y dejarlo listo para construir el **Baseline del pronóstico**. Documentarás cada decisión de depuración en un log de cambios, conectando siempre el tratamiento estadístico con su impacto operativo en el nivel de servicio y la gestión de inventarios.

---

## 3. Objetivos de aprendizaje

Al finalizar esta práctica, podrás:

- [ ] Identificar valores nulos, duplicados, outliers y formatos inconsistentes en un dataset de ventas históricas usando herramientas nativas de Excel.
- [ ] Aplicar el tratamiento de valores nulos mediante el promedio de periodos adyacentes como valor sustituto.
- [ ] Marcar y documentar outliers en una columna auxiliar `Flag_Evento` sin eliminarlos, justificando la decisión desde la perspectiva del forecasting logístico.
- [ ] Estandarizar el formato de fechas en Excel para garantizar la consistencia del dataset.
- [ ] Registrar todas las modificaciones realizadas en una tabla de log de cambios que sirva como trazabilidad del proceso de depuración.

---

## 4. Prerequisitos

### Conocimiento previo
- Haber completado la **Práctica 1** o contar con comprensión básica de la estructura del dataset (8 SKUs × 24 meses).
- Conocimiento del concepto de **outliers** y **periodos atípicos** (tema 1.7 del módulo): eventos extraordinarios como quiebres de stock, promociones no documentadas o errores de registro que distorsionan el Baseline del pronóstico.
- Manejo básico de Excel: navegación por hojas, uso de filtros, formato condicional y funciones como `PROMEDIO`, `SI`, `ESNULO`.

### Acceso y software requerido
- Microsoft Excel 2016 o superior (se recomienda Microsoft 365).
- Archivo `P02_Dataset_Sucio.xlsx` descargado y guardado localmente.
- Permisos de escritura en la carpeta de trabajo para guardar el archivo depurado.

---

## 5. Entorno de laboratorio

### Hardware recomendado

| Componente       | Mínimo                        | Recomendado                   |
|------------------|-------------------------------|-------------------------------|
| Procesador       | Intel Core i5 / AMD Ryzen 5   | Intel Core i7 / AMD Ryzen 7   |
| RAM              | 4 GB                          | 8 GB                          |
| Almacenamiento   | 2 GB libres                   | 5 GB libres                   |
| Resolución       | 1280 × 768                    | 1920 × 1080                   |

### Software requerido

| Software         | Versión mínima   | Notas                                                    |
|------------------|------------------|----------------------------------------------------------|
| Microsoft Excel  | 2016             | Microsoft 365 recomendado para compatibilidad de fórmulas |
| Microsoft Edge / Chrome | 2024    | Para acceso a recursos complementarios si se requiere    |

### Preparación del entorno (antes de iniciar)

1. Descarga el archivo `P02_Dataset_Sucio.xlsx` desde la carpeta compartida del instructor.
2. Guárdalo en tu carpeta de trabajo local (ejemplo: `C:\Forecasting_Lab\Practica02\`).
3. Abre el archivo en Excel. **No lo modifiques todavía.**
4. Haz una copia de seguridad inmediata con el nombre `P02_Dataset_Sucio_BACKUP.xlsx` en la misma carpeta.

> ⚠️ **IMPORTANTE:** Trabaja siempre sobre el archivo original `P02_Dataset_Sucio.xlsx`. La copia de seguridad es solo de referencia. Al finalizar, guardarás el resultado como `P02_Dataset_Depurado.xlsx`.

---

## 6. Instrucciones paso a paso

---

### Paso 1 — Reconocimiento del dataset con problemas de calidad

**Objetivo:** Familiarizarte con la estructura del archivo y localizar visualmente los tipos de problemas presentes antes de aplicar cualquier corrección.

#### Instrucciones

1. Abre `P02_Dataset_Sucio.xlsx`. Observa que el archivo contiene dos hojas:
   - **`Datos_Ventas`**: dataset principal con columnas `Fecha`, `SKU`, `Descripcion`, `Unidades_Vendidas`.
   - **`Log_Cambios`**: hoja vacía con encabezados predefinidos donde registrarás cada modificación.

2. En la hoja `Datos_Ventas`, revisa la columna `Fecha` (columna A). Desplázate por las filas y observa si todas las fechas tienen el mismo formato. Busca celdas que muestren texto en lugar de fechas (por ejemplo: `"enero-23"` en lugar de `01/01/2023`).

3. Revisa la columna `Unidades_Vendidas` (columna D). Busca celdas en blanco o con valor `0` que parezcan anómalos en el contexto del historial del SKU.

4. Selecciona toda la hoja con **Ctrl + A** y observa si hay filas que parecen repetidas.

5. Sin modificar nada, anota mentalmente (o en papel) cuántos problemas detectas visualmente. En los pasos siguientes los tratarás uno por uno.

#### Resultado esperado
Tienes una visión general del dataset. Has identificado al menos: fechas con formato inconsistente en 2 filas, celdas en blanco en la columna `Unidades_Vendidas`, y sospechas de filas duplicadas.

#### Verificación
La hoja `Datos_Ventas` debe tener exactamente **196 filas de datos** (más la fila de encabezado = fila 1). Si ves un número diferente, notifica al instructor.

---

### Paso 2 — Identificación y eliminación de filas duplicadas

**Objetivo:** Eliminar registros duplicados que inflarían artificialmente la demanda histórica y distorsionarían el Baseline del pronóstico.

#### Instrucciones

1. Haz clic en cualquier celda dentro del rango de datos de la hoja `Datos_Ventas`.

2. Ve a la pestaña **Datos** en la cinta de opciones → grupo **Herramientas de datos** → haz clic en **Quitar duplicados**.

3. En el cuadro de diálogo que aparece:
   - Asegúrate de que la opción **"Mis datos tienen encabezados"** esté marcada.
   - Selecciona **todas las columnas** como criterio de comparación: `Fecha`, `SKU`, `Descripcion`, `Unidades_Vendidas`.
   - Haz clic en **Aceptar**.

4. Excel mostrará un mensaje indicando cuántas filas duplicadas fueron eliminadas. El resultado esperado es **2 filas duplicadas eliminadas**.

5. Registra este cambio en la hoja `Log_Cambios`. Completa las columnas de la tabla de log de la siguiente manera:

   | Fila_Log | Fecha_Cambio | Tipo_Problema | SKU_Afectado | Periodo_Afectado | Accion_Tomada | Valor_Original | Valor_Nuevo | Justificacion_Logistica |
   |---|---|---|---|---|---|---|---|---|
   | 1 | [fecha hoy] | Duplicado | Múltiple | Múltiple | Eliminación de 2 filas duplicadas con herramienta "Quitar duplicados" | Fila repetida | Eliminada | Los duplicados inflan artificialmente la demanda histórica y elevan el Baseline, generando sobrecompras y deterioro de la rotación de inventario |

#### Resultado esperado
El dataset ahora tiene **194 filas de datos** (196 originales − 2 duplicados). La tabla de log tiene su primera entrada completada.

#### Verificación
En la celda vacía debajo del último dato de la columna A, escribe temporalmente `=CONTARA(A2:A300)` para contar las filas con datos. El resultado debe ser **194**. Elimina esta fórmula temporal después de verificar.

---

### Paso 3 — Detección de valores nulos con formato condicional

**Objetivo:** Localizar visualmente las 3 celdas con valores en blanco en la columna `Unidades_Vendidas` para proceder a su tratamiento en el paso siguiente.

#### Instrucciones

1. Selecciona el rango completo de la columna `Unidades_Vendidas` (columna D), desde `D2` hasta la última fila con datos (aproximadamente `D195`). Puedes hacer clic en `D2` y luego presionar **Ctrl + Shift + Fin** para llegar al final.

2. Ve a la pestaña **Inicio** → grupo **Estilos** → haz clic en **Formato condicional** → **Nueva regla**.

3. En el cuadro de diálogo:
   - Selecciona **"Aplicar formato únicamente a las celdas que contengan"**.
   - En el primer desplegable elige **"Espacios en blanco"**.
   - Haz clic en **Formato...** → pestaña **Relleno** → selecciona un color llamativo, por ejemplo **amarillo** (`#FFFF00`).
   - Haz clic en **Aceptar** dos veces.

4. Las celdas vacías en la columna `Unidades_Vendidas` quedarán resaltadas en amarillo. Anota en papel o en una celda auxiliar temporal la referencia de cada celda vacía encontrada (fila y el SKU correspondiente en columna B).

5. Deberías encontrar exactamente **3 celdas vacías**. Si encuentras más o menos, revisa con el instructor.

#### Resultado esperado
Tres celdas en la columna `Unidades_Vendidas` están resaltadas en amarillo. Has anotado sus referencias exactas (ejemplo: `D45`, `D112`, `D178`) y el SKU al que corresponde cada una.

#### Verificación
Para confirmar que Excel reconoce esas celdas como vacías, en una celda auxiliar escribe:
```excel
=CONTAR.BLANCO(D2:D195)
```
El resultado debe ser **3**. Elimina la fórmula auxiliar al terminar.

---

### Paso 4 — Tratamiento de valores nulos: sustitución por promedio de periodos adyacentes

**Objetivo:** Reemplazar los valores nulos con el promedio de los meses inmediatamente anterior y posterior del mismo SKU, preservando la tendencia local de la demanda.

#### Instrucciones

1. Para cada celda vacía identificada en el paso anterior, debes calcular el valor sustituto. La fórmula general es el promedio simple entre la celda de la fila anterior y la celda de la fila siguiente **del mismo SKU**.

   > ⚠️ **Verificación previa:** Antes de aplicar la fórmula, confirma que la fila anterior y la fila siguiente corresponden al **mismo SKU** (columna B). Si el dataset está ordenado por SKU y luego por fecha, esto debería cumplirse. Si no, ordena primero por SKU y luego por Fecha.

2. Para ordenar el dataset (si es necesario): haz clic en cualquier celda del rango → **Datos** → **Ordenar** → Nivel 1: columna `SKU` (A a Z) → Agregar nivel → Nivel 2: columna `Fecha` (más antiguo a más reciente) → **Aceptar**.

3. Para cada celda vacía (ejemplo: celda `D45`), haz clic sobre ella y escribe directamente el valor calculado. **No uses una fórmula que referencie otras celdas**; escribe el valor numérico resultante para evitar dependencias circulares. Calcula el promedio de la siguiente manera:

   ```excel
   ' En una celda auxiliar temporal (por ejemplo, en columna F), escribe:
   =PROMEDIO(D44, D46)
   ```
   Copia el valor resultante (solo el número, no la fórmula) y pégalo en la celda vacía con **Pegado especial → Valores** (Ctrl + Alt + V → V → Enter).

4. Repite el proceso para las 3 celdas vacías. Elimina las fórmulas auxiliares de la columna F al terminar.

5. Registra cada sustitución en la hoja `Log_Cambios`. Ejemplo para una de las tres entradas:

   | Fila_Log | Fecha_Cambio | Tipo_Problema | SKU_Afectado | Periodo_Afectado | Accion_Tomada | Valor_Original | Valor_Nuevo | Justificacion_Logistica |
   |---|---|---|---|---|---|---|---|---|
   | 2 | [fecha hoy] | Valor nulo | SKU-003 | Mar-2023 | Sustitución por promedio de periodos adyacentes (Feb-2023 y Abr-2023) | Vacío | 412 | Un nulo en el historial genera un hueco en el Baseline que puede interpretarse como quiebre de stock real, distorsionando el cálculo del stock de seguridad |

#### Resultado esperado
Las 3 celdas anteriormente vacías ahora contienen valores numéricos calculados como promedio de sus periodos adyacentes. El formato condicional amarillo ya no se activa en esas celdas (si el formato condicional detectaba "espacios en blanco", al tener valor ahora no se resaltan).

#### Verificación
Vuelve a ejecutar la fórmula `=CONTAR.BLANCO(D2:D195)` en una celda auxiliar. El resultado ahora debe ser **0**. Elimina la fórmula al terminar.

---

### Paso 5 — Identificación y marcado de outliers con columna auxiliar `Flag_Evento`

**Objetivo:** Detectar los valores extremos (outliers) en el dataset, marcarlos en una columna auxiliar sin eliminarlos, y documentar si corresponden a eventos extraordinarios o errores de registro.

> **Concepto clave:** En forecasting logístico, los outliers **no se eliminan automáticamente**. Primero se investiga su origen: si corresponden a un evento real (promoción, quiebre de stock, catástrofe), se documentan para ser tratados metodológicamente en el Baseline. Si son errores de registro, se corrigen o excluyen con trazabilidad. Eliminar un outlier sin documentarlo es una práctica incorrecta que destruye información valiosa.

#### Instrucciones

1. **Agrega la columna auxiliar `Flag_Evento`:** Haz clic en el encabezado de la columna E (la primera columna vacía después de `Unidades_Vendidas`). Escribe `Flag_Evento` como encabezado en la celda `E1`.

2. **Calcula los límites estadísticos para detectar outliers:** Usarás el método del rango intercuartílico (IQR) o, de forma simplificada para este ejercicio, identificarás valores que se desvíen más de **2 veces la desviación estándar** del promedio del SKU correspondiente.

   Para simplificar en el tiempo disponible, el instructor habrá indicado previamente cuáles son las **5 filas con outliers** en el dataset (3 por quiebre de stock con demanda anormalmente baja, y 2 por promoción no documentada con demanda anormalmente alta). Localiza esas filas.

   > Si no tienes la lista del instructor, puedes detectarlos visualmente: ordena el dataset por `Unidades_Vendidas` de mayor a menor y busca los valores más extremos en ambos extremos.

3. **Aplica la fórmula `SI` en la columna `Flag_Evento`** para todas las filas. En la celda `E2` escribe:

   ```excel
   =SI(D2="","NULO_PENDIENTE","OK")
   ```

   Esto establece un valor base. Luego, para las filas que identificaste como outliers, **sobreescribe manualmente** el valor de la celda `Flag_Evento` con una de las siguientes etiquetas estándar:

   | Etiqueta              | Cuándo usarla                                                                 |
   |-----------------------|-------------------------------------------------------------------------------|
   | `QUIEBRE_STOCK`       | Demanda anormalmente baja causada por falta de producto disponible            |
   | `PROMO_NO_DOCUMENTADA`| Demanda anormalmente alta causada por una promoción no registrada formalmente |
   | `ERROR_REGISTRO`      | Valor que claramente no corresponde a la realidad (ej: 1 unidad para un SKU de alta rotación) |
   | `OK`                  | Valor normal dentro del comportamiento esperado del SKU                       |

4. Para las filas identificadas como outliers, escribe directamente la etiqueta correspondiente en la celda de la columna `Flag_Evento`. **No elimines ni modifiques el valor de `Unidades_Vendidas`.**

5. Aplica **formato condicional** a la columna `Flag_Evento` para resaltar visualmente los outliers:
   - Selecciona el rango `E2:E195`.
   - **Inicio** → **Formato condicional** → **Resaltar reglas de celdas** → **Texto que contiene...**
   - Escribe `QUIEBRE_STOCK` → elige formato de relleno **rojo claro**.
   - Repite el proceso para `PROMO_NO_DOCUMENTADA` → relleno **naranja claro**.
   - Repite para `ERROR_REGISTRO` → relleno **rojo oscuro / texto blanco**.

6. Registra cada outlier en la hoja `Log_Cambios`. Ejemplo:

   | Fila_Log | Fecha_Cambio | Tipo_Problema | SKU_Afectado | Periodo_Afectado | Accion_Tomada | Valor_Original | Valor_Nuevo | Justificacion_Logistica |
   |---|---|---|---|---|---|---|---|---|
   | 5 | [fecha hoy] | Outlier - Demanda baja | SKU-001 | Jun-2023 | Marcado como QUIEBRE_STOCK en Flag_Evento. Valor conservado. | 18 | Sin cambio | Demanda de 18 uds. es 87% inferior al promedio del SKU (138 uds.). Corresponde a periodo de quiebre de stock confirmado. Incluir en Baseline distorsionaría a la baja el pronóstico y elevaría el riesgo de subcompra. |
   | 6 | [fecha hoy] | Outlier - Demanda alta | SKU-005 | Nov-2022 | Marcado como PROMO_NO_DOCUMENTADA en Flag_Evento. Valor conservado. | 892 | Sin cambio | Demanda de 892 uds. es 340% del promedio del SKU (263 uds.). Sin registro de campaña en el periodo. Se excluirá del Baseline en Práctica 3 y se analizará como evento extraordinario. |

#### Resultado esperado
La columna `Flag_Evento` está completa para todas las filas. Las filas con outliers están marcadas con las etiquetas correspondientes y resaltadas con colores diferenciados. Los valores originales de `Unidades_Vendidas` no han sido modificados. El log de cambios tiene una entrada por cada outlier identificado.

#### Verificación
Aplica un filtro a la columna `Flag_Evento` y selecciona solo los valores distintos de `OK`. Deberías ver exactamente **5 filas marcadas** (3 como `QUIEBRE_STOCK` y 2 como `PROMO_NO_DOCUMENTADA` o `ERROR_REGISTRO`, según la distribución indicada por el instructor).

---

### Paso 6 — Estandarización del formato de fechas

**Objetivo:** Corregir las 2 filas con fechas en formato de texto inconsistente (`"enero-23"`, `"feb-23"`) para que Excel las reconozca como valores de fecha válidos, garantizando que las funciones de series de tiempo funcionen correctamente en prácticas posteriores.

#### Instrucciones

1. Aplica formato condicional a la columna `Fecha` (columna A) para identificar las celdas que contienen texto en lugar de fechas:
   - Selecciona el rango `A2:A195`.
   - **Inicio** → **Formato condicional** → **Nueva regla** → **"Usar una fórmula que determine las celdas para aplicar formato"**.
   - Escribe la fórmula:
     ```excel
     =ESTEXTO(A2)
     ```
   - Elige un relleno de color **morado claro** para distinguirlo de los otros formatos condicionales.
   - Haz clic en **Aceptar**.

2. Las celdas con fechas en formato texto quedarán resaltadas en morado. Deberías ver **2 celdas resaltadas**.

3. Para convertir cada celda de texto a fecha:
   - Haz clic en la primera celda resaltada (ejemplo: contiene `"enero-23"`).
   - Borra el contenido actual.
   - Escribe la fecha correcta en formato `DD/MM/AAAA`. Ejemplo: si el texto decía `"enero-23"` y corresponde a enero de 2023, escribe `01/01/2023`.
   - Presiona **Enter**.
   - Confirma que Excel reconoció la entrada como fecha: la celda debe alinearse a la derecha (las fechas se alinean a la derecha por defecto en Excel; el texto se alinea a la izquierda).

4. Repite el proceso para la segunda celda resaltada.

5. Si Excel no reconoce la fecha automáticamente (la celda sigue alineada a la izquierda), puedes forzar la conversión con la función `FECHANUMERO`:
   - En una celda auxiliar temporal (columna F), escribe:
     ```excel
     =FECHANUMERO("01/01/2023")
     ```
   - Copia el valor resultante y pégalo como **Pegado especial → Valores** en la celda problemática.
   - Luego aplica el formato de fecha: **Inicio** → **Número** → selecciona **Fecha corta** o **Fecha larga** según el formato del resto del dataset.

6. Verifica que el formato condicional morado ya no se activa en esas celdas (al ser ahora fechas válidas, `ESTEXTO()` devuelve `FALSO` y el formato no aplica).

7. Registra los cambios en `Log_Cambios`:

   | Fila_Log | Fecha_Cambio | Tipo_Problema | SKU_Afectado | Periodo_Afectado | Accion_Tomada | Valor_Original | Valor_Nuevo | Justificacion_Logistica |
   |---|---|---|---|---|---|---|---|---|
   | 7 | [fecha hoy] | Formato fecha incorrecto | SKU-002 | Ene-2023 | Conversión de texto "enero-23" a fecha 01/01/2023 con formato DD/MM/AAAA | "enero-23" (texto) | 01/01/2023 (fecha) | Las fechas en formato texto impiden que Excel construya series de tiempo correctas. Las funciones de pronóstico (PRONOSTICO.LINEAL, gráficos de tendencia) requieren valores de fecha numéricos. |

#### Resultado esperado
Las 2 celdas con fechas en formato texto ahora contienen valores de fecha válidos reconocidos por Excel. El formato condicional morado no se activa en ninguna celda del rango `A2:A195`. El dataset tiene fechas uniformes en toda la columna A.

#### Verificación
En una celda auxiliar, escribe:
```excel
=SUMAPRODUCTO(--(ESTEXTO(A2:A195)))
```
El resultado debe ser **0** (ninguna celda de fecha es texto). Elimina la fórmula auxiliar al terminar.

---

### Paso 7 — Revisión final y guardado del dataset depurado

**Objetivo:** Realizar una revisión integral del dataset depurado, confirmar que todos los problemas han sido tratados y guardar el archivo con el nombre correcto para que sirva como insumo de las prácticas siguientes.

#### Instrucciones

1. **Revisión de integridad final.** En una hoja nueva llamada `Resumen_Validacion`, crea una pequeña tabla de verificación con las siguientes fórmulas:

   | Verificación | Fórmula | Resultado esperado |
   |---|---|---|
   | Filas de datos totales | `=CONTARA(Datos_Ventas!A2:A300)-1` | 194 |
   | Celdas nulas en Unidades_Vendidas | `=CONTAR.BLANCO(Datos_Ventas!D2:D195)` | 0 |
   | Fechas en formato texto | `=SUMAPRODUCTO(--(ESTEXTO(Datos_Ventas!A2:A195)))` | 0 |
   | Registros con Flag distinto de OK | `=CONTAR.SI(Datos_Ventas!E2:E195,"<>OK")` | 5 |

   Escribe estas fórmulas en las celdas B2:B5 de la hoja `Resumen_Validacion`, con las etiquetas descriptivas en las celdas A2:A5.

2. Confirma que todos los valores en la columna `Resultado esperado` coinciden con los resultados reales de tus fórmulas. Si algún valor no coincide, regresa al paso correspondiente y revisa.

3. Revisa la hoja `Log_Cambios` y confirma que tiene al menos **7 entradas** (1 por duplicados + 3 por nulos + al menos 2 por outliers + 2 por fechas). Si falta alguna entrada, complétala ahora.

4. Guarda el archivo con el nombre `P02_Dataset_Depurado.xlsx`:
   - **Archivo** → **Guardar como** → navega a tu carpeta de trabajo → escribe el nombre `P02_Dataset_Depurado` → tipo de archivo: **Libro de Excel (.xlsx)** → **Guardar**.

5. Cierra el archivo `P02_Dataset_Sucio.xlsx` sin guardar cambios adicionales (el backup ya está guardado).

#### Resultado esperado
El archivo `P02_Dataset_Depurado.xlsx` está guardado en tu carpeta de trabajo con:
- Hoja `Datos_Ventas`: 194 filas de datos, sin nulos, sin duplicados, con columna `Flag_Evento` completa y fechas estandarizadas.
- Hoja `Log_Cambios`: tabla completa con al menos 7 entradas documentadas.
- Hoja `Resumen_Validacion`: tabla de verificación con todos los indicadores en verde (valores esperados confirmados).

#### Verificación
Cierra y vuelve a abrir el archivo `P02_Dataset_Depurado.xlsx`. Navega a la hoja `Resumen_Validacion` y confirma que las fórmulas siguen mostrando los valores correctos. Esto confirma que el archivo fue guardado correctamente y está listo para la Práctica 3.

---

## 7. Validación y prueba integral

Al finalizar todos los pasos, verifica los siguientes criterios de aceptación antes de declarar la práctica completada:

| # | Criterio de validación | Cómo verificarlo | Estado |
|---|---|---|---|
| 1 | El dataset tiene exactamente 194 filas de datos | `=CONTARA(A2:A300)` = 194 en hoja `Datos_Ventas` | ☐ |
| 2 | No hay celdas vacías en `Unidades_Vendidas` | `=CONTAR.BLANCO(D2:D195)` = 0 | ☐ |
| 3 | No hay fechas en formato texto | `=SUMAPRODUCTO(--(ESTEXTO(A2:A195)))` = 0 | ☐ |
| 4 | Exactamente 5 registros tienen `Flag_Evento` distinto de `OK` | `=CONTAR.SI(E2:E195,"<>OK")` = 5 | ☐ |
| 5 | Los valores de `Unidades_Vendidas` en filas con outliers no fueron modificados | Comparar con `P02_Dataset_Sucio_BACKUP.xlsx` en esas filas | ☐ |
| 6 | El log de cambios tiene al menos 7 entradas con justificación logística | Revisar hoja `Log_Cambios` manualmente | ☐ |
| 7 | El archivo está guardado como `P02_Dataset_Depurado.xlsx` | Verificar nombre en explorador de archivos | ☐ |

> **Criterio de aprobación:** Todos los 7 criterios deben estar marcados como cumplidos para que el archivo sea válido como insumo de la Práctica 3.

---

## 8. Resolución de problemas

### Problema 1 — Excel no elimina los duplicados esperados o elimina filas incorrectas

**Síntoma:** Al ejecutar "Quitar duplicados", Excel reporta 0 duplicados eliminados (cuando deberían ser 2), o elimina más filas de las esperadas.

**Causa probable:** Las filas duplicadas pueden tener diferencias invisibles en alguna columna, como espacios adicionales al inicio o al final de un valor de texto en la columna `SKU` o `Descripcion` (por ejemplo: `"SKU-001"` vs `"SKU-001 "` con un espacio al final). Excel las considera diferentes y no las detecta como duplicados. Alternativamente, si se seleccionaron columnas incorrectas como criterio, puede eliminar filas que no son verdaderos duplicados.

**Solución:**
1. Antes de ejecutar "Quitar duplicados", aplica la función `ESPACIOS()` (TRIM en inglés) a las columnas de texto para eliminar espacios extra. En una columna auxiliar temporal (columna F), escribe `=ESPACIOS(B2)` y copia hacia abajo. Luego copia esos valores y pégalos como "Valores" sobre la columna B original.
2. Repite para la columna `Descripcion` (columna C).
3. Vuelve a ejecutar "Quitar duplicados". Ahora debería detectar las 2 filas duplicadas correctamente.
4. Si el problema persiste, identifica manualmente las filas duplicadas usando la fórmula `=CONTAR.SI.CONJUNTO($A$2:$A$195,A2,$B$2:$B$195,B2,$D$2:$D$195,D2)` en una columna auxiliar. Las filas con valor > 1 son duplicadas; elimina manualmente la segunda ocurrencia.

---

### Problema 2 — Las fechas no se convierten correctamente y siguen mostrándose como texto

**Síntoma:** Después de escribir la fecha en formato `DD/MM/AAAA` en la celda, Excel sigue mostrando el valor alineado a la izquierda (indicando que lo interpreta como texto) o muestra un número de serie en lugar de la fecha formateada.

**Causa probable:** La configuración regional del sistema operativo o de Excel puede estar usando el formato de fecha `MM/DD/AAAA` (formato estadounidense) en lugar de `DD/MM/AAAA` (formato latinoamericano). Cuando escribes `01/01/2023`, Excel puede interpretarlo correctamente en ambos casos (1 de enero), pero fechas como `15/06/2023` fallarían porque no existe el mes 15 en formato MM/DD. Alternativamente, el formato de la celda puede estar configurado explícitamente como "Texto", lo que impide que Excel interprete cualquier entrada como fecha.

**Solución:**
1. Selecciona la celda problemática → **Inicio** → grupo **Número** → desplegable de formato → selecciona **General** (para quitar el formato "Texto"). Luego vuelve a escribir la fecha.
2. Si el problema persiste por configuración regional, usa la función `FECHA()` para construir la fecha explícitamente. En una celda auxiliar escribe `=FECHA(2023,1,1)` para el 1 de enero de 2023. Copia el valor resultante y pégalo como "Pegado especial → Valores" en la celda original. Luego aplica el formato de fecha deseado desde **Inicio → Número → Fecha corta**.
3. Para verificar que el valor es una fecha válida, la celda debe alinearse a la derecha y la fórmula `=ESTEXTO(celda)` debe devolver `FALSO`.

---

## 9. Limpieza del entorno

Al finalizar la práctica, realiza las siguientes acciones de limpieza para mantener el entorno ordenado y preparar el espacio para la Práctica 3:

1. **Elimina todas las fórmulas auxiliares temporales** que hayas creado en columnas F, G u otras columnas fuera del esquema oficial del dataset. Verifica que el dataset solo tenga las columnas: `A: Fecha`, `B: SKU`, `C: Descripcion`, `D: Unidades_Vendidas`, `E: Flag_Evento`.

2. **Elimina la hoja `Resumen_Validacion`** si el instructor indica que no es necesaria como entregable (o mantenla si forma parte del entregable). Por defecto, **consérvala**: sirve como evidencia de validación.

3. **Elimina los formatos condicionales de trabajo** que ya no son necesarios (los colores amarillo para nulos y morado para fechas texto) para evitar confusión visual en prácticas posteriores. Mantén únicamente el formato condicional de colores en la columna `Flag_Evento` (rojo para `QUIEBRE_STOCK`, naranja para `PROMO_NO_DOCUMENTADA`), ya que es parte del entregable.
   - Para eliminar un formato condicional: selecciona el rango → **Inicio** → **Formato condicional** → **Borrar reglas** → **Borrar reglas de las celdas seleccionadas**.

4. **Guarda el archivo una vez más** después de la limpieza: **Ctrl + S**.

5. **Entrega el archivo** `P02_Dataset_Depurado.xlsx` al instructor según el método definido para el curso (carpeta compartida, correo electrónico, plataforma LMS).

6. Mantén el archivo `P02_Dataset_Sucio_BACKUP.xlsx` en tu carpeta local como referencia. **No lo compartas ni lo uses en prácticas posteriores.**

---

## 10. Resumen y conexión con el contexto logístico

### Lo que aprendiste en esta práctica

En esta práctica aplicaste el proceso completo de **depuración de datos para forecasting logístico**:

- **Duplicados:** Aprendiste que filas duplicadas inflan artificialmente la demanda histórica, lo que eleva el Baseline del pronóstico y genera sobrecompras. El resultado directo en operaciones es un aumento del inventario promedio, menor rotación y capital inmovilizado innecesariamente.

- **Valores nulos:** Comprendiste que un hueco en el historial de demanda puede interpretarse erróneamente como un período de demanda cero, lo que subestima el Baseline y aumenta el riesgo de quiebres de stock en el futuro. La sustitución por promedio de periodos adyacentes es una técnica conservadora que preserva la tendencia local.

- **Outliers:** Internalizaste que los outliers no se eliminan automáticamente. Un quiebre de stock que generó demanda artificialmente baja **no representa la demanda real del cliente**; si se incluye en el Baseline sin marcar, el pronóstico subestimará la demanda real y el nivel de servicio caerá. Una promoción no documentada que generó demanda artificialmente alta tampoco es representativa del comportamiento regular; si se incluye sin marcar, el pronóstico sobreestimará la demanda habitual. La columna `Flag_Evento` es la herramienta que permite tratar estos casos metodológicamente en pasos posteriores.

- **Formato de fechas:** Entendiste que las funciones de series de tiempo en Excel (y en Power BI) requieren valores de fecha numéricos válidos. Fechas en formato texto rompen las funciones de pronóstico, los gráficos de tendencia y los cálculos de estacionalidad.

### Conexión con el Baseline del pronóstico

La estructura metodológica del pronóstico que verás en prácticas posteriores se define como:

> **Pronóstico = Baseline + Estacionalidad + Iniciativas**

El **Baseline** es la demanda "limpia y regular" del SKU, libre de eventos extraordinarios y errores de registro. Un dataset mal depurado produce un Baseline contaminado que arrastra todos sus errores a las tres componentes del pronóstico. Por eso, la depuración no es un paso burocrático: es la **base de calidad** sobre la que descansa toda la precisión del sistema de forecasting.

### Recursos adicionales

- **Excel Help — Quitar duplicados:** [support.microsoft.com/es-es/office/buscar-y-quitar-duplicados](https://support.microsoft.com/es-es/office/buscar-y-quitar-duplicados-00e35bea-b46a-4d5d-b28e-66a552dc138d)
- **Excel Help — Formato condicional:** [support.microsoft.com/es-es/office/usar-el-formato-condicional](https://support.microsoft.com/es-es/office/usar-el-formato-condicional-para-resaltar-informaci%C3%B3n-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
- **Chopra & Meindl — Supply Chain Management** (Cap. 7: Demand Forecasting in a Supply Chain): referencia conceptual sobre el impacto de la calidad de datos en la precisión del pronóstico.
- **Archivo de checkpoint:** Si no completaste algún paso, solicita al instructor el archivo `P02_Checkpoint_Paso[N].xlsx` correspondiente al paso donde te quedaste.

> 🔗 **Siguiente práctica:** En la **Práctica 3** utilizarás el archivo `P02_Dataset_Depurado.xlsx` que generaste hoy para construir el Baseline del pronóstico, calculando promedios móviles e identificando la componente de tendencia de cada SKU.

---

---

# Práctica 3 — Identificación Visual de Tendencias Mediante Gráficos

## Metadatos

| Campo            | Detalle                                      |
|------------------|----------------------------------------------|
| **Duración**     | 12 minutos                                   |
| **Complejidad**  | Fácil                                        |
| **Nivel Bloom**  | Aplicar                                      |
| **Práctica N°**  | 3 de 10                                      |
| **Archivo base** | `Dataset_Depurado_P2.xlsx` (salida de Práctica 2) |

---

## Descripción General

En esta práctica construirás gráficos de línea en Excel para visualizar el comportamiento histórico de la demanda mensual de los SKUs del dataset depurado. Agregarás líneas de tendencia lineal y de media móvil para identificar la dirección general de la demanda, y anotarás directamente en los gráficos los eventos relevantes (picos, valles, cambios de tendencia). El producto final es un **dashboard visual básico con 3 gráficos anotados** que servirá como insumo de análisis para las prácticas posteriores.

Esta práctica conecta directamente con el concepto central del forecasting logístico: antes de proyectar la demanda futura, es imprescindible entender visualmente qué ha ocurrido en el pasado —qué tendencia sigue la demanda, qué tan estable es y en qué meses se producen los picos que podrían generar quiebres de stock o sobreinventario.

---

## Objetivos de Aprendizaje

Al finalizar esta práctica, serás capaz de:

- [ ] Construir gráficos de línea en Excel para visualizar 24 meses de historial de demanda por SKU.
- [ ] Agregar líneas de tendencia lineal (con ecuación y R²) y de media móvil de 3 periodos para identificar la dirección general y suavizar la variabilidad.
- [ ] Interpretar visualmente los componentes de tendencia, estacionalidad y variabilidad en los gráficos, relacionándolos con implicaciones operativas de inventario y picking.
- [ ] Crear un gráfico comparativo de múltiples SKUs y anotar los meses críticos directamente sobre los gráficos.

---

## Prerrequisitos

### Conocimiento previo
- Haber completado la **Práctica 2** y tener disponible el archivo `Dataset_Depurado_P2.xlsx` con los datos corregidos de 8 SKUs (24 meses).
- Conocimiento básico de inserción de gráficos en Excel (pestaña **Insertar → Gráficos**).
- Comprensión de los componentes del comportamiento de la demanda: tendencia, estacionalidad y variabilidad (Lección 1.1 y tema 1.3).
- Familiaridad con los conceptos de **rotación de inventario**, **nivel de servicio** y **quiebres de stock** revisados en la Lección 1.1.

### Acceso requerido
- Computadora con **Microsoft Excel 2016 o superior** (Microsoft 365 recomendado).
- Archivo `Dataset_Depurado_P2.xlsx` abierto y disponible.
- Si no completaste la Práctica 2, solicita al instructor el archivo de checkpoint correspondiente.

---

## Entorno de Laboratorio

### Hardware recomendado

| Componente        | Mínimo                          | Recomendado                     |
|-------------------|---------------------------------|---------------------------------|
| Procesador        | Intel Core i5 / AMD Ryzen 5     | Intel Core i7 / AMD Ryzen 7     |
| RAM               | 4 GB                            | 8 GB                            |
| Resolución        | 1280 × 768                      | 1920 × 1080                     |
| Almacenamiento    | 2 GB libres                     | 2 GB libres                     |

### Software requerido

| Software                  | Versión mínima       | Notas                                          |
|---------------------------|----------------------|------------------------------------------------|
| Microsoft Excel           | 2016                 | Microsoft 365 recomendado                      |
| Archivo de práctica       | `Dataset_Depurado_P2.xlsx` | Proporcionado por el instructor o generado en Práctica 2 |

### Preparación del entorno

1. Abre Excel y carga el archivo `Dataset_Depurado_P2.xlsx`.
2. Verifica que la hoja `Datos_Depurados` contiene las columnas: `Mes`, `Periodo` (1–24), y una columna de demanda por cada SKU (`SKU-A` a `SKU-H`).
3. Crea una hoja nueva llamada `Gráficos_P3` donde construirás todos los gráficos de esta práctica.
4. Guarda una copia del archivo con el nombre `Dataset_Graficos_P3.xlsx` antes de comenzar.

---

## Instrucciones Paso a Paso

---

### Paso 1 — Identificar los SKUs de trabajo y preparar la columna de media móvil

**Objetivo:** Seleccionar el SKU de mayor volumen (para el Gráfico 1) y el SKU de mayor variabilidad (para el Gráfico 2), y calcular la media móvil de 3 periodos para ambos.

#### Instrucciones

1. En la hoja `Datos_Depurados`, revisa los totales de demanda por SKU. Identifica:
   - **SKU de mayor volumen:** el que tiene la suma más alta en los 24 meses (usaremos este como **SKU_Alto**).
   - **SKU de mayor variabilidad:** el que tiene el coeficiente de variación más alto calculado en la Práctica 2 (usaremos este como **SKU_Variable**). Si no tienes ese dato, elige visualmente el que tenga los valores más irregulares.

   > **Nota del instructor:** En el dataset de referencia, `SKU-A` suele ser el de mayor volumen y `SKU-F` el de mayor variabilidad. Ajusta según tus datos reales.

2. Ve a la hoja `Gráficos_P3` y crea una **tabla auxiliar** a partir de la columna `A`, con el siguiente formato:

   | Col A       | Col B              | Col C                        | Col D              | Col E                        |
   |-------------|---------------------|------------------------------|--------------------|------------------------------|
   | `Periodo`   | `SKU_Alto (demanda)` | `MM3_SKU_Alto`               | `SKU_Variable (demanda)` | `MM3_SKU_Variable`     |
   | 1           | (valor)             | —                            | (valor)            | —                            |
   | 2           | (valor)             | —                            | (valor)            | —                            |
   | 3           | (valor)             | `=PROMEDIO(B2:B4)`           | (valor)            | `=PROMEDIO(D2:D4)`           |
   | 4           | (valor)             | `=PROMEDIO(B3:B5)`           | (valor)            | `=PROMEDIO(D3:D5)`           |
   | …           | …                   | …                            | …                  | …                            |
   | 24          | (valor)             | `=PROMEDIO(B23:B25)`         | (valor)            | `=PROMEDIO(D23:D25)`         |

3. Para llenar las columnas B y D, vincula los datos desde la hoja `Datos_Depurados`:
   - En `B2` escribe: `=Datos_Depurados!C2` (reemplaza `C` por la columna real de tu SKU_Alto).
   - Arrastra la fórmula hasta la fila 25 (periodo 24).
   - Repite para la columna D con el SKU_Variable.

4. En la columna C (MM3_SKU_Alto), deja las celdas C2 y C3 **vacías** (la media móvil de 3 periodos requiere al menos 3 valores). A partir de C4:
   ```
   =PROMEDIO(B2:B4)
   ```
   Arrastra hasta C25.

5. Repite el mismo proceso en la columna E para `MM3_SKU_Variable`.

6. Etiqueta la fila 1 con los encabezados correspondientes. Tu tabla debe tener 25 filas (encabezado + 24 periodos) y 5 columnas.

**Salida esperada:** Una tabla auxiliar en la hoja `Gráficos_P3` con los datos de demanda y media móvil de 3 periodos para los dos SKUs seleccionados. Las celdas C2, C3, E2 y E3 deben estar vacías (sin valor).

**Verificación:** Confirma que la suma de la columna B en `Gráficos_P3` coincide con la suma de la misma columna en `Datos_Depurados`. Si difieren, revisa las referencias de celda en el paso 3.

---

### Paso 2 — Construir el Gráfico 1: SKU de mayor volumen con línea de tendencia lineal

**Objetivo:** Crear un gráfico de línea para el SKU de mayor volumen que muestre los 24 meses de historial, agregar la línea de tendencia lineal con ecuación y R² visible.

#### Instrucciones

1. En la hoja `Gráficos_P3`, selecciona el rango de datos del SKU_Alto: **columnas A y B** (Periodo + demanda), filas 1 a 25 (incluyendo encabezados).

2. Ve a la pestaña **Insertar** → grupo **Gráficos** → **Gráfico de líneas** → selecciona **Línea con marcadores**.

3. Excel insertará el gráfico en la hoja. Haz clic derecho sobre el gráfico y selecciona **Mover gráfico** → **Objeto en: Gráficos_P3**. Reubica el gráfico en el rango aproximado `G1:P14` de la hoja.

4. **Agrega la serie de media móvil al gráfico:**
   - Haz clic derecho sobre el área del gráfico → **Seleccionar datos**.
   - Haz clic en **Agregar** (en la sección "Entradas de leyenda").
   - En **Nombre de la serie**, escribe: `Media Móvil 3P`.
   - En **Valores de la serie**, selecciona el rango C2:C25 de tu tabla auxiliar.
   - Haz clic en **Aceptar**.

   > La serie de media móvil aparecerá ahora como una segunda línea en el gráfico. Visualmente verás cómo suaviza los picos y valles de la demanda real.

5. **Agrega la línea de tendencia lineal:**
   - Haz clic sobre la línea de la serie principal (demanda real del SKU_Alto) para seleccionarla.
   - Haz clic derecho → **Agregar línea de tendencia**.
   - En el panel derecho "Formato de línea de tendencia":
     - Selecciona tipo: **Lineal**.
     - Marca la casilla **Presentar ecuación en el gráfico**.
     - Marca la casilla **Presentar el valor R cuadrado en el gráfico**.
   - Cierra el panel.

6. **Da formato al gráfico:**
   - Título del gráfico: `Demanda Histórica — [Nombre del SKU_Alto] (24 meses)`.
   - Eje horizontal: etiqueta `Periodo (mes)`.
   - Eje vertical: etiqueta `Unidades demandadas`.
   - Cambia el color de la línea de media móvil a **naranja** y hazla discontinua para distinguirla visualmente de la demanda real.
   - Cambia el color de la línea de tendencia a **rojo** y hazla punteada.

7. **Agrega anotaciones directamente en el gráfico:**
   - Identifica visualmente el mes de mayor demanda (pico) y el de menor demanda (valle) en la línea de demanda real.
   - Ve a la pestaña **Insertar** → **Cuadro de texto**. Dibuja un cuadro de texto pequeño dentro del área del gráfico, apuntando al pico, y escribe: `Pico: Mes [N] — ¿Evento estacional?`.
   - Repite para el valle: `Valle: Mes [N] — Revisar causa`.

**Salida esperada:** Un gráfico de línea con tres series visibles: (1) demanda real con marcadores, (2) media móvil de 3 periodos en naranja discontinuo, (3) línea de tendencia lineal en rojo punteado. La ecuación de la tendencia (ej. `y = 12.5x + 430`) y el valor R² (ej. `R² = 0.72`) deben ser visibles en el área del gráfico. Dos cuadros de texto anotan el pico y el valle.

**Verificación:** 
- ¿La ecuación de la tendencia muestra una pendiente positiva, negativa o cercana a cero? Anota mentalmente este dato: lo necesitarás en el Paso 5.
- ¿El valor R² es mayor a 0.5? Si es así, la tendencia lineal explica más del 50% del comportamiento de la demanda. Si es menor a 0.3, la demanda de este SKU tiene baja tendencia lineal y alta variabilidad aleatoria.

---

### Paso 3 — Construir el Gráfico 2: SKU de alta variabilidad con contraste visual

**Objetivo:** Replicar el ejercicio del Paso 2 para el SKU de mayor variabilidad, identificando visualmente las diferencias de comportamiento respecto al SKU_Alto.

#### Instrucciones

1. En la hoja `Gráficos_P3`, selecciona el rango de datos del SKU_Variable: **columnas A y D** (Periodo + demanda), filas 1 a 25. Para seleccionar columnas no contiguas, mantén presionada la tecla **Ctrl** mientras seleccionas.

2. Ve a **Insertar** → **Gráfico de líneas** → **Línea con marcadores**. Mueve el gráfico al rango aproximado `G16:P30`.

3. **Agrega la serie de media móvil:**
   - Haz clic derecho sobre el gráfico → **Seleccionar datos** → **Agregar**.
   - Nombre: `Media Móvil 3P`.
   - Valores: columna E (MM3_SKU_Variable), filas 2 a 25.
   - Acepta.

4. **Agrega la línea de tendencia lineal:**
   - Selecciona la serie de demanda real del SKU_Variable.
   - Clic derecho → **Agregar línea de tendencia** → **Lineal**.
   - Activa: **Presentar ecuación** y **Presentar R²**.

5. **Da formato al gráfico:**
   - Título: `Demanda Histórica — [Nombre del SKU_Variable] (24 meses)`.
   - Misma convención de colores que el Gráfico 1 (naranja para media móvil, rojo para tendencia).
   - Eje vertical: asegúrate de que la escala refleje la amplitud real de la variabilidad (no la ajustes manualmente; deja que Excel la calcule automáticamente).

6. **Agrega anotaciones:**
   - Identifica al menos **2 meses** donde la demanda real se aleja significativamente de la media móvil (más de 20% por encima o por debajo).
   - Inserta cuadros de texto para cada uno: `Mes [N]: Desviación alta — ¿Promoción / evento?`.
   - Si observas un patrón que se repite cada cierto número de meses (por ejemplo, picos en los meses 6, 12, 18, 24), anota: `Posible patrón estacional`.

7. **Comparación visual rápida:**
   - Observa ambos gráficos (Gráfico 1 y Gráfico 2) lado a lado.
   - Responde mentalmente: ¿Cuál SKU tiene una tendencia más clara? ¿Cuál tiene mayor variabilidad? ¿Cuál representaría mayor riesgo de quiebre de stock si no se pronostica correctamente?

**Salida esperada:** Un segundo gráfico de línea con la misma estructura que el Gráfico 1, pero mostrando el comportamiento del SKU de alta variabilidad. La media móvil debe verse notablemente más "suave" que la línea de demanda real, evidenciando el efecto de suavizamiento. El valor R² de la tendencia lineal probablemente será más bajo que en el Gráfico 1 (indicando menor predictibilidad lineal).

**Verificación:** Compara los valores R² de ambos gráficos:
- Gráfico 1 (SKU_Alto): R² = ___
- Gráfico 2 (SKU_Variable): R² = ___
- El SKU_Variable debe tener un R² **menor**, confirmando que su demanda es menos predecible mediante una tendencia lineal simple.

---

### Paso 4 — Construir el Gráfico 3: Panel comparativo de múltiples SKUs

**Objetivo:** Crear un gráfico que muestre en un solo panel la demanda de **4 SKUs simultáneamente** para identificar diferencias de comportamiento entre productos.

#### Instrucciones

1. En la hoja `Gráficos_P3`, crea una segunda tabla auxiliar a partir de la columna `R` (o en un área libre de la hoja), con el siguiente formato:

   | Col R      | Col S     | Col T     | Col U     | Col V     |
   |------------|-----------|-----------|-----------|-----------|
   | `Periodo`  | `SKU-A`   | `SKU-B`   | `SKU-C`   | `SKU-D`   |
   | 1          | (valor)   | (valor)   | (valor)   | (valor)   |
   | …          | …         | …         | …         | …         |
   | 24         | (valor)   | (valor)   | (valor)   | (valor)   |

   > Elige 4 SKUs que representen comportamientos distintos: uno de alto volumen estable, uno con tendencia creciente, uno estacional y uno con alta variabilidad. Vincula los datos desde `Datos_Depurados` usando referencias como `=Datos_Depurados!C2`.

2. Selecciona el rango completo de la segunda tabla auxiliar (columnas R a V, filas 1 a 25).

3. Ve a **Insertar** → **Gráfico de líneas** → **Línea** (sin marcadores, para que el gráfico sea más limpio con 4 series).

4. Mueve el gráfico al rango aproximado `A32:P52` (debajo de los dos gráficos anteriores).

5. **Da formato al gráfico comparativo:**
   - Título: `Comparativo de Demanda — 4 SKUs (24 meses)`.
   - Asigna un color diferente a cada línea (usa la paleta de Excel o asigna manualmente: azul, verde, naranja, morado).
   - Activa la **leyenda** del gráfico (pestaña **Diseño de gráfico** → **Agregar elemento de gráfico** → **Leyenda** → **Abajo**).
   - Eje horizontal: `Periodo (mes)`. Eje vertical: `Unidades demandadas`.

6. **NO agregues líneas de tendencia en este gráfico** (quedaría visualmente saturado). En su lugar, el objetivo es la comparación directa de los patrones de demanda.

7. **Agrega anotaciones clave:**
   - Identifica el mes donde más SKUs coinciden en un pico simultáneo. Inserta un cuadro de texto: `Mes [N]: Pico multi-SKU — Riesgo de congestión en picking`.
   - Identifica el mes de menor demanda generalizada. Anota: `Mes [N]: Demanda baja — Oportunidad para reabastecimiento y reorganización`.
   - Si un SKU tiene una línea notablemente más alta que los demás, anota: `[SKU-X]: Mayor volumen — Prioridad en planificación de espacio`.

8. Ajusta el ancho de las líneas a **1.5 pt** para todos los SKUs para mejorar la legibilidad.

**Salida esperada:** Un gráfico de líneas con 4 series de colores distintos, leyenda visible, título claro y al menos 3 cuadros de texto con anotaciones operativas. El gráfico debe permitir identificar de un vistazo qué SKUs tienen comportamientos similares y cuáles son outliers.

**Verificación:** Verifica que las 4 series estén correctamente identificadas en la leyenda y que los nombres coincidan con los SKUs del dataset. Haz clic en cada línea del gráfico y confirma en la barra de fórmulas que el rango de datos es correcto.

---

### Paso 5 — Interpretar los gráficos y documentar conclusiones operativas

**Objetivo:** Traducir los patrones visuales identificados en los 3 gráficos en implicaciones concretas para la gestión de inventario y la planificación del picking.

#### Instrucciones

1. En la hoja `Gráficos_P3`, debajo de los gráficos (a partir de la fila 55 aproximadamente), crea una tabla de interpretación con el siguiente formato:

   | Gráfico | SKU(s) | Patrón observado | Componente dominante | Implicación para inventario | Implicación para picking |
   |---------|--------|------------------|----------------------|-----------------------------|--------------------------|
   | Gráfico 1 | SKU_Alto | (describir) | (Tendencia / Estacionalidad / Variabilidad) | (describir) | (describir) |
   | Gráfico 2 | SKU_Variable | (describir) | (describir) | (describir) | (describir) |
   | Gráfico 3 | 4 SKUs | (describir) | (describir) | (describir) | (describir) |

2. Para completar la columna **"Patrón observado"**, usa los siguientes criterios visuales como guía:

   | Lo que ves en el gráfico | Interpretación |
   |--------------------------|----------------|
   | Línea con pendiente positiva clara, R² > 0.6 | Tendencia creciente fuerte |
   | Línea con pendiente negativa, R² > 0.6 | Tendencia decreciente (posible descontinuación) |
   | Picos recurrentes cada 6 o 12 meses | Estacionalidad semestral o anual |
   | Media móvil muy separada de la demanda real | Alta variabilidad / demanda errática |
   | Línea relativamente plana, R² < 0.2 | Demanda estable sin tendencia clara |
   | Un solo pico aislado sin repetición | Evento extraordinario (promoción, desabasto) |

3. Para la columna **"Implicación para inventario"**, considera:
   - Si hay tendencia creciente: el stock de seguridad calculado con datos históricos antiguos puede ser **insuficiente** para los próximos meses.
   - Si hay estacionalidad: se debe planificar un **incremento de inventario previo** a los meses de pico.
   - Si hay alta variabilidad: se necesita un **mayor stock de seguridad** para mantener el nivel de servicio objetivo.
   - Si la demanda es estable: se puede operar con **inventarios ajustados** y mayor rotación.

4. Para la columna **"Implicación para picking"**, considera:
   - Meses de pico → mayor carga de trabajo en el área de picking → necesidad de **personal adicional o turnos extendidos**.
   - Meses de valle → oportunidad para **reorganizar ubicaciones** en el almacén y optimizar rutas de picking.
   - Alta variabilidad → mayor dificultad para planificar la operación diaria → conviene establecer **buffers de tiempo** en los procesos.

5. Guarda el archivo `Dataset_Graficos_P3.xlsx` con todos los cambios.

**Salida esperada:** Una tabla de interpretación completada con al menos 2 filas de contenido sólido (Gráficos 1 y 2 son obligatorios; el Gráfico 3 puede ser más general). La tabla debe mostrar que el participante conecta los patrones visuales con decisiones operativas reales, no solo con descripciones estadísticas abstractas.

**Verificación:** Revisa que cada fila de la tabla de interpretación tenga contenido en **todas** las columnas. Una celda vacía en "Implicación para inventario" o "Implicación para picking" indica que la interpretación está incompleta.

---

## Validación y Pruebas

Al finalizar los 5 pasos, verifica los siguientes criterios de aceptación:

| Criterio | Cómo verificar | Estado esperado |
|----------|----------------|-----------------|
| Gráfico 1 construido correctamente | El gráfico muestra 3 series: demanda real, media móvil (naranja), tendencia lineal (roja) | ✅ Completo |
| Ecuación y R² visibles en Gráfico 1 | Texto con `y = ...` y `R² = ...` visible dentro del área del gráfico | ✅ Completo |
| Gráfico 2 construido con SKU diferente | El SKU del Gráfico 2 es distinto al del Gráfico 1 y muestra mayor variabilidad visual | ✅ Completo |
| Gráfico 3 con 4 SKUs y leyenda | Cuatro líneas de colores distintos, leyenda identificada, título claro | ✅ Completo |
| Anotaciones en los 3 gráficos | Al menos 2 cuadros de texto por gráfico con observaciones específicas | ✅ Completo |
| Tabla de interpretación completada | Todas las columnas tienen contenido para al menos 2 gráficos | ✅ Completo |
| Archivo guardado | `Dataset_Graficos_P3.xlsx` guardado con hoja `Gráficos_P3` | ✅ Completo |

**Prueba de coherencia rápida:** Abre el Gráfico 1 y el Gráfico 2 simultáneamente. Pregúntate: ¿el SKU con mayor R² tiene una línea de demanda más "ordenada" visualmente? Si la respuesta es sí, tu análisis es coherente. Si el SKU con menor R² parece más ordenado, revisa si seleccionaste los SKUs correctos.

---

## Solución de Problemas

### Problema 1: La línea de tendencia no aparece o aparece sobre la serie incorrecta

**Síntoma:** Al hacer clic derecho sobre el gráfico, la opción "Agregar línea de tendencia" no está disponible, o la línea de tendencia aparece sobre la media móvil en lugar de la demanda real.

**Causa:** Excel agrega la línea de tendencia a la serie **actualmente seleccionada**. Si antes de hacer clic derecho hiciste clic sobre la línea de media móvil (en lugar de la línea de demanda real), la tendencia se aplicará a la serie incorrecta. Adicionalmente, la opción no está disponible si hiciste clic en el área vacía del gráfico en lugar de directamente sobre una línea de datos.

**Solución:**
1. Haz clic **directamente sobre uno de los puntos de datos** de la línea de demanda real (no sobre la media móvil). Verás que los marcadores de esa serie se resaltan.
2. Haz clic derecho → **Agregar línea de tendencia**.
3. Si la tendencia ya existe en la serie incorrecta, selecciona esa línea de tendencia y presiona **Suprimir** para eliminarla, luego repite el proceso seleccionando la serie correcta.
4. Verifica en el panel "Formato de línea de tendencia" que el campo **"Basado en la serie"** muestra el nombre de la demanda real (no "Media Móvil 3P").

---

### Problema 2: Los valores de la media móvil no coinciden con los datos o aparecen en posiciones incorrectas en el gráfico

**Síntoma:** La línea de media móvil en el gráfico parece estar desplazada hacia la izquierda o hacia la derecha respecto a la demanda real, o los valores calculados en la columna auxiliar son claramente incorrectos (por ejemplo, valores negativos o extremadamente altos).

**Causa:** El rango de la fórmula `=PROMEDIO(B2:B4)` está mal definido. Un error frecuente es comenzar la fórmula en la fila 2 (periodo 1) en lugar de la fila 4 (periodo 3), lo que hace que la media móvil se "adelante" en el tiempo. También puede ocurrir que al agregar la serie al gráfico, se haya seleccionado un rango que incluye las celdas vacías del inicio (C2:C3), generando un desplazamiento visual.

**Solución:**
1. Revisa la fórmula en la columna de media móvil: la primera celda con valor debe ser la fila correspondiente al **periodo 3** (no al periodo 1 ni al 2).
   - Correcto: celda C4 = `=PROMEDIO(B2:B4)` → representa el promedio de los periodos 1, 2 y 3.
   - Incorrecto: celda C2 = `=PROMEDIO(B2:B4)` → estaría calculando el promedio de periodos futuros.
2. Las celdas C2 y C3 deben estar **vacías** (no contener cero ni ningún valor).
3. Al agregar la serie de media móvil al gráfico, verifica que el rango seleccionado sea `C2:C25` (incluyendo las celdas vacías del inicio para que el alineamiento temporal sea correcto). Excel interpretará las celdas vacías como ausencia de dato y no graficará esos puntos, generando el efecto visual correcto de que la media móvil "comienza" en el periodo 3.
4. Si el gráfico sigue mostrando desplazamiento, elimina la serie de media móvil del gráfico y vuelve a agregarla seleccionando manualmente el rango correcto.

---

## Limpieza

Al finalizar la práctica, realiza las siguientes acciones:

1. **Guarda el archivo final:** Asegúrate de que `Dataset_Graficos_P3.xlsx` esté guardado con la hoja `Gráficos_P3` que contiene los 3 gráficos anotados y la tabla de interpretación.

2. **Verifica el nombre del archivo de salida:** El archivo que entregarás o usarás en prácticas posteriores debe llamarse exactamente `Dataset_Graficos_P3.xlsx`. Este archivo será referenciado en la Práctica 4.

3. **No modifiques la hoja `Datos_Depurados`:** Esta hoja es el insumo compartido de las prácticas 3 a 6. Cualquier modificación accidental podría afectar los cálculos de prácticas posteriores. Si detectaste algún error en los datos durante esta práctica, notifica al instructor en lugar de corregirlo directamente.

4. **Cierra los paneles de formato:** Cierra cualquier panel lateral de Excel ("Formato de línea de tendencia", "Formato de serie de datos") que haya quedado abierto, para liberar espacio en pantalla.

5. **Opcional — Exportar como PDF:** Si el instructor lo solicita, exporta la hoja `Gráficos_P3` como PDF: **Archivo → Exportar → Crear documento PDF/XPS**. Nombra el archivo `Graficos_P3_[TuNombre].pdf`.

---

## Resumen

En esta práctica construiste un **dashboard visual básico de demanda** con tres gráficos anotados en Excel, aplicando herramientas fundamentales de análisis visual de series de tiempo logísticas:

- **Gráfico 1:** Demanda del SKU de mayor volumen con línea de tendencia lineal (ecuación + R²) y media móvil de 3 periodos, que te permitió evaluar la fuerza y dirección de la tendencia.
- **Gráfico 2:** Demanda del SKU de mayor variabilidad, donde la brecha entre la demanda real y la media móvil evidenció el nivel de incertidumbre operativa de ese producto.
- **Gráfico 3:** Panel comparativo de 4 SKUs que reveló diferencias de comportamiento entre productos y permitió identificar meses de riesgo operativo simultáneo (picos multi-SKU).

La tabla de interpretación que completaste en el Paso 5 es el elemento más importante de esta práctica: traduce el análisis visual en **lenguaje operativo logístico**, conectando los patrones de demanda con decisiones concretas de inventario (cuánto stock mantener, cuándo reponer) y de picking (cuándo anticipar mayor carga de trabajo).

Este análisis visual es el punto de partida del ciclo de forecasting: antes de calcular cualquier modelo estadístico, es imprescindible entender qué comportamiento tiene la demanda. Un analista que construye modelos sin haber visualizado los datos primero corre el riesgo de aplicar el modelo equivocado al problema equivocado.

### Conexión con las próximas prácticas

- **Práctica 4:** Usarás los patrones identificados visualmente aquí para calcular los indicadores estadísticos (promedio, mediana, desviación estándar, coeficiente de variación) que cuantifican lo que acabas de observar de forma visual.
- **Práctica 5:** Clasificarás los SKUs según su estabilidad de demanda, usando como referencia los patrones visuales identificados en esta práctica para validar que la clasificación estadística es coherente con el comportamiento visual.

### Pregunta de reflexión final

> *¿Cuál de los SKUs analizados representa el mayor riesgo operativo para la empresa: el de mayor volumen o el de mayor variabilidad? ¿Por qué? ¿Cómo cambiaría tu respuesta si el SKU de mayor variabilidad también fuera el de mayor volumen?*

Esta pregunta no tiene una respuesta única, pero reflexionar sobre ella prepara el terreno para la discusión sobre clasificación de inventarios y estrategias diferenciadas de abastecimiento que desarrollaremos en las prácticas siguientes.

### Recursos de referencia

- [Documentación oficial de Microsoft: Agregar una línea de tendencia a un gráfico](https://support.microsoft.com/es-es/office/agregar-una-línea-de-tendencia-o-de-media-móvil-a-un-gráfico-fa59f86c-5852-4b68-a6d4-901a745842ad)
- [Documentación oficial de Microsoft: Tipos de líneas de tendencia disponibles en Office](https://support.microsoft.com/es-es/office/elegir-la-mejor-línea-de-tendencia-para-los-datos-1bb3c9e7-0280-45b5-9ab0-d0c62a622a08)
- Chopra, S. & Meindl, P. — *Supply Chain Management: Strategy, Planning, and Operation* — Capítulo 7: Demand Forecasting in a Supply Chain.

---

---

# Práctica 4 — Identificación Manual de Periodos Estacionales

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Herramienta principal** | Microsoft Excel 2016 o superior |
| **Práctica secuencial** | Requiere archivos de salida de Prácticas 1, 2 y 3 |

---

## 2. Descripción General

En esta práctica calcularás índices estacionales mensuales utilizando el método de **razón al promedio** (*ratio-to-moving-average*) sobre 24 meses de historial para 8 SKUs. Identificarás qué productos presentan estacionalidad fuerte y cuáles son prácticamente estables, aplicando el ajuste de Baseline para obtener una demanda desestacionalizada que sirva como insumo directo para el pronóstico logístico. Esta habilidad conecta directamente con la gestión de inventarios: un índice estacional bien calculado permite anticipar picos de demanda antes de que ocurran, evitando quiebres de stock y sobreinventario en temporadas bajas.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Calcular índices estacionales mensuales mediante el método de razón al promedio usando las funciones `PROMEDIO` y `PROMEDIO.SI` en Excel
- [ ] Construir una tabla de estacionalidad (12 meses × 8 SKUs) con formato condicional de escala de colores para visualizar patrones recurrentes de demanda
- [ ] Distinguir entre variabilidad aleatoria y estacionalidad genuina comparando índices contra umbrales operativos (> 1.3 o < 0.7)
- [ ] Calcular la demanda desestacionalizada (Baseline) y graficarla contra la demanda original para verificar el efecto del ajuste estacional

---

## 4. Prerrequisitos

### Conocimiento previo
- Haber completado las **Prácticas 1, 2 y 3** del curso (datos depurados y clasificados disponibles)
- Comprensión del concepto de estacionalidad e índices estacionales (Tema 1.4 del curso)
- Manejo de **referencias absolutas** en Excel (`$A$1`, `$B$2`) para anclar denominadores en fórmulas de división
- Comprensión básica del rol del pronóstico en inventarios y nivel de servicio (Lección 1.1)

### Acceso requerido
- Archivo de datos de práctica: **`Practica_04_Estacionalidad.xlsx`** (proporcionado por el instructor, derivado del dataset base de 24 meses para 8 SKUs)
- Microsoft Excel 2016 o superior (Microsoft 365 recomendado)
- Acceso al archivo de salida de la Práctica 3 (clasificación ABC/XYZ de los 8 SKUs)

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| Memoria RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1280 × 768 | 1920 × 1080 |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior | Microsoft 365 recomendado |
| Archivo de práctica | `Practica_04_Estacionalidad.xlsx` | Distribuido por el instructor |

### Configuración inicial

Antes de comenzar, verifica que el archivo de práctica esté disponible y abre Excel:

```
1. Abre el archivo: Practica_04_Estacionalidad.xlsx
2. Confirma que contiene las siguientes hojas:
   - "Datos_Historicos"  → 24 meses de demanda para 8 SKUs
   - "Indices_Estacionales" → hoja en blanco para tus cálculos
   - "Desestacionalizacion"  → hoja en blanco para el ajuste de Baseline
3. Guarda una copia de trabajo con tu nombre:
   Archivo > Guardar como > Practica_04_TuNombre.xlsx
```

### Estructura esperada del archivo de datos

La hoja **`Datos_Historicos`** debe tener la siguiente estructura:

| Columna | Contenido |
|---|---|
| A | Mes (1 a 24, o fecha formato MM/AAAA) |
| B | Año (2022 o 2023) |
| C | Mes calendario (1 = Enero … 12 = Diciembre) |
| D | SKU-A |
| E | SKU-B |
| F | SKU-C |
| G | SKU-D |
| H | SKU-E |
| I | SKU-F |
| J | SKU-G |
| K | SKU-H |

> **Nota del instructor:** Si el archivo tiene una estructura diferente, ajusta las referencias de columna en las fórmulas indicadas en cada paso.

---

## 6. Procedimiento Paso a Paso

---

### Paso 1 — Calcular el Promedio Mensual General de Cada SKU

**Objetivo:** Obtener la demanda media mensual de cada SKU a lo largo de los 24 meses. Este valor será el **denominador** de todos los índices estacionales.

#### Instrucciones

1. Ve a la hoja **`Indices_Estacionales`**.

2. En la fila 1, crea los encabezados de la tabla de trabajo:

```
A1: "Concepto"
B1: "SKU-A"
C1: "SKU-B"
D1: "SKU-C"
E1: "SKU-D"
F1: "SKU-E"
G1: "SKU-F"
H1: "SKU-G"
I1: "SKU-H"
```

3. En la celda **A2**, escribe: `Promedio General (24 meses)`

4. En la celda **B2**, ingresa la fórmula para calcular el promedio de los 24 meses del SKU-A:

```excel
=PROMEDIO(Datos_Historicos!D2:D25)
```

5. Copia la fórmula hacia la derecha hasta la celda **I2**, ajustando la columna de referencia para cada SKU:

```excel
C2: =PROMEDIO(Datos_Historicos!E2:E25)   ← SKU-B
D2: =PROMEDIO(Datos_Historicos!F2:F25)   ← SKU-C
E2: =PROMEDIO(Datos_Historicos!G2:G25)   ← SKU-D
F2: =PROMEDIO(Datos_Historicos!H2:H25)   ← SKU-E
G2: =PROMEDIO(Datos_Historicos!I2:I25)   ← SKU-F
H2: =PROMEDIO(Datos_Historicos!J2:J25)   ← SKU-G
I2: =PROMEDIO(Datos_Historicos!K2:K25)   ← SKU-H
```

6. Aplica formato de número con **2 decimales** a las celdas B2:I2 (Inicio > Número > 0.00).

#### Resultado esperado

La fila 2 mostrará 8 valores numéricos positivos representando la demanda promedio mensual de cada SKU. Ejemplo orientativo (los valores dependen de tu dataset):

| SKU-A | SKU-B | SKU-C | SKU-D | SKU-E | SKU-F | SKU-G | SKU-H |
|---|---|---|---|---|---|---|---|
| 342.50 | 187.25 | 510.83 | 95.42 | 278.17 | 423.67 | 156.33 | 612.08 |

#### Verificación

✅ Cada valor debe ser **positivo y razonable** (dentro del rango de datos que observaste en la Práctica 1).
✅ Ninguna celda debe mostrar `#¡REF!` ni `#¡VALOR!`.
✅ El promedio general debe estar aproximadamente en el **rango medio** de los datos históricos del SKU (no debe ser el valor más alto ni el más bajo).

---

### Paso 2 — Calcular el Promedio de Cada Mes Calendario por SKU

**Objetivo:** Calcular cuánto se demanda en promedio en cada mes del año (enero, febrero, …, diciembre) combinando los datos de los 2 años disponibles. Este es el **numerador** del índice estacional.

#### Instrucciones

1. En la hoja **`Indices_Estacionales`**, a partir de la fila 4, construye la tabla de promedios mensuales:

```
A4: "Mes"
B4: "SKU-A"
C4: "SKU-B"
... (mismos encabezados hasta I4)
```

2. En la columna A (celdas A5 a A16), ingresa los nombres de los meses:

```
A5:  Enero
A6:  Febrero
A7:  Marzo
A8:  Abril
A9:  Mayo
A10: Junio
A11: Julio
A12: Agosto
A13: Septiembre
A14: Octubre
A15: Noviembre
A16: Diciembre
```

3. En la celda **B5**, ingresa la fórmula `PROMEDIO.SI` para calcular el promedio de SKU-A solo en los meses de enero (mes calendario = 1):

```excel
=PROMEDIO.SI(Datos_Historicos!$C$2:$C$25, 1, Datos_Historicos!D$2:D$25)
```

> **Explicación de la fórmula:**
> - `Datos_Historicos!$C$2:$C$25` → rango de criterio: columna de mes calendario (fija con `$`)
> - `1` → criterio: mes número 1 (enero)
> - `Datos_Historicos!D$2:D$25` → rango de suma/promedio: columna del SKU-A

4. Para los meses de febrero a diciembre (celdas B6 a B16), cambia el criterio numérico (2, 3, 4… 12):

```excel
B6:  =PROMEDIO.SI(Datos_Historicos!$C$2:$C$25, 2, Datos_Historicos!D$2:D$25)
B7:  =PROMEDIO.SI(Datos_Historicos!$C$2:$C$25, 3, Datos_Historicos!D$2:D$25)
...
B16: =PROMEDIO.SI(Datos_Historicos!$C$2:$C$25, 12, Datos_Historicos!D$2:D$25)
```

5. **Copia el bloque B5:B16** hacia la derecha hasta la columna I, ajustando la columna del SKU en el tercer argumento de `PROMEDIO.SI`:

```excel
C5:  =PROMEDIO.SI(Datos_Historicos!$C$2:$C$25, 1, Datos_Historicos!E$2:E$25)
D5:  =PROMEDIO.SI(Datos_Historicos!$C$2:$C$25, 1, Datos_Historicos!F$2:F$25)
... (continúa hasta SKU-H en columna I)
```

> **Consejo de eficiencia:** Usa el autocompletado de Excel arrastrando el controlador de relleno (cuadro verde en la esquina inferior derecha de la celda) hacia abajo y luego hacia la derecha.

#### Resultado esperado

Una tabla de 12 filas × 8 columnas con los promedios históricos de cada mes calendario para cada SKU. Los valores deben ser coherentes con los datos históricos: los meses de alta demanda mostrarán valores superiores al promedio general calculado en el Paso 1.

#### Verificación

✅ Cada SKU debe tener exactamente **12 valores** (uno por mes).
✅ La **suma de los 12 promedios mensuales / 12** debe ser aproximadamente igual al promedio general de la fila 2 (tolerancia ±5%).
✅ Si alguna celda muestra `#¡DIV/0!`, significa que no hay datos para ese mes en ese SKU — verifica la columna C de `Datos_Historicos`.

---

### Paso 3 — Calcular los Índices Estacionales

**Objetivo:** Dividir el promedio de cada mes entre el promedio general para obtener el índice estacional. Un índice de 1.0 significa demanda igual al promedio; un índice de 1.4 indica un mes con 40% más demanda que el promedio anual.

#### Instrucciones

1. En la hoja **`Indices_Estacionales`**, crea una nueva sección a partir de la fila 19:

```
A19: "Mes"
B19: "IE SKU-A"
C19: "IE SKU-B"
... hasta I19: "IE SKU-H"
```

2. En las celdas A20:A31, replica los nombres de los meses (enero a diciembre).

3. En la celda **B20**, ingresa la fórmula del índice estacional para SKU-A / Enero:

```excel
=B5/B$2
```

> **Explicación:**
> - `B5` = promedio de SKU-A en enero (numerador, fila relativa para que cambie al copiar hacia abajo)
> - `B$2` = promedio general de SKU-A (denominador, **fila fija** con `$2` para que no cambie al copiar)

4. Copia la fórmula **hacia abajo** (B20:B31) para los 12 meses del SKU-A:

```excel
B21: =B6/B$2
B22: =B7/B$2
...
B31: =B16/B$2
```

5. Copia el bloque **B20:B31** hacia la derecha hasta la columna I para los 8 SKUs.

6. Aplica formato de número con **4 decimales** a todas las celdas de índices (B20:I31).

7. **Verificación de suma:** En la fila 33, calcula la suma de cada columna de índices:

```excel
A33: "Suma IE (debe ≈ 12.00)"
B33: =SUMA(B20:B31)
```

Copia hacia la derecha hasta I33. La suma de los 12 índices de cada SKU debe ser aproximadamente **12.00** (tolerancia ±0.05).

#### Resultado esperado

Una tabla de índices estacionales con valores alrededor de 1.0. Ejemplo orientativo:

| Mes | IE SKU-A | IE SKU-B | IE SKU-C |
|---|---|---|---|
| Enero | 0.72 | 1.05 | 0.68 |
| Febrero | 0.78 | 0.98 | 0.71 |
| Marzo | 1.42 | 1.02 | 1.38 |
| … | … | … | … |
| Diciembre | 1.35 | 1.01 | 1.41 |
| **Suma** | **12.00** | **12.00** | **12.00** |

#### Verificación

✅ La suma de cada columna de índices debe ser **≈ 12.00**.
✅ Ningún índice debe ser **0 ni negativo** (indica error en los datos fuente).
✅ Los índices deben estar en un rango **razonable** (típicamente entre 0.4 y 2.0 para datos reales de distribución).

---

### Paso 4 — Aplicar Formato Condicional de Escala de Colores

**Objetivo:** Visualizar rápidamente los patrones estacionales usando una escala de colores que destaque los meses de alta demanda (rojo/naranja) y baja demanda (verde/azul).

#### Instrucciones

1. Selecciona el rango de índices estacionales: **B20:I31** (12 meses × 8 SKUs).

2. Ve a **Inicio > Formato condicional > Escalas de color**.

3. Selecciona la escala **Verde - Amarillo - Rojo** (la primera opción en el submenú):
   - **Verde** → valores bajos (demanda por debajo del promedio, índice < 1.0)
   - **Amarillo** → valores medios (demanda cercana al promedio, índice ≈ 1.0)
   - **Rojo** → valores altos (demanda por encima del promedio, índice > 1.0)

   > **Alternativa:** Si prefieres que el rojo indique picos de demanda (más intuitivo para alertas de inventario), usa la escala **Rojo - Amarillo - Verde** invertida: Formato condicional > Administrar reglas > Editar regla > intercambia los colores mínimo y máximo.

4. Ajusta el ancho de las columnas B:I para que todos los valores sean visibles: selecciona las columnas, clic derecho > **Ancho de columna** > escribe `12`.

5. Añade un borde exterior al rango B20:I31: selecciona el rango > **Inicio > Bordes > Borde de cuadro grueso**.

6. **Etiqueta la tabla:** En la celda A18, escribe:

```
"TABLA DE ÍNDICES ESTACIONALES (IE = Promedio Mes / Promedio General)"
```

Aplica negrita y color de fuente azul oscuro.

#### Resultado esperado

Una tabla visualmente clara donde los patrones estacionales son inmediatamente reconocibles. Los SKUs con estacionalidad fuerte mostrarán columnas con alternancia clara de colores rojos e intensos (meses pico) y verdes (meses valle). Los SKUs estables mostrarán una paleta uniforme de amarillos.

#### Verificación

✅ El formato condicional debe aplicarse **solo al rango B20:I31**, no a los encabezados ni a la fila de sumas.
✅ Deben ser visibles al menos **2 tonos claramente diferenciados** de color en cada SKU con estacionalidad.
✅ Si toda la tabla aparece en un solo color, revisa que los datos de origen no sean todos iguales (posible error en el dataset).

---

### Paso 5 — Identificar Meses Pico y Valle por SKU

**Objetivo:** Para cada SKU, identificar los 3 meses de mayor demanda y los 3 de menor demanda según los índices estacionales. Esta información es crítica para la planificación de inventario de seguridad y la negociación con proveedores.

#### Instrucciones

1. En la hoja **`Indices_Estacionales`**, crea una nueva sección a partir de la fila 36:

```
A36: "SKU"
B36: "Mes Pico 1 (mayor IE)"
C36: "Mes Pico 2"
D36: "Mes Pico 3"
E36: "Mes Valle 1 (menor IE)"
F36: "Mes Valle 2"
G36: "Mes Valle 3"
H36: "IE Máximo"
I36: "IE Mínimo"
J36: "Rango IE (Max-Min)"
K36: "¿Estacionalidad Fuerte?"
```

2. En las celdas A37:A44, escribe los nombres de los 8 SKUs (SKU-A a SKU-H).

3. Para identificar los meses pico, usa la función `K.ESIMO.MAYOR` combinada con `INDICE` y `COINCIDIR`. Para SKU-A (columna B, índices en B20:B31):

```excel
H37: =K.ESIMO.MAYOR(B20:B31, 1)   ← IE máximo de SKU-A
I37:  =K.ESIMO.MENOR(B20:B31, 1)   ← IE mínimo de SKU-A
J37:  =H37-I37                      ← Rango de variación
```

4. Para identificar el **nombre del mes** con mayor IE (columna B37):

```excel
B37: =INDICE($A$20:$A$31, COINCIDIR(K.ESIMO.MAYOR(B20:B31,1), B20:B31, 0))
C37: =INDICE($A$20:$A$31, COINCIDIR(K.ESIMO.MAYOR(B20:B31,2), B20:B31, 0))
D37: =INDICE($A$20:$A$31, COINCIDIR(K.ESIMO.MAYOR(B20:B31,3), B20:B31, 0))
E37: =INDICE($A$20:$A$31, COINCIDIR(K.ESIMO.MENOR(B20:B31,1), B20:B31, 0))
F37: =INDICE($A$20:$A$31, COINCIDIR(K.ESIMO.MENOR(B20:B31,2), B20:B31, 0))
G37: =INDICE($A$20:$A$31, COINCIDIR(K.ESIMO.MENOR(B20:B31,3), B20:B31, 0))
```

> **Nota:** Si dos meses tienen el mismo IE (empate), `COINCIDIR` devolverá el primero que encuentre. Esto es aceptable para el propósito de esta práctica.

5. Para la columna **K** (¿Estacionalidad Fuerte?), aplica la siguiente lógica: un SKU tiene estacionalidad fuerte si su IE máximo supera 1.3 **o** su IE mínimo es menor que 0.7:

```excel
K37: =SI(O(H37>1.3, I37<0.7), "SÍ — Estacional", "NO — Estable")
```

6. Copia todas las fórmulas de la fila 37 hacia abajo hasta la fila 44 (un SKU por fila), ajustando las referencias de columna para cada SKU:

```
Fila 37 → SKU-A: columna B (índices en B20:B31)
Fila 38 → SKU-B: columna C (índices en C20:C31)
Fila 39 → SKU-C: columna D (índices en D20:D31)
... y así sucesivamente hasta SKU-H en columna I
```

7. Aplica **formato condicional** a la columna K (K37:K44):
   - Regla 1: Si el texto contiene "SÍ" → fondo **naranja**, texto en negrita
   - Regla 2: Si el texto contiene "NO" → fondo **verde claro**

#### Resultado esperado

Una tabla resumen que muestra, para cada SKU, los meses de mayor y menor demanda, el rango de variación de los índices y una clasificación binaria de estacionalidad. Esta tabla es el insumo directo para la planificación de inventario de seguridad diferenciado por temporada.

#### Verificación

✅ Todos los nombres de mes en las columnas B-G deben ser nombres válidos (Enero, Febrero, etc.), no números.
✅ La columna J (Rango IE) debe ser **positiva** para todos los SKUs.
✅ Los SKUs clasificados como "SÍ — Estacional" deben coincidir visualmente con los patrones de color más intensos en la tabla del Paso 4.

---

### Paso 6 — Calcular la Demanda Desestacionalizada (Baseline)

**Objetivo:** Dividir cada dato de demanda real entre su índice estacional correspondiente para obtener la demanda "limpia" de efectos estacionales. Esta demanda desestacionalizada es el **Baseline** que usaremos en prácticas posteriores para construir el pronóstico.

#### Instrucciones

1. Ve a la hoja **`Desestacionalizacion`**.

2. Crea los encabezados:

```
A1: "Mes"
B1: "Año"
C1: "Mes Cal."
D1: "Demanda Real SKU-A"
E1: "IE SKU-A"
F1: "Baseline SKU-A"
G1: "Demanda Real SKU-B"
H1: "IE SKU-B"
I1: "Baseline SKU-B"
```

> **Nota:** Por razones de tiempo, en esta práctica trabajaremos con **SKU-A y SKU-B** como ejemplos representativos. El instructor puede solicitar extender el ejercicio a los 8 SKUs como tarea.

3. Copia los datos de mes, año y mes calendario desde `Datos_Historicos`:

```excel
A2: =Datos_Historicos!A2
B2: =Datos_Historicos!B2
C2: =Datos_Historicos!C2
```

Copia estas tres fórmulas hasta la fila 25 (24 meses).

4. Copia la demanda real del SKU-A:

```excel
D2: =Datos_Historicos!D2
```

Copia hasta D25.

5. En la columna E (IE SKU-A), usa `BUSCARV` o `INDICE/COINCIDIR` para traer el índice estacional correspondiente al mes calendario de cada fila. Dado que los índices están en `Indices_Estacionales!B20:B31` y los meses en `A20:A31`, usa:

```excel
E2: =INDICE(Indices_Estacionales!$B$20:$B$31, C2)
```

> **Explicación:** `C2` contiene el número del mes calendario (1-12). Como los índices están ordenados de enero (fila 1) a diciembre (fila 12), `INDICE` con `C2` como número de posición devuelve directamente el índice del mes correcto.

Copia E2 hasta E25.

6. Calcula el **Baseline (demanda desestacionalizada)** dividiendo la demanda real entre el índice estacional:

```excel
F2: =D2/E2
```

Copia F2 hasta F25.

7. Repite los pasos 4-6 para SKU-B en las columnas G, H e I:

```excel
G2: =Datos_Historicos!E2          ← Demanda real SKU-B
H2: =INDICE(Indices_Estacionales!$C$20:$C$31, C2)  ← IE SKU-B
I2: =G2/H2                        ← Baseline SKU-B
```

8. Aplica formato de número con **0 decimales** a las columnas D, F, G e I (las demandas reales y desestacionalizadas son unidades enteras).

#### Resultado esperado

Una tabla de 24 filas con la demanda real y la demanda desestacionalizada para SKU-A y SKU-B. En los SKUs con estacionalidad fuerte, el Baseline debe ser **más suave** que la demanda real (menos picos y valles pronunciados).

#### Verificación

✅ Ninguna celda del Baseline debe ser **negativa**.
✅ El **promedio del Baseline** de cada SKU debe ser aproximadamente igual al promedio general calculado en el Paso 1 (±5%).
✅ Si el Baseline es idéntico a la demanda real, verifica que los índices estacionales no sean todos 1.0 (posible error en el dataset o en el cálculo del Paso 3).

---

### Paso 7 — Graficar Demanda Original vs. Demanda Desestacionalizada

**Objetivo:** Crear un gráfico de líneas comparativo que muestre visualmente el efecto del ajuste estacional, confirmando que el Baseline es una serie más estable que la demanda original.

#### Instrucciones

1. En la hoja **`Desestacionalizacion`**, selecciona el rango **A1:A25** (meses) manteniendo `Ctrl` presionado, luego selecciona también **D1:D25** (Demanda Real SKU-A) y **F1:F25** (Baseline SKU-A).

2. Inserta un gráfico de líneas:
   - **Insertar > Gráficos > Línea > Línea con marcadores**

3. Configura el gráfico:
   - **Título del gráfico:** `SKU-A: Demanda Real vs. Baseline Desestacionalizado`
   - **Eje X:** Meses (1 a 24)
   - **Eje Y:** Unidades demandadas
   - **Serie 1 (Demanda Real):** línea azul sólida con marcadores circulares
   - **Serie 2 (Baseline):** línea naranja discontinua con marcadores cuadrados

4. Para cambiar los colores:
   - Clic derecho en la línea de Demanda Real > **Formato de serie de datos** > Color de línea: **Azul**
   - Clic derecho en la línea de Baseline > **Formato de serie de datos** > Color de línea: **Naranja** > Tipo de guión: **Guión**

5. Agrega una **línea de referencia** horizontal en el valor del promedio general:
   - Crea una columna auxiliar en la hoja con el valor del promedio general repetido 24 veces:
     ```excel
     K2:K25: =Indices_Estacionales!$B$2
     ```
   - Agrega esta serie al gráfico como una tercera línea de color gris punteado con etiqueta "Promedio General".

6. Ubica el gráfico en la hoja `Desestacionalizacion`, debajo de la tabla de datos (aproximadamente desde la fila 28).

#### Resultado esperado

Un gráfico de líneas con dos series claramente diferenciadas:
- La **línea azul** (Demanda Real) muestra los picos y valles estacionales
- La **línea naranja discontinua** (Baseline) es visualmente más plana y estable
- La **línea gris** marca el promedio general como referencia

En un SKU con estacionalidad fuerte, la distancia entre ambas líneas debe ser claramente visible en los meses pico y valle.

#### Verificación

✅ El gráfico debe mostrar **exactamente 24 puntos** en el eje X.
✅ La línea del Baseline debe ser **visualmente más suave** que la demanda real (menos variabilidad).
✅ Si ambas líneas son prácticamente idénticas, el SKU seleccionado no tiene estacionalidad significativa — considera graficar un SKU con estacionalidad fuerte identificado en el Paso 5.

---

## 7. Validación y Prueba Final

Una vez completados los 7 pasos, realiza las siguientes verificaciones integrales:

### Lista de verificación de entregables

| # | Verificación | Criterio de éxito |
|---|---|---|
| 1 | Fila de promedios generales (Paso 1) | 8 valores positivos, formato 2 decimales |
| 2 | Tabla de promedios mensuales (Paso 2) | 12 × 8 = 96 celdas con valores, sin `#¡DIV/0!` |
| 3 | Tabla de índices estacionales (Paso 3) | Suma de cada columna ≈ 12.00 (±0.05) |
| 4 | Formato condicional (Paso 4) | Escala de colores visible, al menos 2 tonos diferenciados |
| 5 | Tabla de meses pico/valle (Paso 5) | 8 SKUs clasificados, columna K con "SÍ" o "NO" |
| 6 | Tabla de Baseline (Paso 6) | 24 filas para SKU-A y SKU-B, valores positivos |
| 7 | Gráfico comparativo (Paso 7) | Dos líneas diferenciadas, Baseline más suave |

### Pregunta de reflexión logística

> **¿Cómo impacta este análisis en una decisión operativa de logística?**
>
> Responde brevemente (2-3 oraciones) en la celda **A50** de la hoja `Indices_Estacionales`:
> - ¿Qué SKUs requieren un stock de seguridad diferenciado por temporada?
> - ¿En qué meses debería el equipo de compras anticipar sus órdenes de reposición?
> - ¿Cómo cambia la planificación del picking si conoces los meses de mayor índice estacional?

### Prueba de coherencia cruzada

Verifica que los resultados de esta práctica sean coherentes con la clasificación obtenida en la Práctica 3:
- Los SKUs clasificados como **X** (alta variabilidad) en la Práctica 3 deberían mostrar índices estacionales más extremos (> 1.3 o < 0.7) en esta práctica.
- Los SKUs clasificados como **Z** (muy alta variabilidad) podrían tener índices extremos **o** simplemente alta variabilidad aleatoria — esta práctica te ayuda a distinguir cuál es el caso.

---

## 8. Solución de Problemas

### Problema 1: La suma de índices estacionales no es ≈ 12.00

**Síntoma:** La fila de suma (Paso 3, celda B33) muestra un valor significativamente diferente de 12.00 (por ejemplo, 8.5 o 15.3).

**Causa probable:** El rango de la función `PROMEDIO.SI` en el Paso 2 no cubre exactamente los 24 meses, o la columna de mes calendario (columna C en `Datos_Historicos`) tiene valores fuera del rango 1-12 (por ejemplo, fechas en formato fecha de Excel en lugar de números enteros).

**Solución:**
1. Ve a `Datos_Historicos` y verifica la columna C: todos los valores deben ser números enteros del 1 al 12.
2. Si la columna C contiene fechas en formato `MM/AAAA`, extrae el mes con la función `MES()`:
   ```excel
   C2: =MES(A2)
   ```
   Copia hasta C25 y verifica que los valores sean 1-12.
3. Regresa a `Indices_Estacionales` y recalcula el libro completo: **Fórmulas > Calcular ahora** (o presiona `F9`).
4. Verifica que la suma sea ≈ 12.00. Una tolerancia de ±0.10 es aceptable por redondeo.

---

### Problema 2: El gráfico del Paso 7 muestra el Baseline idéntico o casi idéntico a la demanda real

**Síntoma:** Ambas líneas del gráfico se superponen o están muy cercanas, haciendo imposible distinguir el efecto del ajuste estacional.

**Causa probable:** (a) El SKU seleccionado genuinamente no tiene estacionalidad (todos sus índices son ≈ 1.0), o (b) los índices estacionales se calcularon incorrectamente con el denominador equivocado (por ejemplo, se usó el promedio del mes en lugar del promedio general en la fórmula del Paso 3).

**Solución:**
1. **Verifica primero la causa:** Revisa la tabla de clasificación del Paso 5. Si el SKU seleccionado está marcado como "NO — Estable", el comportamiento del gráfico es **correcto** — el SKU simplemente no tiene estacionalidad.
2. **Si el problema es el cálculo:** Revisa la fórmula del Paso 3. La celda B20 debe ser `=B5/B$2`, donde `B$2` es el promedio general (fila fija). Si la fórmula es `=B5/B5` (división por sí misma), el índice siempre será 1.0 — corrige el denominador a `B$2`.
3. **Solución práctica:** Selecciona para el gráfico un SKU que la tabla del Paso 5 haya clasificado como "SÍ — Estacional" (por ejemplo, el SKU con el mayor rango de IE en la columna J). Actualiza las series del gráfico con las columnas correctas de ese SKU.

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza las siguientes acciones:

1. **Guarda el archivo de trabajo:**
   ```
   Ctrl + S  (o Archivo > Guardar)
   Nombre: Practica_04_TuNombre_COMPLETA.xlsx
   ```

2. **Verifica que el archivo contenga las tres hojas trabajadas:**
   - `Datos_Historicos` (datos originales, sin modificaciones)
   - `Indices_Estacionales` (tabla de IE, tabla pico/valle, respuesta de reflexión)
   - `Desestacionalizacion` (tabla de Baseline y gráfico comparativo)

3. **Entrega al instructor** (según el método indicado: carpeta compartida, correo o plataforma del curso) el archivo `Practica_04_TuNombre_COMPLETA.xlsx`.

4. **Importante para la siguiente práctica:** La hoja `Indices_Estacionales` (específicamente el rango B20:I31 con los índices estacionales) será utilizada directamente en la **Práctica 5** para calcular el pronóstico ajustado por estacionalidad. Asegúrate de que el archivo esté guardado y accesible antes de cerrar Excel.

5. **No cierres Excel** si el instructor indica continuar directamente con la Práctica 5.

---

## 10. Resumen y Recursos Adicionales

### Conceptos clave aplicados en esta práctica

| Concepto | Definición operativa | Aplicación logística |
|---|---|---|
| **Índice estacional (IE)** | Razón entre el promedio del mes y el promedio general (IE = Prom_mes / Prom_general) | Indica cuánto más o menos demanda habrá en un mes respecto al promedio anual |
| **Estacionalidad fuerte** | IE > 1.3 en meses pico o IE < 0.7 en meses valle | Requiere stock de seguridad diferenciado y órdenes de compra anticipadas |
| **Baseline (demanda desestacionalizada)** | Demanda real dividida entre el IE del mes correspondiente | Serie más estable que permite identificar la tendencia subyacente del SKU |
| **Método razón al promedio** | Técnica de cálculo de IE usando promedios históricos por mes calendario | Método estándar en logística para estacionalidad con 2+ años de datos |

### Conexión con el pronóstico logístico

El trabajo realizado en esta práctica completa el segundo componente de la estructura metodológica del pronóstico que veremos en el curso:

```
Pronóstico = Baseline × Índice Estacional × Factor de Iniciativas
                ↑                ↑                      ↑
           (Práctica 7)    (Esta práctica)          (Prácticas 8-10)
```

Los índices estacionales calculados hoy son el insumo que permitirá, en las prácticas siguientes, transformar un pronóstico "plano" en un pronóstico que respeta los ritmos reales de la demanda — exactamente como el analista de la empresa distribuidora de alimentos del ejemplo de la Lección 1.1, que anticipó el pico de marzo con dos semanas de anticipación y mantuvo el nivel de servicio en 97%.

### Recursos de referencia

- **Chopra & Meindl** — *Supply Chain Management: Strategy, Planning, and Operation* (Pearson): Capítulo 7 — Demand Forecasting in a Supply Chain
- **Ballou, R.H.** — *Business Logistics/Supply Chain Management* (Pearson): Capítulo 9 — Forecasting Supply Chain Requirements
- **Excel Help** — Función `PROMEDIO.SI`: [support.microsoft.com/es-es/office/función-promedio-si](https://support.microsoft.com/es-es/office/promedio-si-función-promedio-si-faec8e2e-0dec-4308-af69-f5576d8ac642)
- **Excel Help** — Formato condicional con escala de colores: [support.microsoft.com — Agregar, cambiar o borrar formatos condicionales](https://support.microsoft.com/es-es/office/agregar-cambiar-o-borrar-formatos-condicionales-164d9f03-1f61-4a53-8b7f-c9954c2caf36)

---

---

# Práctica 5 — Identificación de Productos de Alta Rotación y Análisis de su Impacto en las Operaciones de Picking

## Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Herramientas principales** | Microsoft Excel 2016 o superior |
| **Archivos requeridos** | Dataset depurado de Práctica 1–4 (24 meses, 8 SKUs) |
| **Producto entregable** | Tabla de clasificación ABC con recomendaciones operativas de picking |

---

## Descripción General

En esta práctica aplicarás un análisis de rotación y clasificación ABC sobre los 8 SKUs del dataset trabajado en las prácticas anteriores. Utilizarás funciones de Excel para calcular participaciones acumuladas, asignar categorías A, B y C, y traducir esos resultados en recomendaciones concretas para la organización de zonas de picking, asignación de recursos y planificación de turnos. El análisis se ancla directamente en el concepto central del curso: un mejor pronóstico de demanda permite anticipar qué productos necesitan mayor disponibilidad operativa, reduciendo quiebres de stock y elevando el nivel de servicio al cliente.

---

## Objetivos de Aprendizaje

Al finalizar esta práctica, serás capaz de:

- [ ] Calcular el volumen total y la participación porcentual acumulada de cada SKU para aplicar la regla ABC (80-15-5).
- [ ] Asignar la clasificación ABC a cada SKU utilizando criterios cuantitativos en Excel con formato condicional.
- [ ] Estimar el volumen de líneas de picking por mes para los SKUs Clase A como insumo de planificación operativa.
- [ ] Relacionar los picos estacionales identificados en la Práctica 4 con la planificación de turnos para los productos de mayor rotación.
- [ ] Documentar recomendaciones operativas de almacén (zonas prime, media y remota) fundamentadas en el análisis de rotación.

---

## Prerrequisitos

### Conocimiento previo
- Haber completado las **Prácticas 1 a 4** del curso y contar con el archivo Excel resultante con datos depurados de 24 meses para 8 SKUs.
- Comprensión del impacto del pronóstico en la rotación de inventario, nivel de servicio y quiebres de stock (Lección 1.1).
- Comprensión básica de indicadores operativos vinculados al pronóstico (tema 1.5) y su impacto en la planificación del picking (tema 1.6).
- Familiaridad con funciones Excel: `SUMA`, `SUMAR.SI`, porcentajes, ordenamiento de datos y formato condicional.

### Acceso y archivos
- Computadora con **Microsoft Excel 2016 o superior** (se recomienda Microsoft 365).
- Archivo de trabajo de la Práctica 4 guardado como `Practica4_[TuNombre].xlsx` o el archivo checkpoint provisto por el instructor.
- Conexión a internet disponible (no requerida en este paso, pero recomendada para Prácticas 7–8).

> **Nota para el instructor:** Si algún participante no completó la Práctica 4, proporcionar el archivo checkpoint `Checkpoint_P4_Completo.xlsx` antes de iniciar esta sesión.

---

## Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 2 GB | 2 GB |
| Resolución de pantalla | 1366×768 | 1920×1080 |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior | Microsoft 365 recomendado |
| Archivo de práctica | `Practica4_[TuNombre].xlsx` | Generado en Práctica 4 |

### Configuración inicial (antes de comenzar)

1. Abre el archivo `Practica4_[TuNombre].xlsx` en Excel.
2. Guarda inmediatamente una copia con el nombre `Practica5_[TuNombre].xlsx` usando **Archivo → Guardar como**.
3. Verifica que el archivo contiene la hoja `Datos_Limpios` con 24 filas de datos mensuales y columnas para los 8 SKUs (SKU-01 a SKU-08).
4. Asegúrate de que la barra de fórmulas esté visible: **Vista → Barra de fórmulas** (activada).

---

## Procedimiento Paso a Paso

---

### Paso 1 — Crear la Hoja de Trabajo para el Análisis ABC

**Objetivo:** Preparar una hoja dedicada con la estructura base para el análisis de rotación y clasificación ABC.

**Instrucciones:**

1. En el archivo `Practica5_[TuNombre].xlsx`, haz clic derecho sobre la pestaña `Datos_Limpios` y selecciona **Mover o copiar**.
2. Selecciona **Crear una copia** y ubícala al final. Renombra la nueva hoja como `Analisis_ABC`.
3. En la hoja `Analisis_ABC`, inserta una nueva hoja en blanco adicional y nómbrala `Reporte_Picking`. La usarás en el Paso 5.
4. Regresa a la hoja `Analisis_ABC`. Desplázate hacia la derecha de los datos existentes y ubica una columna vacía (por ejemplo, columna **K** si los datos terminan en la columna **J**).
5. En la fila 1 de esa columna vacía, escribe el encabezado: `TOTAL_24M` (total de unidades vendidas en 24 meses por SKU).

> **Tip de orientación:** Si tu dataset tiene la estructura Fila = Mes, Columna = SKU, el total por SKU se calculará sumando verticalmente cada columna de datos. Confirma con tu instructor la orientación exacta de tu archivo.

**Resultado esperado:** El archivo tiene dos hojas nuevas — `Analisis_ABC` (copia de los datos limpios) y `Reporte_Picking` (en blanco). La hoja `Analisis_ABC` tiene un encabezado `TOTAL_24M` preparado para los cálculos.

**Verificación:** Confirma que puedes ver las tres pestañas: `Datos_Limpios`, `Analisis_ABC` y `Reporte_Picking` en la barra inferior de Excel.

---

### Paso 2 — Calcular el Volumen Total por SKU y Ordenar de Mayor a Menor

**Objetivo:** Obtener el total de unidades vendidas en 24 meses para cada SKU y ordenarlos de mayor a menor para preparar el análisis de participación.

**Instrucciones:**

1. En la hoja `Analisis_ABC`, crea una **tabla resumen independiente** a partir de la fila 30 (o en una zona despejada) con la siguiente estructura de columnas:

   | Columna | Encabezado |
   |---|---|
   | A | SKU |
   | B | Total_Unidades_24M |
   | C | Part_Individual_% |
   | D | Part_Acumulada_% |
   | E | Clasificacion_ABC |
   | F | Promedio_Mensual |
   | G | Lineas_Picking_Mes |
   | H | Zona_Picking |

2. En la columna **A** (desde A31), escribe los nombres de los 8 SKUs: `SKU-01`, `SKU-02`, ... `SKU-08`.

3. En la celda **B31**, escribe la fórmula para sumar las 24 filas de datos del SKU-01. Ejemplo (ajusta el rango según tu dataset):
   ```
   =SUMA(Datos_Limpios!B2:B25)
   ```
   Repite para cada SKU en B32 a B38, referenciando la columna correspondiente en `Datos_Limpios`.

4. Una vez completados los 8 valores en B31:B38, **ordena la tabla completa** (filas 31 a 38) de mayor a menor según la columna B:
   - Selecciona el rango A31:H38.
   - Ve a **Datos → Ordenar**.
   - Ordenar por: **Total_Unidades_24M**, de **Mayor a Menor**.
   - Haz clic en **Aceptar**.

5. En una celda aparte (por ejemplo, **B40**), calcula el total general:
   ```
   =SUMA(B31:B38)
   ```
   Etiqueta esta celda en A40 como `TOTAL GENERAL`.

**Resultado esperado:** Una tabla con los 8 SKUs ordenados de mayor a menor volumen de ventas, con el total general calculado en B40. El SKU con mayor volumen aparece en la fila 31.

**Verificación:** La suma de B31:B38 debe ser igual al valor en B40. Verifica con `=SUMA(B31:B38)-B40` → el resultado debe ser `0`.

---

### Paso 3 — Calcular Participaciones Individual y Acumulada

**Objetivo:** Determinar el porcentaje que cada SKU representa del volumen total y construir el acumulado para aplicar los cortes A (80%), B (15%) y C (5%).

**Instrucciones:**

1. En la celda **C31**, calcula la participación individual del primer SKU (el de mayor volumen):
   ```
   =B31/$B$40
   ```
   Aplica formato de porcentaje con 2 decimales: **Inicio → Número → %** (o `Ctrl+Shift+%`).

2. Copia la fórmula de C31 hacia abajo hasta **C38** (los 8 SKUs). Verifica que la referencia `$B$40` permanezca fija (signo `$`).

3. En la celda **D31**, la participación acumulada del primer SKU es igual a su participación individual:
   ```
   =C31
   ```

4. En la celda **D32**, comienza el acumulado real:
   ```
   =D31+C32
   ```
   Copia esta fórmula hacia abajo hasta **D38**. El valor en D38 debe ser exactamente **100%** (o muy cercano, con diferencias de redondeo menores a 0.01%).

5. Aplica formato de porcentaje con 2 decimales a todo el rango **C31:D38**.

**Resultado esperado:** La columna C muestra el porcentaje individual de cada SKU (todos suman 100%). La columna D muestra el acumulado progresivo, terminando en 100% en D38. Podrás ver visualmente en qué fila el acumulado supera el 80% (corte Clase A) y el 95% (corte Clase B).

**Verificación:** Confirma que `=SUMA(C31:C38)` = 100% (puede mostrar 100.00% o 1.0000 según formato). Si el resultado difiere en más de 0.1%, revisa las fórmulas de SUMA del Paso 2.

---

### Paso 4 — Asignar la Clasificación ABC con Fórmula Condicional

**Objetivo:** Etiquetar automáticamente cada SKU como Clase A, B o C según los umbrales de participación acumulada (80% / 95% / 100%).

**Instrucciones:**

1. En la celda **E31**, escribe la siguiente fórmula con `SI` anidado:
   ```
   =SI(D31<=0.80,"A",SI(D31<=0.95,"B","C"))
   ```

2. Copia esta fórmula hacia abajo hasta **E38**.

3. Revisa los resultados: los primeros SKUs (acumulado ≤ 80%) deben mostrar **"A"**; los siguientes (acumulado entre 80.01% y 95%) deben mostrar **"B"**; los últimos (acumulado > 95%) deben mostrar **"C"**.

4. Aplica **Formato Condicional** para resaltar visualmente las clases:
   - Selecciona el rango **E31:E38**.
   - Ve a **Inicio → Formato condicional → Resaltar reglas de celdas → Es igual a...**
   - Crea tres reglas:
     - Valor = `A` → Relleno **verde claro** (representa alta rotación / alta prioridad).
     - Valor = `B` → Relleno **amarillo** (rotación media).
     - Valor = `C` → Relleno **rojo claro** (baja rotación / baja prioridad).
   - Aplica también el mismo formato condicional al rango **A31:A38** para que los nombres de SKU se coloreen según su clase.

5. Selecciona también el rango completo **A31:H38** y aplica bordes de tabla: **Inicio → Bordes → Todos los bordes**.

**Resultado esperado:** La columna E muestra "A", "B" o "C" con colores diferenciados (verde/amarillo/rojo). Los SKUs Clase A son los que concentran el 80% del volumen de ventas y aparecen en las primeras filas de la tabla ordenada.

**Verificación:** Cuenta manualmente cuántos SKUs quedaron en cada clase. En un dataset de 8 SKUs con distribución típica, es normal encontrar 2–3 SKUs en Clase A, 2–3 en Clase B y 2–3 en Clase C. Si todos los SKUs quedaron en Clase A, verifica que el total general en B40 sea correcto y que las fórmulas de participación acumulada en D31:D38 estén bien construidas.

---

### Paso 5 — Calcular Promedio Mensual y Estimación de Líneas de Picking

**Objetivo:** Calcular el promedio mensual de unidades y estimar las líneas de picking por mes para cada SKU, con énfasis en los SKUs Clase A.

**Instrucciones:**

1. En la celda **F31**, calcula el promedio mensual de unidades para el primer SKU:
   ```
   =B31/24
   ```
   Copia hacia abajo hasta **F38**. Aplica formato de número con 1 decimal.

2. En la celda **G31**, calcula las líneas de picking estimadas por mes. Bajo el supuesto de que **cada unidad vendida representa 1 línea de picking**:
   ```
   =REDONDEAR(F31,0)
   ```
   Copia hacia abajo hasta **G38**. Esto redondea el promedio mensual al entero más cercano para facilitar la planificación.

   > **Nota conceptual:** En operaciones reales, el ratio líneas/unidad puede variar (un pedido puede incluir varias unidades del mismo SKU). Para esta práctica usamos el supuesto simplificado 1 unidad = 1 línea de picking. Tu instructor puede indicar un ratio diferente si el dataset lo justifica.

3. En la columna **H**, asigna la zona de picking sugerida según la clasificación ABC:
   - En **H31**, escribe:
     ```
     =SI(E31="A","Zona Prime (Alta frecuencia)",SI(E31="B","Zona Media (Frecuencia moderada)","Zona Remota (Baja frecuencia)"))
     ```
   - Copia hacia abajo hasta **H38**.

4. Ajusta el ancho de las columnas para que todo el texto sea legible: selecciona las columnas A a H, haz doble clic en el borde de cualquier encabezado de columna para **autoajustar**.

**Resultado esperado:** La tabla completa (A31:H38) muestra para cada SKU: nombre, total 24 meses, participación individual, participación acumulada, clase ABC, promedio mensual, líneas de picking estimadas por mes y zona de picking sugerida. Los SKUs Clase A tienen asignada "Zona Prime".

**Verificación:** Suma la columna G (`=SUMA(G31:G38)`) y compárala mentalmente con el total general dividido entre 24. El resultado debe ser coherente (diferencias menores de ±5 unidades son aceptables por redondeo).

---

### Paso 6 — Cruzar con Estacionalidad (Resultados de Práctica 4) para Planificación de Turnos

**Objetivo:** Identificar en qué meses los SKUs Clase A alcanzan sus picos de demanda para ajustar la planificación de turnos operativos de picking.

**Instrucciones:**

1. Ve a la hoja `Reporte_Picking` (creada en el Paso 1).

2. Crea la siguiente estructura de tabla a partir de la celda **A1**:

   | Columna | Encabezado |
   |---|---|
   | A | SKU |
   | B | Clase |
   | C | Promedio_Mensual_Unidades |
   | D | Mes_Pico_1 |
   | E | Mes_Pico_2 |
   | F | Unidades_Mes_Pico |
   | G | Incremento_vs_Promedio_% |
   | H | Recomendacion_Turno |

3. Copia únicamente los SKUs clasificados como **Clase A** desde la hoja `Analisis_ABC` hacia las columnas A, B y C de esta tabla. Usa referencias directas:
   ```
   =Analisis_ABC!A31
   ```
   (ajusta la fila según qué SKUs sean Clase A en tu dataset).

4. Para las columnas **D, E y F** (meses pico), consulta los resultados de tu Práctica 4 (índices estacionales o gráficos de demanda mensual). Identifica manualmente los 2 meses con mayor demanda para cada SKU Clase A e ingrésalos como texto (ejemplo: `"Marzo"`, `"Septiembre"`).

5. En la columna **G**, calcula el incremento porcentual del mes pico respecto al promedio:
   ```
   =(F2/C2)-1
   ```
   Aplica formato de porcentaje con 1 decimal.

6. En la columna **H**, escribe manualmente la recomendación de turno según el incremento:
   - Si el incremento es **> 50%**: `"Activar turno adicional / horas extra"`
   - Si el incremento es entre **20% y 50%**: `"Reforzar turno existente con personal de apoyo"`
   - Si el incremento es **< 20%**: `"Turno estándar suficiente"`

   > **Alternativa con fórmula (opcional):**
   > ```
   > =SI(G2>0.5,"Activar turno adicional / horas extra",SI(G2>=0.2,"Reforzar turno existente con personal de apoyo","Turno estándar suficiente"))
   > ```

7. Agrega un título en la celda **A1** (encima de la tabla, inserta una fila): escribe `PLANIFICACIÓN OPERATIVA DE PICKING — SKUs CLASE A` y aplica negrita y color de fondo azul oscuro con texto blanco.

**Resultado esperado:** La hoja `Reporte_Picking` contiene una tabla de planificación para los SKUs Clase A, mostrando sus meses pico, el incremento porcentual de demanda respecto al promedio y la recomendación de turno operativo correspondiente.

**Verificación:** Revisa que el número de filas en esta tabla coincida exactamente con la cantidad de SKUs clasificados como Clase A en la hoja `Analisis_ABC`. Cada SKU Clase A debe tener una fila completa con todos los campos llenados.

---

### Paso 7 — Documentar Recomendaciones Operativas del Almacén

**Objetivo:** Consolidar los hallazgos del análisis ABC y de picking en un bloque de recomendaciones operativas escritas, conectando el análisis cuantitativo con decisiones logísticas concretas.

**Instrucciones:**

1. En la hoja `Reporte_Picking`, debajo de la tabla de planificación (deja 3 filas de separación), crea una sección titulada **"RECOMENDACIONES OPERATIVAS"** en negrita.

2. Escribe (o adapta según tus resultados) las siguientes recomendaciones en celdas combinadas de texto libre. Puedes usar la función de combinar celdas para mayor legibilidad: selecciona A[fila]:H[fila] → **Inicio → Combinar y centrar** → cambia a **Combinar celdas** (sin centrar, para texto largo).

   Incluye al menos los siguientes 5 bloques de recomendación:

   **Recomendación 1 — Zonificación del almacén:**
   > *"Los SKUs clasificados como Clase A ([lista tus SKUs A]) deben ubicarse en la Zona Prime del almacén: área de mayor accesibilidad, menor distancia al punto de despacho y pasillos de mayor ancho para facilitar el flujo continuo de picking. Esto reduce el tiempo de desplazamiento por línea y aumenta la productividad del operario."*

   **Recomendación 2 — Reabastecimiento interno prioritario:**
   > *"Los SKUs Clase A deben tener prioridad en el reabastecimiento interno de posiciones de picking (reposición desde reserva a zona activa). Se recomienda programar reposiciones al inicio de cada turno y al mediodía para evitar posiciones vacías durante las horas pico de preparación."*

   **Recomendación 3 — Planificación de turnos en meses pico:**
   > *"Durante los meses de mayor estacionalidad identificados ([insertar meses de tu Práctica 4]), se recomienda activar [turno adicional / personal de apoyo] para los SKUs Clase A. El incremento estimado de [X]% sobre el promedio mensual implica un aumento equivalente en líneas de picking que el turno estándar no puede absorber sin afectar el nivel de servicio."*

   **Recomendación 4 — Monitoreo de disponibilidad:**
   > *"Los SKUs Clase A representan el 80% del volumen de ventas. Un quiebre de stock en cualquiera de estos productos tiene impacto directo y significativo en el nivel de servicio global. Se recomienda revisar el stock disponible de estos SKUs con frecuencia diaria durante períodos de alta demanda."*

   **Recomendación 5 — SKUs Clase C:**
   > *"Los SKUs Clase C ([lista tus SKUs C]) representan solo el 5% del volumen. Pueden ubicarse en zonas remotas del almacén con menor frecuencia de reposición. Sin embargo, deben monitorearse para detectar si alguno experimenta un cambio de comportamiento (incremento súbito de demanda) que justifique reclasificación."*

3. Al final de la sección, agrega una celda con la fecha de análisis:
   ```
   =HOY()
   ```
   Etiquétala como `Fecha de análisis:`.

**Resultado esperado:** La hoja `Reporte_Picking` contiene tanto la tabla de planificación cuantitativa como el bloque de recomendaciones operativas textuales. El documento está listo para ser presentado a un jefe de operaciones o gerente de almacén.

**Verificación:** Lee las recomendaciones en voz alta y verifica que cada una responde a la pregunta: *"¿Cómo impacta este análisis en una decisión operativa de logística?"*. Si alguna recomendación no tiene un vínculo claro con una acción operativa concreta, revísala y fortalécela.

---

### Paso 8 — Guardar y Preparar el Entregable Final

**Objetivo:** Guardar el archivo correctamente y verificar que el producto final esté completo y bien presentado.

**Instrucciones:**

1. Aplica los siguientes ajustes de presentación finales:
   - En la hoja `Analisis_ABC`, selecciona la tabla A30:H38 y aplica un **estilo de tabla** desde **Inicio → Dar formato como tabla** (elige un estilo con encabezados en azul o verde).
   - En la hoja `Reporte_Picking`, ajusta todas las columnas con **autoajuste de ancho**.
   - Congela la primera fila en ambas hojas: **Vista → Inmovilizar paneles → Inmovilizar fila superior**.

2. Agrega una hoja final llamada `Resumen_Ejecutivo` con la siguiente información resumida en formato de tabla simple (puedes escribirla manualmente):

   | Indicador | Valor |
   |---|---|
   | Total SKUs analizados | 8 |
   | SKUs Clase A (80% volumen) | [número] |
   | SKUs Clase B (15% volumen) | [número] |
   | SKUs Clase C (5% volumen) | [número] |
   | Total líneas de picking estimadas/mes (todos los SKUs) | [valor de SUMA(G31:G38)] |
   | Líneas de picking/mes solo SKUs Clase A | [calcular] |
   | % de líneas de picking concentradas en Clase A | [calcular] |
   | Meses críticos identificados (picos estacionales) | [de Práctica 4] |

3. Guarda el archivo como `Practica5_[TuNombre].xlsx`:
   - **Archivo → Guardar como → Excel Workbook (.xlsx)**
   - Confirma que el archivo tiene 4 hojas: `Datos_Limpios`, `Analisis_ABC`, `Reporte_Picking`, `Resumen_Ejecutivo`.

**Resultado esperado:** Un archivo Excel completo con 4 hojas, presentación profesional, tabla ABC con formato condicional, tabla de planificación de picking y bloque de recomendaciones operativas. El `Resumen_Ejecutivo` consolida los indicadores clave del análisis.

**Verificación:** Cierra el archivo y vuelve a abrirlo. Navega por las 4 hojas y confirma que todos los datos, fórmulas y formatos se conservaron correctamente.

---

## Validación y Pruebas del Entregable

Antes de considerar la práctica completada, verifica los siguientes criterios de calidad:

| # | Criterio de validación | Cómo verificar | ✓ |
|---|---|---|---|
| 1 | La suma de participaciones individuales (columna C) es 100% | `=SUMA(C31:C38)` → debe ser 1.0000 | ☐ |
| 2 | La participación acumulada termina exactamente en 100% en el último SKU | Valor en D38 debe ser ≥ 99.99% | ☐ |
| 3 | Todos los SKUs tienen asignada una clase (A, B o C) sin celdas vacías | Revisar rango E31:E38 | ☐ |
| 4 | El formato condicional muestra verde/amarillo/rojo correctamente | Inspección visual de E31:E38 | ☐ |
| 5 | La hoja `Reporte_Picking` contiene solo SKUs Clase A | Contar filas de datos en la tabla | ☐ |
| 6 | Las recomendaciones operativas están redactadas con datos específicos del análisis (no genéricas) | Leer cada recomendación | ☐ |
| 7 | El `Resumen_Ejecutivo` tiene todos los campos completados | Revisar que no haya celdas en blanco | ☐ |
| 8 | El archivo se guardó como `.xlsx` con las 4 hojas | Verificar en el explorador de archivos | ☐ |

> **Criterio de aprobación:** 7 de 8 criterios verificados correctamente.

---

## Solución de Problemas

### Problema 1 — La participación acumulada no llega al 100% (muestra 99.97% o valores similares)

**Síntoma:** La celda D38 (última fila del acumulado) muestra un valor como 0.9997 o 99.97% en lugar de 100.00%.

**Causa:** Este es un comportamiento normal causado por el **redondeo de punto flotante** en Excel al trabajar con divisiones de decimales. No indica un error en las fórmulas.

**Solución:**
- Si la diferencia es menor a 0.05%, es aceptable para efectos de esta práctica. Documenta la nota: *"Diferencia de redondeo menor a 0.05%, aceptable para análisis operativo."*
- Si la diferencia supera 0.1%, verifica que la fórmula en B40 sume exactamente los mismos 8 SKUs que aparecen en B31:B38 y que no haya filas duplicadas o faltantes en el dataset de origen.
- Como alternativa, puedes forzar el cierre al 100% modificando la fórmula del último SKU en D38 a `=1` (solo para presentación visual), pero mantén la fórmula original en una columna auxiliar para el cálculo real.

---

### Problema 2 — Todos los SKUs quedan clasificados como Clase A (la columna E muestra "A" en todas las filas)

**Síntoma:** Después de aplicar la fórmula `=SI(D31<=0.80,"A",SI(D31<=0.95,"B","C"))`, todas las celdas de E31:E38 muestran "A".

**Causa más frecuente:** La fórmula en la columna D (participación acumulada) no está calculando el acumulado progresivo correctamente. Es probable que la fórmula en D32:D38 esté referenciando mal las celdas, o que la columna D esté mostrando valores en formato de número entero (0–100) en lugar de decimal (0.00–1.00), haciendo que la comparación `<=0.80` sea siempre verdadera porque todos los valores son menores a 0.80 cuando deberían ser menores a 80.

**Solución:**
1. Verifica el formato de la columna D: haz clic en D31 y observa si el valor es `0.45` (formato decimal correcto) o `45` (formato entero que requiere ajuste). Si está en formato entero, ajusta la fórmula de clasificación:
   ```
   =SI(D31<=80,"A",SI(D31<=95,"B","C"))
   ```
   O bien, selecciona D31:D38 y aplica formato de porcentaje para estandarizar.
2. Si el problema persiste, verifica que D32 contenga `=D31+C32` (no `=C31+C32` ni otra variación). La fórmula debe sumar el acumulado anterior más la participación individual del SKU actual.
3. Confirma que la tabla esté ordenada de mayor a menor en la columna B antes de calcular los acumulados. Si el ordenamiento se realizó después de escribir las fórmulas de acumulado, es posible que las referencias se hayan desplazado incorrectamente.

---

## Limpieza y Cierre

Al finalizar la práctica, realiza las siguientes acciones:

1. **Guarda el archivo final** una última vez con `Ctrl+S` y confirma el nombre `Practica5_[TuNombre].xlsx`.

2. **Entrega o comparte el archivo** según las instrucciones del instructor (carpeta compartida, correo electrónico o plataforma del curso).

3. **Cierra Excel** o minimiza la ventana si continuarás con la Práctica 6 inmediatamente.

4. **No elimines ni modifiques** el archivo después de guardarlo: la Práctica 6 utilizará el dataset enriquecido provisto por el instructor, pero las clasificaciones ABC y los índices estacionales de esta práctica serán referenciados en las Prácticas 7 y 8.

> **Nota de secuencialidad:** El archivo `Practica5_[TuNombre].xlsx` es insumo directo para las Prácticas 7 y 8. Guárdalo en una ubicación de fácil acceso y con nombre claro.

---

## Resumen

### Lo que aprendiste en esta práctica

En esta práctica aplicaste el **análisis de clasificación ABC** — una de las herramientas más utilizadas en logística para priorizar recursos operativos según el impacto real de cada producto en el volumen de operaciones. Los pasos que ejecutaste siguieron una lógica directamente conectada con los conceptos de la Lección 1.1:

1. **Rotación de inventario:** Los SKUs Clase A son los de mayor rotación. Mantenerlos bien abastecidos y en zonas de fácil acceso es crítico para sostener la velocidad operativa.

2. **Nivel de servicio:** Concentrar el esfuerzo operativo (zonas prime, reposición prioritaria, turnos reforzados) en los SKUs Clase A garantiza que el 80% del volumen de pedidos se atienda con máxima eficiencia, elevando el nivel de servicio global.

3. **Quiebres de stock:** Los SKUs Clase A son los que generan mayor impacto cuando se produce un quiebre. El cruce con la estacionalidad de la Práctica 4 permite anticipar los períodos de mayor riesgo y activar medidas preventivas.

4. **Planificación de picking:** La estimación de líneas de picking por mes transforma el análisis estadístico en un insumo operativo directo: cuántas personas necesito, en qué turno y en qué zona del almacén.

### Conexión con las próximas prácticas

| Práctica | Cómo utiliza los resultados de esta práctica |
|---|---|
| Práctica 6 | Los SKUs Clase A identificados serán los foco del análisis de eventos extraordinarios (COVID, promociones) en el dataset de 36 meses |
| Práctica 7 | Los modelos de pronóstico se construirán prioritariamente para los SKUs Clase A, donde el error de pronóstico tiene mayor impacto operativo |
| Práctica 8 | Copilot generará recomendaciones de reabastecimiento diferenciadas por clase ABC, usando la segmentación construida en esta práctica |

### Recursos de referencia

- **Chopra, S. & Meindl, P.** — *Supply Chain Management: Strategy, Planning, and Operation* (Pearson) — Capítulo 11: Managing Uncertainty in a Supply Chain.
- **Ballou, R. H.** — *Business Logistics/Supply Chain Management* (Pearson) — Capítulo 9: Inventory Management Fundamentals.
- **CSCMP Glosario** — Definiciones de clasificación ABC, rotación de inventario y nivel de servicio: [cscmp.org](https://cscmp.org/CSCMP/Educate/SCM_Definitions_and_Glossary_of_Terms.aspx)
- **Instituto Aragonés de Fomento** — Guía práctica de gestión de inventarios y previsión de demanda: sección "Clasificación ABC de artículos".

---

> **Reflexión final para el participante:** Antes de cerrar el archivo, responde mentalmente: *¿Si mañana tuvieras que reorganizar el almacén de tu empresa, qué productos moverías a la zona de mayor acceso y por qué? ¿Cómo cambiaría esa decisión si los datos de demanda cambian el próximo trimestre?* Esta es la pregunta que el forecasting y la clasificación ABC buscan responder de forma continua y basada en datos.

---

---LAB_START---
LAB_ID: 01-00-06
---MARKDOWN---
# Práctica 6 — Detección de Eventos Extraordinarios (Promociones, Pandemia, Quiebres de Stock)

## Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 1 — Fundamentos de Forecasting Logístico |
| **Práctica número** | 6 de 10 |
| **Archivo de entrada** | `Dataset_Practica6_36meses.xlsx` (proporcionado por el instructor) |
| **Archivo de salida** | `Dataset_Practica6_Baseline_Limpio.xlsx` |

---

## Descripción General

En las prácticas anteriores construiste indicadores estadísticos y clasificaste SKUs según su estabilidad de demanda. Sin embargo, ese historial puede contener meses atípicos —promociones agresivas, el impacto del COVID-19, o periodos donde el producto simplemente no estaba disponible en bodega— que, si no se tratan, contaminan el Baseline y producen pronósticos distorsionados. En esta práctica aplicarás una metodología sistemática para detectar, clasificar y tratar esos eventos extraordinarios antes de construir el pronóstico final. El producto resultante es un **dataset depurado** con un Baseline limpio, listo para ser utilizado en las Prácticas 7 y 8.

---

## Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Identificar meses atípicos en el historial de demanda utilizando el criterio estadístico de **promedio ± 2 desviaciones estándar**.
- [ ] Clasificar cada evento extraordinario detectado en una de tres categorías: Promoción/Iniciativa, Evento Externo/Disrupción o Quiebre de Stock.
- [ ] Aplicar el tratamiento correcto para cada tipo de evento: sustitución por promedio móvil (quiebres), marcado con flag (promociones) o evaluación caso a caso (eventos externos).
- [ ] Cuantificar el impacto del ajuste comparando el Baseline sin tratamiento versus el Baseline depurado.
- [ ] Explicar por qué incluir eventos extraordinarios sin tratamiento distorsiona el pronóstico y genera errores operativos en inventario y nivel de servicio.

---

## Prerrequisitos

### Conocimiento previo

| Requisito | Descripción |
|---|---|
| Prácticas 1–5 completadas | El participante debe tener familiaridad con el dataset base de 8 SKUs y los indicadores estadísticos calculados en prácticas anteriores. |
| Componentes del comportamiento de la demanda | Comprensión de tendencia, estacionalidad, variabilidad y eventos extraordinarios (temas 1.3 y 1.4 del curso). |
| Funciones condicionales en Excel | Uso operativo de `SI()`, `Y()`, `O()`, `PROMEDIO()`, `DESVEST.M()` y `ABS()`. |
| Concepto de Baseline de pronóstico | Entender que el Baseline representa la demanda "normal" libre de efectos extraordinarios. |

### Acceso y archivos necesarios

| Recurso | Estado requerido |
|---|---|
| Microsoft Excel 2016 o superior | ✅ Instalado y funcional |
| `Dataset_Practica6_36meses.xlsx` | ✅ Descargado desde la plataforma del curso |
| `Registro_Eventos_Negocio.xlsx` | ✅ Descargado desde la plataforma del curso |
| Conexión a internet (opcional en esta práctica) | Para acceso a Copilot si se requiere apoyo en interpretación |

---

## Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior (Microsoft 365 recomendado) | Las fórmulas usan `DESVEST.M` y `PRONOSTICO.LINEAL`; verificar compatibilidad |
| Navegador web | Chrome o Edge (versión 2024) | Solo si se usa Copilot como apoyo |

### Configuración inicial — Antes de comenzar

Ejecuta los siguientes pasos de preparación **antes** de iniciar el cronómetro de 12 minutos:

1. Abre el archivo `Dataset_Practica6_36meses.xlsx`.
2. Abre el archivo `Registro_Eventos_Negocio.xlsx` en una segunda ventana de Excel.
3. Verifica que el dataset principal tenga la siguiente estructura:
   - **Columna A:** `Mes` (formato `YYYY-MM`, desde `2020-01` hasta `2022-12`, 36 filas de datos).
   - **Columnas B a I:** Demanda mensual para 8 SKUs (`SKU_01` a `SKU_08`).
   - **Fila 1:** Encabezados.
   - **Filas 2 a 37:** Datos de demanda (36 meses).
4. Verifica que el `Registro_Eventos_Negocio.xlsx` tenga columnas: `Mes`, `SKU`, `Tipo_Evento`, `Descripcion`.
5. Asegúrate de que la hoja activa del dataset principal se llame `Historial_Demanda`.

> **Nota del instructor:** Si el participante no completó las Prácticas 1–5, puede usar el archivo checkpoint `Checkpoint_P5_Completado.xlsx` que contiene los indicadores calculados hasta la Práctica 5.

---

## Pasos del Laboratorio

### Paso 1 — Preparar la hoja de trabajo para el análisis de eventos

**Objetivo:** Crear el espacio de trabajo estructurado donde se realizará toda la detección y clasificación de eventos extraordinarios.

**Instrucciones:**

1. En el archivo `Dataset_Practica6_36meses.xlsx`, haz clic en la pestaña de la hoja `Historial_Demanda`.
2. Crea una nueva hoja haciendo clic derecho sobre la pestaña → **Insertar** → **Hoja de cálculo**. Nómbrala `Analisis_Eventos`.
3. En la hoja `Analisis_Eventos`, construye la siguiente estructura de encabezados en la **Fila 1**:

```
A1: Mes
B1: SKU
C1: Demanda_Real
D1: Promedio_SKU
E1: DesvEst_SKU
F1: Limite_Superior
G1: Limite_Inferior
H1: Es_Atipico (SI/NO)
I1: Tipo_Evento
J1: Descripcion_Evento
K1: Demanda_Ajustada
L1: Flag_Iniciativa
```

4. Aplica **negrita** a toda la Fila 1 y un color de relleno azul claro (`RGB: 173, 216, 230`) para distinguir los encabezados.
5. Congela la fila de encabezados: **Vista** → **Inmovilizar paneles** → **Inmovilizar fila superior**.

**Resultado esperado:** Una hoja `Analisis_Eventos` con 12 columnas encabezadas, formato visual claro y fila superior inmovilizada.

**Verificación:** Confirma que puedes desplazarte hacia abajo en la hoja y los encabezados permanecen visibles. La hoja debe estar vacía debajo de la Fila 1 (los datos se ingresarán en los pasos siguientes).

---

### Paso 2 — Calcular Promedio y Desviación Estándar por SKU

**Objetivo:** Establecer los parámetros estadísticos que definen el rango de demanda "normal" para cada SKU, base del criterio de detección de atípicos.

**Instrucciones:**

1. Regresa a la hoja `Historial_Demanda`.
2. Debajo de los 36 meses de datos (en la **Fila 39**), crea una sección de estadísticos. En la celda `A39` escribe: `Promedio_36m`. En `A40` escribe: `DesvEst_36m`.
3. Para el **SKU_01** (columna B), ingresa las siguientes fórmulas:

```excel
B39: =PROMEDIO(B2:B37)
B40: =DESVEST.M(B2:B37)
```

4. Copia las fórmulas de `B39:B40` hacia las columnas `C` hasta `I` (para SKU_02 a SKU_08). Selecciona `B39:B40`, copia con `Ctrl+C`, selecciona el rango `C39:I40` y pega con `Ctrl+V`.
5. Formatea las celdas `B39:I40` con **2 decimales** y color de relleno amarillo claro para distinguirlas de los datos históricos.
6. En la **Fila 41** calcula el **Límite Superior** y en la **Fila 42** el **Límite Inferior**:

```excel
B41: =B39+(2*B40)    ' Límite Superior = Promedio + 2 * DesvEst
B42: =B39-(2*B40)    ' Límite Inferior = Promedio - 2 * DesvEst
```

7. Copia `B41:B42` hacia las columnas `C` a `I`.
8. En `A41` escribe: `Limite_Superior_(+2σ)`. En `A42` escribe: `Limite_Inferior_(-2σ)`.

> **Concepto clave:** El criterio **Promedio ± 2 desviaciones estándar** captura aproximadamente el 95% de los valores de una distribución normal. Cualquier mes fuera de este rango es estadísticamente inusual y merece revisión. En logística, esto no significa automáticamente que el valor sea "incorrecto" —puede ser una promoción legítima— pero sí que requiere clasificación antes de incluirlo en el Baseline.

**Resultado esperado:** Las filas 39 a 42 de la hoja `Historial_Demanda` muestran promedio, desviación estándar, límite superior e inferior para cada uno de los 8 SKUs.

**Verificación:** Para el SKU_01, verifica manualmente que `B41 = B39 + 2*B40`. Si el promedio de SKU_01 es, por ejemplo, 850 unidades y la desviación estándar es 120, el límite superior debe ser 1,090 y el límite inferior 610. Los valores exactos dependen del dataset proporcionado por el instructor.

---

### Paso 3 — Identificar automáticamente los meses atípicos con fórmula SI

**Objetivo:** Usar una fórmula condicional para marcar automáticamente cada mes donde la demanda de un SKU esté fuera del rango normal establecido en el Paso 2.

**Instrucciones:**

1. Regresa a la hoja `Historial_Demanda`.
2. En la columna **J** (o la primera columna disponible después de los datos de SKUs), en la celda `J1` escribe el encabezado: `SKU_01_Atipico`.
3. En la celda `J2` ingresa la siguiente fórmula:

```excel
=SI(O(B2>B$41, B2<B$42), "ATIPICO", "Normal")
```

> **Explicación de la fórmula:**
> - `O(B2>B$41, B2<B$42)`: Evalúa si la demanda del mes está **por encima** del límite superior **O por debajo** del límite inferior.
> - `B$41` y `B$42`: Referencia absoluta en fila (con `$`) para que al copiar hacia abajo, siempre apunte a las filas de límites calculadas en el Paso 2.
> - Si la condición es verdadera, devuelve `"ATIPICO"`; si no, devuelve `"Normal"`.

4. Copia la fórmula de `J2` hacia abajo hasta `J37` (los 36 meses).
5. Repite el proceso para los demás SKUs en columnas `K` a `Q`, ajustando la referencia de columna:
   - `K1`: `SKU_02_Atipico` → fórmula: `=SI(O(C2>C$41, C2<C$42), "ATIPICO", "Normal")`
   - `L1`: `SKU_03_Atipico` → fórmula: `=SI(O(D2>D$41, D2<D$42), "ATIPICO", "Normal")`
   - Continúa hasta `Q1: SKU_08_Atipico`

6. Aplica **Formato Condicional** para resaltar visualmente los atípicos:
   - Selecciona el rango `J2:Q37`.
   - Ve a **Inicio** → **Formato condicional** → **Nueva regla**.
   - Selecciona **"Dar formato únicamente a las celdas que contengan"**.
   - Configura: **Texto específico** → **que contiene** → `ATIPICO`.
   - Establece el formato: relleno **rojo claro** (`RGB: 255, 199, 206`) y fuente en **negrita roja**.
   - Haz clic en **Aceptar**.

**Resultado esperado:** Las columnas J a Q muestran "Normal" o "ATIPICO" para cada mes y SKU. Las celdas con "ATIPICO" aparecen resaltadas en rojo, permitiendo identificar visualmente los periodos problemáticos de un solo vistazo.

**Verificación:** Cuenta el número total de celdas marcadas como "ATIPICO" en todo el rango. En un dataset de 36 meses × 8 SKUs (288 celdas), se espera encontrar entre 8 y 20 atípicos considerando el impacto del COVID-19 (2020-2021) y las promociones documentadas. Si encuentras más de 30 atípicos, verifica que los límites del Paso 2 estén correctamente calculados.

---

### Paso 4 — Cruzar atípicos con el Registro de Eventos para clasificar

**Objetivo:** Para cada mes marcado como "ATIPICO", determinar la causa raíz cruzando con el archivo de registro de eventos del negocio y asignar una de las tres categorías de clasificación.

**Instrucciones:**

1. Abre el archivo `Registro_Eventos_Negocio.xlsx`. Revisa su contenido. Debe listar eventos como:
   - `2020-03 | SKU_03 | Evento_Externo | Inicio pandemia COVID-19 — caída abrupta de demanda`
   - `2020-04 | SKU_05 | Evento_Externo | COVID-19 — quiebre de suministro proveedor`
   - `2021-06 | SKU_01 | Promocion | Campaña descuento 30% — pico de demanda`
   - `2021-11 | SKU_02 | Quiebre_Stock | Falta de materia prima — demanda reprimida`

2. Regresa a la hoja `Analisis_Eventos` del archivo principal. En las columnas `A` a `C`, copia manualmente (o usa `BUSCARV`) los datos de los meses identificados como "ATIPICO":
   - **Columna A:** Mes
   - **Columna B:** SKU
   - **Columna C:** Demanda real del mes atípico

   Para poblar esta tabla eficientemente, usa la siguiente fórmula en una celda auxiliar para extraer los atípicos (puedes hacerlo manualmente si son pocos):

   > **Opción manual (recomendada por tiempo):** Filtra cada columna de atípicos en la hoja `Historial_Demanda` usando **Datos** → **Filtro**, selecciona solo las celdas con "ATIPICO" y copia los registros correspondientes a la hoja `Analisis_Eventos`.

3. Una vez que tengas la lista de atípicos en `Analisis_Eventos`, completa las columnas de clasificación para cada registro:

   **Columna I — Tipo_Evento:** Ingresa manualmente el tipo según el cruce con el Registro de Eventos:
   - `Promocion` — si el registro de eventos documenta una acción comercial (descuento, campaña, lanzamiento).
   - `Evento_Externo` — si corresponde a una disrupción fuera del control de la empresa (pandemia, crisis de suministro, catástrofe natural).
   - `Quiebre_Stock` — si la demanda es anormalmente **baja** y hay documentación de falta de disponibilidad de producto.
   - `Sin_Clasificar` — si el mes es atípico estadísticamente pero no aparece en el registro de eventos (requiere investigación adicional).

   **Columna J — Descripcion_Evento:** Copia la descripción del Registro de Eventos o escribe una descripción breve.

4. Para los meses atípicos que **no aparecen** en el Registro de Eventos, revisa si la demanda es alta (posible promoción no documentada) o baja (posible quiebre no registrado) y clasifícalos como `Sin_Clasificar` con una nota en la columna J.

> **Regla operativa de clasificación:**
> - Demanda **muy alta** (> Límite Superior) + evento documentado = **Promoción/Iniciativa**
> - Demanda **muy alta** + sin evento documentado = **Sin_Clasificar** (investigar)
> - Demanda **muy baja** (< Límite Inferior) + quiebre documentado = **Quiebre_Stock**
> - Demanda **muy baja** + pandemia/crisis = **Evento_Externo**
> - Demanda **muy baja** + sin explicación = **Sin_Clasificar** (investigar)

**Resultado esperado:** La tabla en `Analisis_Eventos` tiene todos los meses atípicos clasificados en una de las cuatro categorías, con descripción del evento asociado.

**Verificación:** Verifica que todos los eventos del `Registro_Eventos_Negocio.xlsx` que corresponden a meses marcados como "ATIPICO" estén clasificados. No debe haber registros con `Tipo_Evento` vacío. Los eventos de tipo `Sin_Clasificar` son aceptables si el registro de eventos no los documenta, pero deben tener una nota en la columna J.

---

### Paso 5 — Aplicar el tratamiento correcto según el tipo de evento

**Objetivo:** Para cada tipo de evento clasificado, aplicar la transformación correspondiente sobre la demanda para construir un Baseline limpio que refleje la demanda "estructural" del SKU.

**Instrucciones:**

#### Tratamiento A — Quiebres de Stock: Sustitución por promedio de los 3 meses previos

Para cada registro clasificado como `Quiebre_Stock`:

1. En la columna `K` (`Demanda_Ajustada`) de la hoja `Analisis_Eventos`, calcula el promedio de los 3 meses inmediatamente anteriores al quiebre. Usa la siguiente fórmula (ajusta las referencias según la posición real de los datos):

```excel
' Ejemplo: si el quiebre es en el mes de la fila actual,
' y los 3 meses previos están en filas anteriores de la hoja Historial_Demanda:
=PROMEDIO(INDIRECTO("Historial_Demanda!B"&(COINCIDIR(A2,Historial_Demanda!A:A,0)-3)&":B"&(COINCIDIR(A2,Historial_Demanda!A:A,0)-1)))
```

> **Alternativa simplificada (recomendada por tiempo):** Busca manualmente en la hoja `Historial_Demanda` los 3 meses anteriores al quiebre para el SKU correspondiente, y escribe directamente el promedio calculado en la columna `K`.

**Ejemplo concreto:** Si SKU_02 tuvo un quiebre en `2021-11` con demanda real de 120 unidades, y los meses anteriores fueron: `2021-08 = 780`, `2021-09 = 810`, `2021-10 = 795`, entonces:
```
Demanda_Ajustada = (780 + 810 + 795) / 3 = 795
```

2. En la columna `L` (`Flag_Iniciativa`), escribe `0` para los quiebres (no son iniciativas comerciales).

#### Tratamiento B — Promociones/Iniciativas: Mantener valor real con flag

Para cada registro clasificado como `Promocion`:

1. En la columna `K` (`Demanda_Ajustada`), **copia el valor real** de la columna `C` (`Demanda_Real`). El valor NO se modifica.

```excel
=C2   ' Copia el valor real sin modificación
```

2. En la columna `L` (`Flag_Iniciativa`), escribe `1`. Este flag indica que el mes debe ser usado en la **capa de Iniciativas** del pronóstico (no en el Baseline), y será excluido del cálculo del Baseline en el Paso 6.

> **Razón operativa:** Las promociones representan demanda real que ocurrió, pero no es la demanda "base" del producto. Si la incluimos en el Baseline, el pronóstico futuro sobreestimará la demanda en meses normales, generando sobreinventario. Sin embargo, el valor se guarda con flag para usarlo en la planificación de futuras campañas similares.

#### Tratamiento C — Eventos Externos/Disrupciones: Evaluación caso a caso

Para cada registro clasificado como `Evento_Externo`:

1. Evalúa si el evento fue **temporal y no repetible** (ej. confinamiento por COVID-19 en 2020) o si podría repetirse (ej. crisis de suministro recurrente).
2. Para eventos **no repetibles** (como el inicio del COVID-19 en marzo-junio 2020):
   - En columna `K`, aplica el mismo tratamiento que los quiebres: **promedio de los 3 meses previos**.
   - En columna `L`, escribe `0`.
3. Para eventos **potencialmente repetibles** o de duración prolongada (ej. demanda alterada durante todo 2020):
   - En columna `K`, copia el valor real con `=C2`.
   - En columna `L`, escribe `2` (flag especial para "Evento Externo mantenido").
   - Agrega una nota en columna `J` indicando la razón de mantener el valor.

> **Guía práctica para el periodo COVID (2020):** Para este dataset, aplica el siguiente criterio simplificado:
> - **2020-03 a 2020-06:** Tratar como eventos no repetibles → sustituir por promedio de 3 meses previos.
> - **2020-07 a 2020-12:** Si la demanda se normalizó, mantener el valor real con flag `2`.

**Resultado esperado:** La columna `K` (`Demanda_Ajustada`) tiene valores en todos los registros atípicos:
- Quiebres → valor sustituido (promedio 3 meses previos).
- Promociones → valor real conservado (con flag `1` en columna `L`).
- Eventos externos no repetibles → valor sustituido.
- Eventos externos mantenidos → valor real (con flag `2` en columna `L`).

**Verificación:** Confirma que ninguna celda de la columna `K` esté vacía para los registros atípicos. Para los quiebres de stock, el valor ajustado debe ser **mayor** que la demanda real registrada (porque el quiebre suprimió la demanda real). Para las promociones, el valor ajustado debe ser **igual** a la demanda real.

---

### Paso 6 — Reconstruir el Baseline limpio y comparar con el Baseline sin tratamiento

**Objetivo:** Construir el dataset final depurado reemplazando los valores atípicos tratados en el historial original, y cuantificar el impacto del ajuste sobre el Baseline de cada SKU.

**Instrucciones:**

1. Crea una nueva hoja en el archivo y nómbrala `Baseline_Limpio`.

2. Copia la estructura completa de la hoja `Historial_Demanda` (encabezados + 36 meses) a la nueva hoja `Baseline_Limpio`. Usa **Pegado especial → Solo valores** para no arrastrar fórmulas.

3. En la hoja `Baseline_Limpio`, reemplaza los valores de los meses atípicos con sus valores ajustados:
   - Para cada registro de la hoja `Analisis_Eventos` que tenga `Flag_Iniciativa = 1` (Promociones): **reemplaza el valor** en la hoja `Baseline_Limpio` con el **promedio de los 3 meses previos** (igual que el tratamiento de quiebres). Esto excluye la demanda promocional del Baseline.
   - Para los registros con `Flag_Iniciativa = 0` (Quiebres y Eventos Externos no repetibles): el valor ya fue ajustado en el Paso 5; reemplaza el valor en `Baseline_Limpio` con el valor de la columna `K` de `Analisis_Eventos`.
   - Para los registros con `Flag_Iniciativa = 2` (Eventos Externos mantenidos): **no modifiques** el valor en `Baseline_Limpio`; se mantiene el valor real.

4. Una vez completado el reemplazo, crea una tabla comparativa en la parte inferior de la hoja `Baseline_Limpio` (a partir de la Fila 42). Construye la siguiente comparación por SKU:

```
A42: SKU
B42: Promedio_Baseline_Original
C42: Promedio_Baseline_Limpio
D42: Diferencia_Absoluta
E42: Diferencia_Porcentual
F42: Impacto_Operativo
```

5. Para el SKU_01 (columna B del historial), calcula:

```excel
B43: =PROMEDIO(Historial_Demanda!B2:B37)          ' Baseline original (con atípicos)
C43: =PROMEDIO(B2:B37)                             ' Baseline limpio (en esta hoja)
D43: =ABS(C43-B43)                                 ' Diferencia absoluta en unidades
E43: =ABS((C43-B43)/B43)*100                       ' Diferencia porcentual
F43: =SI(E43>5,"ALTO - Revisar pronóstico",SI(E43>2,"MEDIO - Monitorear","BAJO - Aceptable"))
```

6. Copia las fórmulas de la Fila 43 para los SKUs 02 a 08 (columnas C a I), ajustando las referencias de columna.

7. Formatea la columna `F` con Formato Condicional:
   - `"ALTO"` → relleno rojo.
   - `"MEDIO"` → relleno amarillo.
   - `"BAJO"` → relleno verde.

> **Interpretación operativa:** Una diferencia porcentual superior al 5% entre el Baseline original y el Baseline limpio indica que los eventos extraordinarios estaban distorsionando significativamente el pronóstico. En términos logísticos, esto significa que las órdenes de compra calculadas sobre el Baseline original estarían sobredimensionadas (si había muchas promociones) o subdimensionadas (si había muchos quiebres sin ajustar), afectando directamente el nivel de servicio y la rotación de inventario.

**Resultado esperado:** La hoja `Baseline_Limpio` contiene 36 meses de demanda depurada para los 8 SKUs, con la tabla comparativa mostrando el impacto cuantificado del ajuste. Al menos 2 o 3 SKUs deben mostrar diferencias superiores al 2% dada la inclusión del periodo COVID en el dataset.

**Verificación:** Suma la demanda total de los 36 meses en la hoja `Historial_Demanda` para un SKU y compárala con la suma en `Baseline_Limpio`. Las sumas **no deben ser iguales** si hubo ajustes (esto confirma que los reemplazos se aplicaron correctamente). Si las sumas son idénticas, revisa que los valores atípicos hayan sido efectivamente reemplazados y no solo copiados.

---

### Paso 7 — Documentar el proceso y preparar el archivo de salida

**Objetivo:** Consolidar toda la documentación del proceso de detección y tratamiento en un formato que sirva de insumo para las Prácticas 7 y 8.

**Instrucciones:**

1. Crea una última hoja llamada `Resumen_Eventos` con la siguiente tabla resumen:

```
A1: Resumen de Eventos Extraordinarios Detectados y Tratados
A3: Total de meses analizados:          B3: =36*8   ' 288 combinaciones mes-SKU
A4: Total de meses atípicos detectados: B4: [contar manualmente desde Analisis_Eventos]
A5: % de atípicos sobre total:          B5: =B4/B3*100
A7: Desglose por tipo:
A8: Promociones/Iniciativas:            B8: [contar registros con Flag=1]
A9: Quiebres de Stock:                  B9: [contar registros con Tipo=Quiebre_Stock]
A10: Eventos Externos:                  B10: [contar registros con Tipo=Evento_Externo]
A11: Sin Clasificar:                    B11: [contar registros con Tipo=Sin_Clasificar]
A13: SKU con mayor impacto de ajuste:   B13: [SKU con mayor diferencia porcentual]
A14: Diferencia porcentual máxima:      B14: [valor de la diferencia máxima]
```

2. Guarda el archivo con el nombre `Dataset_Practica6_Baseline_Limpio.xlsx` usando **Guardar como** (`Ctrl+Mayús+S`).

3. Verifica que el archivo guardado contiene exactamente las siguientes hojas:
   - `Historial_Demanda` (datos originales, sin modificar).
   - `Analisis_Eventos` (tabla de clasificación y tratamiento).
   - `Baseline_Limpio` (dataset depurado listo para pronóstico).
   - `Resumen_Eventos` (tabla resumen del proceso).

**Resultado esperado:** El archivo `Dataset_Practica6_Baseline_Limpio.xlsx` está guardado con las 4 hojas documentadas, listo para ser utilizado como insumo en la Práctica 7 (construcción del modelo de pronóstico).

**Verificación:** Cierra el archivo y vuelve a abrirlo. Confirma que las 4 hojas están presentes y que los datos de `Baseline_Limpio` no contienen los valores atípicos originales en los meses que fueron ajustados.

---

## Validación y Pruebas

Una vez completados los 7 pasos, realiza las siguientes verificaciones de integridad del trabajo:

### Lista de verificación final

| # | Verificación | Criterio de éxito |
|---|---|---|
| V1 | Estadísticos calculados | Las filas 39-42 de `Historial_Demanda` tienen valores de promedio, desviación estándar y límites para los 8 SKUs. |
| V2 | Identificación de atípicos | Las columnas J-Q de `Historial_Demanda` muestran "ATIPICO" o "Normal" para cada mes, con formato condicional activo. |
| V3 | Clasificación completa | Todos los registros en `Analisis_Eventos` tienen un valor en la columna `Tipo_Evento` (sin celdas vacías). |
| V4 | Tratamiento aplicado | La columna `Demanda_Ajustada` tiene valores en todos los registros atípicos. Los quiebres tienen valores mayores a la demanda real registrada. |
| V5 | Baseline depurado | La hoja `Baseline_Limpio` tiene valores distintos a `Historial_Demanda` en al menos los meses correspondientes a quiebres y eventos COVID. |
| V6 | Impacto cuantificado | La tabla comparativa (Fila 42 en adelante de `Baseline_Limpio`) muestra diferencias porcentuales para los 8 SKUs. Al menos un SKU debe tener diferencia > 2%. |
| V7 | Archivo guardado correctamente | El archivo `Dataset_Practica6_Baseline_Limpio.xlsx` tiene 4 hojas y está guardado en la carpeta de trabajo del curso. |

### Pregunta de reflexión para el cierre

> **¿Cómo impacta este análisis en una decisión operativa de logística?**
>
> Considera el siguiente escenario: si el SKU_03 tuvo una demanda media de 950 unidades/mes en el Baseline original (con quiebres de stock incluidos), pero el Baseline limpio muestra una demanda real de 1,150 unidades/mes, ¿qué implicación tiene esto para el cálculo del stock de seguridad y el punto de reorden? ¿Qué error operativo se habría cometido si se hubiera usado el Baseline sin depurar?

*Respuesta esperada:* El Baseline sin depurar subestimaría la demanda real en 200 unidades/mes. El stock de seguridad calculado sería insuficiente, el punto de reorden se activaría demasiado tarde y la empresa experimentaría quiebres de stock recurrentes, deteriorando el nivel de servicio. El error se perpetuaría en cada ciclo de planificación porque el modelo "aprendería" una demanda artificialmente baja.

---

## Solución de Problemas

### Problema 1 — La fórmula `DESVEST.M` devuelve `#¡NOMBRE?` o `#NAME?`

**Síntoma:** Al ingresar `=DESVEST.M(B2:B37)` en la hoja `Historial_Demanda`, la celda muestra el error `#¡NOMBRE?` (o `#NAME?` en versiones en inglés).

**Causa:** La versión de Excel instalada es anterior a Excel 2010, o el idioma de las funciones no está configurado en español. La función `DESVEST.M` corresponde a `STDEV.S` en inglés y no existe en versiones anteriores a Excel 2010.

**Solución:**
1. Verifica la versión de Excel: **Archivo** → **Cuenta** → **Acerca de Excel**.
2. Si la versión es Excel 2007 o anterior, usa la función equivalente:
   ```excel
   =DESVEST(B2:B37)    ' Equivalente en versiones antiguas
   ```
3. Si la versión es correcta pero el idioma de las funciones es inglés (instalación en inglés), usa:
   ```excel
   =STDEV.S(B2:B37)    ' Nombre en inglés para Excel en inglés
   ```
4. Si el problema persiste, verifica que el rango `B2:B37` contiene valores numéricos y no texto. Usa **Datos** → **Texto en columnas** para convertir si es necesario.

---

### Problema 2 — Los valores en `Baseline_Limpio` son idénticos a `Historial_Demanda` después del reemplazo

**Síntoma:** Al comparar la suma total de demanda entre `Historial_Demanda` y `Baseline_Limpio` para un SKU con quiebres documentados, los valores son exactamente iguales. La tabla comparativa muestra 0% de diferencia para todos los SKUs.

**Causa:** Al copiar la hoja `Historial_Demanda` a `Baseline_Limpio`, se usó **Pegado normal** (con fórmulas) en lugar de **Pegado especial → Solo valores**. Las celdas en `Baseline_Limpio` siguen referenciando los valores originales de `Historial_Demanda` a través de fórmulas vinculadas, por lo que cualquier modificación en `Baseline_Limpio` no tiene efecto real, o bien los valores reemplazados están siendo sobreescritos por las fórmulas vinculadas.

**Solución:**
1. Elimina la hoja `Baseline_Limpio` actual.
2. Regresa a `Historial_Demanda`, selecciona el rango de datos completo (`A1:I37`), copia con `Ctrl+C`.
3. Crea una nueva hoja `Baseline_Limpio`.
4. En la celda `A1` de la nueva hoja, usa **Inicio** → **Pegar** → **Pegado especial** → selecciona **Valores** (o presiona `Ctrl+Alt+V` → selecciona `V` → `Aceptar`).
5. Confirma que las celdas no tienen fórmulas (haz clic en cualquier celda y verifica que la barra de fórmulas muestra un número, no una fórmula como `=Historial_Demanda!B2`).
6. Ahora procede a reemplazar los valores atípicos según el Paso 6.

---

## Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos de cierre:

1. **Guarda el archivo final** una última vez con `Ctrl+G` para asegurar que todos los cambios están guardados.
2. **Cierra el archivo** `Registro_Eventos_Negocio.xlsx` sin modificarlo (este archivo es de solo lectura para esta práctica).
3. **Verifica la ubicación del archivo de salida:** Confirma que `Dataset_Practica6_Baseline_Limpio.xlsx` está guardado en la carpeta `Curso_Forecasting/Practica6/` (o la ruta indicada por el instructor).
4. **No elimines** el archivo `Dataset_Practica6_36meses.xlsx` original; será necesario para comparaciones en prácticas posteriores.
5. Si utilizaste Microsoft Copilot u otra herramienta de IA como apoyo, cierra las pestañas del navegador relacionadas para liberar memoria.

> **Para el instructor:** El archivo `Dataset_Practica6_Baseline_Limpio.xlsx` generado en esta práctica es el insumo directo de la **Práctica 7** (selección y evaluación de modelos de pronóstico). Asegúrese de que todos los participantes tengan este archivo guardado correctamente antes de iniciar la siguiente sesión. Tenga disponible el archivo checkpoint `Checkpoint_P6_Completado.xlsx` para participantes que no lograron completar el Paso 6.

---

## Resumen

### Lo que lograste en esta práctica

En esta práctica aplicaste una metodología estructurada de **detección, clasificación y tratamiento de eventos extraordinarios** en el historial de demanda, un paso crítico que separa un pronóstico logístico profesional de una simple proyección estadística.

| Actividad realizada | Herramienta/Técnica aplicada | Impacto logístico |
|---|---|---|
| Cálculo de rangos de demanda normal | `PROMEDIO()` + `DESVEST.M()` ± 2σ | Define qué es "normal" para cada SKU |
| Identificación automática de atípicos | Fórmula `SI(O(...))` + Formato Condicional | Detecta meses que distorsionan el Baseline |
| Clasificación por tipo de evento | Cruce con Registro de Eventos | Determina el tratamiento correcto |
| Sustitución de quiebres | Promedio de 3 meses previos | Recupera la demanda real reprimida |
| Flagging de promociones | Flag `1` en columna L | Preserva el valor para capa de Iniciativas |
| Tratamiento de eventos externos | Evaluación caso a caso (COVID-19) | Evita que la pandemia distorsione el Baseline |
| Comparación Baseline original vs. limpio | `ABS()` + diferencia porcentual | Cuantifica el impacto del ajuste |

### Conexión con el modelo de pronóstico (Baseline + Estacionalidad + Iniciativas)

Esta práctica operacionaliza la separación de capas del modelo de pronóstico:

- El **Baseline limpio** generado hoy es la base sobre la que se calculará la tendencia estructural del SKU.
- Los meses con **Flag = 1 (Promociones)** serán utilizados en la capa de **Iniciativas** para estimar el uplift de futuras campañas.
- Los meses con **Flag = 2 (Eventos Externos mantenidos)** serán analizados en la Práctica 8 para evaluar si deben incluirse en escenarios de riesgo.

### Recursos adicionales

- **Chopra & Meindl** — *Supply Chain Management* (Cap. 7: Demand Forecasting): Marco conceptual para el tratamiento de la demanda histórica.
- **CSCMP Glosario** — Definiciones operativas de "quiebre de stock", "nivel de servicio" y "demanda reprimida": [cscmp.org](https://cscmp.org/CSCMP/Educate/SCM_Definitions_and_Glossary_of_Terms.aspx)
- **MIT OpenCourseWare 15.762** — Supply Chain Planning: Metodología de limpieza de datos históricos para forecasting: [ocw.mit.edu](https://ocw.mit.edu/courses/15-762j-supply-chain-planning-spring-2011/)
- **Microsoft Support** — Función `DESVEST.M` en Excel: [support.microsoft.com/es-es](https://support.microsoft.com/es-es/office/desvest-m-función-desvest-m-d081f5e3-0efd-45b3-a5a4-abb4b7b8e3e2)

---

> **Pregunta de cierre para reflexión grupal:** ¿Qué habría ocurrido si hubiéramos construido el pronóstico de los próximos 12 meses usando directamente el historial con los quiebres de stock sin ajustar? ¿Qué decisión de compra incorrecta habría generado ese pronóstico distorsionado, y cómo habría afectado el nivel de servicio al cliente?

---
LAB_END---

---

---LAB_START---
LAB_ID: 01-00-07
---MARKDOWN---
# Práctica 7 — Construcción Conceptual del Pronóstico (Baseline + Estacionalidad + Iniciativas)

## 1. Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Alta |
| **Nivel Bloom** | Crear (Síntesis e integración de componentes) |
| **Práctica dentro del curso** | Práctica 7 de 10 (requiere insumos de Prácticas 2–6) |
| **Herramientas principales** | Microsoft Excel 2016+, IA Generativa (Copilot / ChatGPT / Gemini) |

---

## 2. Descripción General

Esta práctica integra los tres componentes de la estructura metodológica del pronóstico —**Baseline**, **Estacionalidad** e **Iniciativas**— en un modelo funcional construido en Excel. Aplicarás la fórmula central del curso:

> **Pronóstico = Baseline × Índice Estacional × Factor de Iniciativa**

Proyectarás la demanda de los próximos 6 meses para un SKU representativo, validarás la lógica del modelo con una herramienta de IA Generativa, y traducirás los resultados en indicadores operativos concretos: rotación proyectada, riesgo de quiebre de stock y volumen de picking esperado. Esta práctica es el primer pronóstico completo del curso y el punto de convergencia de todo el trabajo analítico previo.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Integrar los tres componentes del pronóstico (Baseline, Estacionalidad e Iniciativas) en una tabla estructurada en Excel con trazabilidad completa de cada capa.
- [ ] Calcular el pronóstico mensual para los próximos 6 meses aplicando la fórmula `Pronóstico = Baseline × Índice Estacional × Factor de Iniciativa`.
- [ ] Utilizar una herramienta de IA Generativa para validar la lógica del modelo construido e identificar factores adicionales aplicables al contexto logístico.
- [ ] Relacionar el pronóstico resultante con indicadores operativos: rotación proyectada, estimación de quiebres de stock y volumen de picking esperado.

---

## 4. Prerrequisitos

### Conocimiento previo requerido

- Haber completado las Prácticas 1 a 6 del curso (o contar con los archivos de checkpoint provistos por el instructor).
- Comprensión de la estructura metodológica Baseline + Estacionalidad + Iniciativas (Tema 1.4).
- Manejo básico de fórmulas en Excel: referencias absolutas (`$`), funciones `PROMEDIO`, `PRONOSTICO.LINEAL` o `TENDENCIA`.
- Comprensión de los indicadores logísticos: rotación de inventario, nivel de servicio y quiebres de stock (Lección 1.1).

### Acceso requerido

| Recurso | Requerimiento |
|---|---|
| Archivo Excel de trabajo | `Lab07_Pronostico_Completo.xlsx` (provisto por el instructor o generado en Prácticas anteriores) |
| Hoja `Baseline_Depurado` | Resultado de la Práctica 6 (tendencia calculada) |
| Hoja `Indices_Estacionales` | Resultado de la Práctica 4 (índices por mes) |
| IA Generativa | Copilot (copilot.microsoft.com), ChatGPT (chat.openai.com) o Gemini (gemini.google.com) — acceso activo verificado |
| Conexión a internet | Mínimo 10 Mbps para acceso a IA |

> **⚠️ Nota para el instructor:** Si algún participante no completó las Prácticas 4 o 6, debe recibir el archivo de checkpoint `Lab07_Checkpoint_Inicio.xlsx` que incluye el Baseline con tendencia calculada y los índices estacionales ya disponibles. No avanzar sin estos insumos.

---

## 5. Entorno del Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 2 GB | 2 GB+ |
| Resolución de pantalla | 1366×768 | 1920×1080 |
| Conexión a internet | 10 Mbps | 20 Mbps+ |

### Software requerido

| Software | Versión mínima | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior | `PRONOSTICO.LINEAL` no disponible en versiones anteriores a 2016 |
| Navegador web | Chrome / Edge (versión 2024) | Para acceso a IA Generativa |
| IA Generativa | Copilot, ChatGPT o Gemini | Cualquier versión gratuita es suficiente |

### Configuración inicial (antes de comenzar)

1. Abre el archivo `Lab07_Pronostico_Completo.xlsx` (o `Lab07_Checkpoint_Inicio.xlsx` si corresponde).
2. Verifica que el archivo contiene las siguientes hojas:
   - `Datos_Historicos` — 24 meses de historial para los 8 SKUs
   - `Baseline_Depurado` — Baseline con parámetros de tendencia (pendiente e intercepto)
   - `Indices_Estacionales` — Tabla con 12 índices estacionales (uno por mes del año)
   - `Pronostico_Final` — Hoja en blanco donde construirás el modelo en esta práctica
3. Abre una pestaña del navegador con la herramienta de IA Generativa seleccionada y verifica que puedes enviar mensajes.
4. Asegúrate de tener ambas ventanas visibles simultáneamente (Excel y navegador) o usa dos monitores si están disponibles.

---

## 6. Pasos del Laboratorio

---

### Paso 1: Revisar los Insumos Disponibles (Baseline y Estacionalidad)

**Objetivo:** Confirmar que los datos de entrada de las Prácticas 4 y 6 están correctos y listos para ser usados en la construcción del pronóstico.

#### Instrucciones

1. En el archivo Excel, navega a la hoja **`Baseline_Depurado`**.

2. Localiza los parámetros de la línea de tendencia del SKU seleccionado para esta práctica (el instructor indicará cuál usar; se recomienda el SKU con comportamiento de demanda más estable, identificado en la Práctica 3). Debes encontrar:
   - **Pendiente (m):** el incremento mensual promedio de la demanda base.
   - **Intercepto (b):** el valor inicial de la línea de tendencia.

   Ejemplo de referencia (tus valores pueden diferir):
   | Parámetro | Valor de referencia |
   |---|---|
   | Pendiente (m) | 12.5 unidades/mes |
   | Intercepto (b) | 820 unidades |

3. Navega a la hoja **`Indices_Estacionales`** y localiza la tabla con los 12 índices estacionales. Verifica que:
   - Los índices corresponden a los meses del año (Enero = 1, Febrero = 2, … Diciembre = 12).
   - La suma de los 12 índices es aproximadamente igual a **12.00** (o el promedio es ≈ 1.00). Si la suma difiere en más de ±0.05, consulta al instructor antes de continuar.

   Ejemplo de referencia:
   | Mes | Índice Estacional |
   |---|---|
   | Enero | 0.85 |
   | Febrero | 0.78 |
   | Marzo | 1.15 |
   | Abril | 1.02 |
   | Mayo | 0.95 |
   | Junio | 0.88 |
   | Julio | 0.92 |
   | Agosto | 1.05 |
   | Septiembre | 1.20 |
   | Octubre | 1.08 |
   | Noviembre | 1.18 |
   | Diciembre | 0.94 |
   | **Suma** | **12.00** |

4. Anota en papel o en una celda auxiliar los valores de pendiente e intercepto. Los necesitarás en el Paso 2.

#### Resultado esperado

Tienes confirmados: (a) los parámetros de tendencia del SKU seleccionado y (b) los 12 índices estacionales con suma ≈ 12.00.

#### Verificación

✅ La suma de los índices estacionales está entre 11.95 y 12.05.
✅ Los parámetros de tendencia (pendiente e intercepto) son valores numéricos positivos y razonables para el contexto de demanda del SKU.

---

### Paso 2: Proyectar el Baseline para los Próximos 6 Meses

**Objetivo:** Calcular el Baseline proyectado para los meses 25 al 30 (los 6 meses siguientes al historial disponible) usando la ecuación de tendencia lineal.

#### Instrucciones

1. Navega a la hoja **`Pronostico_Final`** (debe estar en blanco).

2. Construye la siguiente estructura de tabla comenzando en la celda **A1**:

   | Columna | Encabezado |
   |---|---|
   | A | `Mes_Num` |
   | B | `Mes_Nombre` |
   | C | `Baseline_Proyectado` |
   | D | `Indice_Estacional` |
   | E | `Pronostico_Estacional` |
   | F | `Factor_Iniciativa` |
   | G | `Pronostico_Final` |
   | H | `Inventario_Actual_Estimado` |
   | I | `Rotacion_Proyectada` |
   | J | `Alerta_Quiebre` |

3. En la columna **A** (filas 2 a 7), ingresa los números de período correspondientes a los próximos 6 meses. Si el historial disponible cubre los meses 1 a 24, los períodos futuros son:

   ```
   A2: 25
   A3: 26
   A4: 27
   A5: 28
   A6: 29
   A7: 30
   ```

4. En la columna **B** (filas 2 a 7), ingresa los nombres de los meses correspondientes. El instructor indicará desde qué mes del año comienza el período 25. Por ejemplo, si el período 1 fue enero del primer año, el período 25 sería enero del tercer año:

   ```
   B2: Enero (Año 3)
   B3: Febrero (Año 3)
   B4: Marzo (Año 3)
   B5: Abril (Año 3)
   B6: Mayo (Año 3)
   B7: Junio (Año 3)
   ```

5. En la columna **C**, calcula el **Baseline proyectado** usando la función `PRONOSTICO.LINEAL` referenciando los datos históricos de la hoja `Baseline_Depurado`:

   ```excel
   =PRONOSTICO.LINEAL(A2, Baseline_Depurado!$B$2:$B$25, Baseline_Depurado!$A$2:$A$25)
   ```

   > **Nota de compatibilidad:** Si usas Excel 2016 con versión anterior a la actualización de 2018, prueba con la función equivalente:
   > ```excel
   > =TENDENCIA(Baseline_Depurado!$B$2:$B$25, Baseline_Depurado!$A$2:$A$25, A2)
   > ```
   > Si ninguna funciona, aplica la fórmula manual usando los parámetros anotados en el Paso 1:
   > ```excel
   > = $pendiente * A2 + $intercepto
   > ```
   > Reemplaza `$pendiente` e `$intercepto` con los valores numéricos reales. Por ejemplo: `= 12.5 * A2 + 820`

6. Copia la fórmula de C2 hacia abajo hasta C7.

7. Formatea la columna C con **0 decimales** (los pronósticos de unidades se expresan en números enteros).

#### Resultado esperado

La columna C muestra 6 valores de Baseline proyectado con tendencia creciente (o estable si la pendiente es cercana a cero). Los valores deben ser coherentes con el rango histórico del SKU —no deben ser negativos ni excesivamente mayores al promedio histórico sin justificación.

#### Verificación

✅ Los 6 valores de Baseline proyectado son positivos.
✅ La diferencia entre el valor del mes 25 y el del mes 30 refleja la tendencia (si la pendiente es +12.5, la diferencia entre C2 y C7 debe ser ≈ 62–63 unidades).
✅ Los valores son coherentes con el rango del historial (±30% del promedio histórico es aceptable para una proyección de corto plazo).

---

### Paso 3: Aplicar los Índices Estacionales

**Objetivo:** Ajustar el Baseline proyectado con los índices estacionales para obtener el pronóstico ajustado por estacionalidad.

#### Instrucciones

1. En la columna **D** (`Indice_Estacional`), ingresa manualmente los índices estacionales correspondientes a cada mes futuro, tomándolos de la hoja `Indices_Estacionales`. Usa el número de mes del año (no el número de período):

   ```
   D2: [Índice de Enero]    → ejemplo: 0.85
   D3: [Índice de Febrero]  → ejemplo: 0.78
   D4: [Índice de Marzo]    → ejemplo: 1.15
   D5: [Índice de Abril]    → ejemplo: 1.02
   D6: [Índice de Mayo]     → ejemplo: 0.95
   D7: [Índice de Junio]    → ejemplo: 0.88
   ```

   **Alternativa con fórmula de referencia dinámica** (opcional, para usuarios avanzados):
   ```excel
   =INDICE(Indices_Estacionales!$B$2:$B$13, MOD(A2-1,12)+1)
   ```
   Esta fórmula extrae automáticamente el índice correspondiente al mes del año según el número de período.

2. En la columna **E** (`Pronostico_Estacional`), calcula el pronóstico ajustado por estacionalidad:

   ```excel
   =REDONDEAR(C2 * D2, 0)
   ```

   Copia hacia abajo hasta E7.

3. Verifica que los meses con índice > 1.0 producen un pronóstico estacional **mayor** que el Baseline, y los meses con índice < 1.0 producen un pronóstico **menor**. Esto es la validación lógica fundamental de este paso.

#### Resultado esperado

La columna E muestra el pronóstico con el efecto estacional visible. Por ejemplo, si Marzo tiene índice 1.15 y el Baseline proyectado es 1,100 unidades, el pronóstico estacional debe ser ≈ 1,265 unidades.

#### Verificación

✅ Para todo mes con índice > 1.0: `E_n > C_n`
✅ Para todo mes con índice < 1.0: `E_n < C_n`
✅ Para todo mes con índice = 1.0: `E_n = C_n`
✅ Ningún valor en la columna E es negativo.

---

### Paso 4: Incorporar el Factor de Iniciativa

**Objetivo:** Agregar la capa de iniciativas comerciales o eventos planificados que modifican la demanda esperada por encima de la estacionalidad normal.

#### Instrucciones

1. En la columna **F** (`Factor_Iniciativa`), asigna el factor correspondiente a cada mes según el calendario de iniciativas comerciales provisto por el instructor. Si no se ha provisto un calendario específico, usa la siguiente asignación de referencia:

   | Mes | Factor de Iniciativa | Justificación |
   |---|---|---|
   | Enero (Año 3) | 1.00 | Sin iniciativas planificadas |
   | Febrero (Año 3) | 1.00 | Sin iniciativas planificadas |
   | Marzo (Año 3) | 1.20 | Campaña de temporada (limpieza de inicio de año) |
   | Abril (Año 3) | 1.10 | Promoción de precio planificada |
   | Mayo (Año 3) | 1.00 | Sin iniciativas planificadas |
   | Junio (Año 3) | 1.15 | Lanzamiento de nuevo canal de distribución |

   Ingresa los valores directamente:
   ```
   F2: 1.00
   F3: 1.00
   F4: 1.20
   F5: 1.10
   F6: 1.00
   F7: 1.15
   ```

   > **Regla operativa:** El Factor de Iniciativa siempre es ≥ 1.00. Un valor de 1.00 significa "sin efecto adicional". Valores entre 1.05 y 1.50 representan incrementos esperados por acciones comerciales planificadas. Si una iniciativa implica reducción de demanda (por ejemplo, descontinuación de un canal), se puede usar un factor < 1.00, pero esto debe documentarse explícitamente.

2. En la columna **G** (`Pronostico_Final`), calcula el pronóstico completo con las tres capas:

   ```excel
   =REDONDEAR(C2 * D2 * F2, 0)
   ```

   Copia hacia abajo hasta G7.

   > **Nota:** Esta fórmula es equivalente a `E2 * F2`, pero se recomienda usar la fórmula completa `C2 * D2 * F2` para mantener la trazabilidad de los tres componentes directamente visibles.

3. Agrega una fila de **totales** en la fila 8:
   ```
   A8: TOTAL 6 MESES
   C8: =SUMA(C2:C7)
   E8: =SUMA(E2:E7)
   G8: =SUMA(G2:G7)
   ```

4. Aplica formato visual para distinguir las tres capas:
   - Columna C (Baseline): fondo **azul claro** (`#DDEEFF`)
   - Columna E (con Estacionalidad): fondo **verde claro** (`#DDFFDD`)
   - Columna G (Pronóstico Final): fondo **naranja claro** (`#FFE8CC`), texto en negrita

#### Resultado esperado

La tabla muestra las tres capas del pronóstico con trazabilidad completa. El Pronóstico Final (columna G) es siempre ≥ al Pronóstico Estacional (columna E) cuando el Factor de Iniciativa ≥ 1.0.

Ejemplo de tabla completa con valores de referencia:

| Mes_Num | Mes_Nombre | Baseline | Índice Est. | Pron. Estacional | Factor Inic. | **Pron. Final** |
|---|---|---|---|---|---|---|
| 25 | Enero Año 3 | 1,113 | 0.85 | 946 | 1.00 | **946** |
| 26 | Febrero Año 3 | 1,125 | 0.78 | 878 | 1.00 | **878** |
| 27 | Marzo Año 3 | 1,138 | 1.15 | 1,309 | 1.20 | **1,571** |
| 28 | Abril Año 3 | 1,150 | 1.02 | 1,173 | 1.10 | **1,290** |
| 29 | Mayo Año 3 | 1,163 | 0.95 | 1,105 | 1.00 | **1,105** |
| 30 | Junio Año 3 | 1,175 | 0.88 | 1,034 | 1.15 | **1,189** |
| **TOTAL** | | **6,864** | | **6,445** | | **6,979** |

#### Verificación

✅ Para todo mes con `F_n = 1.00`: `G_n = E_n`
✅ Para todo mes con `F_n > 1.00`: `G_n > E_n`
✅ El total del Pronóstico Final (G8) es mayor o igual al total del Pronóstico Estacional (E8).
✅ La diferencia porcentual entre G8 y C8 refleja el efecto combinado de estacionalidad e iniciativas (puede ser positiva o negativa según los meses).

---

### Paso 5: Validar el Modelo con IA Generativa

**Objetivo:** Usar una herramienta de IA Generativa para validar la lógica del modelo construido e identificar factores adicionales relevantes para logística de distribución.

#### Instrucciones

1. Abre la herramienta de IA Generativa en el navegador (Copilot en `copilot.microsoft.com`, ChatGPT en `chat.openai.com`, o Gemini en `gemini.google.com`).

2. Copia y envía el siguiente **prompt principal** tal como está escrito (puedes personalizar los valores con los tuyos reales):

   ```
   Tengo un modelo de pronóstico de demanda para logística con tres capas:
   
   1. BASELINE: Proyección de tendencia lineal calculada con regresión sobre 24 meses 
      de historial de ventas. El Baseline para los próximos 6 meses varía entre 
      1,113 y 1,175 unidades/mes con una pendiente de +12.5 unidades/mes.
   
   2. ESTACIONALIDAD: Índices estacionales calculados como promedio de ratios 
      mensuales sobre el promedio anual. Los índices varían entre 0.78 (Febrero) 
      y 1.20 (Septiembre).
   
   3. INICIATIVAS: Factores multiplicadores entre 1.00 y 1.20 para meses con 
      promociones o lanzamientos planificados.
   
   La fórmula es: Pronóstico = Baseline × Índice Estacional × Factor de Iniciativa
   
   Mis preguntas son:
   a) ¿Cómo puedo validar que mi Baseline es representativo de la demanda normal 
      y no está distorsionado por eventos extraordinarios?
   b) ¿Qué factores adicionales debería considerar para mejorar este modelo en 
      un contexto de logística de distribución?
   c) ¿Cuáles son los riesgos más comunes de este tipo de modelo multiplicativo 
      cuando se usa para planificar inventarios?
   ```

3. Lee la respuesta completa de la IA. Identifica y **subraya o resalta** (mentalmente o en papel) las recomendaciones que son directamente aplicables a tu modelo actual.

4. Si la respuesta de la IA no menciona alguno de los siguientes temas, envía un **prompt de seguimiento**:

   ```
   ¿Cómo debería ajustar este modelo si tengo productos con alta variabilidad 
   de demanda (coeficiente de variación > 0.5)? ¿El modelo multiplicativo sigue 
   siendo válido o debo considerar un enfoque diferente?
   ```

5. En la hoja `Pronostico_Final` de Excel, agrega una sección de documentación debajo de la tabla de pronóstico (a partir de la fila 12). Crea la siguiente estructura:

   | Celda | Contenido |
   |---|---|
   | A12 | `VALIDACIÓN CON IA GENERATIVA` (en negrita) |
   | A13 | `Herramienta utilizada:` |
   | B13 | [Nombre de la IA que usaste] |
   | A14 | `Fecha de consulta:` |
   | B14 | [Fecha actual] |
   | A15 | `Recomendación 1 aplicable:` |
   | B15 | [Primera recomendación relevante de la IA] |
   | A16 | `Recomendación 2 aplicable:` |
   | B16 | [Segunda recomendación relevante de la IA] |
   | A17 | `Recomendación 3 aplicable:` |
   | B17 | [Tercera recomendación relevante de la IA] |
   | A18 | `Recomendaciones NO aplicables (justificación):` |
   | B18 | [Recomendaciones que la IA dio pero no aplican a tu contexto, con razón] |

#### Resultado esperado

Tienes documentadas al menos 3 recomendaciones de la IA aplicables al modelo, y has identificado al menos 1 recomendación que no aplica directamente con su justificación. Esta diferenciación crítica demuestra comprensión del modelo y del contexto logístico.

**Recomendaciones típicas que la IA suele generar (referencia para el instructor):**
- Verificar que el Baseline excluye períodos atípicos (COVID, quiebres de stock, lanzamientos únicos).
- Considerar el lead time del proveedor al traducir el pronóstico en órdenes de compra.
- Revisar si los índices estacionales son estables entre años o muestran variación creciente.
- Agregar un intervalo de confianza o rango de pronóstico (optimista/pesimista) para gestión de riesgo.
- Considerar la correlación entre SKUs si son productos complementarios o sustitutos.

#### Verificación

✅ La sección de validación en Excel está completa con nombre de herramienta, fecha y al menos 3 recomendaciones.
✅ Al menos una recomendación está marcada como "no aplicable" con justificación escrita.
✅ Puedes explicar verbalmente por qué cada recomendación aplicable mejora el modelo.

---

### Paso 6: Calcular Indicadores Operativos del Pronóstico

**Objetivo:** Traducir el pronóstico construido en indicadores operativos concretos que conecten el análisis estadístico con decisiones de logística: rotación proyectada, riesgo de quiebre de stock y volumen de picking.

#### Instrucciones

1. En la hoja `Pronostico_Final`, agrega una segunda tabla a partir de la fila 22 con el siguiente encabezado:

   ```
   A22: INDICADORES OPERATIVOS DEL PRONÓSTICO
   ```

2. **Rotación proyectada:** En las celdas A23:B28, calcula la rotación mensual proyectada para cada uno de los 6 meses. Usa el inventario actual estimado provisto por el instructor (o el valor de referencia de 2,500 unidades si no se ha provisto):

   ```
   A23: Mes
   B23: Pronóstico Final (unidades)
   C23: Inventario Inicial Estimado (unidades)
   D23: Rotación Proyectada (veces)
   ```

   En la columna **D** (filas 24 a 29), calcula la rotación usando la fórmula:

   ```excel
   =REDONDEAR(G2 / C24, 2)
   ```

   Donde `G2` es el Pronóstico Final del mes correspondiente y `C24` es el inventario inicial estimado para ese mes.

   > **Simplificación para esta práctica:** Usa el mismo inventario inicial para todos los meses (por ejemplo, 2,500 unidades). En un modelo más avanzado, el inventario inicial de cada mes sería el inventario final del mes anterior.

   > **Recordatorio de la Lección 1.1:**
   > `Rotación = Demanda del período / Inventario promedio del período`
   > Para esta práctica usamos el inventario inicial como aproximación del inventario promedio.

3. **Estimación de quiebres de stock:** En la columna **E** (filas 24 a 29), agrega una alerta de quiebre:

   ```excel
   =SI(C24 < G2, "⚠️ RIESGO QUIEBRE", "✅ STOCK SUFICIENTE")
   ```

   Esta fórmula compara el inventario disponible con el pronóstico: si el inventario es menor que la demanda proyectada, hay riesgo de quiebre de stock.

4. **Volumen de picking esperado:** En la columna **F** (filas 24 a 29), estima el volumen diario de picking asumiendo que el mes tiene 22 días hábiles:

   ```excel
   =REDONDEAR(G2 / 22, 0)
   ```

   Esto da las unidades promedio por día hábil que el área de picking debe procesar.

5. Agrega una fila de resumen en la fila 30:

   ```
   A30: RESUMEN 6 MESES
   B30: =SUMA(G2:G7)          → Total unidades a despachar
   D30: =PROMEDIO(D24:D29)    → Rotación promedio proyectada
   F30: =SUMA(F24:F29)        → Total días-picking estimados
   ```

6. Agrega un comentario en celda E30 con el conteo de meses en riesgo:

   ```excel
   =CONTAR.SI(E24:E29,"*RIESGO*") & " de 6 meses con riesgo de quiebre de stock"
   ```

#### Resultado esperado

Una tabla de indicadores operativos que muestre, para cada uno de los 6 meses del pronóstico:
- La demanda proyectada (unidades).
- La rotación esperada.
- La alerta de quiebre de stock (sí/no).
- El volumen diario de picking.

Ejemplo con valores de referencia:

| Mes | Pron. Final | Inv. Inicial | Rotación | Alerta | Picking/día |
|---|---|---|---|---|---|
| Enero Año 3 | 946 | 2,500 | 0.38 | ✅ STOCK SUFICIENTE | 43 |
| Febrero Año 3 | 878 | 2,500 | 0.35 | ✅ STOCK SUFICIENTE | 40 |
| Marzo Año 3 | 1,571 | 2,500 | 0.63 | ✅ STOCK SUFICIENTE | 71 |
| Abril Año 3 | 1,290 | 2,500 | 0.52 | ✅ STOCK SUFICIENTE | 59 |
| Mayo Año 3 | 1,105 | 2,500 | 0.44 | ✅ STOCK SUFICIENTE | 50 |
| Junio Año 3 | 1,189 | 2,500 | 0.48 | ✅ STOCK SUFICIENTE | 54 |

> **Reflexión operativa (discutir con el grupo):** ¿Qué pasaría si el inventario inicial fuera de 1,000 unidades en lugar de 2,500? ¿Cuántos meses estarían en riesgo de quiebre? ¿Cómo cambiaría la decisión de reabastecimiento?

#### Verificación

✅ La columna de Rotación muestra valores entre 0.1 y 2.0 para un contexto de distribución mensual (valores fuera de este rango indican un posible error de fórmula o un dato de inventario irreal).
✅ La alerta de quiebre funciona correctamente: cambia a "⚠️ RIESGO QUIEBRE" si modificas el inventario inicial a un valor menor que el pronóstico.
✅ El volumen de picking por día es un número entero positivo razonable (entre 20 y 200 unidades/día para un SKU típico de distribución).

---

## 7. Validación y Pruebas del Modelo Completo

Una vez completados los 6 pasos, realiza las siguientes pruebas de consistencia para validar que el modelo funciona correctamente:

### Prueba 1: Prueba de sensibilidad del Factor de Iniciativa

1. Cambia temporalmente el Factor de Iniciativa de Marzo (F4) de **1.20 a 1.00**.
2. Verifica que el Pronóstico Final de Marzo (G4) disminuye exactamente un 16.67% (es decir, `G4_nuevo = G4_original / 1.20`).
3. Observa cómo cambia la alerta de quiebre en la tabla de indicadores operativos.
4. **Restaura el valor a 1.20** antes de continuar.

### Prueba 2: Coherencia entre capas

Verifica manualmente para al menos un mes que la fórmula multiplicativa es correcta:
```
Pronóstico Final = Baseline × Índice Estacional × Factor Iniciativa
Ejemplo Marzo: 1,138 × 1.15 × 1.20 = 1,570.44 ≈ 1,571 ✓
```

### Prueba 3: Validación lógica del impacto estacional

Identifica el mes con el índice estacional más alto y el más bajo en tu tabla de 6 meses. Confirma que:
- El mes con índice más alto tiene el Pronóstico Estacional más alto (si el Baseline es similar entre meses).
- La diferencia porcentual entre ambos meses es coherente con la diferencia entre sus índices.

### Prueba 4: Coherencia con el historial

Compara el total del Pronóstico Final de los 6 meses con el promedio de los 6 meses equivalentes del año anterior (disponibles en la hoja `Datos_Historicos`). La diferencia no debería superar el 25% sin una justificación clara (tendencia creciente + iniciativas). Si supera ese umbral, revisa los parámetros de tendencia y los factores de iniciativa.

---

## 8. Resolución de Problemas

### Problema 1: `PRONOSTICO.LINEAL` devuelve `#¡VALOR!` o `#N/A`

**Síntoma:** La función `PRONOSTICO.LINEAL` en la columna C muestra un error en lugar de un valor numérico.

**Causa probable:** Los rangos de referencia en la hoja `Baseline_Depurado` contienen celdas vacías, texto en lugar de números, o los rangos de `x` (períodos) y `y` (demanda) tienen tamaños diferentes.

**Solución:**
1. Navega a la hoja `Baseline_Depurado` y verifica que las columnas de período (eje X) y demanda depurada (eje Y) tienen exactamente el mismo número de filas con datos numéricos sin celdas vacías intermedias.
2. Selecciona el rango de datos y usa `Datos → Texto en columnas → Finalizar` para forzar la conversión a formato numérico si las celdas muestran números alineados a la izquierda (señal de que están en formato texto).
3. Si el problema persiste, usa la fórmula manual alternativa: `= pendiente * A2 + intercepto` con los valores numéricos directamente ingresados.
4. Como última alternativa, usa `=TENDENCIA(Baseline_Depurado!$B$2:$B$25, Baseline_Depurado!$A$2:$A$25, A2)` que es más robusta en algunas versiones de Excel.

---

### Problema 2: La herramienta de IA Generativa no está disponible o la conexión falla

**Síntoma:** No se puede acceder a Copilot, ChatGPT o Gemini durante el Paso 5. La página no carga o el servicio devuelve errores de conexión.

**Causa probable:** Restricciones de red corporativa o universitaria, límite de uso de la cuenta gratuita alcanzado, o fallo temporal del servicio.

**Solución:**
1. **Primera alternativa:** Intenta con otra herramienta de IA. Si Copilot falla, prueba con `chat.openai.com` (ChatGPT) o `gemini.google.com`. Al menos una de las tres suele estar disponible.
2. **Segunda alternativa:** El instructor debe tener preparada una captura de pantalla de una respuesta de referencia generada previamente con el mismo prompt. Úsala como base para completar la sección de documentación en Excel.
3. **Tercera alternativa (sin internet):** Usa el siguiente texto de referencia como respuesta simulada de la IA para completar el ejercicio:

   > *"Para validar que el Baseline es representativo: (1) verifica que los períodos usados para calcular la tendencia no incluyen meses con quiebres de stock conocidos ni eventos extraordinarios como pandemia; (2) compara el Baseline con la mediana del historial —si difieren en más del 15%, revisa los outliers; (3) considera usar una ventana móvil de los últimos 12 meses si la tendencia es muy reciente. Factores adicionales para logística de distribución: lead time del proveedor, capacidad de almacenamiento, costo de oportunidad del sobreinventario, y correlación con otros SKUs del portafolio."*

4. Documenta en la celda B13 de la sección de validación cuál fue la alternativa utilizada y por qué.

---

## 9. Limpieza y Entrega

### Antes de cerrar el archivo

1. **Guarda el archivo** con el nombre: `Lab07_[TuNombre]_Pronostico_Final.xlsx`
   - Usa `Ctrl + Shift + S` (Guardar como) para evitar sobreescribir el archivo original.

2. **Verifica que el archivo contiene:**
   - [ ] Hoja `Pronostico_Final` con la tabla de 6 meses completa (columnas A–G, fila de totales).
   - [ ] Sección de validación con IA (filas 12–18) con al menos 3 recomendaciones documentadas.
   - [ ] Tabla de indicadores operativos (filas 22–30) con rotación, alertas de quiebre y volumen de picking.
   - [ ] Formato de color aplicado para distinguir las tres capas del pronóstico.

3. **No elimines** las hojas `Baseline_Depurado` ni `Indices_Estacionales` — son insumos que se usarán en la Práctica 8.

4. Cierra la pestaña del navegador con la IA Generativa (opcional, pero recomendado para liberar memoria).

### Entrega al instructor

Sube el archivo `Lab07_[TuNombre]_Pronostico_Final.xlsx` a la plataforma del curso o compártelo según las instrucciones del instructor. Este archivo es el **insumo principal de la Práctica 8**, donde se calculará el error del pronóstico (MAE y MAPE) comparando el modelo construido con datos reales.

---

## 10. Resumen y Recursos Adicionales

### Lo que construiste en esta práctica

En esta práctica integraste por primera vez los tres componentes del pronóstico en un modelo operativo completo:

| Componente | Fuente | Función en el modelo |
|---|---|---|
| **Baseline** | Tendencia lineal (Práctica 6) | Representa la demanda "normal" sin efectos estacionales ni comerciales |
| **Estacionalidad** | Índices calculados (Práctica 4) | Amplifica o reduce el Baseline según el patrón histórico del mes |
| **Iniciativas** | Calendario comercial (esta práctica) | Refleja el impacto esperado de acciones planificadas por el negocio |

La fórmula `Pronóstico = Baseline × Índice Estacional × Factor de Iniciativa` no es solo una ecuación matemática: es una representación explícita de cómo el negocio entiende su demanda. Cada capa tiene una fuente de datos diferente y una responsabilidad diferente en la organización: el equipo de análisis construye el Baseline, el historial define la estacionalidad, y el equipo comercial define las iniciativas.

### Conexión con la Lección 1.1

Como viste en la Lección 1.1, el forecasting existe para **reducir la incertidumbre** y mejorar decisiones logísticas. Los indicadores que calculaste en el Paso 6 son exactamente los que la lección describe como críticos:

- **Rotación proyectada** → indica si el inventario actual está sobredimensionado o subdimensionado para la demanda esperada.
- **Alerta de quiebre de stock** → anticipa el problema antes de que ocurra, permitiendo una respuesta proactiva (no reactiva).
- **Volumen de picking diario** → permite planificar la capacidad operativa del almacén con anticipación.

### Pregunta de reflexión final

> *¿En qué situación el Factor de Iniciativa podría ser el componente más importante del pronóstico, superando incluso el efecto de la estacionalidad? ¿Qué implicaciones tendría esto para la planificación de inventarios?*

Discute esta pregunta con tu grupo antes de iniciar la Práctica 8.

### Recursos adicionales

| Recurso | Descripción |
|---|---|
| [Chopra & Meindl — Supply Chain Management (Pearson)](https://www.pearson.com/en-us/subject-catalog/p/supply-chain-management/P200000005934) | Capítulo 7: Demand Forecasting in a Supply Chain — fundamentos del modelo multiplicativo |
| [MIT OpenCourseWare 15.762](https://ocw.mit.edu/courses/15-762j-supply-chain-planning-spring-2011/) | Módulo de forecasting aplicado a cadena de suministro |
| [Microsoft Support — PRONOSTICO.LINEAL](https://support.microsoft.com/es-es/office/pronostico-lineal-función-pronostico-lineal-bbf8d4a1-4b28-4de1-8645-4e6bfdd3f0f6) | Documentación oficial de la función en Excel |
| [Council of Supply Chain Management Professionals — Glosario](https://cscmp.org/CSCMP/Educate/SCM_Definitions_and_Glossary_of_Terms.aspx) | Definiciones de rotación de inventario, nivel de servicio y quiebre de stock |

---

> **⚡ Próxima práctica:** En la **Práctica 8** calcularás el error del pronóstico construido hoy usando las métricas MAE (Error Absoluto Medio) y MAPE (Error Porcentual Absoluto Medio), y seleccionarás el modelo con menor error entre tres alternativas. El archivo que guardaste en esta práctica es el punto de partida.

---
LAB_END---

---

# Práctica 8 — Simulación Cualitativa del Impacto de una Promoción en el Pronóstico

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Alta |
| **Nivel Bloom** | Crear (Create) |
| **Práctica anterior requerida** | Lab 01-00-07 (modelo de pronóstico completo) |
| **Herramientas principales** | Microsoft Excel · IA Generativa (Copilot / ChatGPT / Gemini) |
| **Archivo de entrada** | `Practica7_Pronostico_Completado.xlsx` |
| **Archivo de salida** | `Practica8_Simulacion_Promocion.xlsx` |

---

## 2. Descripción General

En esta práctica aplicarás el modelo de pronóstico construido en la Práctica 7 para simular el impacto de una campaña promocional real sobre la demanda proyectada. Modificarás el **Factor de Iniciativa** del modelo para los meses de promoción, calcularás el inventario adicional necesario, estimarás el incremento en operaciones de picking y analizarás el efecto de canibalización post-promocional. Finalmente, utilizarás una herramienta de IA Generativa para enriquecer el análisis de riesgos y construirás una tabla de decisiones logísticas accionable.

Esta práctica cierra el ciclo metodológico del curso conectando directamente el pronóstico con decisiones operativas concretas: **cuánto inventario preparar, cuántos pickers asignar y qué hacer después de la promoción** para evitar el sobreinventario.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica, serás capaz de:

- [ ] **Modificar** el Factor de Iniciativa en el modelo de pronóstico para simular el impacto de una promoción del 40% durante 2 meses y observar el cambio en la demanda proyectada total.
- [ ] **Calcular** el inventario adicional requerido y el incremento en líneas de picking durante los meses promocionales, vinculando el pronóstico con decisiones operativas de recursos humanos.
- [ ] **Analizar** el efecto de canibalización post-promoción ajustando el factor del mes siguiente y cuantificando la caída esperada de demanda.
- [ ] **Utilizar** una herramienta de IA Generativa para generar escenarios alternativos de riesgo logístico y evaluar su razonabilidad frente al modelo propio.
- [ ] **Construir** una tabla de decisiones logísticas (antes / durante / después de la promoción) fundamentada en los resultados de la simulación.

---

## 4. Prerrequisitos

### Conocimiento previo requerido

| Tema | Descripción |
|---|---|
| Práctica 7 completada | Modelo de pronóstico con columnas Baseline, Estacionalidad, Iniciativa y Pronóstico Total disponible en Excel |
| Componente Iniciativas (Tema 1.4) | Comprensión del Factor de Iniciativa como multiplicador que ajusta el pronóstico base ante eventos planificados del negocio |
| Impacto del pronóstico en picking (Temas 1.5–1.6) | Relación entre volumen de demanda proyectada, líneas de picking y dimensionamiento de recursos humanos |
| Nivel de servicio y quiebres de stock (Lección 1.1) | Comprensión de cómo un pico de demanda no anticipado genera quiebres de stock y deteriora el nivel de servicio |

### Acceso requerido

| Recurso | Estado requerido |
|---|---|
| Archivo `Practica7_Pronostico_Completado.xlsx` | Disponible y abierto en Excel |
| Microsoft Excel 2016 o superior | Activo |
| Herramienta de IA Generativa | Sesión iniciada (Copilot, ChatGPT o Gemini) |
| Conexión a internet | Activa (mínimo 10 Mbps) |

> **⚠️ Nota para el instructor:** Si un participante no completó la Práctica 7, proporcionarle el archivo checkpoint `Practica7_Checkpoint_Completo.xlsx` antes de iniciar esta práctica. Sin ese archivo, los pasos 1 al 4 no pueden ejecutarse.

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Almacenamiento libre | 500 MB | 2 GB |

### Software requerido

| Software | Versión | Propósito en esta práctica |
|---|---|---|
| Microsoft Excel | 2016 o superior (Microsoft 365 recomendado) | Modificación del modelo, cálculos de inventario y picking |
| Microsoft 365 Copilot / ChatGPT / Gemini | Versión actualizada 2024 | Generación de escenarios de riesgo logístico |
| Microsoft Edge / Google Chrome | Versión 2024 | Acceso a herramienta de IA |

### Preparación del entorno (antes de iniciar)

```
1. Abrir el archivo Practica7_Pronostico_Completado.xlsx en Excel
2. Guardar una copia inmediatamente con el nombre: Practica8_Simulacion_Promocion.xlsx
3. Abrir el navegador y acceder a la herramienta de IA Generativa:
   - Microsoft Copilot: https://copilot.microsoft.com
   - ChatGPT:          https://chat.openai.com
   - Gemini:           https://gemini.google.com
4. Verificar que ambas ventanas (Excel y navegador) sean visibles simultáneamente
   en pantalla (modo ventanas lado a lado recomendado)
```

---

## 6. Desarrollo Paso a Paso

---

### Paso 1: Revisar el Modelo de Pronóstico Base y Preparar la Hoja de Simulación

**Objetivo:** Familiarizarse con la estructura del modelo recibido de la Práctica 7 e identificar las columnas que serán modificadas durante la simulación.

#### Instrucciones

1. Abre el archivo `Practica8_Simulacion_Promocion.xlsx` (copia del archivo de la Práctica 7).

2. Localiza la hoja de trabajo principal del modelo de pronóstico. Verifica que contenga las siguientes columnas para el SKU Clase A de mayor rotación (el que utilizarás en esta práctica):

   ```
   Columna A: Mes (1 al 6 para el horizonte de pronóstico)
   Columna B: Baseline (demanda base proyectada)
   Columna C: Índice de Estacionalidad
   Columna D: Factor de Iniciativa (actualmente = 1.00 para todos los meses)
   Columna E: Pronóstico Total = B * C * D
   ```

3. Haz clic derecho sobre la pestaña de la hoja y selecciona **Mover o copiar**. Marca la casilla **Crear una copia** y nómbrala `Simulacion_Promocion`. Esta hoja será donde realizarás todos los cambios sin alterar el modelo original.

4. En la hoja `Simulacion_Promocion`, agrega un bloque de parámetros en las celdas **H1:I5** con los supuestos de la simulación:

   ```
   H1: PARÁMETROS DE LA PROMOCIÓN
   H2: SKU simulado
   H3: Meses de promoción
   H4: Factor de Iniciativa (promoción)
   H5: Factor de Iniciativa (post-promoción)

   I2: [Nombre del SKU Clase A de mayor rotación]
   I3: Meses 4 y 5
   I4: 1.40
   I5: 0.85
   ```

5. Añade también los parámetros operativos en **H7:I9**:

   ```
   H7: Productividad picker (líneas/hora)
   H8: Horas laborales por día
   H9: Días laborales por mes

   I7: 80
   I8: 8
   I9: 22
   ```

**Resultado esperado:** La hoja `Simulacion_Promocion` contiene el modelo base intacto más un bloque de parámetros documentados. El pronóstico total en la columna E aún no ha cambiado (todos los factores de iniciativa = 1.00).

**Verificación:** Confirma que la suma del Pronóstico Total para los 6 meses en la hoja `Simulacion_Promocion` es idéntica a la suma en la hoja original. Si difieren, hay un error en la copia.

---

### Paso 2: Aplicar el Factor de Iniciativa Promocional y Observar el Impacto en el Pronóstico

**Objetivo:** Modificar el Factor de Iniciativa para los meses 4, 5 y 6 del horizonte de pronóstico para simular la promoción y el efecto de canibalización post-promocional.

#### Instrucciones

1. En la hoja `Simulacion_Promocion`, localiza la columna D (Factor de Iniciativa).

2. **Modifica los factores** para los tres meses clave según el escenario definido:

   ```
   Celda D4 (Mes 4): Cambiar de 1.00 a  ==$I$4   → resultado: 1.40
   Celda D5 (Mes 5): Cambiar de 1.00 a  ==$I$4   → resultado: 1.40
   Celda D6 (Mes 6): Cambiar de 1.00 a  ==$I$5   → resultado: 0.85
   ```

   > **Nota:** Usar referencias a las celdas de parámetros (`$I$4` y `$I$5`) en lugar de valores fijos permite modificar los supuestos fácilmente en análisis de sensibilidad posteriores.

3. Verifica que la fórmula en la columna E (Pronóstico Total) se actualice automáticamente. Si la fórmula es `=B*C*D`, los meses 4 y 5 mostrarán valores un 40% superiores al escenario sin promoción, y el mes 6 mostrará un valor un 15% inferior.

4. Agrega una columna comparativa. En la columna F, escribe el encabezado `Pronóstico SIN Promoción` y copia los valores originales de la columna E **antes de los cambios** (puedes referenciar la hoja original):

   ```
   F2: Pronóstico SIN Promoción
   F3: =[NombreHojaOriginal]!E3
   F4: =[NombreHojaOriginal]!E4
   F5: =[NombreHojaOriginal]!E5
   F6: =[NombreHojaOriginal]!E6
   F7: =[NombreHojaOriginal]!E7
   F8: =[NombreHojaOriginal]!E8
   ```

5. En la columna G, calcula la diferencia:

   ```
   G2: Diferencia (unidades)
   G3: =E3-F3
   G4: =E4-F4
   ... (hasta G8)
   ```

6. Agrega una fila de totales al final:

   ```
   Fila 9:
   Columna E: =SUMA(E3:E8)   → Total demanda CON promoción
   Columna F: =SUMA(F3:F8)   → Total demanda SIN promoción
   Columna G: =SUMA(G3:G8)   → Diferencia neta total
   ```

**Resultado esperado:** La tabla comparativa muestra que los meses 4 y 5 tienen un incremento de ~40% en la demanda proyectada, el mes 6 tiene una reducción de ~15%, y la diferencia neta total del período de 6 meses es positiva (la promoción genera demanda adicional neta, aunque menor al 40% porque el mes 6 canibaliza parte de esa ganancia).

**Verificación:** Divide el valor de G4 (diferencia mes 4) entre F4 (base sin promoción). El resultado debe ser aproximadamente 0.40 (40%). Si obtienes un valor muy diferente, revisa que la fórmula de Pronóstico Total en la columna E multiplique correctamente los tres factores.

---

### Paso 3: Calcular el Inventario Adicional Necesario para los Meses Promocionales

**Objetivo:** Traducir el incremento de demanda proyectado en unidades adicionales de inventario requeridas, vinculando el pronóstico con una decisión operativa de abastecimiento.

#### Instrucciones

1. Crea una nueva sección en la misma hoja, comenzando en la celda **A12**. Titúlala `ANÁLISIS DE INVENTARIO ADICIONAL REQUERIDO`.

2. Construye la siguiente tabla en el rango **A13:E17**:

   ```
   A13: Concepto
   B13: Mes 4 (Promoción)
   C13: Mes 5 (Promoción)
   D13: Mes 6 (Post-Promoción)
   E13: Total Período

   A14: Demanda proyectada CON promoción
   A15: Demanda proyectada SIN promoción
   A16: Inventario adicional requerido
   A17: Stock de seguridad adicional (20% del incremento)
   ```

3. Completa las fórmulas de referencia:

   ```
   B14: =E4    (Pronóstico mes 4 CON promoción)
   C14: =E5    (Pronóstico mes 5 CON promoción)
   D14: =E6    (Pronóstico mes 6 CON promoción)
   E14: =SUMA(B14:D14)

   B15: =F4    (Pronóstico mes 4 SIN promoción)
   C15: =F5
   D15: =F6
   E15: =SUMA(B15:D15)

   B16: =B14-B15   (Unidades adicionales a tener disponibles en mes 4)
   C16: =C14-C15
   D16: =D14-D15   (Nota: este valor será negativo en mes 6)
   E16: =SUMA(B16:D16)

   B17: =MAX(B16*0.20, 0)   (Stock de seguridad solo si hay incremento positivo)
   C17: =MAX(C16*0.20, 0)
   D17: =MAX(D16*0.20, 0)
   E17: =SUMA(B17:D17)
   ```

4. Agrega una fila de interpretación en **A19:B19**:

   ```
   A19: Inventario total a preparar ANTES de la promoción (mes 4):
   B19: =B16+B17
   ```

   > **Interpretación logística:** Este valor representa las unidades adicionales que deben estar físicamente disponibles en bodega al inicio del mes 4. Si el lead time del proveedor es de 3 semanas, la orden de compra adicional debe colocarse en la semana 1 del mes 3.

5. En la celda **A21**, agrega la siguiente nota de decisión:

   ```
   A21: DECISIÓN DE ABASTECIMIENTO:
   A22: "Colocar orden de compra adicional de [=B19] unidades con al menos
         3 semanas de anticipación al inicio del mes 4 para garantizar
         disponibilidad sin generar quiebre de stock durante la promoción."
   ```

**Resultado esperado:** La tabla muestra claramente cuántas unidades adicionales se necesitan en cada mes. El mes 6 muestra un valor negativo en la fila 16, lo que indica que se espera una reducción de demanda (canibalización) y que **no se debe realizar reposición adicional** en ese mes.

**Verificación:** El valor en B19 debe ser mayor que B16. Si B19 es menor o igual a B16, verifica que la fórmula de B17 esté calculando el 20% correctamente. El stock de seguridad adicional siempre es positivo o cero, nunca negativo.

---

### Paso 4: Estimar el Incremento en Líneas de Picking y su Impacto en Recursos Humanos

**Objetivo:** Calcular cuántas líneas de picking adicionales generará la promoción y cuántos pickers adicionales serán necesarios para procesarlas, conectando el pronóstico con la planificación operativa del almacén.

#### Instrucciones

1. Crea una nueva sección en **A24**, titulada `ANÁLISIS DE CAPACIDAD DE PICKING`.

2. Construye la tabla en el rango **A25:D32**:

   ```
   A25: Parámetro
   B25: Mes 4 (Promoción)
   C25: Mes 5 (Promoción)
   D25: Mes 6 (Post-Prom.)

   A26: Demanda proyectada CON promoción (unidades)
   A27: Demanda proyectada SIN promoción (unidades)
   A28: Líneas de picking adicionales por mes
   A29: Líneas de picking adicionales por día
   A30: Horas de picking adicionales por día
   A31: Pickers adicionales requeridos (turno 8h)
   A32: Pickers adicionales requeridos (redondeado)
   ```

3. Completa las fórmulas:

   ```
   B26: =B14    C26: =C14    D26: =D14
   B27: =B15    C27: =C15    D27: =D15

   B28: =MAX(B26-B27, 0)
   C28: =MAX(C26-C27, 0)
   D28: =MAX(D26-D27, 0)

   B29: =B28/$I$9     (dividir entre 22 días laborales)
   C29: =C28/$I$9
   D29: =D28/$I$9

   B30: =B29/$I$7     (dividir entre 80 líneas/hora)
   C30: =C29/$I$7
   D30: =D29/$I$7

   B31: =B30/$I$8     (dividir entre 8 horas por turno)
   C31: =C30/$I$8
   D31: =D30/$I$8

   B32: =REDONDEAR.MAS(B31,0)
   C32: =REDONDEAR.MAS(C31,0)
   D32: =REDONDEAR.MAS(D31,0)
   ```

4. Agrega una celda de interpretación en **A34**:

   ```
   A34: DECISIÓN DE RECURSOS HUMANOS:
   A35: "Durante los meses 4 y 5 se requerirán [=B32] y [=C32] pickers
         adicionales respectivamente para procesar el incremento de líneas
         generado por la promoción, manteniendo la productividad estándar
         de 80 líneas/hora. Considerar contratación temporal o horas extra."
   ```

> **Reflexión logística (ancla al Tema 1.1):** Un pronóstico preciso no solo determina cuánto inventario comprar; también define cuántas personas necesitas para mover ese inventario. La subestimación del pico promocional generaría cuellos de botella en el almacén, retrasos en despacho y deterioro del nivel de servicio, exactamente el escenario que el forecasting busca evitar.

**Resultado esperado:** Las celdas B32 y C32 muestran un número entero positivo de pickers adicionales (típicamente entre 1 y 4 dependiendo del volumen del SKU). La celda D32 debe mostrar 0 (no se necesitan pickers adicionales en el mes post-promoción porque la demanda cae).

**Verificación:** Multiplica el valor de B32 × 80 líneas/hora × 8 horas × 22 días. El resultado debe ser mayor o igual a B28 (líneas adicionales del mes 4). Si es menor, el número de pickers calculado es insuficiente y debes revisar la fórmula de redondeo.

---

### Paso 5: Construir el Gráfico Comparativo de Pronóstico Base vs. Pronóstico con Promoción

**Objetivo:** Visualizar el impacto de la promoción y el efecto de canibalización en un gráfico de líneas que facilite la comunicación de los resultados a la gerencia.

#### Instrucciones

1. Selecciona el rango de datos para el gráfico: **A2:A8** (meses) junto con las columnas **E** y **F** (pronóstico con y sin promoción). Para seleccionar rangos no contiguos, mantén presionada la tecla `Ctrl` mientras seleccionas.

   ```
   Selección: A2:A8 + E2:E8 + F2:F8
   ```

2. Inserta un gráfico de líneas:
   ```
   Menú Insertar → Gráficos → Líneas → Línea con marcadores
   ```

3. Configura el gráfico con los siguientes elementos:

   ```
   Título del gráfico:    "Pronóstico Base vs. Pronóstico con Promoción — SKU [Nombre]"
   Eje X (horizontal):    Meses 1 al 6
   Eje Y (vertical):      Unidades demandadas
   Serie 1 (línea azul):  "Pronóstico CON Promoción"
   Serie 2 (línea gris):  "Pronóstico SIN Promoción"
   ```

4. Agrega anotaciones visuales al gráfico:
   - Haz clic derecho sobre el punto del mes 4 en la serie azul → **Agregar etiqueta de datos**
   - Repite para el mes 5 y el mes 6
   - En el área del gráfico, inserta un cuadro de texto con la nota: `"Meses 4-5: Promoción +40% | Mes 6: Canibalización -15%"`

5. Mueve el gráfico a una posición visible en la hoja (debajo de la tabla de picking, aproximadamente en la fila 40).

**Resultado esperado:** El gráfico muestra claramente dos picos en los meses 4 y 5 (línea azul por encima de la gris) y una caída en el mes 6 (línea azul por debajo de la gris). La forma visual es característica de un evento promocional con canibalización: pico seguido de valle.

**Verificación:** La línea azul (con promoción) debe cruzar hacia abajo de la línea gris (sin promoción) exactamente en el mes 6. Si ambas líneas se comportan igual en el mes 6, verifica que la celda D6 tenga el factor 0.85 y no 1.00.

---

### Paso 6: Consultar a la IA Generativa para Análisis de Riesgos Logísticos

**Objetivo:** Utilizar una herramienta de IA Generativa para obtener perspectivas adicionales sobre los riesgos operativos de una promoción con pico de demanda del 40% y evaluar su razonabilidad frente al análisis propio.

#### Instrucciones

1. Abre tu herramienta de IA Generativa preferida en el navegador:
   - **Copilot:** https://copilot.microsoft.com
   - **ChatGPT:** https://chat.openai.com
   - **Gemini:** https://gemini.google.com

2. Copia y pega el siguiente **prompt exacto** en la herramienta de IA:

   ```
   PROMPT PRINCIPAL:

   "En logística, cuando una promoción genera un pico de demanda del 40%,
   ¿cuáles son los principales riesgos operativos y cómo se mitigan desde
   la planificación del inventario y el picking?"
   ```

3. Lee la respuesta completa de la IA. Luego, realiza una **segunda consulta de profundización** con el siguiente prompt:

   ```
   PROMPT DE PROFUNDIZACIÓN:

   "Específicamente para un almacén de distribución que maneja 8 SKUs con
   picking manual, con una productividad estándar de 80 líneas por hora
   por picker, ¿qué acciones preventivas concretas recomendarías implementar
   2 semanas antes de una promoción que duplica el volumen de picking en
   los meses 4 y 5 de un horizonte de 6 meses?"
   ```

4. En Excel, crea una nueva hoja llamada `Analisis_IA` e inserta una tabla en el rango **A1:C15** con la siguiente estructura:

   ```
   A1: Riesgo Identificado por IA
   B1: Nivel de Relevancia para Nuestro Escenario (Alto/Medio/Bajo)
   C1: ¿Está contemplado en nuestro modelo? (Sí/Parcialmente/No)
   ```

5. Completa la tabla con **al menos 5 riesgos** identificados en la respuesta de la IA, evaluando cada uno según tu criterio:

   ```
   Ejemplo de fila:
   A2: Quiebre de stock por subestimación del pico de demanda
   B2: Alto
   C2: Sí — el modelo calcula stock de seguridad adicional del 20%

   A3: Colapso de capacidad de picking por falta de personal
   B3: Alto
   C3: Sí — el modelo calcula pickers adicionales necesarios

   A4: Sobreinventario post-promoción por pedidos excesivos
   B4: Alto
   C4: Parcialmente — el modelo reduce el factor en mes 6 pero no
       calcula el volumen de devoluciones posibles

   A5: [Completar con riesgo identificado en la respuesta de la IA]
   ...
   ```

6. Al final de la tabla, en la celda **A17**, escribe una conclusión personal de 2-3 oraciones:

   ```
   A17: EVALUACIÓN DE LA RESPUESTA DE LA IA:
   A18: [Tu evaluación: ¿Qué agregó la IA que no habías considerado?
         ¿Qué riesgos mencionó la IA que tu modelo ya contempla?
         ¿Hay alguna recomendación de la IA que cambiaría tu análisis?]
   ```

**Resultado esperado:** La hoja `Analisis_IA` contiene una tabla con mínimo 5 riesgos evaluados y una conclusión personal que demuestra pensamiento crítico sobre la utilidad y las limitaciones de la IA como herramienta de análisis logístico.

**Verificación:** Revisa que al menos 2 de los riesgos identificados tengan nivel "Alto" y que la columna C muestre una mezcla de "Sí", "Parcialmente" y "No" — si todo dice "Sí", probablemente la evaluación no es suficientemente crítica.

> **Nota sobre la IA como herramienta:** La IA Generativa no tiene acceso a tus datos específicos de demanda; sus respuestas son generales basadas en conocimiento de logística. Tu valor como analista está en **conectar** esas recomendaciones generales con los números específicos que calculaste en los pasos anteriores. Esta es la diferencia entre usar la IA como oráculo (error) y usarla como asistente de pensamiento (correcto).

---

### Paso 7: Construir la Tabla de Decisiones Logísticas Antes / Durante / Después

**Objetivo:** Sintetizar todos los resultados de la simulación en una tabla de decisiones accionables que conecte el pronóstico con acciones operativas concretas en cada fase de la promoción.

#### Instrucciones

1. Crea una nueva hoja llamada `Tabla_Decisiones`.

2. En el rango **A1:D20**, construye la siguiente tabla de decisiones. Completa las celdas marcadas con `[CALCULAR]` referenciando los valores calculados en los pasos anteriores:

   ```
   A1: TABLA DE DECISIONES LOGÍSTICAS — SIMULACIÓN PROMOCIONAL
   A2: SKU: [Nombre del SKU Clase A]  |  Período: Meses 4-5 (promoción) + Mes 6 (post)

   A4:  FASE
   B4:  ACCIÓN LOGÍSTICA
   C4:  VALOR / REFERENCIA DEL MODELO
   D4:  RESPONSABLE SUGERIDO

   ─── ANTES DE LA PROMOCIÓN (Mes 3) ───

   A5:  ANTES
   B5:  Colocar orden de compra adicional al proveedor
   C5:  [CALCULAR] unidades adicionales requeridas
        (referenciar celda B19 de hoja Simulacion_Promocion)
   D5:  Jefe de Compras / Planeador de Demanda

   A6:  ANTES
   B6:  Reservar espacio adicional en bodega para recepción de mercancía
   C6:  Estimado: [CALCULAR] m² basado en volumen adicional
        (asumir 0.05 m² por unidad adicional)
   D6:  Jefe de Almacén

   A7:  ANTES
   B7:  Contratar o programar pickers adicionales para meses 4 y 5
   C7:  [CALCULAR] pickers adicionales requeridos
        (referenciar celdas B32 y C32 de hoja Simulacion_Promocion)
   D7:  Supervisor de Operaciones / RRHH

   A8:  ANTES
   B8:  Comunicar el plan de demanda al proveedor para asegurar disponibilidad
   C8:  Incremento esperado: 40% sobre demanda base
   D8:  Planeador de Demanda / Compras

   ─── DURANTE LA PROMOCIÓN (Meses 4-5) ───

   A9:  DURANTE
   B9:  Monitorear el nivel de inventario diariamente
   C9:  Alerta si stock disponible cae por debajo del stock de seguridad
        (nivel crítico: [CALCULAR] unidades — 20% del inventario adicional)
   D9:  Analista de Inventario

   A10: DURANTE
   B10: Controlar la productividad real de picking vs. estándar (80 líneas/hora)
   C10: Si productividad cae >10%, activar turno adicional o horas extra
   D10: Supervisor de Almacén

   A11: DURANTE
   B11: Registrar la demanda real diaria para comparar con el pronóstico
   C11: Actualizar el modelo si la demanda real difiere >15% del pronóstico
   D11: Planeador de Demanda

   A12: DURANTE
   B12: Evitar reposición adicional no planificada (riesgo de sobreinventario)
   C12: El incremento del 40% ya fue considerado en la orden anticipada
   D12: Jefe de Compras

   ─── DESPUÉS DE LA PROMOCIÓN (Mes 6) ───

   A13: DESPUÉS
   B13: NO colocar orden de reposición adicional en el mes 6
   C13: Demanda esperada en mes 6: [CALCULAR] unidades
        (Factor de canibalización = 0.85 — reducción del 15%)
        (referenciar celda E6 de hoja Simulacion_Promocion)
   D13: Jefe de Compras / Planeador de Demanda

   A14: DESPUÉS
   B14: Liberar al personal temporal contratado para la promoción
   C14: Demanda mes 6 vuelve a nivel normal o inferior al base
   D14: Supervisor de Operaciones / RRHH

   A15: DESPUÉS
   B15: Evaluar el sobreinventario potencial y activar acciones de rotación
   C15: Si el stock al final del mes 5 supera [CALCULAR] unidades,
        considerar promoción de liquidación o transferencia a otra sede
   D15: Jefe de Almacén / Comercial

   A16: DESPUÉS
   B16: Documentar los errores del pronóstico (demanda real vs. proyectada)
   C16: Calcular MAE y MAPE post-evento para mejorar el modelo para
        la próxima campaña promocional
   D16: Planeador de Demanda / Analista de Forecasting
   ```

3. Aplica formato visual a la tabla:
   - Fila 4 (encabezados): fondo azul oscuro, texto blanco, negrita
   - Filas "ANTES" (5-8): fondo amarillo claro
   - Filas "DURANTE" (9-12): fondo naranja claro
   - Filas "DESPUÉS" (13-16): fondo verde claro
   - Columna C: fuente en cursiva para distinguir los valores calculados

4. En la celda **A18**, agrega el siguiente bloque de síntesis:

   ```
   A18: SÍNTESIS EJECUTIVA:
   A19: "La simulación demuestra que una promoción del 40% durante 2 meses
         genera una demanda neta adicional de [=Simulacion_Promocion!G9]
         unidades en el período completo (considerando la canibalización del
         mes 6). Para ejecutarla sin quiebres de stock ni colapso operativo,
         se requiere: (1) una orden de compra anticipada de
         [=Simulacion_Promocion!B19] unidades adicionales, (2) la asignación
         de [=Simulacion_Promocion!B32] pickers adicionales en los meses
         pico, y (3) una reducción deliberada de reposición en el mes 6
         para evitar sobreinventario. La IA Generativa identificó
         [completar con número] riesgos adicionales, de los cuales
         [completar] ya estaban contemplados en el modelo."
   ```

**Resultado esperado:** La hoja `Tabla_Decisiones` contiene una tabla completa con 12 acciones logísticas distribuidas en tres fases, con valores numéricos referenciados del modelo y responsables asignados. La síntesis ejecutiva conecta todos los resultados en un párrafo coherente.

**Verificación:** Cuenta las acciones en cada fase: debe haber exactamente 4 acciones en "Antes", 4 en "Durante" y 4 en "Después". Verifica que todas las celdas marcadas como `[CALCULAR]` tengan fórmulas de referencia, no valores escritos manualmente (de lo contrario, si cambias los parámetros de la simulación, la tabla no se actualizará automáticamente).

---

## 7. Validación y Pruebas

Una vez completados los 7 pasos, realiza las siguientes verificaciones integrales:

### Lista de Verificación Final

| # | Verificación | Cómo comprobar | Estado |
|---|---|---|---|
| 1 | El archivo tiene 4 hojas: hoja original, `Simulacion_Promocion`, `Analisis_IA` y `Tabla_Decisiones` | Revisar pestañas en la parte inferior de Excel | ☐ |
| 2 | Los factores de iniciativa en la hoja `Simulacion_Promocion` son: meses 1-3 = 1.00, mes 4 = 1.40, mes 5 = 1.40, mes 6 = 0.85 | Revisar columna D, filas 3-8 | ☐ |
| 3 | El gráfico de líneas muestra dos picos en meses 4-5 y una caída en mes 6 para la serie "CON Promoción" | Inspección visual del gráfico | ☐ |
| 4 | El inventario adicional requerido en el mes 4 (celda B16) es positivo y representa aproximadamente el 40% de la demanda base del mes 4 | Dividir B16/B15 ≈ 0.40 | ☐ |
| 5 | El número de pickers adicionales (celdas B32 y C32) es un número entero positivo | Inspección visual | ☐ |
| 6 | La hoja `Analisis_IA` contiene mínimo 5 riesgos con evaluación de relevancia y cobertura del modelo | Contar filas de la tabla | ☐ |
| 7 | La hoja `Tabla_Decisiones` tiene 12 acciones (4 por fase) con valores referenciados del modelo | Contar filas y verificar fórmulas | ☐ |
| 8 | La síntesis ejecutiva en A19 contiene valores numéricos calculados (no texto genérico) | Revisar que las referencias de celda funcionen | ☐ |

### Prueba de Sensibilidad (opcional si el tiempo lo permite)

Cambia el Factor de Iniciativa en la celda `$I$4` de 1.40 a 1.60 (simulando una promoción más agresiva del 60%) y observa:
- ¿Cuántas unidades adicionales se necesitan ahora?
- ¿Cuántos pickers adicionales se requieren?
- ¿Cambia la síntesis ejecutiva automáticamente?

Si todas las tablas y el gráfico se actualizan automáticamente, el modelo está correctamente construido con referencias dinámicas. Restaura el valor a 1.40 antes de guardar el archivo final.

---

## 8. Solución de Problemas

### Problema 1: El gráfico no muestra la diferencia entre las dos series (ambas líneas son iguales o muy similares)

**Síntoma:** Al insertar el gráfico comparativo en el Paso 5, ambas líneas (CON y SIN promoción) se superponen o muestran valores prácticamente idénticos en todos los meses.

**Causa probable:** Las celdas de la columna F (Pronóstico SIN Promoción) están referenciando la misma hoja `Simulacion_Promocion` en lugar de la hoja original. Si los factores de iniciativa ya se modificaron en la hoja original por error, ambas series mostrarán los mismos valores.

**Solución:**
1. Haz clic en la celda F3 y revisa la barra de fórmulas. La referencia debe apuntar a la hoja **original** (por ejemplo, `=[Pronostico_Base]!E3`), no a la hoja `Simulacion_Promocion`.
2. Si la hoja original también fue modificada, cierra el archivo sin guardar, vuelve a abrir la copia de seguridad del archivo de la Práctica 7 y repite el Paso 1 (crear la copia).
3. Para evitar este error en el futuro: **nunca modifiques la hoja original**; siempre trabaja sobre la copia `Simulacion_Promocion`.

---

### Problema 2: La herramienta de IA Generativa no está disponible o la respuesta es demasiado genérica para completar la tabla de riesgos

**Síntoma:** No hay acceso a internet, la herramienta de IA está caída, o la respuesta obtenida es tan general que no permite identificar 5 riesgos específicos para completar la tabla en la hoja `Analisis_IA`.

**Causa probable:** Problemas de conectividad, límite de consultas gratuitas alcanzado, o prompt demasiado abierto que genera respuestas superficiales.

**Solución:**
1. **Si no hay acceso a la IA:** El instructor proporcionará la captura de pantalla de referencia con una respuesta típica de la IA para este prompt. Utiliza esa referencia para completar la tabla de riesgos.
2. **Si la respuesta es demasiado genérica:** Reformula el prompt agregando más contexto específico:
   ```
   PROMPT ALTERNATIVO MÁS ESPECÍFICO:
   "Soy el planeador de demanda de un centro de distribución de consumo masivo.
   Tenemos un SKU Clase A con demanda base de [X] unidades/mes. Vamos a hacer
   una promoción de 2 meses que esperamos incremente la demanda un 40%.
   Dame exactamente 5 riesgos operativos específicos para el almacén y
   5 acciones preventivas concretas para el equipo de picking."
   ```
3. **Si la herramienta sigue sin funcionar:** Completa la tabla con los siguientes 5 riesgos de referencia basados en literatura logística estándar (Chopra & Meindl, 2016):
   - Quiebre de stock por subestimación del pico
   - Colapso de capacidad de picking por falta de personal
   - Sobreinventario post-promoción por pedidos excesivos
   - Deterioro del nivel de servicio por errores en el procesamiento de órdenes bajo presión
   - Efecto látigo (bullwhip effect) hacia los proveedores por variabilidad extrema en los pedidos

---

## 9. Limpieza y Entrega

### Antes de cerrar el archivo

1. **Guarda el archivo** con el nombre exacto: `Practica8_Simulacion_Promocion.xlsx`
   ```
   Ctrl + S  (o Archivo → Guardar)
   ```

2. **Verifica que el archivo tiene exactamente 4 hojas** y que todas están correctamente nombradas:
   - Hoja del modelo original (nombre heredado de la Práctica 7)
   - `Simulacion_Promocion`
   - `Analisis_IA`
   - `Tabla_Decisiones`

3. **Restaura los parámetros de simulación** a los valores del escenario base si realizaste la prueba de sensibilidad opcional:
   ```
   Celda I4 (Factor de Iniciativa promoción): 1.40
   Celda I5 (Factor post-promoción):          0.85
   ```

4. **Cierra la herramienta de IA Generativa** en el navegador si ya no la necesitas.

### Archivo de salida para la siguiente práctica

| Archivo | Contenido clave para prácticas posteriores |
|---|---|
| `Practica8_Simulacion_Promocion.xlsx` | Modelo de pronóstico con simulación promocional, tabla de decisiones logísticas y análisis de riesgos de IA |

> **Nota:** Esta es la última práctica de la secuencia de forecasting (Prácticas 1-8). El archivo generado puede ser utilizado como insumo de referencia para las Prácticas 9 y 10 si el instructor lo indica.

---

## 10. Resumen y Recursos Adicionales

### Puntos Clave de esta Práctica

| Concepto | Lo que aprendiste |
|---|---|
| **Factor de Iniciativa como palanca de simulación** | Modificar un único parámetro (el Factor de Iniciativa) en el modelo permite simular escenarios complejos como promociones, sin reconstruir el modelo desde cero. Esta es la potencia de un modelo bien estructurado. |
| **Inventario adicional como consecuencia del pronóstico** | El incremento de demanda del 40% no es solo un número estadístico: se traduce directamente en unidades físicas que deben estar en bodega antes de que la promoción inicie. El pronóstico es el insumo, la decisión de compra es el resultado. |
| **Picking como recurso limitado** | El almacén no puede procesar infinitas líneas de picking. Un pico de demanda del 40% puede requerir entre 1 y 4 pickers adicionales, lo que tiene implicaciones de costo, contratación y capacitación que deben planificarse con anticipación. |
| **Canibalización post-promocional** | Las promociones no solo generan demanda adicional; también "roban" demanda futura. El factor 0.85 en el mes 6 representa este efecto y es crítico para evitar sobreinventario en el período post-promoción. |
| **IA Generativa como asistente, no como oráculo** | La IA puede identificar riesgos que el modelo cuantitativo no contempla, pero no tiene acceso a tus datos específicos. El valor del analista está en conectar los riesgos cualitativos de la IA con los números concretos del modelo. |
| **Tabla de decisiones: el puente entre el pronóstico y la operación** | Un pronóstico sin tabla de decisiones es información sin acción. La tabla Antes/Durante/Después convierte el análisis en un plan operativo concreto con responsables y valores de referencia. |

### Conexión con la Lección 1.1

Esta práctica materializa los tres conceptos fundamentales de la Lección 1.1:

- **Rotación de inventario:** El sobreinventario post-promocional (mes 6) reduce la rotación. La tabla de decisiones incluye acciones para mitigarlo.
- **Nivel de servicio:** Sin la planificación anticipada de inventario y picking, el pico del 40% generaría quiebres de stock y deterioro del nivel de servicio exactamente como ilustra el caso de la distribuidora de productos de limpieza de la Lección 1.1.
- **Quiebres de stock:** El stock de seguridad adicional del 20% calculado en el Paso 3 es precisamente el "colchón" que absorbe la variabilidad entre el pronóstico y la demanda real durante la promoción.

### Recursos Adicionales

| Recurso | Enlace / Referencia |
|---|---|
| Chopra & Meindl — *Supply Chain Management* (Capítulo 7: Demand Forecasting) | [Pearson](https://www.pearson.com/en-us/subject-catalog/p/supply-chain-management/P200000005934) |
| CSCMP — Glosario de términos: Promotional Lift, Safety Stock, Picking Productivity | [CSCMP Glossary](https://cscmp.org/CSCMP/Educate/SCM_Definitions_and_Glossary_of_Terms.aspx) |
| MIT OpenCourseWare — Supply Chain Planning (15.762): Demand Management | [MIT OCW](https://ocw.mit.edu/courses/15-762j-supply-chain-planning-spring-2011/) |
| Microsoft Support — Crear gráficos de líneas en Excel | [support.microsoft.com](https://support.microsoft.com/es-es/office/crear-un-grafico-de-lineas-4ffe9b49-0a3d-4b0c-8b44-6d4c5e8e1c1f) |

---

> **Pregunta de cierre para reflexión grupal:** ¿Qué habría pasado en este escenario si el equipo de logística no hubiera tenido un modelo de pronóstico y hubiera descubierto el pico de demanda solo cuando el stock comenzó a agotarse en la semana 2 del mes 4? ¿Cuánto tiempo tendría para reaccionar considerando un lead time típico de 3 semanas con el proveedor?

---
