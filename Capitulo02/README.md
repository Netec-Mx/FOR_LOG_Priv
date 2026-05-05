# Práctica 9 — Cálculo de Estadísticas Básicas Utilizando Funciones de Excel

## Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 11 minutos |
| **Complejidad** | Fácil |
| **Nivel Bloom** | Aplicar |
| **Módulo** | Módulo 2 — Estadística Descriptiva para Forecasting Logístico |
| **Herramientas principales** | Microsoft Excel (2016 o superior) |
| **Archivo de entrada** | `dataset_depurado_modulo1.xlsx` (proporcionado por el instructor) |
| **Archivo de salida** | `practica9_estadisticas_basicas.xlsx` |

---

## Descripción General

En esta práctica aplicarás directamente los conceptos de medidas de tendencia central estudiados en la lección 2.1 para calcular estadísticas descriptivas básicas sobre el historial de demanda mensual de 8 SKUs. Utilizarás las funciones nativas de Excel `PROMEDIO`, `MEDIANA`, `MAX`, `MIN` y `ABS` para construir una tabla de análisis completa que te permitirá identificar qué productos tienen demanda estable y cuáles presentan valores atípicos que distorsionan la media. Adicionalmente, introducirás el concepto de **error de pronóstico** calculando la diferencia entre la demanda real y el promedio histórico, lo cual sienta las bases para las métricas MAE y MAPE que se desarrollarán en prácticas posteriores.

> **Ancla logística:** Cada estadística que calcules en esta práctica responde a una pregunta operativa concreta: el promedio orienta el nivel de reposición, la mediana revela la demanda típica libre de distorsiones, el rango indica el colchón de seguridad necesario, y el error de pronóstico anticipa cuán confiable será cualquier modelo que uses para planificar tu inventario.

---

## Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Calcular el promedio y la mediana de la demanda mensual de cada SKU usando `=PROMEDIO()` y `=MEDIANA()` en Excel, e interpretar sus diferencias en términos de confiabilidad del pronóstico.
- [ ] Obtener los valores máximo y mínimo de cada SKU con `=MAX()` y `=MIN()`, y calcular el rango como medida simple de dispersión para evaluar la variabilidad operativa del inventario.
- [ ] Construir una tabla comparativa Promedio vs. Mediana y aplicar formato condicional para identificar SKUs con diferencias superiores al 15%, señal de presencia de valores atípicos.
- [ ] Calcular el Error de Pronóstico mensual (Demanda Real − Promedio) y el Error Absoluto Medio (MAE) como introducción práctica al ciclo de evaluación de modelos de forecasting.
- [ ] Visualizar la distribución de errores de pronóstico en un gráfico de barras para determinar si los errores son aleatorios o siguen un patrón sistemático.

---

## Prerrequisitos

### Conocimientos previos
- Haber completado las prácticas del Módulo 1 (depuración de datos) o contar con el archivo `dataset_depurado_modulo1.xlsx` proporcionado por el instructor como checkpoint.
- Comprensión conceptual de promedio y mediana (lección 2.1 de este módulo).
- Familiaridad básica con la interfaz de Excel: navegación entre hojas, selección de rangos y escritura de fórmulas.

### Acceso y archivos necesarios
- Archivo `dataset_depurado_modulo1.xlsx` abierto en Excel (24 meses de historial para 8 SKUs, sin valores nulos ni errores de captura).
- Excel 2016 o superior instalado y funcional.
- No se requiere conexión a internet para esta práctica.

---

## Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |

### Software requerido

| Software | Versión mínima | Notas |
|---|---|---|
| Microsoft Excel | 2016 | Microsoft 365 recomendado |
| Archivo de datos | `dataset_depurado_modulo1.xlsx` | Proporcionado por el instructor |

### Configuración inicial (antes de comenzar)

1. Abre el archivo `dataset_depurado_modulo1.xlsx`.
2. Verifica que la hoja `Demanda_Historica` contiene datos en el rango `A1:Z9` (fila 1 = encabezados, columnas B–Y = meses 1 al 24, columna A = código SKU, 8 filas de datos).
3. Crea una copia del archivo con el nombre `practica9_estadisticas_basicas.xlsx` usando **Archivo → Guardar como** antes de realizar cualquier modificación.
4. En el archivo copiado, añade una nueva hoja llamada `Estadisticas` haciendo clic derecho sobre la pestaña de hojas → **Insertar → Hoja de cálculo**.

> **Nota sobre la estructura del dataset:** El dataset depurado contiene 8 SKUs (SKU-01 al SKU-08) con 24 meses de historial de demanda mensual. La columna A contiene los códigos de SKU, las columnas B a Y contienen los valores de demanda mensual (Mes 1 a Mes 24), y la fila 1 contiene los encabezados. Ajusta los rangos de las fórmulas si tu archivo tiene una estructura diferente.

---

## Instrucciones Paso a Paso

---

### Paso 1 — Construir la Tabla de Encabezados en la Hoja `Estadisticas`

**Objetivo:** Preparar la estructura de la tabla de estadísticas descriptivas donde se consolidarán todos los cálculos de esta práctica.

#### Instrucciones

1. Navega a la hoja `Estadisticas` del archivo `practica9_estadisticas_basicas.xlsx`.

2. Escribe los siguientes encabezados en la fila 1, comenzando en la celda `A1`:

   | Celda | Texto del encabezado |
   |---|---|
   | `A1` | `SKU` |
   | `B1` | `Promedio` |
   | `C1` | `Mediana` |
   | `D1` | `Máximo` |
   | `E1` | `Mínimo` |
   | `F1` | `Rango (Max-Min)` |
   | `G1` | `Dif. Abs. Prom-Med` |
   | `H1` | `% Dif. sobre Promedio` |
   | `I1` | `MAE (Error Abs. Medio)` |
   | `J1` | `Señal de Alerta` |

3. En las celdas `A2:A9`, escribe los códigos de SKU: `SKU-01`, `SKU-02`, `SKU-03`, `SKU-04`, `SKU-05`, `SKU-06`, `SKU-07`, `SKU-08`.

4. Aplica **negrita** a la fila 1 completa: selecciona `A1:J1` y presiona `Ctrl + N`.

5. Aplica un color de fondo gris claro a la fila de encabezados: con `A1:J1` seleccionado, usa **Inicio → Color de relleno → Gris, Énfasis 3, Claro 60%** (o cualquier gris claro disponible).

#### Salida esperada

Una tabla con 10 columnas encabezadas correctamente y 8 filas de códigos SKU listos para recibir las fórmulas.

#### Verificación

Confirma que la celda `A1` contiene el texto `SKU` y la celda `J1` contiene `Señal de Alerta`. Los códigos de SKU en `A2:A9` deben coincidir exactamente con los de la hoja `Demanda_Historica`.

---

### Paso 2 — Calcular el Promedio de Demanda por SKU

**Objetivo:** Aplicar la función `=PROMEDIO()` para obtener la demanda media mensual de cada SKU a partir de los 24 meses de historial.

#### Instrucciones

1. Haz clic en la celda `B2` de la hoja `Estadisticas`.

2. Escribe la siguiente fórmula para calcular el promedio del SKU-01:

   ```excel
   =PROMEDIO(Demanda_Historica!B2:Y2)
   ```

   > **Explicación:** `Demanda_Historica!B2:Y2` hace referencia a los 24 meses de demanda del SKU-01 en la hoja `Demanda_Historica`. Las columnas B a Y corresponden a los meses 1 al 24.

3. Presiona `Enter` para confirmar la fórmula.

4. Haz clic nuevamente en `B2`. Posiciona el cursor sobre la esquina inferior derecha de la celda hasta que aparezca la cruz negra (+) y arrastra hacia abajo hasta `B9` para copiar la fórmula a los 8 SKUs.

5. Verifica que las referencias de fila se ajusten automáticamente: `B3` debe contener `=PROMEDIO(Demanda_Historica!B3:Y3)`, `B4` debe contener `=PROMEDIO(Demanda_Historica!B4:Y4)`, y así sucesivamente.

6. Selecciona `B2:B9` y aplica formato de número con **2 decimales**: clic derecho → **Formato de celdas → Número → Decimales: 2**.

#### Salida esperada

La columna `B` muestra el promedio mensual de demanda para cada uno de los 8 SKUs, expresado con 2 decimales. Los valores deben ser coherentes con el rango general de demanda del dataset (típicamente entre 50 y 500 unidades según el dataset de práctica).

#### Verificación

Selecciona manualmente el rango `Demanda_Historica!B2:Y2` y observa el valor `Promedio` en la barra de estado de Excel (parte inferior derecha de la pantalla). Debe coincidir con el valor calculado en `B2`.

---

### Paso 3 — Calcular la Mediana de Demanda por SKU

**Objetivo:** Aplicar la función `=MEDIANA()` para obtener el valor central de la distribución de demanda de cada SKU, resistente a la influencia de valores atípicos.

#### Instrucciones

1. Haz clic en la celda `C2` de la hoja `Estadisticas`.

2. Escribe la siguiente fórmula para el SKU-01:

   ```excel
   =MEDIANA(Demanda_Historica!B2:Y2)
   ```

3. Presiona `Enter` y arrastra la fórmula hacia abajo hasta `C9` usando el mismo procedimiento del Paso 2.

4. Aplica formato de número con **2 decimales** al rango `C2:C9`.

5. Una vez completadas ambas columnas, realiza una **inspección visual rápida**: compara mentalmente los valores de la columna B (Promedio) con los de la columna C (Mediana) para cada SKU. Anota mentalmente cuáles SKUs muestran diferencias notables.

#### Salida esperada

La columna `C` muestra la mediana mensual para cada SKU. Para SKUs con demanda estable, los valores de promedio y mediana deberán ser muy cercanos. Para SKUs con eventos atípicos, la mediana será notablemente menor que el promedio.

#### Verificación

Selecciona el rango `Demanda_Historica!B2:Y2` y verifica el valor `Mediana` en la barra de estado (si no aparece, haz clic derecho sobre la barra de estado y activa `Mediana`). El valor debe coincidir con `C2`.

---

### Paso 4 — Calcular Máximo, Mínimo y Rango por SKU

**Objetivo:** Identificar los valores extremos de demanda de cada SKU y calcular el rango como medida simple de dispersión para estimar la amplitud del colchón de seguridad necesario.

#### Instrucciones

1. En la celda `D2`, escribe la fórmula para el valor máximo del SKU-01:

   ```excel
   =MAX(Demanda_Historica!B2:Y2)
   ```

2. En la celda `E2`, escribe la fórmula para el valor mínimo del SKU-01:

   ```excel
   =MIN(Demanda_Historica!B2:Y2)
   ```

3. Arrastra ambas fórmulas hacia abajo hasta las filas 9 (`D9` y `E9` respectivamente).

4. En la celda `F2`, calcula el **Rango** como la diferencia entre el máximo y el mínimo:

   ```excel
   =D2-E2
   ```

5. Arrastra la fórmula de `F2` hacia abajo hasta `F9`.

6. Aplica formato de número entero (sin decimales) a los rangos `D2:F9`: selecciona el rango → clic derecho → **Formato de celdas → Número → Decimales: 0**.

7. **Reflexión operativa (30 segundos):** Identifica el SKU con el mayor rango (Max − Min). Este producto requiere el mayor stock de seguridad porque su demanda oscila más. Anota mentalmente este SKU.

#### Salida esperada

| Columna | Contenido |
|---|---|
| D | Valor máximo de demanda mensual por SKU (entero) |
| E | Valor mínimo de demanda mensual por SKU (entero) |
| F | Rango = Max − Min por SKU (entero) |

El SKU con el rango más alto es el candidato principal para revisar su política de stock de seguridad.

#### Verificación

Para cualquier SKU, verifica manualmente que `D[fila] ≥ B[fila] ≥ C[fila] ≥ E[fila]` (el máximo ≥ promedio ≥ mediana ≥ mínimo no siempre se cumple en orden exacto, pero el máximo siempre debe ser ≥ promedio y el mínimo siempre ≤ promedio). Si alguna celda muestra un valor de promedio mayor que el máximo, hay un error en los rangos de referencia.

---

### Paso 5 — Calcular la Diferencia Promedio vs. Mediana y Aplicar Señal de Alerta

**Objetivo:** Cuantificar la diferencia entre el promedio y la mediana para cada SKU e identificar automáticamente con formato condicional aquellos productos cuya diferencia supera el umbral del 15%, señal de valores atípicos que distorsionan la media.

#### Instrucciones

1. En la celda `G2`, calcula la diferencia absoluta entre promedio y mediana:

   ```excel
   =ABS(B2-C2)
   ```

2. Arrastra la fórmula hacia abajo hasta `G9`.

3. En la celda `H2`, calcula el porcentaje que representa esa diferencia sobre el promedio:

   ```excel
   =SI(B2>0, G2/B2, 0)
   ```

   > **Nota:** La condición `SI(B2>0, ...)` evita un error de división por cero en caso de que algún SKU tenga promedio cero.

4. Arrastra la fórmula hacia abajo hasta `H9`.

5. Selecciona `H2:H9` y aplica formato de **Porcentaje con 1 decimal**: clic derecho → **Formato de celdas → Porcentaje → Decimales: 1**.

6. En la celda `J2`, escribe la fórmula de señal de alerta:

   ```excel
   =SI(H2>0.15, "⚠ REVISAR OUTLIERS", "✓ OK")
   ```

7. Arrastra la fórmula hacia abajo hasta `J9`.

8. **Aplicar formato condicional a la columna J:**
   - Selecciona `J2:J9`.
   - Ve a **Inicio → Formato condicional → Nueva regla**.
   - Selecciona **"Aplicar formato únicamente a las celdas que contengan"**.
   - Configura: **Texto específico → que contenga → REVISAR**.
   - Haz clic en **Formato** → pestaña **Relleno** → selecciona **amarillo** → **Aceptar → Aceptar**.
   - Repite el proceso para las celdas con "✓ OK" aplicando relleno **verde claro**.

9. **Aplicar formato condicional adicional a la columna H:**
   - Selecciona `H2:H9`.
   - Ve a **Inicio → Formato condicional → Escalas de color**.
   - Selecciona la escala **Verde-Amarillo-Rojo** (los porcentajes bajos en verde, los altos en rojo).

#### Salida esperada

- La columna `G` muestra la diferencia absoluta en unidades entre promedio y mediana.
- La columna `H` muestra el porcentaje de diferencia con escala de color (verde = diferencia pequeña, rojo = diferencia grande).
- La columna `J` muestra `⚠ REVISAR OUTLIERS` en amarillo para SKUs con diferencia > 15% y `✓ OK` en verde para los demás.

#### Verificación

Identifica al menos un SKU con la etiqueta `⚠ REVISAR OUTLIERS`. Si ningún SKU supera el 15%, verifica con el instructor que el dataset incluye los valores atípicos intencionados. Si todos los SKUs muestran alerta, revisa que la fórmula en `H2` divida correctamente `G2/B2` y no `B2/G2`.

---

### Paso 6 — Calcular el Error de Pronóstico Mensual (Modelo Naïve)

**Objetivo:** Introducir el concepto de error de pronóstico calculando, para cada mes, la diferencia entre la demanda real y el promedio histórico del SKU. Este cálculo simula el error de un modelo naïve que siempre predice el promedio histórico.

#### Instrucciones

1. En el archivo `practica9_estadisticas_basicas.xlsx`, crea una nueva hoja llamada `Errores_Pronostico`: clic derecho sobre una pestaña → **Insertar → Hoja de cálculo** → renombra como `Errores_Pronostico`.

2. En la hoja `Errores_Pronostico`, construye la siguiente estructura de encabezados en la fila 1:

   | Celda | Encabezado |
   |---|---|
   | `A1` | `SKU` |
   | `B1` | `Mes 1` |
   | `C1` | `Mes 2` |
   | ... | ... |
   | `Y1` | `Mes 24` |
   | `Z1` | `MAE` |

   > **Tip de eficiencia:** Para crear los encabezados de mes rápidamente, escribe `Mes 1` en `B1`, luego escribe `Mes 2` en `C1`. Selecciona `B1:C1` y arrastra el controlador de relleno hacia la derecha hasta `Y1`. Excel completará automáticamente la serie.

3. Copia los códigos de SKU de la hoja `Estadisticas`: en `A2` escribe `=Estadisticas!A2` y arrastra hasta `A9`.

4. En la celda `B2`, escribe la fórmula de error de pronóstico para el SKU-01, Mes 1:

   ```excel
   =Demanda_Historica!B2 - Estadisticas!$B2
   ```

   > **Explicación:** `Demanda_Historica!B2` es la demanda real del SKU-01 en el Mes 1. `Estadisticas!$B2` es el promedio histórico del SKU-01 (la columna B está fija con `$B` para que no se desplace al copiar horizontalmente, pero la fila 2 sí puede variar al copiar verticalmente).

5. Copia la fórmula de `B2` hacia la derecha hasta `Y2` (24 meses), y luego hacia abajo hasta la fila 9 (8 SKUs). El bloque `B2:Y9` debe quedar completamente lleno con los errores de pronóstico.

6. Aplica formato de número con **1 decimal** al rango `B2:Y9`.

7. **Observación rápida:** Revisa visualmente el bloque de errores. Los valores positivos indican meses donde la demanda real superó el promedio (infrapredicción); los negativos indican meses donde la demanda fue menor que el promedio (sobrepredicción).

#### Salida esperada

Una tabla de 8 filas × 24 columnas con los errores de pronóstico (Demanda Real − Promedio) para cada SKU y cada mes. Los errores deben oscilar tanto en positivo como en negativo alrededor de cero para SKUs con demanda estable.

#### Verificación

Para el SKU-01, suma manualmente los errores del Mes 1 al Mes 24 usando `=SUMA(B2:Y2)`. Para un modelo naïve basado en el promedio histórico, esta suma debe ser igual a cero (o muy cercana a cero), ya que el promedio es justamente el valor que minimiza la suma de errores. Si el resultado es significativamente diferente de cero, verifica que `Estadisticas!$B2` hace referencia correcta al promedio del SKU-01.

---

### Paso 7 — Calcular el MAE (Error Absoluto Medio) por SKU

**Objetivo:** Calcular el Error Absoluto Medio (MAE) como indicador de la magnitud promedio del error del modelo naïve, introduciendo la métrica fundamental de evaluación de pronósticos que se usará en prácticas posteriores.

#### Instrucciones

1. En la hoja `Errores_Pronostico`, haz clic en la celda `Z2`.

2. Escribe la fórmula del MAE para el SKU-01:

   ```excel
   =PROMEDIO(ABS(B2:Y2))
   ```

   > **¡Importante!** Esta es una **fórmula matricial**. Debes confirmarla presionando **`Ctrl + Shift + Enter`** en lugar de solo `Enter`. Excel mostrará la fórmula entre llaves `{=PROMEDIO(ABS(B2:Y2))}` para indicar que es matricial.
   >
   > **Alternativa para Excel 365/2019:** En versiones modernas de Excel con soporte de matrices dinámicas, puedes confirmar con `Enter` normal y la función funcionará correctamente sin llaves.

3. Arrastra la fórmula hacia abajo hasta `Z9`.

4. Aplica formato de número con **2 decimales** al rango `Z2:Z9`.

5. Regresa a la hoja `Estadisticas` y en la columna `I2`, enlaza el MAE calculado:

   ```excel
   =Errores_Pronostico!Z2
   ```

6. Arrastra hacia abajo hasta `I9`.

7. **Interpretación logística:** El SKU con el MAE más alto es el que tiene mayor error promedio al usar el promedio histórico como pronóstico. En términos operativos, este SKU genera mayor incertidumbre en la planificación del inventario y podría requerir un modelo de pronóstico más sofisticado.

#### Salida esperada

La columna `Z` de la hoja `Errores_Pronostico` y la columna `I` de la hoja `Estadisticas` muestran el MAE de cada SKU. El MAE está expresado en las mismas unidades que la demanda (unidades por mes), lo que facilita su interpretación operativa.

#### Verificación

Verifica que el MAE de cada SKU sea un número positivo y que sea menor o igual al Rango (Max − Min) calculado en el Paso 4. Un MAE mayor que el rango sería matemáticamente imposible y señalaría un error en las fórmulas. Compara también el MAE con el Promedio: un MAE que representa más del 30% del promedio indica alta variabilidad.

---

### Paso 8 — Visualizar la Distribución de Errores con un Gráfico de Barras

**Objetivo:** Crear un gráfico de barras que muestre los errores de pronóstico mes a mes para un SKU seleccionado, permitiendo identificar visualmente si los errores son aleatorios (patrón deseable) o siguen un patrón sistemático (señal de que el modelo naïve es inadecuado).

#### Instrucciones

1. En la hoja `Errores_Pronostico`, selecciona el rango de encabezados de mes y los errores del **SKU con el MAE más alto** (identificado en el Paso 7). Por ejemplo, si el SKU-03 tiene el MAE más alto, selecciona `B1:Y1` (encabezados) manteniendo `Ctrl` y luego `B4:Y4` (errores del SKU-03).

   > **Tip:** Si no puedes seleccionar rangos no contiguos con el gráfico, crea una pequeña tabla auxiliar debajo de los datos principales (por ejemplo, a partir de la fila 12) con los encabezados de mes en una fila y los errores del SKU seleccionado en la siguiente.

2. Con el rango seleccionado, ve a **Insertar → Gráficos → Gráfico de barras/columnas → Columna agrupada 2D**.

3. Haz clic derecho sobre el gráfico → **Mover gráfico** → selecciona **Nueva hoja** → escribe el nombre `Grafico_Errores` → **Aceptar**.

4. Personaliza el gráfico:
   - **Título del gráfico:** Haz doble clic sobre el título y escribe: `Distribución de Errores de Pronóstico — [SKU seleccionado]`
   - **Eje Y (vertical):** Haz clic derecho sobre el eje → **Dar formato a eje** → en **Número**, selecciona formato de número con 0 decimales.
   - **Línea de referencia en cero:** Ve a **Diseño de gráfico → Agregar elemento de gráfico → Líneas → Línea de valor cero** (o equivalente según versión de Excel). Esto añade una línea horizontal en y=0 que facilita identificar meses con error positivo vs. negativo.
   - **Colores:** Selecciona las barras → clic derecho → **Dar formato a serie de datos** → aplica un color azul para las barras positivas. Para diferenciar las negativas, puedes usar **Formato condicional** de colores en la tabla auxiliar si tu versión de Excel lo permite en gráficos.

5. **Análisis visual (1 minuto):** Observa el gráfico y responde mentalmente:
   - ¿Las barras alternan entre positivo y negativo de forma irregular? → Errores aleatorios (buen signo: el modelo no tiene sesgo sistemático).
   - ¿Hay una secuencia de barras todas positivas o todas negativas? → Patrón sistemático (el promedio subestima o sobreestima consistentemente, posible tendencia o estacionalidad no capturada).
   - ¿Hay una barra claramente más grande que las demás? → Posible outlier en ese mes específico.

#### Salida esperada

Un gráfico de columnas en la hoja `Grafico_Errores` que muestra los 24 errores mensuales del SKU seleccionado, con una línea de referencia en cero. La interpretación visual del patrón de errores es el producto principal de este paso.

#### Verificación

Confirma que el gráfico muestra exactamente 24 barras (una por mes) y que los valores de las barras coinciden con los valores de la tabla de errores. Haz clic sobre cualquier barra; el tooltip debe mostrar el valor del error para ese mes.

---

## Validación y Pruebas

Una vez completados todos los pasos, realiza las siguientes verificaciones de integridad para confirmar que los cálculos son correctos:

### Lista de verificación final

| # | Verificación | Cómo comprobarla | Resultado esperado |
|---|---|---|---|
| 1 | El promedio de cada SKU coincide con la barra de estado de Excel | Selecciona el rango de demanda de cualquier SKU en `Demanda_Historica` y observa `Promedio` en la barra de estado | Coincidencia exacta |
| 2 | La mediana es ≤ promedio cuando hay valores atípicos positivos | Compara columnas B y C en la hoja `Estadisticas` para SKUs con alerta | Mediana < Promedio en SKUs con `⚠ REVISAR OUTLIERS` |
| 3 | El Rango (Max − Min) es positivo para todos los SKUs | Revisa columna F en `Estadisticas` | Todos los valores > 0 |
| 4 | La suma de errores por SKU es ≈ 0 | En `Errores_Pronostico`, calcula `=SUMA(B2:Y2)` para cualquier SKU | Valor muy cercano a 0 (puede haber diferencia por redondeo) |
| 5 | El MAE es positivo para todos los SKUs | Revisa columna Z en `Errores_Pronostico` | Todos los valores > 0 |
| 6 | Al menos un SKU muestra `⚠ REVISAR OUTLIERS` | Revisa columna J en `Estadisticas` | Mínimo 1 SKU con alerta amarilla |
| 7 | El gráfico muestra 24 barras | Cuenta visualmente las barras en `Grafico_Errores` | Exactamente 24 barras |

### Tabla de estadísticas de referencia (valores aproximados)

> **Nota:** Los valores exactos dependen del dataset específico proporcionado por el instructor. La siguiente tabla muestra rangos típicos esperados para un dataset de práctica estándar.

| Estadística | Rango típico esperado |
|---|---|
| Promedio mensual por SKU | 80 – 450 unidades |
| Mediana mensual por SKU | 75 – 440 unidades |
| Rango (Max − Min) por SKU | 30 – 300 unidades |
| % Diferencia Prom-Med | 0% – 35% |
| MAE del modelo naïve | 15 – 120 unidades |

---

## Resolución de Problemas

### Problema 1: La fórmula `=PROMEDIO(ABS(...))` no funciona correctamente y devuelve un error o el valor incorrecto

**Síntoma:** La celda `Z2` muestra `#¡VALOR!`, `#REF!`, o un número que parece ser el promedio sin aplicar el valor absoluto (igual al resultado de `=PROMEDIO(B2:Y2)`).

**Causa:** En Excel 2016 y versiones anteriores a las matrices dinámicas (Excel 365/2019), la función `ABS()` no puede aplicarse directamente a un rango dentro de `PROMEDIO()` sin confirmar la fórmula como matricial. Si se presionó solo `Enter` en lugar de `Ctrl + Shift + Enter`, Excel interpreta `ABS(B2:Y2)` solo sobre la primera celda del rango (`B2`), ignorando el resto.

**Solución:**
1. Haz clic en la celda `Z2`.
2. Observa la barra de fórmulas: si la fórmula NO aparece entre llaves `{}`, no fue confirmada como matricial.
3. Haz clic en la barra de fórmulas para entrar en modo de edición.
4. Presiona `Ctrl + Shift + Enter` para confirmar correctamente.
5. La fórmula debe aparecer ahora como `{=PROMEDIO(ABS(B2:Y2))}`.
6. Si usas Excel 365, la fórmula funciona con `Enter` normal; verifica que el resultado sea diferente al promedio simple (debe ser mayor o igual, ya que los valores absolutos nunca son negativos).

**Alternativa sin fórmula matricial** (compatible con todas las versiones):
```excel
=PROMEDIO(SI(B2:Y2>=0, B2:Y2, -B2:Y2))
```
Confirmar también con `Ctrl + Shift + Enter` en versiones antiguas. O usar una columna auxiliar que calcule `=ABS(B2)`, `=ABS(C2)`, etc., y luego calcular el promedio de esa columna auxiliar.

---

### Problema 2: Las referencias cruzadas entre hojas muestran `#¡REF!` al copiar fórmulas hacia abajo

**Síntoma:** Al arrastrar hacia abajo la fórmula `=Demanda_Historica!B2 - Estadisticas!$B2` desde `B2` hasta `B9` en la hoja `Errores_Pronostico`, algunas celdas muestran `#¡REF!` en lugar del error calculado.

**Causa:** La referencia a la columna del promedio en la hoja `Estadisticas` no tiene correctamente fijada la columna con el signo `$`. Si la fórmula se escribió como `=Demanda_Historica!B2 - Estadisticas!B2` (sin el `$` antes de `B`), al arrastrar hacia la derecha (para los 24 meses), la referencia al promedio también se desplaza horizontalmente (`C2`, `D2`, `E2`...) apuntando a columnas que no contienen el promedio sino otros estadísticos o celdas vacías.

**Solución:**
1. Haz clic en la celda `B2` de la hoja `Errores_Pronostico`.
2. En la barra de fórmulas, verifica que la referencia al promedio sea `Estadisticas!$B2` (con `$` antes de `B`, fijando la columna, pero no la fila).
3. Si la referencia es incorrecta, edita la fórmula y agrega el `$`:
   ```excel
   =Demanda_Historica!B2 - Estadisticas!$B2
   ```
4. Confirma con `Enter`.
5. Selecciona `B2` nuevamente y arrastra primero **hacia la derecha** hasta `Y2` (los 24 meses del SKU-01), luego selecciona `B2:Y2` y arrastra **hacia abajo** hasta la fila 9.
6. Verifica que `B3` contenga `=Demanda_Historica!B3 - Estadisticas!$B3` (la fila de `Estadisticas` debe haberse ajustado a 3 para apuntar al promedio del SKU-02).

---

## Limpieza y Entrega

Al finalizar la práctica, realiza los siguientes pasos de cierre:

1. **Guarda el archivo final:** Presiona `Ctrl + S` para guardar todos los cambios en `practica9_estadisticas_basicas.xlsx`.

2. **Verifica la estructura de hojas:** El archivo debe contener exactamente 3 hojas:
   - `Demanda_Historica` (datos originales, sin modificaciones)
   - `Estadisticas` (tabla de estadísticas descriptivas completa con formato condicional)
   - `Errores_Pronostico` (tabla de errores mensuales y MAE por SKU)
   - `Grafico_Errores` (gráfico de distribución de errores)

3. **Nombra el archivo con tu identificación:** Renombra el archivo como `practica9_[TuNombre]_estadisticas.xlsx` antes de entregarlo al instructor.

4. **Archivos que NO deben modificarse:** El archivo `dataset_depurado_modulo1.xlsx` original debe permanecer intacto. Toda la práctica se realizó sobre la copia `practica9_estadisticas_basicas.xlsx`.

5. **Este archivo es insumo para prácticas posteriores:** La hoja `Estadisticas` (especialmente las columnas de Promedio, MAE y la clasificación de alertas) será utilizada como punto de partida en la Práctica 10. No elimines ninguna columna ni cambies la estructura de la tabla.

---

## Resumen

En esta práctica aplicaste las funciones estadísticas básicas de Excel para construir un análisis descriptivo completo de la demanda histórica de 8 SKUs. Los conceptos clave que pusiste en práctica fueron:

| Concepto | Función Excel | Uso logístico |
|---|---|---|
| **Promedio** | `=PROMEDIO()` | Estimación del nivel de reposición base |
| **Mediana** | `=MEDIANA()` | Demanda típica robusta frente a eventos atípicos |
| **Máximo / Mínimo** | `=MAX()` / `=MIN()` | Límites del rango de variación para stock de seguridad |
| **Rango** | `=MAX()-MIN()` | Amplitud de la variabilidad operativa |
| **Diferencia Prom-Med** | `=ABS(PROM-MED)` | Indicador de presencia de outliers (umbral: >15%) |
| **Error de pronóstico** | `=Demanda-Promedio` | Base del ciclo de evaluación de modelos |
| **MAE** | `=PROMEDIO(ABS(...))` | Magnitud promedio del error del modelo naïve |

### Conclusión logística

La diferencia entre promedio y mediana es tu primera herramienta de diagnóstico: cuando supera el 15%, el promedio no es un indicador confiable para planificar el inventario de ese SKU sin antes depurar o ajustar los datos. El MAE del modelo naïve (que siempre predice el promedio) establece el **benchmark mínimo de referencia**: cualquier modelo de pronóstico más sofisticado debe producir un MAE menor para justificar su complejidad adicional.

### Recursos de apoyo

- [Función PROMEDIO — Soporte Microsoft](https://support.microsoft.com/es-es/office/promedio-funci%C3%B3n-promedio-047bac88-d466-426c-a32b-8f33eb960cf6)
- [Función MEDIANA — Soporte Microsoft](https://support.microsoft.com/es-es/office/mediana-funci%C3%B3n-mediana-d0916313-4753-414c-8537-ce85bdd967d2)
- [Fórmulas matriciales en Excel — Soporte Microsoft](https://support.microsoft.com/es-es/office/instrucciones-y-ejemplos-de-f%C3%B3rmulas-de-matriz-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
- [Hyndman & Athanasopoulos — *Forecasting: Principles and Practice*, Cap. 3: Herramientas de pronóstico](https://otexts.com/fpp3/toolbox.html)

---

> **Próxima práctica — Práctica 10:** Aplicarás la desviación estándar y el coeficiente de variación para clasificar los 8 SKUs según su estabilidad de demanda, utilizando como base la tabla de estadísticas construida en esta práctica. Los SKUs clasificados como de alta variabilidad recibirán un tratamiento diferenciado en la selección del modelo de pronóstico.

---

# Práctica 10 — Cálculo de Desviación Estándar y Análisis de Variabilidad de la Demanda

## 1. Metadatos

| Campo             | Detalle                                      |
|-------------------|----------------------------------------------|
| **Duración**      | 11 minutos                                   |
| **Complejidad**   | Media                                        |
| **Nivel Bloom**   | Aplicar (Apply)                              |
| **Práctica N°**   | 10 de 10                                     |
| **Módulo**        | 2 — Estadística Descriptiva para Forecasting |
| **Herramientas**  | Microsoft Excel 2016 o superior              |

---

## 2. Descripción General

Esta práctica cierra el ciclo estadístico del módulo 2. Partiendo de las estadísticas básicas calculadas en la Práctica 9 (promedio, mediana) y de la clasificación ABC obtenida en la Práctica 5, el participante calculará la **desviación estándar muestral** y el **coeficiente de variación (CV)** para los 8 SKUs del dataset, construirá una segmentación por estabilidad de demanda y estimará el **inventario de seguridad básico** para los SKUs de Clase A.

El hilo conductor es la pregunta logística central: *¿cuánto varía la demanda de cada producto y qué implica esa variabilidad para el pronóstico, el stock de seguridad y el nivel de servicio?* Al finalizar, el participante podrá vincular directamente la variabilidad estadística con decisiones operativas concretas de inventario y abastecimiento.

---

## 3. Objetivos de Aprendizaje

- [ ] Calcular la desviación estándar muestral (`DESVEST.M`) de la demanda histórica por SKU e interpretar su significado en el contexto del forecasting logístico.
- [ ] Calcular el coeficiente de variación (CV) para cada SKU y utilizarlo como criterio de clasificación de productos según su estabilidad de demanda.
- [ ] Segmentar los SKUs en categorías de estabilidad (Alta, Media, Baja) basándose en el CV calculado y derivar implicaciones para la estrategia de pronóstico.
- [ ] Calcular el inventario de seguridad básico para los SKUs Clase A y evaluar su razonabilidad frente al promedio mensual de demanda.
- [ ] Relacionar la variabilidad de la demanda con el error de pronóstico esperado, cerrando el ciclo conceptual iniciado en la Práctica 9.

---

## 4. Prerrequisitos

### Conocimiento previo
- Haber completado la **Práctica 9** (estadísticas básicas: promedio, mediana, máximo, mínimo calculados por SKU).
- Haber completado la **Práctica 5** (clasificación ABC disponible en el archivo de trabajo).
- Comprensión de promedio y mediana (Lección 2.1): saber interpretar la diferencia entre ambas como señal de presencia de valores atípicos.
- Comprensión introductoria de desviación estándar, coeficiente de variación y error de pronóstico (Lecciones 2.3–2.5).

### Acceso requerido
- Archivo Excel de práctica con el dataset de 24 meses para 8 SKUs (proporcionado por el instructor), con la hoja `Práctica 9` completada y la hoja `ABC_Clasificación` disponible.
- Microsoft Excel 2016 o superior instalado y funcional.
- No se requiere conexión a internet para esta práctica.

> **Nota de secuencialidad:** Si no completaste la Práctica 9, solicita al instructor el archivo *checkpoint* correspondiente antes de continuar.

---

## 5. Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente        | Mínimo                        | Recomendado                    |
|-------------------|-------------------------------|--------------------------------|
| Procesador        | Intel Core i5 / AMD Ryzen 5   | Intel Core i7 / AMD Ryzen 7    |
| RAM               | 4 GB                          | 8 GB                           |
| Almacenamiento    | 2 GB libres en disco          | 2 GB libres en disco           |
| Resolución        | 1366 × 768                    | 1920 × 1080                    |

### Software requerido

| Software               | Versión mínima         | Notas                                              |
|------------------------|------------------------|----------------------------------------------------|
| Microsoft Excel        | 2016 o superior        | `DESVEST.M` no disponible en versiones anteriores  |
| Archivo de práctica    | Provisto por instructor| Debe incluir hojas de Práctica 9 y ABC_Clasificación|

### Preparación del entorno (antes de comenzar)

1. Abre el archivo Excel de práctica provisto por el instructor.
2. Verifica que el archivo contiene las siguientes hojas:
   - `Datos_Historicos` — datos crudos de 24 meses para 8 SKUs.
   - `Práctica_9` — promedio y mediana ya calculados por SKU.
   - `ABC_Clasificación` — tabla con la clase A/B/C asignada a cada SKU.
3. Crea una nueva hoja llamada `Práctica_10` haciendo clic derecho en las pestañas → **Insertar** → **Hoja de cálculo**.
4. Guarda el archivo inmediatamente con `Ctrl + S`.

---

## 6. Procedimiento Paso a Paso

> **Gestión del tiempo:** Esta práctica tiene 11 minutos. Los pasos 1–3 deben completarse en los primeros 4 minutos; los pasos 4–5 en los siguientes 4 minutos; los pasos 6–7 en los últimos 3 minutos.

---

### Paso 1 — Construir la tabla base en la hoja `Práctica_10`

**Objetivo:** Crear la estructura de trabajo que consolidará todos los cálculos de esta práctica en un solo lugar.

**Instrucciones:**

1. Activa la hoja `Práctica_10`.
2. En la fila 1, escribe los siguientes encabezados en las columnas A a J:

   | Col | Encabezado           |
   |-----|----------------------|
   | A   | `SKU`                |
   | B   | `Descripción`        |
   | C   | `Clase ABC`          |
   | D   | `Promedio (μ)`       |
   | E   | `Desv. Estándar (σ)` |
   | F   | `CV (%)`             |
   | G   | `Clasificación CV`   |
   | H   | `IS (unidades)`      |
   | I   | `IS / Promedio (%)`  |
   | J   | `Riesgo Operativo`   |

3. En las columnas A, B y C (filas 2 a 9), ingresa los 8 SKUs con su descripción y clase ABC. Copia estos valores directamente desde la hoja `ABC_Clasificación` para garantizar consistencia:
   - Selecciona la tabla en `ABC_Clasificación`, copia las columnas SKU, Descripción y Clase.
   - Pega en `Práctica_10` a partir de la celda A2 usando **Pegado especial → Solo valores** (`Ctrl + Alt + V` → seleccionar *Valores* → Aceptar).

4. En la columna D (Promedio), copia los promedios ya calculados en `Práctica_9`. Usa referencias directas a esa hoja para mantener la trazabilidad. Ejemplo para el primer SKU en D2:
   ```excel
   =Práctica_9!D2
   ```
   Ajusta la referencia según la ubicación real del promedio en tu hoja de Práctica 9.

**Resultado esperado:** Una tabla con 8 filas de datos (una por SKU) y los encabezados correctamente definidos. Las columnas A, B, C y D deben estar completas; las columnas E a J estarán vacías por ahora.

**Verificación:** Confirma que los valores de la columna D (Promedio) coinciden exactamente con los calculados en la Práctica 9. Si hay discrepancia, revisa la referencia de celda.

---

### Paso 2 — Calcular la Desviación Estándar Muestral con `DESVEST.M`

**Objetivo:** Calcular la dispersión de la demanda mensual de cada SKU usando la función estadística correcta para datos muestrales.

**Instrucciones:**

1. Haz clic en la celda **E2** (primera fila de datos, columna Desv. Estándar).

2. Escribe la siguiente fórmula, adaptando el rango al bloque de datos del primer SKU en la hoja `Datos_Historicos`. Asumiendo que los 24 meses del SKU-01 están en la columna B (filas 2 a 25) de esa hoja:
   ```excel
   =DESVEST.M(Datos_Historicos!B2:B25)
   ```

3. Repite el proceso para los 8 SKUs (E2 a E9), apuntando al rango de datos correspondiente a cada SKU en `Datos_Historicos`.

   > **Nota técnica importante:** Usa `DESVEST.M` (desviación estándar *muestral*), **no** `DESVEST.P` (poblacional). En forecasting logístico trabajamos con muestras históricas, no con la población completa de todos los períodos posibles. `DESVEST.M` aplica el factor de corrección de Bessel (divide entre *n-1*), lo que produce una estimación más conservadora y apropiada para dimensionar el stock de seguridad.

4. Formatea las celdas E2:E9 con **2 decimales** (selecciona el rango → `Ctrl + 1` → Número → 2 decimales).

**Resultado esperado:** La columna E muestra un valor de desviación estándar para cada SKU. Los valores deben ser positivos y, en general, menores que el promedio correspondiente (aunque en productos con demanda muy errática pueden acercarse o superar al promedio).

**Verificación:** Selecciona el rango de datos de un SKU directamente en `Datos_Historicos` y observa el valor de `DESVEST` que Excel muestra en la barra de estado (parte inferior derecha). Debe coincidir con lo calculado en la columna E.

---

### Paso 3 — Calcular el Coeficiente de Variación (CV)

**Objetivo:** Normalizar la desviación estándar respecto al promedio para obtener una medida de variabilidad relativa comparable entre SKUs con diferentes escalas de demanda.

**Instrucciones:**

1. Haz clic en la celda **F2**.

2. Escribe la fórmula del CV expresado en porcentaje:
   ```excel
   =E2/D2*100
   ```
   Esta fórmula implementa: **CV (%) = (Desviación Estándar / Promedio) × 100**

3. Copia la fórmula hacia abajo hasta F9 (para los 8 SKUs).

4. Formatea las celdas F2:F9 con **1 decimal** y agrega el símbolo `%` manualmente en el encabezado (ya está en la columna F del encabezado que creaste en el Paso 1).

   > **Concepto clave — ¿Por qué el CV y no solo la desviación estándar?**
   > Imagina dos productos: el SKU-A tiene σ = 50 unidades sobre un promedio de 500 (CV = 10%), y el SKU-B tiene σ = 30 unidades sobre un promedio de 60 (CV = 50%). La desviación estándar de SKU-A es mayor en valor absoluto, pero su demanda es *proporcionalmente* mucho más estable. El CV permite esta comparación justa entre productos de diferente escala, lo que es esencial para clasificarlos y priorizar la estrategia de pronóstico.

   > **Conexión con Lección 2.1:** Recuerda que en la Lección 2.1 aprendimos que una diferencia entre promedio y mediana superior al 10–15% es señal de valores atípicos. El CV complementa ese diagnóstico: un CV alto confirma que la demanda es intrínsecamente variable, mientras que una brecha promedio-mediana alta puede indicar un evento puntual que distorsiona el promedio. Ambos análisis juntos dan una imagen completa del comportamiento de la demanda.

**Resultado esperado:** La columna F muestra el CV en porcentaje para cada SKU. Valores típicos esperados: entre 8% y 70% dependiendo del tipo de producto.

**Verificación:** Calcula manualmente el CV del primer SKU con una calculadora o en una celda aparte: `=E2/D2*100`. Debe coincidir exactamente con F2.

---

### Paso 4 — Clasificar los SKUs por Estabilidad de Demanda

**Objetivo:** Asignar una categoría de estabilidad a cada SKU usando los umbrales del CV y una fórmula condicional en Excel.

**Instrucciones:**

1. Haz clic en la celda **G2**.

2. Escribe la siguiente fórmula anidada con `SI`:
   ```excel
   =SI(F2<20,"Estable",SI(F2<=50,"Moderada","Errática"))
   ```
   Esta fórmula implementa la siguiente tabla de clasificación:

   | Rango CV          | Clasificación | Implicación para el pronóstico                          |
   |-------------------|---------------|---------------------------------------------------------|
   | CV < 20%          | **Estable**   | Fácil de pronosticar; modelos simples funcionan bien    |
   | 20% ≤ CV ≤ 50%    | **Moderada**  | Requiere ajuste estacional; modelos con tendencia       |
   | CV > 50%          | **Errática**  | Difícil de pronosticar; requiere stock de seguridad alto|

3. Copia la fórmula hacia abajo hasta G9.

4. Aplica **formato condicional con semáforo** al rango G2:G9:
   - Selecciona G2:G9.
   - Ve a **Inicio → Formato condicional → Nueva regla**.
   - Crea tres reglas de tipo *"Aplicar formato únicamente a las celdas que contengan"* → *Texto específico*:
     - Texto contiene `"Estable"` → Relleno verde claro (`#C6EFCE`), fuente verde oscuro (`#276221`).
     - Texto contiene `"Moderada"` → Relleno amarillo (`#FFEB9C`), fuente naranja oscuro (`#9C5700`).
     - Texto contiene `"Errática"` → Relleno rojo claro (`#FFC7CE`), fuente rojo oscuro (`#9C0006`).

   > **Alternativa rápida:** Si el tiempo es ajustado, aplica el formato condicional con escala de colores sobre la columna F (valores numéricos del CV) en lugar de sobre el texto. Selecciona F2:F9 → Formato condicional → Escalas de color → Verde-Amarillo-Rojo (invertido, ya que menor CV = mejor).

**Resultado esperado:** La columna G muestra las etiquetas "Estable", "Moderada" o "Errática" con colores de semáforo. La tabla debería mostrar una distribución variada entre los 8 SKUs.

**Verificación:** Revisa manualmente que los SKUs con CV < 20 tengan etiqueta "Estable" y color verde. Si algún SKU tiene CV exactamente en 20 o 50, verifica que la fórmula lo clasifique correctamente (el umbral 20 debe caer en "Moderada" según la condición `F2<=50` del segundo SI).

---

### Paso 5 — Calcular el Inventario de Seguridad para SKUs Clase A

**Objetivo:** Aplicar la fórmula de inventario de seguridad básico para los SKUs de mayor valor comercial y evaluar el impacto de la variabilidad en el dimensionamiento del stock.

**Instrucciones:**

1. La fórmula a aplicar es:

   > **IS = Z × σ × √(Lead Time)**

   Donde:
   - **Z = 1.65** → factor para nivel de servicio del 95% (distribución normal estándar).
   - **σ** = desviación estándar de la demanda mensual (columna E).
   - **Lead Time = 2 semanas = 0.5 meses** → expresado en la misma unidad temporal que σ (meses).

   > **Nota sobre la unidad del Lead Time:** Si σ está calculada sobre datos mensuales, el Lead Time debe expresarse en meses. 2 semanas = 0.5 meses. Por eso la fórmula usa `RAIZ(0.5)`.

2. Haz clic en la celda **H2** y escribe:
   ```excel
   =SI(C2="A", 1.65*E2*RAIZ(0.5), "N/A")
   ```
   Esta fórmula calcula el IS solo para SKUs Clase A; para el resto muestra "N/A".

3. Copia la fórmula hacia abajo hasta H9.

4. Para las celdas que muestren un valor numérico (SKUs Clase A), formatea con **0 decimales** (el inventario de seguridad se expresa en unidades enteras). Para las celdas con "N/A", el formato de texto es correcto.

   > **¿Por qué solo Clase A?** En la práctica logística, el cálculo de inventario de seguridad con modelos estadísticos se prioriza para los SKUs de mayor impacto en ventas o rotación (Clase A), ya que un quiebre de stock en estos productos tiene el mayor costo operativo y comercial. Los SKUs Clase B y C pueden gestionarse con reglas más simples (mínimos fijos, revisión periódica).

**Resultado esperado:** La columna H muestra valores numéricos de IS solo para los SKUs Clase A (típicamente 2 o 3 SKUs). Los demás muestran "N/A". Los valores de IS deben ser positivos y expresados en unidades enteras.

**Verificación:** Toma el primer SKU Clase A y calcula manualmente: `1.65 × σ × √0.5`. Compara con el valor en H. Recuerda que `√0.5 ≈ 0.7071`.

---

### Paso 6 — Evaluar la Razonabilidad del Inventario de Seguridad

**Objetivo:** Comparar el IS calculado con el promedio mensual de demanda para determinar si el dimensionamiento es operativamente razonable.

**Instrucciones:**

1. Haz clic en la celda **I2** y escribe:
   ```excel
   =SI(C2="A", H2/D2*100, "N/A")
   ```
   Esta fórmula calcula el IS como porcentaje del promedio mensual de demanda.

2. Copia hasta I9. Formatea las celdas numéricas con **1 decimal**.

3. Interpreta los resultados usando la siguiente guía de referencia:

   | IS / Promedio (%) | Interpretación operativa                                                     |
   |-------------------|------------------------------------------------------------------------------|
   | < 15%             | IS bajo: aceptable para demanda estable; puede ser insuficiente si CV > 30% |
   | 15% – 40%         | IS moderado: razonable para la mayoría de contextos logísticos              |
   | > 40%             | IS alto: señal de demanda muy errática; evaluar si el costo de holding es justificable |

4. Anota en una celda de comentario (fuera de la tabla, por ejemplo en la celda L2) tu interpretación del SKU Clase A con mayor IS/Promedio. Usa el formato:
   ```
   SKU [código]: IS = [valor] unidades ([%] del promedio mensual).
   Clasificación CV: [Estable/Moderada/Errática].
   Recomendación: [tu interpretación en 1 línea].
   ```

**Resultado esperado:** La columna I muestra el porcentaje IS/Promedio para los SKUs Clase A. Al menos un SKU Clase A debería mostrar un porcentaje que genere discusión (ya sea muy bajo o muy alto).

**Verificación:** Suma mentalmente: si el IS de un SKU es mayor que su promedio mensual (IS/Promedio > 100%), algo está mal — revisa si el Lead Time está en la unidad correcta o si hay un error en el rango de `DESVEST.M`.

---

### Paso 7 — Cruzar Clasificación ABC con Clasificación por CV e Identificar Combinaciones Críticas

**Objetivo:** Construir la visión integrada del riesgo operativo por SKU, combinando el valor comercial (ABC) con la estabilidad de la demanda (CV).

**Instrucciones:**

1. Haz clic en la celda **J2** y escribe la siguiente fórmula que combina ambas clasificaciones:
   ```excel
   =SI(Y(C2="A",G2="Errática"),"🔴 Máximo Riesgo",
     SI(Y(C2="A",G2="Moderada"),"🟠 Riesgo Alto",
       SI(Y(C2="A",G2="Estable"),"🟡 Riesgo Medio",
         SI(Y(C2="B",G2="Errática"),"🟠 Riesgo Alto",
           SI(Y(C2="C",G2="Errática"),"🟡 Riesgo Medio","🟢 Riesgo Bajo")))))
   ```

   > **Nota de compatibilidad:** Si tu versión de Excel no soporta emojis en fórmulas, reemplaza los emojis por texto simple: `"MAXIMO RIESGO"`, `"RIESGO ALTO"`, `"RIESGO MEDIO"`, `"RIESGO BAJO"`.

2. Copia la fórmula hasta J9.

3. Lee la siguiente tabla de interpretación operativa para cada combinación crítica:

   | Clase ABC | Clasificación CV | Riesgo Operativo   | Implicación logística                                                                 |
   |-----------|------------------|--------------------|--------------------------------------------------------------------------------------|
   | A         | Errática         | 🔴 Máximo Riesgo   | Quiebre de stock probable; IS alto obligatorio; revisar pronóstico semanalmente      |
   | A         | Moderada         | 🟠 Riesgo Alto     | Ajuste estacional necesario; IS moderado; monitoreo quincenal                        |
   | A         | Estable          | 🟡 Riesgo Medio    | Pronóstico confiable; IS estándar; revisión mensual suficiente                       |
   | B         | Errática         | 🟠 Riesgo Alto     | IS proporcional; considerar revisión de política de reabastecimiento                 |
   | C         | Errática         | 🟡 Riesgo Medio    | Evaluar si justifica mantener stock o trabajar bajo pedido                           |
   | B/C       | Estable/Moderada | 🟢 Riesgo Bajo     | Gestión estándar; modelos simples de pronóstico son suficientes                      |

4. Identifica y anota (en la celda L5) cuántos SKUs caen en la categoría "Máximo Riesgo" o "Riesgo Alto". Este número representa la cantidad de productos que requieren atención prioritaria en la estrategia de inventario.

**Resultado esperado:** La columna J muestra una etiqueta de riesgo operativo para cada SKU. La combinación de colores y etiquetas permite identificar visualmente los productos que requieren mayor atención en la planificación.

**Verificación:** Verifica manualmente la lógica para 2 SKUs: toma uno Clase A con demanda Errática (debe ser "Máximo Riesgo") y uno Clase C con demanda Estable (debe ser "Riesgo Bajo"). Si la fórmula no produce el resultado esperado, revisa que los textos en las columnas C y G coincidan exactamente (mayúsculas/minúsculas) con los usados en la fórmula.

---

## 7. Validación y Pruebas

Una vez completados los 7 pasos, realiza las siguientes verificaciones de integridad para confirmar que los resultados son correctos y consistentes:

### Verificación 1 — Consistencia estadística
En una celda libre (por ejemplo, L8), escribe:
```excel
=SI(Y(E2>0, F2>0, F2=E2/D2*100), "✓ Consistente", "✗ Revisar fórmulas")
```
Copia para los 8 SKUs. Todos deben mostrar "✓ Consistente". Si alguno muestra error, revisa las fórmulas de las columnas E y F para ese SKU.

### Verificación 2 — Rango razonable del CV
El CV de ningún SKU debería superar el 150% en un dataset de demanda comercial típica. En L10 escribe:
```excel
=CONTAR.SI(F2:F9,">150")
```
El resultado debe ser `0`. Si es mayor, revisa si hay datos en cero o negativos en el dataset original que estén distorsionando el cálculo.

### Verificación 3 — Coherencia IS vs. CV
Un SKU con clasificación "Estable" (CV < 20%) no debería tener un IS/Promedio superior al 20%. Verifica visualmente que los SKUs Clase A con etiqueta "Estable" en la columna G tengan valores razonablemente bajos en la columna I.

### Verificación 4 — Tabla resumen de distribución
En el rango L12:M15, construye una mini-tabla de conteo:
```excel
L12: Clasificación CV      M12: Cantidad de SKUs
L13: Estable               M13: =CONTAR.SI(G2:G9,"Estable")
L14: Moderada              M14: =CONTAR.SI(G2:G9,"Moderada")
L15: Errática              M15: =CONTAR.SI(G2:G9,"Errática")
```
La suma de M13+M14+M15 debe ser exactamente 8 (todos los SKUs clasificados).

### Verificación 5 — Pregunta de cierre conceptual
Responde mentalmente (o anota en L17) la siguiente pregunta antes de cerrar la práctica:

> *"Si el promedio y la mediana de un SKU difieren en más del 15% (señal aprendida en la Lección 2.1), ¿esperarías que ese SKU tenga un CV alto o bajo? ¿Por qué?"*

**Respuesta esperada:** Un SKU con gran diferencia promedio-mediana probablemente tiene valores atípicos que inflan el promedio. Estos outliers también aumentan la desviación estándar, por lo que el CV tenderá a ser más alto. Sin embargo, si el outlier es un evento único y no recurrente, el CV puede ser alto sin que la demanda sea intrínsecamente errática — lo que refuerza la necesidad de depurar los datos antes de calcular el CV (conexión con la Práctica 2).

---

## 8. Solución de Problemas

### Problema 1 — `DESVEST.M` devuelve `#¡DIV/0!` o `#¡VALOR!`

**Síntoma:** La celda en la columna E muestra un error `#¡DIV/0!` o `#¡VALOR!` en lugar de un número.

**Causa probable:**
- El rango especificado en `DESVEST.M` contiene celdas vacías, texto o ceros que Excel no puede procesar como números válidos para el cálculo de desviación estándar. Otra causa frecuente es que el rango solo contiene 1 valor (DESVEST.M requiere mínimo 2 valores para calcular la varianza muestral).

**Solución:**
1. Haz clic en la celda con error y revisa la barra de fórmulas para confirmar que el rango apunta correctamente a los datos del SKU en `Datos_Historicos`.
2. Ve a la hoja `Datos_Historicos` y selecciona el rango manualmente. Observa en la barra de estado si Excel muestra "Recuento" (indica que hay texto) o "Promedio" (indica que son números).
3. Si hay celdas vacías en el rango, verifica con el instructor si representan meses sin demanda (deben ser `0`) o meses sin datos (deben excluirse del rango).
4. Si el problema es texto en lugar de números, selecciona el rango afectado → **Datos → Texto en columnas → Finalizar** para forzar la conversión a formato numérico.
5. Una vez corregido el rango de datos, la fórmula `DESVEST.M` debería calcular correctamente.

---

### Problema 2 — La fórmula de Riesgo Operativo (columna J) devuelve siempre "Riesgo Bajo" para todos los SKUs

**Síntoma:** Todos los SKUs muestran "🟢 Riesgo Bajo" en la columna J, incluso para SKUs que visualmente son Clase A con demanda Errática.

**Causa probable:**
- Discrepancia de texto entre los valores en las columnas C (Clase ABC) y G (Clasificación CV) y los textos usados dentro de la función `SI` anidada. Excel distingue entre mayúsculas y minúsculas en comparaciones de texto dentro de `SI`, y un espacio extra o una tilde diferente puede hacer que la condición `Y(C2="A", G2="Errática")` evalúe como `FALSO`.

**Solución:**
1. Haz clic en una celda de la columna C que debería ser "A" y observa el contenido exacto en la barra de fórmulas. Confirma que es exactamente `A` sin espacios adicionales.
2. Repite para la columna G: haz clic en una celda que debería decir "Errática" y verifica el texto exacto.
3. Si hay espacios extra, aplica `=ESPACIOS(C2)` en una celda temporal para limpiarlos, o usa la función `=SUSTITUIR()` para corregir caracteres problemáticos.
4. Alternativamente, modifica la fórmula en J2 para usar `MAYUSC()` y eliminar la sensibilidad a mayúsculas/minúsculas:
   ```excel
   =SI(Y(MAYUSC(ESPACIOS(C2))="A",MAYUSC(ESPACIOS(G2))="ERRÁTICA"),
     "Máximo Riesgo", ...)
   ```
5. Si el problema persiste con la tilde en "Errática", prueba escribiendo el texto sin tilde: `"Erratica"` tanto en la fórmula de la columna G como en la de la columna J.

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza los siguientes pasos para dejar el archivo en condiciones óptimas para las sesiones de revisión:

1. **Guarda el archivo** con `Ctrl + S`. Confirma que el nombre del archivo incluye tu identificador de participante (según las instrucciones del instructor).

2. **Ajusta el ancho de columnas** para que la tabla sea legible sin desplazamiento horizontal: selecciona todas las columnas (A a J) → doble clic en el borde de cualquier encabezado de columna para autoajustar.

3. **Congela la fila de encabezados:** Ve a **Vista → Inmovilizar paneles → Inmovilizar fila superior** para que los encabezados sean visibles al hacer scroll.

4. **Verifica que no haya celdas con errores** (`#¡REF!`, `#¡VALOR!`, `#¡DIV/0!`) en el rango A1:J9. Si las hay, corrígelas antes de cerrar.

5. **No elimines** las hojas `Práctica_9` ni `ABC_Clasificación`. La hoja `Práctica_10` que construiste hoy será referenciada en la sesión de revisión del módulo.

6. **Cierra Excel** normalmente. Si el sistema pregunta si deseas guardar cambios, confirma con "Guardar".

> **Entrega:** El instructor indicará si el archivo debe subirse a la plataforma del curso o compartirse por otro medio. Asegúrate de que el archivo contiene la hoja `Práctica_10` completa con las columnas A a J y la mini-tabla de verificación en el rango L8:M15.

---

## 10. Resumen

### Lo que construiste en esta práctica

En 11 minutos construiste un sistema completo de análisis de variabilidad de la demanda que integra:

| Componente                     | Herramienta / Fórmula          | Propósito logístico                                              |
|--------------------------------|--------------------------------|------------------------------------------------------------------|
| Desviación estándar muestral   | `DESVEST.M`                    | Medir la dispersión absoluta de la demanda por SKU              |
| Coeficiente de variación       | `=E/D*100`                     | Comparar variabilidad entre SKUs de diferente escala            |
| Clasificación por estabilidad  | `SI` anidado + formato semáforo| Segmentar productos por dificultad de pronóstico                |
| Inventario de seguridad básico | `=1.65*σ*RAIZ(0.5)`            | Dimensionar el colchón de stock para nivel de servicio del 95%  |
| Riesgo operativo combinado     | Cruce ABC × CV                 | Identificar SKUs que requieren atención prioritaria             |

### Conexión con el ciclo de forecasting logístico

Esta práctica cierra el ciclo estadístico del módulo 2 y conecta directamente con los módulos posteriores:

- **La desviación estándar** es el insumo directo para calcular el inventario de seguridad y para entender por qué el error de pronóstico (MAE, MAPE) varía entre SKUs.
- **El CV** es el criterio que justifica la selección del modelo de pronóstico: un SKU Estable puede pronosticarse con promedio móvil simple; un SKU Errático requiere modelos más sofisticados o incluso estrategias de gestión de demanda.
- **La clasificación ABC × CV** es la matriz de priorización que debe guiar las decisiones de inversión en precisión del pronóstico: no tiene sentido gastar tiempo en afinar el modelo de un SKU Clase C con demanda Estable, pero sí es crítico hacerlo para un SKU Clase A Errático.
- **El inventario de seguridad** calculado hoy es la traducción directa de la variabilidad estadística en una decisión operativa concreta: cuántas unidades adicionales mantener en bodega para proteger el nivel de servicio frente a la incertidumbre de la demanda.

### Pregunta de reflexión final

> *"Un SKU Clase A tiene CV = 65% (demanda Errática). El IS calculado representa el 48% de su demanda mensual promedio. ¿Es correcto simplemente aumentar el stock de seguridad, o hay otras acciones logísticas que deberían evaluarse primero?"*

**Pista:** Considera si el problema es la variabilidad intrínseca del producto o si hay eventos extraordinarios no gestionados (promociones no planificadas, pedidos especiales) que podrían controlarse con mejor coordinación entre ventas y logística. A veces, reducir la variabilidad en la fuente es más eficiente que aumentar el stock de seguridad.

---

### Recursos de Referencia

- [Microsoft Support — Función DESVEST.M en Excel](https://support.microsoft.com/es-es/office/desvest-m-funci%C3%B3n-desvest-m-d1688e10-2a17-4280-8b55-5aa04e3c0c58)
- [Microsoft Support — Función RAIZ en Excel](https://support.microsoft.com/es-es/office/raiz-funci%C3%B3n-raiz-654975c2-05c4-4831-9a24-2c65e4040fdf)
- [Silver, E. A., Pyke, D. F. & Thomas, D. J. (2017). *Inventory and Production Management in Supply Chains* (4ª ed.) — CRC Press](https://www.crcpress.com/Inventory-and-Production-Management-in-Supply-Chains/Silver-Pyke-Thomas/p/book/9781466558083)
- [Hyndman, R. J. & Athanasopoulos, G. — *Forecasting: Principles and Practice*, Cap. 3: Time series decomposition](https://otexts.com/fpp3/decomposition.html)
- [Khan Academy en Español — Desviación estándar y varianza](https://es.khanacademy.org/math/statistics-probability/summarizing-quantitative-data/variance-standard-deviation-population/a/calculating-standard-deviation-step-by-step)

---

---

---LAB_START---
LAB_ID: 02-00-03
---MARKDOWN---
# Práctica 11 — Clasificación de productos según estabilidad de la demanda

## 1. Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 11 minutos |
| **Complejidad** | Fácil |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 2 — Estadística descriptiva para pronóstico logístico |
| **Herramienta principal** | Microsoft Excel 2016 o superior |
| **Archivo de entrada** | `Dataset_Base_8SKU_24meses.xlsx` (proporcionado por el instructor) |
| **Archivo de salida** | `Lab11_Clasificacion_CV.xlsx` |

---

## 2. Descripción General

En esta práctica aplicarás las medidas estadísticas vistas en el Módulo 2 —promedio, mediana, máximo, mínimo y desviación estándar— para calcular el **coeficiente de variación (CV)** de un conjunto de 8 a 10 productos con 12 meses de historial de demanda. A partir del CV, clasificarás cada producto en una de tres categorías de estabilidad (estable, moderado, inestable) utilizando umbrales predefinidos y la función `SI` anidada de Excel. La práctica cierra con una reflexión operativa sobre qué productos son más predecibles para fines de pronóstico y planificación de inventarios.

> **Conexión logística:** La clasificación por CV es el primer filtro para decidir qué modelo de pronóstico aplicar a cada SKU y qué nivel de stock de seguridad se requiere. Un producto con CV alto necesita más cobertura de inventario y un modelo de pronóstico más robusto que uno con demanda estable.

---

## 3. Objetivos de Aprendizaje

- [ ] Calcular el coeficiente de variación (CV) para cada SKU utilizando las funciones `PROMEDIO`, `MEDIANA`, `MAX`, `MIN` y `DESVEST.M` en Excel.
- [ ] Clasificar cada producto en categorías de estabilidad de demanda (estable, moderado, inestable) aplicando umbrales de CV con la función `SI` anidada.
- [ ] Aplicar formato condicional para visualizar de forma inmediata las categorías de estabilidad sobre la tabla de resultados.
- [ ] Interpretar los resultados de la clasificación en términos de predecibilidad y su impacto en decisiones operativas de inventario y picking.

---

## 4. Prerrequisitos

### Conocimiento previo
- Comprensión conceptual de promedio, mediana y valores extremos (Lecciones 2.1 y 2.2 del Módulo 2).
- Noción básica de qué es el coeficiente de variación: mide cuánta variabilidad tiene la demanda **relativa a su promedio**. Un CV bajo indica demanda predecible; un CV alto indica demanda errática.
- Familiaridad con la sintaxis de fórmulas básicas en Excel (`=PROMEDIO()`, `=DESVEST.M()`, `=SI()`).

### Acceso y archivos
- Archivo `Dataset_Base_8SKU_24meses.xlsx` abierto en Excel (proporcionado por el instructor antes de la sesión).
- Microsoft Excel 2016 o superior instalado y funcional.
- No se requiere conexión a internet para esta práctica.

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Espacio en disco | 500 MB libres | 2 GB libres |

### Software requerido

| Software | Versión mínima | Notas |
|---|---|---|
| Microsoft Excel | 2016 | `DESVEST.M` y `SI` anidado disponibles desde Excel 2010; `PRONOSTICO.LINEAL` desde 2016 |
| Archivo de práctica | — | `Dataset_Base_8SKU_24meses.xlsx` |

### Preparación del entorno (antes de iniciar el cronómetro)

1. Abre el archivo `Dataset_Base_8SKU_24meses.xlsx`.
2. Localiza la hoja **`Demanda_Histórica`**. Debe contener:
   - Columna A: Código de SKU (SKU-01 a SKU-08 o SKU-10).
   - Columna B: Nombre del producto.
   - Columnas C a N: Demanda mensual de los meses 1 a 12 (encabezados: `Mes_01`, `Mes_02`, … `Mes_12`).
3. Crea una hoja nueva llamada **`Análisis_CV`** (clic derecho sobre la pestaña de hoja → *Insertar* → *Hoja de cálculo*).
4. Guarda el archivo con el nombre `Lab11_Clasificacion_CV.xlsx` en tu carpeta de trabajo.

> **Nota:** Si el archivo contiene 24 meses de historial, trabaja únicamente con los primeros 12 meses (columnas C a N) para esta práctica, tal como indica el alcance.

---

## 6. Procedimiento Paso a Paso

---

### Paso 1 — Estructurar la tabla de análisis en la hoja `Análisis_CV`

**Objetivo:** Crear la estructura base donde se calcularán todos los indicadores estadísticos y la clasificación final.

#### Instrucciones

1. En la hoja **`Análisis_CV`**, escribe los siguientes encabezados en la **fila 1**, comenzando en la celda `A1`:

   | Celda | Encabezado |
   |---|---|
   | A1 | `SKU` |
   | B1 | `Producto` |
   | C1 | `Promedio` |
   | D1 | `Mediana` |
   | E1 | `Máximo` |
   | F1 | `Mínimo` |
   | G1 | `Desv_Std` |
   | H1 | `CV (%)` |
   | I1 | `Clasificación` |

2. Aplica **negrita** a toda la fila 1 (selecciona A1:I1 → `Ctrl + N`).

3. Copia los valores de las columnas A y B de la hoja `Demanda_Histórica` (SKU y Nombre del producto) a las columnas A y B de `Análisis_CV`, filas 2 en adelante. Puedes usar **Pegado especial → Solo valores** para evitar arrastrar formatos innecesarios.

   ```
   En Demanda_Histórica: selecciona A2:B9 → Ctrl+C
   En Análisis_CV: clic en A2 → Ctrl+Shift+V → Valores → Aceptar
   ```
   *(En Excel para Windows: clic derecho → Pegado especial → Valores → Aceptar)*

#### Salida esperada
La hoja `Análisis_CV` debe tener 9 filas visibles: 1 de encabezado + 8 filas de datos (una por SKU), con las columnas A y B completadas y las columnas C a I vacías, listas para recibir fórmulas.

#### Verificación
✔ La fila 1 contiene exactamente los 9 encabezados indicados, en negrita.
✔ Las columnas A y B tienen los mismos códigos y nombres que la hoja de origen.
✔ No hay filas en blanco entre los datos.

---

### Paso 2 — Calcular Promedio, Mediana, Máximo y Mínimo

**Objetivo:** Obtener las cuatro medidas descriptivas básicas para los 12 meses de demanda de cada SKU, referenciando los datos de la hoja original.

#### Instrucciones

1. Haz clic en la celda **`C2`** de la hoja `Análisis_CV`.

2. Escribe la siguiente fórmula para calcular el **promedio** del primer SKU:

   ```excel
   =PROMEDIO(Demanda_Histórica!C2:N2)
   ```

   > **Explicación:** `Demanda_Histórica!C2:N2` referencia las 12 columnas de demanda mensual (Mes_01 a Mes_12) del primer SKU en la hoja de origen. Ajusta el rango si tus datos están en columnas distintas.

3. En la celda **`D2`**, escribe la fórmula para la **mediana**:

   ```excel
   =MEDIANA(Demanda_Histórica!C2:N2)
   ```

4. En la celda **`E2`**, escribe la fórmula para el **máximo**:

   ```excel
   =MAX(Demanda_Histórica!C2:N2)
   ```

5. En la celda **`F2`**, escribe la fórmula para el **mínimo**:

   ```excel
   =MIN(Demanda_Histórica!C2:N2)
   ```

6. **Copia las cuatro fórmulas hacia abajo** para todos los SKUs: selecciona el rango `C2:F2`, luego arrastra el controlador de relleno (pequeño cuadrado en la esquina inferior derecha de la selección) hasta la última fila de datos (ej. `C9:F9` si tienes 8 SKUs).

   Alternativa con teclado: selecciona `C2:F2` → `Ctrl+C` → selecciona `C3:F9` → `Ctrl+V`.

7. **Redondea los resultados** para facilitar la lectura: selecciona `C2:F9` → clic en el botón *Disminuir decimales* en la barra de herramientas hasta mostrar **2 decimales**.

#### Salida esperada
Las columnas C a F deben mostrar valores numéricos para todos los SKUs. Ejemplo orientativo (los valores dependen del dataset del instructor):

| SKU | Producto | Promedio | Mediana | Máximo | Mínimo |
|---|---|---|---|---|---|
| SKU-01 | Producto A | 245.33 | 238.50 | 312.00 | 198.00 |
| SKU-02 | Producto B | 87.17 | 85.00 | 180.00 | 42.00 |
| … | … | … | … | … | … |

#### Verificación
✔ Ninguna celda en C2:F9 muestra `#¡REF!`, `#¡VALOR!` o está vacía.
✔ Para cada SKU, el valor de Máximo ≥ Promedio ≥ Mediana (no siempre, pero Máximo siempre debe ser ≥ Promedio) y Mínimo ≤ Promedio.
✔ Compara manualmente uno de los SKUs: suma las 12 celdas de demanda en la hoja origen y divide entre 12; el resultado debe coincidir con el valor de la columna C.

> **Reflexión rápida (30 segundos):** Para los SKUs donde el Promedio y la Mediana difieren más del 15%, anota mentalmente que puede haber valores atípicos en ese producto. Esto es exactamente lo que aprendiste en la Lección 2.1: una diferencia significativa entre ambas medidas es señal de outliers.

---

### Paso 3 — Calcular la Desviación Estándar

**Objetivo:** Obtener la dispersión absoluta de la demanda mensual para cada SKU usando `DESVEST.M` (desviación estándar muestral).

#### Instrucciones

1. Haz clic en la celda **`G2`**.

2. Escribe la siguiente fórmula:

   ```excel
   =DESVEST.M(Demanda_Histórica!C2:N2)
   ```

   > **¿Por qué `DESVEST.M` y no `DESVEST.P`?** Usamos la versión **muestral** (`DESVEST.M`) porque los 12 meses son una *muestra* del comportamiento del producto, no el universo completo de su historia. En logística, casi siempre trabajamos con muestras históricas, por lo que `DESVEST.M` es la elección correcta.

3. Copia la fórmula hacia abajo hasta la última fila de datos (`G3:G9`).

4. Mantén **2 decimales** en la columna G (consistente con las columnas anteriores).

#### Salida esperada
La columna G muestra la desviación estándar en unidades de demanda para cada SKU. Valores más altos indican mayor variabilidad absoluta.

#### Verificación
✔ Todos los valores en G2:G9 son positivos (la desviación estándar nunca puede ser negativa ni cero, a menos que todos los meses tengan exactamente la misma demanda).
✔ Si algún SKU tiene todos los meses con el mismo valor de demanda, `DESVEST.M` devolverá `0` — esto es correcto y ese producto tendrá CV = 0%.

---

### Paso 4 — Calcular el Coeficiente de Variación (CV)

**Objetivo:** Expresar la variabilidad de la demanda en términos relativos al promedio, obteniendo el CV como porcentaje.

#### Instrucciones

1. Haz clic en la celda **`H2`**.

2. Escribe la siguiente fórmula:

   ```excel
   =G2/C2
   ```

   > **Fórmula del CV:** CV = Desviación Estándar / Promedio. Al expresarlo como porcentaje, permite comparar la variabilidad entre productos con niveles de demanda muy distintos (por ejemplo, un producto que vende 50 unidades/mes vs. uno que vende 5,000).

3. Copia la fórmula hacia abajo hasta `H9`.

4. **Formatea la columna H como porcentaje con 1 decimal:**
   - Selecciona `H2:H9`.
   - En la pestaña *Inicio*, haz clic en el botón **%** del grupo *Número*.
   - Luego haz clic una vez en *Aumentar decimales* para mostrar 1 decimal (ej. 23.4%).

   Alternativamente: clic derecho → *Formato de celdas* → *Porcentaje* → 1 decimal → Aceptar.

#### Salida esperada
La columna H muestra el CV en porcentaje para cada SKU. Ejemplo orientativo:

| SKU | CV (%) |
|---|---|
| SKU-01 | 12.3% |
| SKU-02 | 38.7% |
| SKU-03 | 65.2% |
| … | … |

#### Verificación
✔ Los valores de CV están expresados en porcentaje (no como decimales como 0.123).
✔ El CV es siempre ≥ 0%. Si obtienes valores negativos, revisa que el Promedio en columna C sea positivo y que la fórmula apunte a las celdas correctas.
✔ Verifica manualmente un SKU: toma la Desviación Estándar (columna G) y divídela entre el Promedio (columna C) con una calculadora; el resultado debe coincidir con el CV mostrado.

---

### Paso 5 — Clasificar productos con la función SI anidada

**Objetivo:** Asignar automáticamente una etiqueta de categoría a cada SKU según los siguientes umbrales de CV:

| Rango de CV | Categoría | Interpretación logística |
|---|---|---|
| CV < 20% | `Estable` | Demanda predecible; modelos simples funcionan bien; stock de seguridad bajo |
| 20% ≤ CV ≤ 50% | `Moderado` | Variabilidad manejable; requiere revisión periódica del pronóstico |
| CV > 50% | `Inestable` | Demanda errática; modelos avanzados necesarios; mayor stock de seguridad |

#### Instrucciones

1. Haz clic en la celda **`I2`**.

2. Escribe la siguiente fórmula con `SI` anidado:

   ```excel
   =SI(H2<0.2,"Estable",SI(H2<=0.5,"Moderado","Inestable"))
   ```

   > **Explicación de la fórmula:**
   > - La función `SI` evalúa primero si el CV es menor al 20% (0.2 en formato decimal).
   > - Si es verdadero → devuelve `"Estable"`.
   > - Si es falso → evalúa el segundo `SI`: ¿es el CV ≤ 50% (0.5)?
   >   - Si es verdadero → devuelve `"Moderado"`.
   >   - Si es falso → devuelve `"Inestable"`.
   >
   > **Importante:** Aunque la columna H muestra el CV en formato porcentaje (ej. 23.4%), el valor interno de la celda es un decimal (0.234). Por eso los umbrales en la fórmula se escriben como `0.2` y `0.5`, no como `20` y `50`.

3. Copia la fórmula hacia abajo hasta `I9`.

4. **Verifica que las etiquetas sean coherentes** con los valores de CV que calculaste en el paso anterior.

#### Salida esperada
La columna I muestra una de las tres etiquetas para cada SKU:

| SKU | CV (%) | Clasificación |
|---|---|---|
| SKU-01 | 12.3% | Estable |
| SKU-02 | 38.7% | Moderado |
| SKU-03 | 65.2% | Inestable |
| … | … | … |

#### Verificación
✔ Solo aparecen tres valores posibles en la columna I: `Estable`, `Moderado` o `Inestable`.
✔ Ninguna celda muestra `FALSO`, `VERDADERO` o un error.
✔ Comprueba manualmente los casos límite: un SKU con CV exactamente igual a 20.0% debe clasificarse como `Moderado` (la condición `H2<0.2` es falsa, y `H2<=0.5` es verdadera).

---

### Paso 6 — Aplicar Formato Condicional para visualizar las categorías

**Objetivo:** Resaltar visualmente cada categoría con un color diferente para facilitar la lectura rápida de la tabla.

#### Instrucciones

1. Selecciona el rango **`I2:I9`** (columna de Clasificación).

2. Ve a la pestaña **Inicio** → grupo **Estilos** → clic en **Formato condicional** → **Nueva regla**.

3. **Regla 1 — Estable (verde):**
   - Selecciona *"Aplicar formato únicamente a las celdas que contengan"*.
   - Configura: *Valor de celda* → *igual a* → escribe `Estable`.
   - Clic en **Formato** → pestaña **Relleno** → selecciona **verde claro** (ej. color hex `#C6EFCE`).
   - Clic en **Aceptar** → **Aceptar**.

4. **Regla 2 — Moderado (amarillo):**
   - Repite el proceso: *Formato condicional* → *Nueva regla*.
   - Configura: *igual a* → `Moderado`.
   - Formato → Relleno → **amarillo claro** (ej. `#FFEB9C`).
   - Aceptar → Aceptar.

5. **Regla 3 — Inestable (rojo):**
   - Repite: *Formato condicional* → *Nueva regla*.
   - Configura: *igual a* → `Inestable`.
   - Formato → Relleno → **rojo claro** (ej. `#FFC7CE`).
   - Aceptar → Aceptar.

6. **Extiende el formato condicional a toda la fila** (opcional, si el tiempo lo permite):
   - Selecciona el rango completo `A2:I9`.
   - Repite las tres reglas anteriores, pero esta vez usando una **fórmula** como criterio:
     - Para verde: `=$I2="Estable"`
     - Para amarillo: `=$I2="Moderado"`
     - Para rojo: `=$I2="Inestable"`
   - Esto coloreará toda la fila según la clasificación del SKU, haciendo la tabla más legible.

#### Salida esperada
La columna I (o toda la tabla si aplicaste el paso 6 opcional) muestra:
- 🟢 **Verde** para productos `Estable`
- 🟡 **Amarillo** para productos `Moderado`
- 🔴 **Rojo** para productos `Inestable`

#### Verificación
✔ Al cambiar manualmente el valor de CV de un SKU (temporalmente), la celda cambia de color automáticamente. Deshaz el cambio con `Ctrl+Z` después de verificar.
✔ Los colores son consistentes con los valores de CV: ningún producto con CV > 50% aparece en verde.

---

### Paso 7 — Crear la tabla resumen de clasificación

**Objetivo:** Consolidar el conteo de productos por categoría y agregar una reflexión operativa sobre la predecibilidad del portafolio.

#### Instrucciones

1. Selecciona una celda en blanco debajo de tu tabla principal, dejando al menos **dos filas de separación** (ej. celda `A12`).

2. Crea la siguiente tabla resumen escribiendo los encabezados y fórmulas:

   | Celda | Contenido |
   |---|---|
   | A12 | `RESUMEN DE CLASIFICACIÓN` (en negrita) |
   | A13 | `Categoría` |
   | B13 | `Cantidad de SKUs` |
   | C13 | `% del portafolio` |
   | A14 | `Estable` |
   | A15 | `Moderado` |
   | A16 | `Inestable` |
   | A17 | `TOTAL` |

3. En la celda **`B14`**, escribe la fórmula para contar los SKUs estables:

   ```excel
   =CONTAR.SI($I$2:$I$9,"Estable")
   ```

4. En **`B15`**:

   ```excel
   =CONTAR.SI($I$2:$I$9,"Moderado")
   ```

5. En **`B16`**:

   ```excel
   =CONTAR.SI($I$2:$I$9,"Inestable")
   ```

6. En **`B17`** (total):

   ```excel
   =SUMA(B14:B16)
   ```

7. En **`C14`**, calcula el porcentaje del portafolio:

   ```excel
   =B14/$B$17
   ```

   Formatea como porcentaje con 0 decimales. Copia hacia abajo hasta `C16`.

8. En **`C17`**:

   ```excel
   =SUMA(C14:C16)
   ```

   Debe mostrar `100%`.

9. **Debajo de la tabla resumen** (ej. fila 19), agrega una celda de texto con la reflexión operativa. En `A19` escribe:

   ```
   Reflexión: Los productos clasificados como "Estable" (CV < 20%) son los más
   predecibles para el pronóstico y requieren menor stock de seguridad.
   Los productos "Inestable" (CV > 50%) demandan modelos de pronóstico más
   sofisticados y mayor cobertura de inventario para mantener el nivel de servicio.
   ```

   *(Puedes escribir esto directamente en la celda y usar `Alt+Enter` para saltos de línea dentro de la misma celda, o distribuirlo en varias filas.)*

#### Salida esperada
Una tabla resumen compacta similar a:

| Categoría | Cantidad de SKUs | % del portafolio |
|---|---|---|
| Estable | 3 | 38% |
| Moderado | 3 | 38% |
| Inestable | 2 | 25% |
| **TOTAL** | **8** | **100%** |

*(Los valores exactos dependen del dataset proporcionado por el instructor.)*

#### Verificación
✔ La columna B suma exactamente el número total de SKUs del dataset (8 o 10).
✔ La columna C suma 100%.
✔ Los conteos de `CONTAR.SI` coinciden con el recuento manual de la columna I.

---

### Paso 8 — Guardar y nombrar el archivo final

**Objetivo:** Preservar el trabajo realizado con un nombre de archivo identificable para las prácticas posteriores.

#### Instrucciones

1. Presiona `Ctrl + S` para guardar.

2. Si Excel solicita confirmación del formato, selecciona **Mantener formato Excel (.xlsx)**.

3. Verifica que el nombre del archivo en la barra de título sea **`Lab11_Clasificacion_CV.xlsx`**.

4. Cierra la hoja `Demanda_Histórica` si está en primer plano y deja activa la hoja `Análisis_CV` como vista principal del archivo.

#### Verificación
✔ El archivo se guarda sin errores.
✔ Al reabrir el archivo, las fórmulas de la hoja `Análisis_CV` siguen calculando correctamente (no se han convertido en valores estáticos por error).

---

## 7. Validación y Pruebas

Al finalizar los 8 pasos, realiza las siguientes comprobaciones para confirmar que el trabajo está completo y correcto:

| # | Verificación | Resultado esperado |
|---|---|---|
| V1 | La hoja `Análisis_CV` contiene exactamente 9 columnas (A a I) con encabezados correctos | ✔ Sí |
| V2 | Las columnas C a G muestran valores numéricos para todos los SKUs (sin celdas vacías ni errores) | ✔ Sí |
| V3 | La columna H muestra el CV en formato porcentaje con 1 decimal | ✔ Sí |
| V4 | La columna I muestra únicamente los valores: `Estable`, `Moderado` o `Inestable` | ✔ Sí |
| V5 | El formato condicional colorea correctamente las tres categorías | ✔ Sí |
| V6 | La tabla resumen (filas 12-17) muestra los conteos y porcentajes correctos, sumando 100% | ✔ Sí |
| V7 | Para al menos un SKU, la diferencia entre Promedio y Mediana supera el 15% del Promedio | ✔ Identificado (o no aplica si los datos son uniformes) |
| V8 | El archivo está guardado como `Lab11_Clasificacion_CV.xlsx` | ✔ Sí |

### Prueba de integridad de fórmulas

Selecciona la celda `C2` y verifica en la barra de fórmulas que la referencia apunta a la hoja correcta:

```excel
=PROMEDIO(Demanda_Histórica!C2:N2)
```

Si la referencia aparece como `=PROMEDIO(C2:N2)` sin el nombre de la hoja, significa que la fórmula fue escrita directamente en la hoja de origen en lugar de referenciarla. Corrígelo reescribiendo la fórmula con la referencia explícita a la hoja.

---

## 8. Resolución de Problemas

### Problema 1 — La fórmula `SI` anidada devuelve `FALSO` o un error inesperado

**Síntoma:** La columna I muestra `FALSO`, `VERDADERO`, `#¡VALOR!` o un texto incorrecto en lugar de `Estable`, `Moderado` o `Inestable`.

**Causa probable:** La columna H contiene el CV como porcentaje visualmente (ej. 23.4%), pero si la fórmula en H2 fue escrita como `=G2/C2*100` en lugar de `=G2/C2`, el valor interno de la celda es `23.4` (no `0.234`). En ese caso, los umbrales `0.2` y `0.5` de la fórmula `SI` nunca se cumplen.

**Solución:**
1. Haz clic en `H2` y revisa la barra de fórmulas. La fórmula debe ser `=G2/C2` (sin multiplicar por 100).
2. Si la fórmula tiene `*100`, elimínalo: escribe `=G2/C2` y presiona Enter.
3. Formatea la celda como porcentaje con el botón **%** de la barra de herramientas (esto solo cambia la *visualización*, no el valor).
4. Copia la fórmula corregida hacia abajo y verifica que la columna I se actualice correctamente.

**Alternativa:** Si prefieres mantener la fórmula con `*100` (valor interno = 23.4), ajusta los umbrales del `SI` anidado:

```excel
=SI(H2<20,"Estable",SI(H2<=50,"Moderado","Inestable"))
```

En este caso, **no** formatees la columna H como porcentaje (déjala como número).

---

### Problema 2 — Las referencias a la hoja `Demanda_Histórica` devuelven `#¡REF!`

**Síntoma:** Las celdas de las columnas C a G muestran el error `#¡REF!` después de copiar las fórmulas hacia abajo.

**Causa probable:** Al copiar las fórmulas, Excel ajustó las referencias de filas correctamente, pero el nombre de la hoja de origen fue escrito con un error tipográfico (ej. `Demanda_Historica` sin tilde, o `DemandaHistórica` sin guión bajo), o la hoja fue renombrada después de escribir las fórmulas.

**Solución:**
1. Haz clic en la celda con error (ej. `C5`) y revisa la barra de fórmulas.
2. Identifica el nombre de la hoja en la fórmula: `=PROMEDIO(Demanda_Histórica!C5:N5)`.
3. Verifica que el nombre coincida exactamente con la pestaña de la hoja de origen (respeta mayúsculas, tildes y espacios).
4. Si hay discrepancia, corrige el nombre en la fórmula. Para evitar errores tipográficos en el futuro, escribe la fórmula haciendo clic directamente en la hoja de origen:
   - En `C2`, escribe `=PROMEDIO(`.
   - Navega a la hoja `Demanda_Histórica` haciendo clic en su pestaña.
   - Selecciona el rango `C2:N2`.
   - Presiona Enter. Excel completará automáticamente la referencia con el nombre correcto.

---

## 9. Limpieza del Entorno

Al finalizar la práctica:

1. **Guarda el archivo final** con `Ctrl + S`. Confirma que el nombre es `Lab11_Clasificacion_CV.xlsx`.

2. **Cierra el archivo de datos original** `Dataset_Base_8SKU_24meses.xlsx` sin guardar cambios (a menos que el instructor indique lo contrario). Esto evita sobrescribir accidentalmente el dataset base que se usará en prácticas posteriores.

3. **Mantén abierto** `Lab11_Clasificacion_CV.xlsx` si la siguiente práctica lo requiere como insumo (consulta la secuencia de prácticas con el instructor).

4. **No elimines** la hoja `Análisis_CV` ni los datos calculados: la clasificación de productos por CV será referenciada en prácticas posteriores para seleccionar modelos de pronóstico por categoría de SKU.

---

## 10. Resumen

### Lo que aprendiste en esta práctica

En 11 minutos aplicaste un flujo completo de clasificación estadística de demanda:

1. **Calculaste medidas descriptivas** (promedio, mediana, máximo, mínimo y desviación estándar) para 8 a 10 SKUs con 12 meses de historial, usando las funciones nativas de Excel.

2. **Obtuviste el Coeficiente de Variación (CV)** como indicador relativo de variabilidad, que permite comparar la estabilidad de productos con niveles de demanda muy distintos.

3. **Clasificaste automáticamente** cada SKU en tres categorías operativas (Estable, Moderado, Inestable) usando umbrales predefinidos y la función `SI` anidada.

4. **Visualizaste la clasificación** con formato condicional y la consolidaste en una tabla resumen con conteos y porcentajes del portafolio.

### Conexión con el ciclo de forecasting logístico

La clasificación por CV no es un ejercicio estadístico abstracto: es la base para decisiones operativas concretas:

| Categoría | Impacto en pronóstico | Impacto en inventario | Impacto en picking |
|---|---|---|---|
| **Estable** (CV < 20%) | Modelos simples (promedio móvil, suavización exponencial básica) son suficientes | Stock de seguridad bajo; reabastecimiento periódico estable | Planificación de picking predecible; menor variabilidad en carga de trabajo |
| **Moderado** (20-50%) | Requiere modelos con ajuste estacional o tendencia | Stock de seguridad intermedio; revisar puntos de reorden | Cierta variabilidad en picking; planificar con buffer de capacidad |
| **Inestable** (CV > 50%) | Modelos avanzados o revisión manual del historial | Stock de seguridad alto o política de pedido bajo demanda | Alta variabilidad; difícil planificar picking; considerar estrategias de postponement |

### Próximos pasos

- La clasificación obtenida en esta práctica se utilizará en prácticas posteriores para **seleccionar el modelo de pronóstico más adecuado** por categoría de SKU.
- Los productos identificados con diferencia significativa entre promedio y mediana (> 15%) serán candidatos a **depuración de outliers** en la Práctica siguiente.
- El CV también es un insumo directo para el **cálculo del stock de seguridad** y la definición del nivel de servicio objetivo por SKU.

### Recursos de referencia

- [Microsoft Support — Función DESVEST.M](https://support.microsoft.com/es-es/office/desvest-m-funci%C3%B3n-desvest-m-d1688e10-2a8a-4e0d-8e8e-f48e4c7e5f6a)
- [Microsoft Support — Función SI (anidada)](https://support.microsoft.com/es-es/office/si-funci%C3%B3n-si-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)
- [Microsoft Support — Formato condicional en Excel](https://support.microsoft.com/es-es/office/usar-el-formato-condicional-para-resaltar-informaci%C3%B3n-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
- Chase, C. W. (2013). *Demand-Driven Forecasting: A Structured Approach to Forecasting*. Wiley Finance Series.
- Hyndman, R. J. & Athanasopoulos, G. — *Forecasting: Principles and Practice* (3ª ed.) — [https://otexts.com/fpp3/](https://otexts.com/fpp3/)

---

> **Pregunta de cierre para discusión en clase:** ¿Qué harías operativamente si el 70% de tus SKUs caen en la categoría "Inestable"? ¿Es un problema del producto, del mercado o de la calidad de los datos históricos?

---
LAB_END---

---

# Análisis comparativo de productos con alta y baja variabilidad

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 11 minutos |
| **Complejidad** | Fácil |
| **Nivel Bloom** | Aplicar (Apply) |
| **Práctica número** | 12 |
| **Módulo** | 2 — Estadística descriptiva para el análisis de demanda |
| **Secuencia** | Requiere haber completado la Práctica 11 |

---

## 2. Descripción General

En esta práctica compararás directamente el comportamiento histórico de demanda de dos productos seleccionados de tu dataset: el producto con el **coeficiente de variación (CV) más bajo** (demanda más estable) y el producto con el **CV más alto** (demanda más inestable), identificados en la Práctica 11. Construirás gráficos de línea, calcularás un conjunto completo de métricas estadísticas y crearás una tabla comparativa lado a lado. Como cierre, aplicarás un **pronóstico naive** (igual al promedio histórico) a ambos productos y calcularás el error frente a los últimos 3 meses reales, obteniendo tu primer acercamiento práctico al concepto de error de pronóstico. Utilizarás IA Generativa para interpretar los factores que podrían explicar la alta variabilidad del producto inestable.

---

## 3. Objetivos de Aprendizaje

- [ ] Seleccionar correctamente el producto de menor y mayor CV a partir de los resultados de la Práctica 11 y justificar la selección.
- [ ] Construir gráficos de línea en Excel que permitan visualizar visualmente las diferencias en el patrón de demanda entre ambos productos.
- [ ] Calcular y comparar lado a lado las métricas: promedio, mediana, desviación estándar, CV, rango y rango intercuartílico (IQR) para ambos productos.
- [ ] Aplicar el pronóstico naive (promedio histórico) y calcular manualmente el error absoluto de los últimos 3 meses como introducción al concepto de error de pronóstico.
- [ ] Formular una pregunta a una herramienta de IA Generativa sobre los factores que explican la alta variabilidad del producto inestable y documentar la respuesta obtenida.

---

## 4. Prerrequisitos

### Conocimiento previo requerido
- Haber completado la **Práctica 11** y contar con la tabla de clasificación de productos con sus respectivos valores de CV calculados.
- Comprensión de las medidas de tendencia central: promedio (`=PROMEDIO()`) y mediana (`=MEDIANA()`), según lo estudiado en la Lección 2.1.
- Conocimiento básico de creación de gráficos de línea en Excel (insertar gráfico desde un rango de datos).
- Familiaridad básica con las funciones `=DESVEST.M()`, `=MAX()`, `=MIN()`, `=CUARTIL.EXC()` en Excel.

### Acceso y recursos requeridos
- Archivo Excel de trabajo de la Práctica 11 con el dataset de 8 SKUs (24 meses de historial) y la tabla de clasificación por CV ya completada.
- Microsoft Excel 2016 o superior (Microsoft 365 recomendado).
- Acceso activo a una herramienta de IA Generativa: Microsoft Copilot, ChatGPT, Gemini o equivalente.
- Conexión a internet estable (mínimo 10 Mbps) para acceso a la IA Generativa.

> ⚠️ **Nota para el instructor:** Si algún participante no completó la Práctica 11, debe recibir el archivo de *checkpoint* correspondiente antes de iniciar esta práctica. La selección incorrecta de los productos de mayor y menor CV invalidará todos los cálculos posteriores.

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 2 GB | 2 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Conexión a internet | 10 Mbps | 20 Mbps |

### Software requerido

| Software | Versión | Observación |
|---|---|---|
| Microsoft Excel | 2016 o superior | Microsoft 365 recomendado |
| Navegador web | Chrome / Edge (versión 2024) | Para acceso a IA Generativa |
| Herramienta de IA Generativa | Copilot / ChatGPT / Gemini | Acceso gratuito o licencia estándar |

### Preparación del entorno (antes de iniciar)

1. Abre el archivo Excel de la Práctica 11. Verifica que la hoja con la tabla de clasificación de productos por CV esté visible y que los valores de CV estén calculados para los 8 SKUs.
2. Abre tu navegador web y accede a tu herramienta de IA Generativa preferida. Verifica que puedas escribir y recibir respuestas antes de comenzar los pasos cronometrados.
3. Crea una **nueva hoja** dentro del mismo archivo Excel y nómbrala `Practica12`.
4. Asegúrate de tener a la vista la tabla de clasificación de la Práctica 11 para identificar rápidamente los productos de menor y mayor CV.

---

## 6. Desarrollo Paso a Paso

---

### Paso 1 — Identificar los dos productos de análisis

**Objetivo:** Seleccionar el producto con CV más bajo (más estable) y el producto con CV más alto (más inestable) a partir de los resultados de la Práctica 11.

#### Instrucciones

1. Navega a la hoja de resultados de la Práctica 11 y localiza la columna de **Coeficiente de Variación (CV)** de los 8 SKUs.
2. Identifica visualmente (o usa `=MIN()` y `=MAX()` sobre la columna de CV) cuál SKU tiene el **CV más bajo** y cuál tiene el **CV más alto**.
3. En la hoja `Practica12`, crea una pequeña tabla de referencia en el rango `B2:C4` con el siguiente formato:

```
| Campo              | Producto Estable | Producto Inestable |
|--------------------|------------------|--------------------|
| Código SKU         | [SKU_X]          | [SKU_Y]            |
| CV (Práctica 11)   | [valor]          | [valor]            |
```

4. Completa la tabla con los valores reales de tu dataset. Por ejemplo, si el SKU-02 tiene CV = 0.08 y el SKU-07 tiene CV = 0.54, anótalos en la tabla.
5. Agrega debajo de la tabla (celda `B6`) la siguiente etiqueta en negrita: **"Regla de interpretación: diferencia promedio vs. mediana > 15% del promedio = señal de outliers (Lección 2.1)"**

> 💡 **Recordatorio de la Lección 2.1:** Un CV bajo (cercano a 0) indica que la demanda se mantiene cerca del promedio — es una demanda estable y predecible. Un CV alto indica alta dispersión respecto al promedio, lo que hace que el pronóstico sea inherentemente menos confiable.

#### Resultado esperado
Una tabla de referencia de 2 columnas claramente identificando el SKU estable y el SKU inestable con sus valores de CV de la Práctica 11.

#### Verificación
✅ Confirma que el SKU "estable" tiene el CV **numéricamente menor** de todos los 8 SKUs del dataset.
✅ Confirma que el SKU "inestable" tiene el CV **numéricamente mayor** de todos los 8 SKUs.
✅ La diferencia entre ambos CV debe ser notoria (en un dataset bien diseñado, al menos 0.20 puntos de diferencia).

---

### Paso 2 — Construir los gráficos de línea de demanda histórica

**Objetivo:** Visualizar el comportamiento mensual de demanda de ambos productos en gráficos de línea separados para identificar patrones diferenciados de forma visual.

#### Instrucciones

1. En la hoja `Practica12`, copia los 24 meses de demanda del **producto estable** en el rango `B10:C33`:
   - Columna B: Etiquetas de período (Mes 1, Mes 2, ... Mes 24) — o las fechas reales si tu dataset las tiene.
   - Columna C: Valores de demanda mensual del producto estable.

2. En el rango `E10:F33`, copia los 24 meses de demanda del **producto inestable**:
   - Columna E: Etiquetas de período (iguales a las de columna B).
   - Columna F: Valores de demanda mensual del producto inestable.

3. **Gráfico 1 — Producto Estable:**
   - Selecciona el rango `B10:C33`.
   - Ve a **Insertar → Gráficos → Gráfico de líneas → Línea con marcadores**.
   - Nombra el gráfico: `"Demanda Histórica - [SKU Estable]"`.
   - Coloca el gráfico debajo de los datos, aproximadamente en el rango `B35:G50`.

4. **Gráfico 2 — Producto Inestable:**
   - Selecciona el rango `E10:F33`.
   - Inserta un segundo gráfico de líneas con marcadores.
   - Nombra el gráfico: `"Demanda Histórica - [SKU Inestable]"`.
   - Coloca el gráfico a la derecha del primero, aproximadamente en el rango `I35:N50`.

5. En ambos gráficos, asegúrate de que:
   - El eje Y muestre el rango de valores de demanda.
   - El eje X muestre los períodos (meses).
   - Ambos gráficos usen la **misma escala en el eje Y** para facilitar la comparación visual. Para ajustar: clic derecho sobre el eje Y → Dar formato al eje → establece el mismo valor mínimo y máximo en ambos gráficos.

> 💡 **Tip visual:** Si ambos gráficos usan la misma escala, la diferencia en la "rugosidad" de la línea entre el producto estable (línea casi plana) y el inestable (línea con picos y valles pronunciados) será inmediatamente evidente.

#### Resultado esperado
Dos gráficos de línea en la hoja `Practica12`: uno que muestra una demanda relativamente uniforme (producto estable) y otro que muestra fluctuaciones pronunciadas (producto inestable). La diferencia visual debe ser clara e intuitiva.

#### Verificación
✅ Ambos gráficos tienen título descriptivo con el nombre/código del SKU correspondiente.
✅ Los ejes tienen la misma escala en el eje Y.
✅ La línea del producto estable visualmente "oscila menos" que la del producto inestable.
✅ Ambos gráficos muestran los 24 períodos completos.

---

### Paso 3 — Calcular las métricas estadísticas comparativas

**Objetivo:** Calcular lado a lado seis métricas estadísticas para ambos productos y construir una tabla comparativa formal.

#### Instrucciones

1. En la hoja `Practica12`, a partir de la celda `B53`, construye la siguiente tabla comparativa. Usa las celdas de la columna C para el producto estable y la columna D para el producto inestable:

```
| Métrica                     | Producto Estable | Producto Inestable |
|-----------------------------|------------------|--------------------|
| Promedio                    |                  |                    |
| Mediana                     |                  |                    |
| Desviación Estándar         |                  |                    |
| Coeficiente de Variación    |                  |                    |
| Rango (Máx - Mín)           |                  |                    |
| Rango Intercuartílico (IQR) |                  |                    |
```

2. Completa la columna **Producto Estable** con las siguientes fórmulas (asumiendo que los datos del producto estable están en `C10:C33`):

```excel
Promedio:            =PROMEDIO(C10:C33)
Mediana:             =MEDIANA(C10:C33)
Desv. Estándar:      =DESVEST.M(C10:C33)
CV:                  =DESVEST.M(C10:C33)/PROMEDIO(C10:C33)
Rango:               =MAX(C10:C33)-MIN(C10:C33)
IQR:                 =CUARTIL.EXC(C10:C33,3)-CUARTIL.EXC(C10:C33,1)
```

3. Completa la columna **Producto Inestable** con las mismas fórmulas, ajustando el rango a `F10:F33`:

```excel
Promedio:            =PROMEDIO(F10:F33)
Mediana:             =MEDIANA(F10:F33)
Desv. Estándar:      =DESVEST.M(F10:F33)
CV:                  =DESVEST.M(F10:F33)/PROMEDIO(F10:F33)
Rango:               =MAX(F10:F33)-MIN(F10:F33)
IQR:                 =CUARTIL.EXC(F10:F33,3)-CUARTIL.EXC(F10:F33,1)
```

4. Formatea los valores de CV como **porcentaje con 1 decimal** (selecciona las celdas de CV → Formato → Porcentaje → 1 decimal).

5. Agrega una fila adicional al final de la tabla llamada **"Diferencia Promedio vs. Mediana (%)"** y calcula:

```excel
=(PROMEDIO(C10:C33)-MEDIANA(C10:C33))/PROMEDIO(C10:C33)
```
   Haz lo mismo para el producto inestable con el rango `F10:F33`. Formatea como porcentaje con 1 decimal.

6. Aplica **formato condicional** a la fila "Diferencia Promedio vs. Mediana (%)":
   - Selecciona las dos celdas con ese valor.
   - Ve a **Inicio → Formato condicional → Resaltar reglas de celdas → Mayor que**.
   - Establece el umbral en `0.15` (15%) con relleno rojo claro.
   - Esto marcará automáticamente en rojo si la diferencia supera el umbral de alerta de la Lección 2.1.

> 💡 **Conexión con la Lección 2.1:** Recuerda que una diferencia superior al 10-15% entre promedio y mediana es una señal de que existen valores atípicos que distorsionan el promedio. Si el producto inestable activa esta alerta, es evidencia de que su demanda contiene eventos extraordinarios que deben investigarse — exactamente lo que explorarás con la IA en el Paso 5.

#### Resultado esperado
Una tabla de 7 filas × 3 columnas completamente calculada, con formato condicional activo en la fila de diferencia promedio/mediana. El producto inestable debe mostrar valores notoriamente más altos en desviación estándar, CV y rango que el producto estable.

#### Verificación
✅ Todas las fórmulas devuelven valores numéricos (sin errores `#DIV/0!` o `#¡VALOR!`).
✅ El CV del producto inestable coincide (o es muy cercano) al valor calculado en la Práctica 11.
✅ El IQR del producto inestable es mayor que el del producto estable.
✅ El formato condicional está activo y funciona correctamente (prueba ingresando un valor mayor a 0.15 manualmente para verificar que se pinta de rojo, luego restaura la fórmula).

---

### Paso 4 — Pronóstico naive y cálculo del error de pronóstico

**Objetivo:** Aplicar el pronóstico naive (promedio histórico de los primeros 21 meses) como estimación para los últimos 3 meses y calcular el error absoluto, introduciendo el concepto de error de pronóstico.

#### Instrucciones

> ⚠️ **Importante:** Para simular un escenario realista de pronóstico, usaremos los **primeros 21 meses** como historial de entrenamiento y los **últimos 3 meses** (meses 22, 23 y 24) como período de evaluación. Esto simula cómo un analista habría pronosticado esos 3 meses usando solo el historial disponible.

1. En la hoja `Practica12`, a partir de la celda `B65`, construye la siguiente sección. Etiquétala como **"Evaluación de Pronóstico Naive (Promedio Histórico)"**:

2. Calcula el **promedio de entrenamiento** (primeros 21 meses) para cada producto:

```excel
Promedio Entrenamiento (Estable):    =PROMEDIO(C10:C30)
Promedio Entrenamiento (Inestable):  =PROMEDIO(F10:F30)
```
   Anota estos valores en las celdas `C67` y `D67` respectivamente.

3. Construye la tabla de evaluación en el rango `B70:G74`:

```
| Período | Demanda Real (Estable) | Pronóstico Naive (Estable) | Error Abs. (Estable) | Demanda Real (Inestable) | Pronóstico Naive (Inestable) | Error Abs. (Inestable) |
|---------|------------------------|---------------------------|----------------------|--------------------------|------------------------------|------------------------|
| Mes 22  |                        |                           |                      |                          |                              |                        |
| Mes 23  |                        |                           |                      |                          |                              |                        |
| Mes 24  |                        |                           |                      |                          |                              |                        |
```

4. Completa la tabla de la siguiente manera:

   - **Demanda Real (Estable):** Copia manualmente (o referencia) los valores de los meses 22, 23 y 24 del producto estable desde `C31:C33`.
   - **Pronóstico Naive (Estable):** En las tres filas, escribe la misma fórmula referenciando el promedio de entrenamiento calculado en `C67`:
     ```excel
     =$C$67
     ```
   - **Error Absoluto (Estable):** Calcula el valor absoluto de la diferencia:
     ```excel
     =ABS([Demanda Real] - [Pronóstico Naive])
     ```
     Por ejemplo: `=ABS(C71-D71)` para el Mes 22.
   
   - Repite el proceso para el producto inestable usando `F31:F33` y `D67`.

5. Agrega una fila de **totales/promedios** debajo de la tabla:

```excel
MAE (Estable):    =PROMEDIO(E71:E73)
MAE (Inestable):  =PROMEDIO(H71:H73)
```

   Etiqueta estas celdas claramente como **"MAE Naive (3 meses)"**.

6. Agrega una celda de interpretación en `B77`:
   ```
   "El producto con mayor MAE es más difícil de pronosticar con el método naive.
    Un MAE alto indica que el promedio histórico NO es una buena estimación de la demanda futura."
   ```

> 💡 **Conexión conceptual:** El pronóstico naive usando el promedio es equivalente a asumir que la demanda futura será idéntica al promedio histórico. Para el producto estable, este supuesto es razonable. Para el producto inestable, el promedio puede quedar muy lejos de los valores reales — especialmente si hay picos o valles en esos 3 meses. Este es el fundamento del concepto de **error de pronóstico** que se desarrollará en prácticas posteriores con métricas MAE y MAPE.

#### Resultado esperado
Una tabla de evaluación de pronóstico con 3 períodos, errores absolutos calculados para ambos productos y el MAE naive de cada uno. El MAE del producto inestable debe ser notoriamente mayor que el del producto estable, evidenciando que la variabilidad dificulta el pronóstico.

#### Verificación
✅ El pronóstico naive es el **mismo valor** para los 3 meses de cada producto (es el promedio de los primeros 21 meses).
✅ Los errores absolutos son todos valores positivos (función `ABS` aplicada correctamente).
✅ El MAE del producto inestable es mayor que el del producto estable (en un dataset bien diseñado, significativamente mayor).
✅ No hay errores de fórmula en ninguna celda de la tabla.

---

### Paso 5 — Consulta a IA Generativa sobre factores de variabilidad

**Objetivo:** Utilizar una herramienta de IA Generativa para obtener una interpretación sobre los posibles factores logísticos y comerciales que explican la alta variabilidad del producto inestable, y documentar la respuesta.

#### Instrucciones

1. Abre tu herramienta de IA Generativa (Copilot, ChatGPT, Gemini o equivalente) en el navegador.

2. Formula la siguiente consulta, **personalizándola con los datos reales de tu análisis** (sustituye los valores entre corchetes con los valores calculados en los pasos anteriores):

```
Prompt sugerido:

"Soy analista logístico y estoy analizando el comportamiento histórico de demanda 
de dos productos durante 24 meses. El producto estable tiene un coeficiente de 
variación del [CV_ESTABLE]% y un MAE naive de [MAE_ESTABLE] unidades. El producto 
inestable tiene un coeficiente de variación del [CV_INESTABLE]% y un MAE naive de 
[MAE_INESTABLE] unidades. Además, la diferencia entre el promedio y la mediana del 
producto inestable es del [DIFERENCIA_%]%, lo que sugiere la presencia de valores 
atípicos.

¿Cuáles son los 5 factores logísticos, comerciales o externos más comunes que 
podrían explicar una demanda tan variable en un contexto de distribución y 
abastecimiento? ¿Cómo impacta esta variabilidad en la planificación del inventario 
y el nivel de servicio?"
```

3. Lee la respuesta de la IA y selecciona los **3 factores más relevantes** para el contexto logístico de tu organización (o del caso de estudio del curso).

4. En la hoja `Practica12`, a partir de la celda `B80`, crea una sección titulada **"Análisis IA: Factores de Variabilidad"** y documenta:
   - El prompt exacto que utilizaste (puedes copiarlo y pegarlo en la celda, ajustando el ancho de columna).
   - Los 3 factores seleccionados de la respuesta de la IA, con una breve justificación de por qué son relevantes.
   - Una celda adicional con tu reflexión personal: **"¿Cómo impacta este factor en la operación de picking y el nivel de servicio?"** (2-3 líneas por factor).

5. Guarda el archivo Excel con el nombre: `Practica12_[TuNombre]_Completada.xlsx`.

> 💡 **Tip para la consulta IA:** Si la herramienta de IA no está disponible, utiliza el siguiente prompt alternativo con una herramienta offline o discute con un compañero: "¿Qué factores conoces del sector logístico que generan picos de demanda impredecibles?" Documenta igualmente las ideas generadas.

> ⚠️ **Nota importante:** La respuesta de la IA es un punto de partida interpretativo, no una verdad absoluta. Tu rol como analista es evaluar críticamente qué factores son aplicables al contexto específico del producto analizado.

#### Resultado esperado
Una sección documentada en la hoja `Practica12` con el prompt utilizado, los 3 factores de variabilidad seleccionados de la respuesta de la IA y la reflexión sobre su impacto operativo en picking y nivel de servicio.

#### Verificación
✅ El prompt incluye los valores numéricos reales del análisis (CV, MAE, diferencia promedio/mediana).
✅ Se documentaron exactamente 3 factores con justificación.
✅ La reflexión operativa conecta cada factor con una consecuencia concreta en inventario, picking o nivel de servicio.
✅ El archivo fue guardado con el nombre correcto.

---

## 7. Validación y Pruebas

Al completar todos los pasos, verifica que tu hoja `Practica12` contenga los siguientes elementos en el orden indicado:

| # | Elemento | Ubicación aproximada | ¿Completado? |
|---|---|---|---|
| 1 | Tabla de referencia de productos seleccionados (SKU estable vs. inestable con CV) | B2:D4 | ☐ |
| 2 | Datos de demanda de ambos productos (24 meses) | B10:C33 y E10:F33 | ☐ |
| 3 | Gráfico de línea — Producto Estable | B35:G50 | ☐ |
| 4 | Gráfico de línea — Producto Inestable | I35:N50 | ☐ |
| 5 | Tabla comparativa de 6 métricas + diferencia promedio/mediana | B53:D61 | ☐ |
| 6 | Formato condicional activo en fila de diferencia promedio/mediana | C61:D61 | ☐ |
| 7 | Tabla de evaluación de pronóstico naive (3 meses) | B70:H74 | ☐ |
| 8 | MAE naive calculado para ambos productos | C75:H75 | ☐ |
| 9 | Sección de análisis IA con prompt y 3 factores documentados | B80:D95 | ☐ |
| 10 | Archivo guardado con nombre correcto | — | ☐ |

### Prueba de coherencia de resultados

Ejecuta las siguientes verificaciones cruzadas para asegurarte de que los resultados son internamente consistentes:

```
VERIFICACIÓN 1: CV en tabla comparativa (Paso 3) debe coincidir con CV de Práctica 11
  → Si difieren: revisa que estás usando el mismo rango de datos (24 meses completos)

VERIFICACIÓN 2: Promedio de entrenamiento (Paso 4, primeros 21 meses) debe ser
  cercano pero no idéntico al promedio de los 24 meses completos
  → Si son idénticos: probablemente estás referenciando el rango incorrecto

VERIFICACIÓN 3: MAE del producto inestable > MAE del producto estable
  → Si el producto estable tiene mayor MAE: revisa que asignaste correctamente
    cuál es el estable y cuál el inestable en la tabla del Paso 1

VERIFICACIÓN 4: IQR ≤ Rango (el IQR siempre es menor o igual al rango total)
  → Si IQR > Rango: hay un error en la función CUARTIL.EXC o en el rango referenciado
```

---

## 8. Solución de Problemas

### Problema 1: La función `CUARTIL.EXC` devuelve `#¡NUM!` o un valor inesperado

**Síntoma:** Al escribir `=CUARTIL.EXC(F10:F33,1)` o `=CUARTIL.EXC(F10:F33,3)`, Excel devuelve el error `#¡NUM!` o un valor que parece incorrecto (por ejemplo, mayor que el máximo del rango).

**Causa:** El error `#¡NUM!` en `CUARTIL.EXC` ocurre cuando el argumento `cuartil` está fuera del rango permitido (debe ser 1, 2 o 3 para esta función), o cuando el rango de datos contiene celdas vacías o texto mezclado con números. También puede ocurrir si el dataset tiene menos de 4 valores, lo que hace que el cálculo de cuartiles exclusivos no sea posible.

**Solución:**
1. Verifica que el segundo argumento de la función sea exactamente `1` (Q1) o `3` (Q3) — no `0` ni `4` (esos son exclusivos de `CUARTIL.INC`).
2. Selecciona el rango de datos (`F10:F33`) y usa **Inicio → Buscar y seleccionar → Ir a Especial → Celdas en blanco** para identificar si hay celdas vacías en el rango.
3. Si hay celdas vacías, rellénalas con el valor correcto del dataset o usa el archivo de checkpoint proporcionado por el instructor.
4. Como alternativa temporal, puedes calcular el IQR usando `CUARTIL.INC` (disponible en todas las versiones de Excel 2016+): `=CUARTIL.INC(F10:F33,3)-CUARTIL.INC(F10:F33,1)`. El resultado será ligeramente diferente pero conceptualmente equivalente para los fines de esta práctica.

---

### Problema 2: Los dos gráficos de línea no tienen la misma escala en el eje Y y la comparación visual no es clara

**Síntoma:** Al colocar los dos gráficos de línea uno al lado del otro, el producto estable parece "más variable" de lo que realmente es porque Excel ajustó automáticamente su eje Y a una escala más pequeña (por ejemplo, 80–160 unidades), mientras que el producto inestable usa una escala más amplia (por ejemplo, 0–500 unidades). Esto distorsiona la comparación visual.

**Causa:** Excel ajusta automáticamente la escala del eje Y de cada gráfico de forma independiente para maximizar el uso del espacio visual. Esto es conveniente para ver detalles dentro de cada gráfico, pero dificulta la comparación entre dos gráficos distintos.

**Solución:**
1. Determina el valor máximo entre ambos datasets: `=MAX(C10:C33,F10:F33)`. Redondea ese valor hacia arriba al siguiente número "limpio" (por ejemplo, si el máximo es 487, usa 500).
2. Determina el valor mínimo entre ambos: `=MIN(C10:C33,F10:F33)`. Si es positivo y mayor que 50, puedes usar 0 como mínimo para simplificar.
3. En el **Gráfico 1 (Producto Estable):** Haz doble clic sobre el eje Y → Panel "Dar formato al eje" → En "Límites", establece el **Mínimo** en 0 (o el valor calculado) y el **Máximo** en el valor determinado en el paso 1.
4. Repite exactamente el mismo proceso para el **Gráfico 2 (Producto Inestable)** usando los mismos valores de mínimo y máximo.
5. Ahora ambos gráficos usan la misma escala y la comparación visual es válida: la línea del producto estable se verá notoriamente más "plana" que la del inestable.

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza las siguientes acciones de cierre:

1. **Guarda el archivo** con el nombre `Practica12_[TuNombre]_Completada.xlsx` en la carpeta de trabajo del curso.
2. **Verifica el archivo guardado:** Ciérralo y vuelve a abrirlo para confirmar que todos los gráficos, fórmulas y el formato condicional se guardaron correctamente (los gráficos de Excel a veces pierden formato si se guarda en un formato incorrecto — asegúrate de guardar como `.xlsx`, no como `.csv`).
3. **Cierra la sesión de IA Generativa** si estás usando una cuenta compartida o un acceso institucional.
4. **No elimines la hoja `Practica12`** del archivo — este archivo será el punto de partida para las prácticas siguientes del módulo.
5. Si el instructor lo solicita, **comparte el archivo** a través del canal designado (correo, Teams, carpeta compartida) antes de que finalice la sesión.

> ⚠️ **Nota:** No cierres ni modifiques el archivo de la Práctica 11. Las prácticas posteriores pueden necesitar referenciar tanto los resultados de la Práctica 11 como los de esta Práctica 12.

---

## 10. Resumen y Recursos Adicionales

### Lo que aprendiste en esta práctica

En esta práctica aplicaste de forma integrada los conceptos estadísticos de las Lecciones 2.1 y 2.2 a un caso de análisis comparativo real:

| Concepto aplicado | Herramienta usada | Resultado logístico |
|---|---|---|
| Selección por CV mínimo/máximo | Tabla de Práctica 11 | Identificación de productos estables e inestables |
| Visualización de patrones de demanda | Gráficos de línea en Excel | Evidencia visual de la variabilidad diferenciada |
| Promedio y mediana comparados | `PROMEDIO()`, `MEDIANA()` | Detección de outliers (regla del 15%) |
| Dispersión y rango | `DESVEST.M()`, `MAX()`, `MIN()`, `CUARTIL.EXC()` | Cuantificación del riesgo de pronóstico |
| Pronóstico naive | Promedio de entrenamiento (21 meses) | Primer cálculo de error de pronóstico (MAE) |
| Interpretación con IA Generativa | Copilot / ChatGPT / Gemini | Factores explicativos de la variabilidad |

### Conexión con el ciclo de forecasting logístico

Esta práctica establece una base conceptual crítica para las prácticas posteriores:

- **El MAE naive calculado hoy** es la línea base de comparación contra la que se medirán los modelos de pronóstico más sofisticados (promedio móvil, suavización exponencial) en prácticas futuras.
- **La diferencia entre el producto estable e inestable** anticipa la necesidad de seleccionar modelos de pronóstico diferentes según el tipo de producto — concepto central de las Prácticas 13-15.
- **La regla del 15% (promedio vs. mediana)** aprendida en la Lección 2.1 se convierte hoy en una herramienta operativa que puedes aplicar directamente en tu trabajo diario para detectar outliers antes de construir un pronóstico.

### Pregunta de reflexión final

> **¿Cómo impacta este análisis en una decisión operativa de logística?**
>
> Si el producto inestable es un SKU de alta rotación en tu operación, ¿qué estrategia de inventario adoptarías para mantener el nivel de servicio deseado a pesar de la alta variabilidad? ¿Sería suficiente con un stock de seguridad basado en el promedio, o necesitarías una estrategia diferente? Discute con tu grupo.

### Recursos adicionales

- [Microsoft Support — Función CUARTIL.EXC](https://support.microsoft.com/es-es/office/cuartil-exc-funci%C3%B3n-cuartil-exc-5a355b7a-840b-4a01-b0f1-f538c2864cad)
- [Microsoft Support — Crear un gráfico de líneas en Excel](https://support.microsoft.com/es-es/office/tipos-de-gr%C3%A1ficos-disponibles-en-office-a6187218-807e-4103-9e0a-27cdb19afb90)
- [Hyndman, R. J. & Athanasopoulos, G. — *Forecasting: Principles and Practice*, Cap. 3: Evaluación de la precisión del pronóstico](https://otexts.com/fpp3/accuracy.html)
- [Khan Academy en Español — Rango intercuartílico (IQR)](https://es.khanacademy.org/math/statistics-probability/summarizing-quantitative-data/interquartile-range-iqr/a/interquartile-range-review)

---
