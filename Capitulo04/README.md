# Práctica 19 — Cálculo de Errores Paso a Paso en Excel

## 1. Metadatos

| Campo            | Detalle                                      |
|------------------|----------------------------------------------|
| **Duración**     | 12 minutos                                   |
| **Complejidad**  | Media                                        |
| **Nivel Bloom**  | Aplicar (Apply)                              |
| **Módulo**       | 4 — Métricas de Error de Pronóstico          |
| **Práctica #**   | 19                                           |
| **Herramienta**  | Microsoft Excel 2016 o superior              |

---

## 2. Descripción General

En esta práctica construirás paso a paso una tabla de evaluación de errores de pronóstico en Excel, aplicando las funciones `ABS`, `PROMEDIO`, `RAIZ` y `SI` para calcular el Error Absoluto, el Error Porcentual, el MAE, el MAPE y el RMSE. Trabajarás con tres conjuntos de datos que representan distintos perfiles de demanda logística: un producto estable (Producto A), un producto variable (Producto B) y un producto con un valor atípico (*outlier*) (Producto C), lo que te permitirá experimentar directamente cómo la naturaleza de la demanda afecta la interpretación de cada métrica. Al finalizar, vincularás los valores obtenidos con decisiones operativas concretas de inventario y nivel de servicio.

---

## 3. Objetivos de Aprendizaje

- [ ] Calcular el Error Absoluto, el Error Porcentual y el Error Cuadrático columna por columna en Excel usando fórmulas explícitas y verificables.
- [ ] Obtener las métricas agregadas MAE, MAPE y RMSE utilizando las funciones `PROMEDIO` y `RAIZ` de Excel.
- [ ] Clasificar el MAPE de cada producto según los rangos estándar de aceptabilidad (< 10 %, 10–20 %, 20–50 %, > 50 %).
- [ ] Interpretar el impacto logístico de cada nivel de error en términos de riesgo de desabasto, sobrestock y nivel de servicio.
- [ ] Identificar cómo un único *outlier* distorsiona el MAPE y qué implicaciones tiene esto para la toma de decisiones de abastecimiento.

---

## 4. Prerrequisitos

### Conocimiento previo

| Tema                                                                 | Dónde se cubrió          |
|----------------------------------------------------------------------|--------------------------|
| Concepto de error de pronóstico con signo y error absoluto           | Lección 4.1 (este módulo)|
| Definición de MAE y MAPE                                             | Módulos 4.3 y 4.4        |
| Funciones básicas de Excel: `ABS`, `PROMEDIO`, `RAIZ`, `SI`          | Módulo 2 / prerequisito  |
| Impacto del error de pronóstico en inventario y nivel de servicio    | Módulos 2.5 y 3          |

### Acceso y archivos requeridos

| Recurso                                              | Estado requerido                                          |
|------------------------------------------------------|-----------------------------------------------------------|
| Microsoft Excel 2016 o superior (Microsoft 365 recomendado) | Instalado y funcional                              |
| Archivo `Lab04_Errores_Base.xlsx` (proporcionado por el instructor) | Descargado en carpeta de trabajo           |
| Conexión a internet (opcional para esta práctica)    | No estrictamente necesaria                                |

> **Nota para el instructor:** Si el archivo base no está disponible, los datos de los tres productos se encuentran en la sección de Entorno de Laboratorio y pueden ingresarse manualmente en menos de 3 minutos.

---

## 5. Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente        | Mínimo                        | Recomendado                  |
|-------------------|-------------------------------|------------------------------|
| Procesador        | Intel Core i5 / AMD Ryzen 5   | Intel Core i7 / AMD Ryzen 7  |
| RAM               | 4 GB                          | 8 GB                         |
| Almacenamiento    | 2 GB libres en disco          | 5 GB libres                  |
| Resolución        | 1366 × 768                    | 1920 × 1080                  |

### Software requerido

| Aplicación                  | Versión mínima   | Notas                                                      |
|-----------------------------|------------------|------------------------------------------------------------|
| Microsoft Excel             | 2016             | `PROMEDIO`, `ABS`, `RAIZ`, `SI` disponibles desde Excel 2010 |
| Sistema operativo           | Windows 10 / macOS 11 | Excel para Mac es compatible con todas las fórmulas usadas |

### Datos de referencia — Tablas de entrada

Utiliza los siguientes datos para construir el archivo si no dispones del archivo base. Ingresa cada producto en una hoja separada del libro Excel.

**Producto A — Demanda Estable (SKU: PROD-A)**

| Período | Demanda Real D(t) | Pronóstico F(t) |
|---------|-------------------|-----------------|
| 1       | 200               | 195             |
| 2       | 205               | 200             |
| 3       | 198               | 202             |
| 4       | 210               | 205             |
| 5       | 203               | 205             |
| 6       | 197               | 200             |
| 7       | 208               | 202             |
| 8       | 201               | 204             |
| 9       | 196               | 199             |
| 10      | 204               | 200             |
| 11      | 207               | 203             |
| 12      | 200               | 202             |

**Producto B — Demanda Variable (SKU: PROD-B)**

| Período | Demanda Real D(t) | Pronóstico F(t) |
|---------|-------------------|-----------------|
| 1       | 150               | 180             |
| 2       | 230               | 180             |
| 3       | 90                | 180             |
| 4       | 310               | 200             |
| 5       | 120               | 200             |
| 6       | 270               | 200             |
| 7       | 95                | 190             |
| 8       | 285               | 190             |
| 9       | 140               | 190             |
| 10      | 260               | 210             |
| 11      | 110               | 210             |
| 12      | 300               | 210             |

**Producto C — Demanda con Outlier (SKU: PROD-C)**

| Período | Demanda Real D(t) | Pronóstico F(t) |
|---------|-------------------|-----------------|
| 1       | 180               | 175             |
| 2       | 185               | 180             |
| 3       | 178               | 182             |
| 4       | 190               | 183             |
| 5       | 182               | 184             |
| 6       | 188               | 184             |
| 7       | 850               | 185             |
| 8       | 183               | 185             |
| 9       | 177               | 183             |
| 10      | 186               | 180             |
| 11      | 184               | 182             |
| 12      | 179               | 181             |

> **Período 7 de Producto C:** El valor 850 representa un evento extraordinario (ej. promoción masiva no planificada o error de registro). Este *outlier* es intencional y es el objeto de análisis en el Paso 4.

### Preparación inicial del archivo

1. Abre Microsoft Excel y crea un libro nuevo (o abre `Lab04_Errores_Base.xlsx`).
2. Renombra las hojas como: `PROD-A`, `PROD-B`, `PROD-C` y `Resumen`.
3. En cada hoja de producto, ingresa los encabezados y datos de la tabla correspondiente, comenzando en la celda `A1`.
4. Guarda el archivo como `Lab04_Errores_[TuNombre].xlsx` en tu carpeta de trabajo.

---

## 6. Procedimiento Paso a Paso

---

### Paso 1 — Construir la Tabla de Errores para Producto A (Demanda Estable)

**Objetivo:** Crear las cuatro columnas de error fila por fila para el Producto A, aplicando las fórmulas correctas de Excel y verificando la lógica de cada cálculo.

#### Instrucciones

1. Ve a la hoja `PROD-A`. Verifica que la columna `A` contiene `Período`, la columna `B` contiene `Demanda Real D(t)` y la columna `C` contiene `Pronóstico F(t)`.

2. En la celda `D1`, escribe el encabezado:
   ```
   Error = D(t) - F(t)
   ```

3. En la celda `E1`, escribe:
   ```
   Error Absoluto = |Error|
   ```

4. En la celda `F1`, escribe:
   ```
   Error Porcentual (%)
   ```

5. En la celda `G1`, escribe:
   ```
   Error Cuadrático
   ```

6. En la celda `D2`, ingresa la fórmula del error con signo:
   ```excel
   =B2-C2
   ```

7. En la celda `E2`, ingresa la fórmula del error absoluto usando la función `ABS`:
   ```excel
   =ABS(B2-C2)
   ```
   > Esta fórmula implementa directamente `|D(t) - F(t)|`, eliminando el signo tal como se explicó en la Lección 4.1.

8. En la celda `F2`, ingresa la fórmula del error porcentual con protección para división por cero mediante la función `SI`:
   ```excel
   =SI(B2=0, 0, ABS(B2-C2)/B2*100)
   ```
   > La función `SI` evita el error `#¡DIV/0!` en períodos donde la demanda real sea cero. En esos casos, el error porcentual se registra como 0 (o puede dejarse en blanco según criterio del analista).

9. En la celda `G2`, ingresa la fórmula del error cuadrático:
   ```excel
   =(B2-C2)^2
   ```

10. Selecciona el rango `D2:G2` y copia las fórmulas hacia abajo hasta la fila 13 (período 12). Puedes hacerlo con `Ctrl+C` y luego seleccionando `D3:G13` y presionando `Ctrl+V`, o arrastrando el controlador de relleno.

11. Aplica formato numérico:
    - Columnas `D`, `E`, `G`: **Número** con 1 decimal.
    - Columna `F`: **Número** con 2 decimales (no usar formato porcentaje de Excel, ya que el valor ya está multiplicado por 100).

#### Resultado esperado

Tu tabla debería verse así (valores aproximados):

| Período | D(t) | F(t) | Error | Err. Absoluto | Err. Porc. (%) | Err. Cuadrático |
|---------|------|------|-------|---------------|----------------|-----------------|
| 1       | 200  | 195  | 5     | 5             | 2.50           | 25              |
| 2       | 205  | 200  | 5     | 5             | 2.44           | 25              |
| 3       | 198  | 202  | -4    | 4             | 2.02           | 16              |
| ...     | ...  | ...  | ...   | ...           | ...            | ...             |

#### Verificación

- Confirma que la columna `D` contiene tanto valores positivos como negativos.
- Confirma que la columna `E` contiene **únicamente valores positivos o cero**.
- Verifica la celda `E2` manualmente: `|200 - 195| = 5` ✓
- Verifica la celda `F2` manualmente: `5 / 200 × 100 = 2.50 %` ✓

---

### Paso 2 — Calcular MAE, MAPE y RMSE para Producto A

**Objetivo:** Agregar las métricas resumen (MAE, MAPE, RMSE) usando funciones `PROMEDIO` y `RAIZ`, y crear una tabla de clasificación del MAPE.

#### Instrucciones

1. En la celda `A15`, escribe el encabezado de sección:
   ```
   MÉTRICAS DE ERROR - PRODUCTO A
   ```
   Aplica negrita y color de fondo gris claro para destacarlo visualmente.

2. En la celda `A16`, escribe `MAE`. En la celda `B16`, ingresa:
   ```excel
   =PROMEDIO(E2:E13)
   ```
   > El MAE (Mean Absolute Error) es el promedio de todos los errores absolutos. Representa cuántas unidades nos equivocamos en promedio por período.

3. En la celda `A17`, escribe `MAPE (%)`. En la celda `B17`, ingresa:
   ```excel
   =PROMEDIO(F2:F13)
   ```
   > El MAPE (Mean Absolute Percentage Error) es el promedio de los errores porcentuales. Permite comparar la precisión del pronóstico entre productos con escalas de demanda distintas.

4. En la celda `A18`, escribe `RMSE`. En la celda `B18`, ingresa:
   ```excel
   =RAIZ(PROMEDIO(G2:G13))
   ```
   > El RMSE (Root Mean Square Error) penaliza más los errores grandes. Es útil como métrica complementaria para detectar períodos con desviaciones extremas.

5. Formatea las celdas `B16` y `B18` como **Número con 2 decimales**. Formatea `B17` como **Número con 2 decimales** (el valor ya está en porcentaje).

6. En la celda `A20`, escribe el encabezado:
   ```
   CLASIFICACIÓN DEL MAPE
   ```

7. Construye la siguiente tabla de clasificación a partir de la celda `A21`:

   | Rango MAPE          | Calificación  | Implicación Logística                                       |
   |---------------------|---------------|-------------------------------------------------------------|
   | < 10 %              | Excelente     | Bajo riesgo de desabasto/sobrestock. Modelo confiable.      |
   | 10 % – 20 %         | Bueno         | Riesgo moderado. Requiere safety stock estándar.            |
   | 20 % – 50 %         | Razonable     | Riesgo elevado. Revisar modelo o aumentar buffer de inventario. |
   | > 50 %              | Pobre         | Modelo no confiable. Replantear método de pronóstico.       |

   Ingresa esta información en las celdas `A21:D24` (una fila por rango).

8. En la celda `A26`, escribe:
   ```
   Clasificación de PROD-A:
   ```
   En la celda `B26`, ingresa la siguiente fórmula para clasificar automáticamente:
   ```excel
   =SI(B17<10,"Excelente",SI(B17<20,"Bueno",SI(B17<50,"Razonable","Pobre")))
   ```

9. En la celda `A27`, escribe la interpretación logística. Ejemplo de texto a ingresar manualmente:
   ```
   Con un MAPE de [valor]%, el pronóstico de PROD-A es [clasificación].
   El error promedio de [MAE] unidades por período representa un riesgo
   [bajo/moderado/alto] de desabasto o sobrestock.
   ```
   Reemplaza los corchetes con los valores reales calculados.

#### Resultado esperado

Para el Producto A (demanda estable), deberías obtener:
- **MAE:** aproximadamente 5–8 unidades
- **MAPE:** aproximadamente 2.5 %–4.0 %
- **RMSE:** ligeramente superior al MAE
- **Clasificación:** **Excelente**

#### Verificación

- Calcula manualmente el MAE sumando los valores de la columna `E` y dividiendo entre 12. Compara con `B16`.
- Verifica que `B26` muestre `"Excelente"` dado el MAPE esperado para datos estables.
- Confirma que el RMSE (`B18`) sea mayor o igual al MAE (`B16`). Esta relación siempre debe cumplirse matemáticamente.

---

### Paso 3 — Replicar el Análisis para Producto B (Demanda Variable)

**Objetivo:** Aplicar el mismo esquema de cálculo al Producto B y comparar los resultados con el Producto A para evidenciar el impacto de la variabilidad de la demanda en las métricas de error.

#### Instrucciones

1. Ve a la hoja `PROD-B`.

2. **Método rápido (recomendado para ahorrar tiempo):** Copia el rango completo `D1:G27` de la hoja `PROD-A` y pégalo en la misma ubicación de la hoja `PROD-B`. Las fórmulas se ajustarán automáticamente a las referencias de la nueva hoja.

   > **Verificación inmediata:** Confirma que `B16` (MAE de PROD-B) muestra un valor diferente al de PROD-A antes de continuar.

3. Actualiza el encabezado en `A15` para que diga:
   ```
   MÉTRICAS DE ERROR - PRODUCTO B
   ```

4. Actualiza el texto de interpretación en `A27` con los valores reales de PROD-B.

5. Observa la columna `D` (Error con signo) de PROD-B: los errores deberían alternar entre positivos y negativos de forma marcada, reflejando la alta variabilidad de la demanda.

6. Aplica formato condicional a la columna `E` (Error Absoluto) para resaltar visualmente los períodos con mayor error:
   - Selecciona `E2:E13`.
   - Ve a **Inicio → Formato condicional → Escalas de color**.
   - Elige la escala **Verde–Amarillo–Rojo** (verde = error bajo, rojo = error alto).

#### Resultado esperado

Para el Producto B (demanda variable), deberías obtener:
- **MAE:** aproximadamente 70–100 unidades (significativamente mayor que PROD-A)
- **MAPE:** aproximadamente 35 %–55 %
- **RMSE:** notablemente mayor que el MAE (indica errores grandes concentrados en pocos períodos)
- **Clasificación:** **Razonable** o **Pobre**

#### Verificación

- Compara `B16` (MAE) de PROD-B con el de PROD-A. El de PROD-B debe ser sustancialmente mayor.
- Verifica que `B26` muestre `"Razonable"` o `"Pobre"`.
- Observa la columna de formato condicional: los períodos con mayor error (ej. período 4 con demanda 310 vs pronóstico 200) deben aparecer en rojo.

---

### Paso 4 — Analizar el Efecto del Outlier en Producto C

**Objetivo:** Calcular las métricas para Producto C, identificar cómo el *outlier* del período 7 distorsiona el MAPE, y documentar las implicaciones logísticas de este fenómeno.

#### Instrucciones

1. Ve a la hoja `PROD-C`.

2. Repite el proceso del Paso 3: copia el rango `D1:G27` desde `PROD-A` y pégalo en `PROD-C`. Actualiza el encabezado en `A15`.

3. Observa la columna `F` (Error Porcentual). Localiza la fila correspondiente al **Período 7** (fila 8 en Excel si los datos empiezan en fila 2):
   - Demanda Real: 850 | Pronóstico: 185
   - Error Absoluto: `|850 - 185| = 665`
   - Error Porcentual: `665 / 850 × 100 = 78.24 %`

   Este único período domina el cálculo del MAPE.

4. En la celda `A29`, escribe:
   ```
   ANÁLISIS DE DISTORSIÓN POR OUTLIER
   ```

5. En la celda `A30`, escribe `MAPE con outlier (período 7 incluido):`. En `B30`, referencia `B17`.

6. En la celda `A31`, escribe `MAPE sin outlier (período 7 excluido):`. En `B31`, ingresa:
   ```excel
   =PROMEDIO(F2:F7,F9:F13)
   ```
   > Esta fórmula calcula el MAPE excluyendo el período 7 para mostrar el MAPE "real" de la demanda normal del producto.

7. En la celda `A32`, escribe `Diferencia (puntos porcentuales):`. En `B32`, ingresa:
   ```excel
   =B30-B31
   ```

8. En la celda `A34`, escribe el siguiente bloque de análisis (ajusta los valores según tus resultados):
   ```
   IMPLICACIÓN LOGÍSTICA DEL OUTLIER:
   El período 7 (demanda 850 vs pronóstico 185) representa un evento
   extraordinario (posible promoción no planificada o error de captura).
   Este único período eleva el MAPE de ~X% a ~Y%, cambiando la
   clasificación del modelo de [A] a [B].
   
   Decisión recomendada: Investigar la causa del período 7 antes de
   usar el MAPE global para tomar decisiones de inventario. Si fue un
   evento irrepetible, considerar excluirlo del cálculo de error o
   registrarlo como evento extraordinario en el historial.
   ```

9. Aplica formato condicional a la celda `F8` (error porcentual del período 7) con **relleno rojo** y **texto blanco** para destacarlo visualmente.

#### Resultado esperado

Para el Producto C, deberías obtener:
- **MAPE con outlier:** aproximadamente 65 %–80 % → Clasificación: **Pobre**
- **MAPE sin outlier:** aproximadamente 3 %–6 % → Clasificación: **Excelente**
- **Diferencia:** más de 60 puntos porcentuales
- **Conclusión:** El modelo es excelente para la demanda normal del producto; el *outlier* del período 7 es el único responsable de la clasificación negativa.

#### Verificación

- Confirma que `B31` (MAPE sin outlier) es significativamente menor que `B30` (MAPE con outlier).
- Verifica manualmente el error porcentual del período 7: `|850-185|/850*100 = 78.24 %` ✓
- Confirma que la celda `F8` aparece destacada en rojo.

---

### Paso 5 — Construir el Tablero Comparativo en la Hoja Resumen

**Objetivo:** Consolidar los resultados de los tres productos en una tabla comparativa única que permita tomar decisiones logísticas informadas sobre cada SKU.

#### Instrucciones

1. Ve a la hoja `Resumen`.

2. En la celda `A1`, escribe el título:
   ```
   RESUMEN COMPARATIVO DE MÉTRICAS DE ERROR DE PRONÓSTICO
   ```
   Aplica negrita, tamaño de fuente 14 y color de fondo azul oscuro con texto blanco.

3. Construye la tabla de encabezados a partir de la celda `A3`:

   | SKU    | MAE (unid.) | MAPE (%) | RMSE (unid.) | Clasificación | Riesgo Logístico        | Acción Recomendada                        |
   |--------|-------------|----------|--------------|---------------|-------------------------|-------------------------------------------|
   | PROD-A |             |          |              |               |                         |                                           |
   | PROD-B |             |          |              |               |                         |                                           |
   | PROD-C (con outlier) | | |          |               |                         |                                           |
   | PROD-C (sin outlier) | | |          |               |                         |                                           |

4. En la celda `B4` (MAE de PROD-A), ingresa una referencia directa a la celda de la hoja correspondiente:
   ```excel
   ='PROD-A'!B16
   ```

5. Repite el patrón para las demás celdas de métricas:
   - `C4` → `='PROD-A'!B17` (MAPE)
   - `D4` → `='PROD-A'!B18` (RMSE)
   - `E4` → `='PROD-A'!B26` (Clasificación)
   - `B5` → `='PROD-B'!B16`
   - `C5` → `='PROD-B'!B17`
   - `D5` → `='PROD-B'!B18`
   - `E5` → `='PROD-B'!B26`
   - `B6` → `='PROD-C'!B16`
   - `C6` → `='PROD-C'!B17`
   - `D6` → `='PROD-C'!B18`
   - `E6` → `='PROD-C'!B26`
   - `B7` → `='PROD-C'!B31` (MAE sin outlier — puedes calcularlo en PROD-C si no lo tienes aún)
   - `C7` → `='PROD-C'!B31`
   - `E7` → fórmula de clasificación basada en `C7`

6. Completa manualmente las columnas `F` (Riesgo Logístico) y `G` (Acción Recomendada) con el siguiente contenido:

   | SKU    | Riesgo Logístico                        | Acción Recomendada                                                |
   |--------|-----------------------------------------|-------------------------------------------------------------------|
   | PROD-A | Bajo — modelo confiable                 | Mantener modelo actual. Revisar trimestralmente.                  |
   | PROD-B | Alto — demanda impredecible             | Aumentar safety stock. Evaluar modelo de suavización exponencial. |
   | PROD-C (con outlier) | Aparente — distorsionado  | Investigar período 7. No tomar decisiones con este MAPE.          |
   | PROD-C (sin outlier) | Bajo — modelo confiable   | Registrar período 7 como evento y excluir del historial base.     |

7. Aplica formato condicional a la columna `C` (MAPE) con escala de colores:
   - Selecciona `C4:C7`.
   - **Inicio → Formato condicional → Escalas de color → Verde–Amarillo–Rojo** (verde = MAPE bajo, rojo = MAPE alto).

8. Guarda el archivo con `Ctrl+S`.

#### Resultado esperado

La hoja `Resumen` debe mostrar una tabla limpia con los cuatro escenarios, donde PROD-A y PROD-C (sin outlier) aparecen en verde, PROD-B en amarillo/rojo, y PROD-C (con outlier) en rojo intenso.

#### Verificación

- Confirma que las referencias entre hojas funcionan correctamente: cambia un valor de pronóstico en `PROD-A` y verifica que `B4` en `Resumen` se actualiza automáticamente. Luego deshaz el cambio con `Ctrl+Z`.
- Verifica que la columna `E` (Clasificación) muestra texto coherente con los valores de MAPE en la columna `C`.

---

## 7. Validación y Pruebas

Ejecuta las siguientes verificaciones antes de dar por completada la práctica:

### Lista de comprobación final

| # | Verificación                                                                                    | Criterio de éxito                                              |
|---|------------------------------------------------------------------------------------------------|----------------------------------------------------------------|
| 1 | Columna `D` (Error con signo) en las tres hojas contiene valores positivos **y** negativos     | Al menos un valor positivo y uno negativo en cada hoja         |
| 2 | Columna `E` (Error Absoluto) en las tres hojas contiene **únicamente valores ≥ 0**             | Ningún valor negativo en toda la columna                       |
| 3 | Fórmula `SI` en columna `F` maneja correctamente la división por cero                          | Cambiar temporalmente una celda D(t) a 0 → F(t) muestra 0, no error |
| 4 | MAE de PROD-A < MAE de PROD-B                                                                  | La demanda estable produce menor error promedio                |
| 5 | MAPE de PROD-C con outlier > MAPE de PROD-C sin outlier en al menos 50 puntos porcentuales     | El outlier distorsiona significativamente la métrica           |
| 6 | RMSE ≥ MAE en las tres hojas                                                                   | Relación matemática siempre válida                             |
| 7 | Hoja `Resumen` muestra referencias dinámicas (no valores copiados manualmente)                  | Cambiar un dato fuente actualiza la tabla de resumen           |
| 8 | Clasificación automática en `B26` es coherente con el MAPE calculado en `B17`                  | "Excelente" cuando MAPE < 10 %, "Pobre" cuando MAPE > 50 %    |

### Prueba de integridad de fórmulas

Ejecuta esta prueba rápida para confirmar que las fórmulas son robustas:

1. En la hoja `PROD-C`, cambia temporalmente el valor de la celda `B8` (Demanda Real período 7) de `850` a `0`.
2. Verifica que la celda `F8` muestra `0` (no `#¡DIV/0!`), confirmando que la función `SI` funciona correctamente.
3. Observa cómo cambia el MAPE en `B17` drásticamente.
4. Restaura el valor `850` con `Ctrl+Z`.

---

## 8. Resolución de Problemas

### Problema 1: La columna de Error Porcentual muestra `#¡DIV/0!` en algunas celdas

**Síntoma:** Algunas celdas de la columna `F` muestran el error `#¡DIV/0!` y la función `PROMEDIO` en `B17` devuelve también un error.

**Causa:** La fórmula de error porcentual fue ingresada como `=ABS(B2-C2)/B2*100` sin la protección de la función `SI`. Cuando `B2 = 0` (demanda real cero), Excel intenta dividir entre cero.

**Solución:**
1. Selecciona la celda `F2`.
2. Reemplaza la fórmula actual con:
   ```excel
   =SI(B2=0, 0, ABS(B2-C2)/B2*100)
   ```
3. Copia la fórmula corregida hacia abajo hasta `F13`.
4. Si el error persiste en `B17`, verifica que no queden celdas con `#¡DIV/0!` en el rango `F2:F13`. Puedes usar `Ctrl+F` → buscar `#¡DIV/0!` para localizarlas rápidamente.

---

### Problema 2: Las referencias entre hojas en la hoja `Resumen` muestran `#¡REF!` o valores incorrectos

**Síntoma:** Las celdas `B4:D7` en la hoja `Resumen` muestran `#¡REF!`, `0` o valores que no coinciden con los calculados en las hojas de producto.

**Causa:** Las hojas de producto fueron renombradas con espacios o caracteres especiales (ej. `PROD A` en lugar de `PROD-A`), o las fórmulas de referencia apuntan a celdas incorrectas porque la estructura de la hoja de producto fue modificada.

**Solución:**
1. Verifica los nombres exactos de las hojas haciendo clic derecho en cada pestaña. Deben ser exactamente `PROD-A`, `PROD-B`, `PROD-C` (con guion, sin espacios).
2. Si los nombres son correctos, ve a la hoja `Resumen`, selecciona `B4` y verifica que la fórmula sea `='PROD-A'!B16`. Si apunta a otra celda, corrígela manualmente.
3. Si renombraste las hojas después de ingresar las fórmulas, Excel debería haber actualizado las referencias automáticamente. Si no lo hizo, ingresa nuevamente las fórmulas de referencia siguiendo las instrucciones del Paso 5.
4. Para verificar que `B16` en `PROD-A` contiene el MAE (y no otra métrica), ve a esa hoja y confirma que `A16` dice `MAE`.

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza las siguientes acciones:

1. **Guarda el archivo final** con `Ctrl+S`. Confirma que el nombre del archivo incluye tu nombre: `Lab04_Errores_[TuNombre].xlsx`.

2. **Verifica la estructura del libro:** El archivo debe contener exactamente 4 hojas: `PROD-A`, `PROD-B`, `PROD-C` y `Resumen`.

3. **Elimina valores temporales de prueba:** Si realizaste la prueba de integridad del Paso 7 (cambiar demanda a 0), confirma que todos los valores originales están restaurados. El período 7 de `PROD-C` debe mostrar `850`.

4. **Exporta la hoja Resumen como imagen de referencia** (opcional, recomendado para el portafolio del curso):
   - Selecciona el rango `A1:G8` en la hoja `Resumen`.
   - Copia con `Ctrl+C`.
   - Abre Paint o Word y pega con `Ctrl+V` para guardar como imagen de referencia.

5. **Cierra Excel** o mantén el archivo abierto si la siguiente práctica lo requiere como insumo.

> **Nota de secuencialidad:** Este archivo (`Lab04_Errores_[TuNombre].xlsx`) será utilizado como referencia en las Prácticas 20 y 21, donde se compararán múltiples modelos de pronóstico usando estas mismas métricas. Asegúrate de conservarlo en tu carpeta de trabajo.

---

## 10. Resumen y Recursos

### Síntesis de la Práctica

En esta práctica aplicaste el ciclo completo de cálculo de errores de pronóstico en Excel, desde el error fila por fila hasta las métricas agregadas MAE, MAPE y RMSE. Los hallazgos clave que debes llevar a tu práctica logística son:

| Concepto                        | Lo que aprendiste                                                                                   |
|---------------------------------|-----------------------------------------------------------------------------------------------------|
| **Error Absoluto vs. con signo**| Los errores con signo se cancelan entre sí y ocultan la magnitud real del error. Siempre trabaja con `ABS()`. |
| **MAE**                         | Expresa el error promedio en unidades del producto. Directamente interpretable en términos de inventario. |
| **MAPE**                        | Permite comparar la precisión entre productos con distintas escalas de demanda. Sensible a outliers. |
| **RMSE**                        | Penaliza errores grandes. Útil para detectar períodos problemáticos. Siempre ≥ MAE.                 |
| **Outliers y MAPE**             | Un único período extraordinario puede elevar el MAPE más de 60 puntos porcentuales, cambiando la clasificación del modelo. Siempre investiga la causa antes de tomar decisiones. |
| **Clasificación del MAPE**      | < 10 % = Excelente, 10–20 % = Bueno, 20–50 % = Razonable, > 50 % = Pobre.                         |

### Conexión con Decisiones Logísticas

Recuerda la pregunta central del curso: **¿Cómo impacta este análisis en una decisión operativa de logística?**

- Un **MAE de 70 unidades** en un producto con demanda promedio de 180 unidades significa que, en promedio, tu inventario podría estar sobredimensionado o subdimensionado en 70 unidades cada período → costo de almacenamiento innecesario o quiebre de stock.
- Un **MAPE del 40 %** en un SKU de alta rotación implica que el safety stock calculado con ese pronóstico es insuficiente para garantizar un nivel de servicio del 95 %. Necesitarás buffers adicionales o un modelo de pronóstico más preciso.
- Un **outlier no investigado** puede llevar a aumentar el safety stock de todo el año basándose en un evento irrepetible, inmovilizando capital innecesariamente.

### Recursos de Referencia

| Recurso                                                                                                    | Tipo          |
|------------------------------------------------------------------------------------------------------------|---------------|
| [Hyndman & Athanasopoulos — *Forecasting: Principles and Practice*, Sección 5.8](https://otexts.com/fpp3/accuracy.html) | Libro digital gratuito |
| [Microsoft Support — Función ABS](https://support.microsoft.com/es-es/office/abs-funci%C3%B3n-abs-3420200f-5628-4e8c-99da-c99d7c87713c) | Documentación oficial |
| [Microsoft Support — Función SI](https://support.microsoft.com/es-es/office/si-funci%C3%B3n-si-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) | Documentación oficial |
| [Microsoft Support — Función PROMEDIO](https://support.microsoft.com/es-es/office/promedio-funci%C3%B3n-promedio-047bac88-d466-426c-a32b-8f33eb960cf6) | Documentación oficial |
| APICS/ASCM — Diccionario de términos de cadena de suministro                                               | Referencia profesional |

---

> **Próxima práctica (20):** Utilizarás las métricas calculadas aquí para comparar dos modelos de pronóstico (promedio móvil vs. suavización exponencial) sobre el mismo dataset, seleccionando el modelo óptimo con base en MAE y MAPE. Conserva tu archivo `Lab04_Errores_[TuNombre].xlsx` como insumo obligatorio.

---

# Práctica 20 — Comparación del Desempeño de Distintos Métodos de Pronóstico

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Alta |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 4 — Métricas de Error y Selección de Modelos de Pronóstico |
| **Práctica número** | 20 |
| **Herramientas principales** | Microsoft Excel · IA Generativa (Copilot / ChatGPT / Gemini) |
| **Archivo de entrada** | `Practica20_Comparacion_Metodos.xlsx` (proporcionado por el instructor) |
| **Archivo de salida** | `Practica20_Matriz_Decision_[TuNombre].xlsx` |

---

## 2. Descripción General

En esta práctica integrarás todo el ciclo de evaluación de modelos de pronóstico aprendido en las prácticas anteriores. Aplicarás los tres métodos estudiados —Promedio Móvil Simple (PMS-3), Promedio Móvil Ponderado (PMP) y Suavización Exponencial Simple (SES)— sobre cuatro productos con perfiles de demanda distintos, calcularás el MAE y MAPE para cada combinación producto-método, y construirás una **matriz de decisión** que sistematiza la selección del mejor modelo según el contexto logístico. Finalmente, validarás tus conclusiones con una herramienta de IA Generativa y reflexionarás sobre cuándo el juicio experto debe complementar el análisis cuantitativo.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Calcular MAE y MAPE para los tres métodos de pronóstico (PMS, PMP, SES) sobre cuatro productos con perfiles de demanda distintos, construyendo una matriz de comparación 4×3.
- [ ] Aplicar formato condicional de escala de colores (verde-amarillo-rojo) a la matriz de MAPE para identificar visualmente las mejores y peores combinaciones producto-método.
- [ ] Construir una matriz de decisión final que integre el perfil de demanda (CV), el mejor método identificado y la recomendación operativa logística para cada producto.
- [ ] Usar una herramienta de IA Generativa para validar las recomendaciones derivadas del análisis cuantitativo y documentar convergencias o divergencias con sus resultados.
- [ ] Justificar operativamente la selección del modelo de pronóstico con menor error en función del impacto sobre inventario, nivel de servicio y planificación de picking.

---

## 4. Prerrequisitos

### Conocimientos previos
- Haber completado las **Prácticas 13, 14 y 15** (construcción de PMS-3, PMP con pesos 0.5/0.3/0.2 y SES con α=0.3 en Excel).
- Haber completado la **Práctica 19** (cálculo de MAE y MAPE para un único método).
- Comprensión del concepto de **error absoluto** (Lección 4.1): diferencia entre error con signo y error absoluto, función `ABS()` en Excel, e interpretación logística del error (sobreinventario vs. quiebre de stock).
- Comprensión de los indicadores estadísticos básicos: promedio, desviación estándar y coeficiente de variación (CV) para clasificar perfiles de demanda (Prácticas 16-18).

### Acceso y archivos necesarios
- Archivo `Practica20_Comparacion_Metodos.xlsx` distribuido por el instructor (contiene la hoja `Datos_Base` con 24 meses de historial para 4 SKUs).
- Acceso activo a una herramienta de IA Generativa: Microsoft Copilot (`copilot.microsoft.com`), ChatGPT (`chat.openai.com`) o Google Gemini (`gemini.google.com`).
- Microsoft Excel 2016 o superior (se requiere `PRONOSTICO.LINEAL`; funciones `MIN`, `INDICE`, `COINCIDIR` disponibles desde Excel 2010).

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1366×768 | 1920×1080 |
| Conexión a internet | 5 Mbps | 10 Mbps o superior |

### Software requerido

| Software | Versión mínima | Notas |
|---|---|---|
| Microsoft Excel | 2016 / Microsoft 365 | Funciones `ABS`, `PROMEDIO`, `MIN`, `INDICE`, `COINCIDIR` |
| Navegador web | Chrome / Edge (2024) | Para acceso a IA Generativa |
| IA Generativa | Copilot / ChatGPT / Gemini | Cuenta gratuita es suficiente |

### Configuración inicial (antes de comenzar)

1. Abre el archivo `Practica20_Comparacion_Metodos.xlsx` desde la carpeta de materiales del curso.
2. Verifica que el archivo contiene las siguientes hojas:
   - `Datos_Base` — historial de demanda de 4 SKUs (24 meses: M1 a M24).
   - `Calculos_PMS`, `Calculos_PMP`, `Calculos_SES` — hojas de pronósticos ya calculados (traídos de prácticas anteriores o precalculados por el instructor como checkpoint).
   - `Practica20_Trabajo` — hoja en blanco donde construirás la matriz de comparación.
3. Abre tu herramienta de IA Generativa en el navegador y verifica que puedas enviar mensajes.
4. Guarda una copia del archivo con tu nombre: `Practica20_Matriz_Decision_[TuNombre].xlsx`.

> **Nota del instructor:** Si algún participante no completó las Prácticas 13-15 y 19, debe utilizar el archivo checkpoint `Practica20_CHECKPOINT_Completo.xlsx` que contiene todos los pronósticos precalculados. Esto garantiza que ningún participante quede bloqueado en esta práctica por dependencias anteriores.

---

## 6. Pasos del Laboratorio

> **Orientación temporal:** Esta práctica tiene 12 minutos de duración. La distribución sugerida es: Pasos 1-2 (3 min) → Pasos 3-4 (4 min) → Pasos 5-6 (3 min) → Paso 7 (2 min). Mantén el ritmo y usa los archivos checkpoint si es necesario.

---

### Paso 1: Revisar el Dataset y Perfiles de Demanda

**Objetivo:** Familiarizarte con los cuatro productos y sus perfiles de demanda antes de construir la matriz de comparación.

#### Instrucciones

1. Ve a la hoja **`Datos_Base`** del archivo.

2. Revisa la tabla de historial. Debe tener esta estructura (los valores son ilustrativos; usa los del archivo real):

   | Mes | SKU-A (Estable) | SKU-B (Tendencia) | SKU-C (Estacional) | SKU-D (Errático) |
   |-----|----------------|------------------|-------------------|-----------------|
   | M1  | 520 | 310 | 180 | 95 |
   | M2  | 535 | 325 | 210 | 340 |
   | M3  | 510 | 338 | 290 | 112 |
   | ... | ... | ... | ... | ... |
   | M24 | 528 | 598 | 175 | 88 |

3. Localiza la fila de **Coeficiente de Variación (CV)** que ya debe estar calculada al final de la tabla (traída de la Práctica 16). Si no está, calcula el CV para cada SKU con la fórmula:
   ```
   =DESVEST(rango_demanda)/PROMEDIO(rango_demanda)
   ```

4. Registra mentalmente (o anota en papel) los perfiles:
   - **SKU-A:** CV < 0.20 → Demanda estable
   - **SKU-B:** CV < 0.30 con tendencia creciente → Demanda con tendencia
   - **SKU-C:** Patrón repetitivo anual → Demanda estacional
   - **SKU-D:** CV > 0.50 → Demanda errática

5. Recuerda el principio de la Lección 4.1: el error absoluto no distingue si nos equivocamos por exceso o defecto. Para SKU-A (estable), esperamos errores pequeños en todos los métodos. Para SKU-D (errático), esperamos errores grandes en todos. La pregunta es: **¿cuál método se equivoca menos en cada caso?**

#### Resultado esperado
Tienes claridad sobre los cuatro perfiles de demanda y el CV de cada producto. Estás listo para construir la matriz de comparación.

#### Verificación
- El CV de SKU-A debe ser el más bajo (< 0.20).
- El CV de SKU-D debe ser el más alto (> 0.50).
- Si los valores de CV no coinciden con estos rangos, consulta con el instructor antes de continuar.

---

### Paso 2: Construir la Estructura de la Matriz de Comparación en Excel

**Objetivo:** Crear la estructura visual de la matriz 4×3 en la hoja de trabajo antes de ingresar fórmulas.

#### Instrucciones

1. Ve a la hoja **`Practica20_Trabajo`**.

2. A partir de la celda **A1**, construye la siguiente estructura de encabezados. Usa **negrita** (Ctrl+N) para todos los encabezados:

   ```
   Fila 1:  [A1] MATRIZ DE COMPARACIÓN DE MÉTODOS DE PRONÓSTICO
   Fila 3:  [A3] Producto  [B3] Perfil  [C3] CV
             [D3] MAE_PMS  [E3] MAE_PMP  [F3] MAE_SES
             [G3] MAPE_PMS [H3] MAPE_PMP [I3] MAPE_SES
   Fila 4:  [A4] SKU-A    [B4] Estable
   Fila 5:  [A5] SKU-B    [B5] Tendencia
   Fila 6:  [A6] SKU-C    [B6] Estacional
   Fila 7:  [A7] SKU-D    [B7] Errático
   ```

3. En la columna **C** (CV), ingresa referencias a los valores de CV calculados en la hoja `Datos_Base`:
   ```
   =Datos_Base!C27   (ajusta la referencia a la celda donde está el CV de SKU-A)
   ```
   Repite para los cuatro productos.

4. Debajo de la primera matriz (a partir de la fila 10), crea una segunda tabla de **RANKING**:

   ```
   Fila 10: [A10] TABLA DE RANKING - MEJOR MÉTODO POR PRODUCTO
   Fila 12: [A12] Producto  [B12] Mejor_MAE  [C12] Valor_MAE  [D12] Mejor_MAPE  [E12] Valor_MAPE
   Fila 13: [A13] SKU-A
   Fila 14: [A14] SKU-B
   Fila 15: [A15] SKU-C
   Fila 16: [A16] SKU-D
   ```

5. Más abajo (fila 19), crea la estructura de la **MATRIZ DE DECISIÓN FINAL**:

   ```
   Fila 19: [A19] MATRIZ DE DECISIÓN FINAL
   Fila 21: [A21] Producto  [B21] CV  [C21] Perfil_Demanda  [D21] Mejor_Método
             [E21] MAPE_Mejor  [F21] Recomendación_Operativa
   Filas 22-25: SKU-A a SKU-D (dejar en blanco por ahora)
   ```

6. Aplica color de fondo gris claro (RGB: 242, 242, 242) a las filas de encabezado (3, 12, 21) para mejorar la legibilidad.

#### Resultado esperado
La hoja `Practica20_Trabajo` tiene tres tablas estructuradas: Matriz de Comparación, Tabla de Ranking y Matriz de Decisión Final, todas con encabezados en negrita y sin fórmulas aún.

#### Verificación
- Cuenta que tienes exactamente 9 columnas de métricas en la Matriz de Comparación (D a I = 6 columnas de métricas + A, B, C = 3 columnas de identificación).
- Las tres tablas deben ser visualmente distinguibles.

---

### Paso 3: Calcular MAE y MAPE para los Tres Métodos

**Objetivo:** Completar la Matriz de Comparación con los valores de MAE y MAPE para cada combinación producto-método, aplicando la lógica del error absoluto de la Lección 4.1.

#### Instrucciones

> **Contexto:** Los pronósticos de PMS, PMP y SES ya están calculados en las hojas `Calculos_PMS`, `Calculos_PMP` y `Calculos_SES`. Cada hoja tiene columnas de demanda real y pronóstico para los 4 productos. El período de evaluación usa los **meses M4 a M24** (21 períodos evaluables), ya que los primeros meses se usan para inicializar los modelos.

1. En la celda **D4** (MAE_PMS para SKU-A), ingresa la siguiente fórmula:
   ```excel
   =PROMEDIO(ABS(Calculos_PMS!C4:C24 - Calculos_PMS!D4:D24))
   ```
   > ⚠️ **Importante:** Esta es una **fórmula matricial**. En Excel 365/2019 se ingresa normalmente con Enter. En Excel 2016, debes presionar **Ctrl+Shift+Enter** para confirmarla como fórmula de matriz. Verás llaves `{}` alrededor de la fórmula si se ingresó correctamente en versiones antiguas.

   > **Alternativa sin fórmula matricial** (compatible con todas las versiones):
   > Si tienes problemas con la fórmula matricial, usa una columna auxiliar en la hoja `Calculos_PMS` con `=ABS(Real-Pronostico)` fila por fila, y luego referencia el `PROMEDIO` de esa columna auxiliar.

2. Ajusta las referencias de columna para cada producto en la hoja `Calculos_PMS`. La estructura típica de esa hoja es:

   | Col | Contenido |
   |-----|-----------|
   | A | Mes |
   | B | Demanda Real SKU-A |
   | C | PMS SKU-A |
   | D | Demanda Real SKU-B |
   | E | PMS SKU-B |
   | F | Demanda Real SKU-C |
   | G | PMS SKU-C |
   | H | Demanda Real SKU-D |
   | I | PMS SKU-D |

   Ajusta las referencias según la estructura real de tu archivo.

3. Completa **todas las celdas de MAE** (D4:F7) referenciando las hojas correctas:
   - Columna D (`MAE_PMS`): referencia a `Calculos_PMS`
   - Columna E (`MAE_PMP`): referencia a `Calculos_PMP`
   - Columna F (`MAE_SES`): referencia a `Calculos_SES`

4. Para calcular el **MAPE** (columnas G, H, I), usa la siguiente fórmula en **G4** (MAPE_PMS para SKU-A):
   ```excel
   =PROMEDIO(ABS((Calculos_PMS!B4:B24 - Calculos_PMS!C4:C24)/Calculos_PMS!B4:B24))*100
   ```
   Esto devuelve el MAPE como **porcentaje** (por ejemplo, 8.3 significa 8.3%).

   > ⚠️ Si algún valor de demanda real es cero, esta fórmula dará error de división por cero. En ese caso, usa:
   > ```excel
   > =PROMEDIO(SI(Calculos_PMS!B4:B24<>0, ABS((Calculos_PMS!B4:B24-Calculos_PMS!C4:C24)/Calculos_PMS!B4:B24)*100, ""))
   > ```
   > (Confirmar con Ctrl+Shift+Enter en Excel 2016)

5. Completa **todas las celdas de MAPE** (G4:I7) para los tres métodos.

6. Formatea las columnas D, E, F (MAE) con **1 decimal** y las columnas G, H, I (MAPE) con **1 decimal y símbolo %**:
   - Selecciona G4:I7 → Formato de celdas → Número → Personalizado → `0.0"%"`

#### Resultado esperado
La Matriz de Comparación tiene 24 valores numéricos (4 productos × 3 métodos × 2 métricas). Los valores de MAPE para SKU-A deben ser los más bajos (demanda estable → todos los métodos funcionan bien). Los valores de MAPE para SKU-D deben ser los más altos (demanda errática → todos los métodos tienen dificultades).

#### Verificación
- Ninguna celda debe mostrar `#¡DIV/0!`, `#¡VALOR!` o `#¡REF!`.
- Los valores de MAE deben estar en las mismas unidades que la demanda (cajas, unidades, etc.).
- Los valores de MAPE deben estar entre 0% y 100% para productos normales (SKU-D podría superar el 50%).
- Verifica que el MAE de PMS para SKU-A sea razonablemente bajo (< 30 unidades si la demanda promedio es ~520).

---

### Paso 4: Aplicar Formato Condicional de Escala de Colores al MAPE

**Objetivo:** Crear una representación visual inmediata de las mejores y peores combinaciones producto-método usando formato condicional.

#### Instrucciones

1. Selecciona el rango **G4:I7** (toda la sección de MAPE en la Matriz de Comparación).

2. Ve a la pestaña **Inicio** → grupo **Estilos** → **Formato condicional** → **Escalas de color**.

3. Selecciona la escala **Verde-Amarillo-Rojo** (la primera opción de la galería de 3 colores, donde el verde representa el valor mínimo).

   > **Verificación visual:** El verde debe aparecer en los valores de MAPE más bajos (mejor desempeño) y el rojo en los más altos (peor desempeño). Confirma que la lógica es: **verde = menor error = mejor método**.

4. Para personalizar los umbrales (opcional, recomendado):
   - Con el rango G4:I7 seleccionado → **Formato condicional** → **Administrar reglas** → **Editar regla**.
   - Cambia el tipo a **Número** en lugar de Porcentaje:
     - Mínimo (verde): `0`
     - Punto medio (amarillo): `15` (15% de MAPE es un umbral razonable en logística)
     - Máximo (rojo): `40` (40% o más indica modelo muy impreciso)
   - Haz clic en **Aceptar**.

5. Observa el patrón visual resultante. Deberías ver:
   - La fila de **SKU-A** predominantemente verde (demanda estable → todos los métodos tienen MAPE bajo).
   - La fila de **SKU-D** predominantemente roja o naranja (demanda errática → todos los métodos tienen MAPE alto).
   - Las filas de **SKU-B y SKU-C** con variación entre métodos (aquí es donde la selección del método importa más).

6. Crea un gráfico de barras agrupadas para comparar visualmente el MAPE:
   - Selecciona el rango **A3:A7** (productos) y luego con **Ctrl** selecciona también **G3:I7** (MAPE de los tres métodos).
   - Ve a **Insertar** → **Gráfico de barras** → **Barras agrupadas** (o Columnas agrupadas).
   - Título del gráfico: `Comparación de MAPE por Producto y Método de Pronóstico`
   - Etiquetas del eje horizontal: SKU-A, SKU-B, SKU-C, SKU-D.
   - Leyenda: PMS-3, PMP, SES.
   - Coloca el gráfico debajo de la Matriz de Decisión Final (aproximadamente en la fila 30).

#### Resultado esperado
La sección de MAPE de la matriz muestra una escala de colores verde-amarillo-rojo que permite identificar de un vistazo las mejores combinaciones. El gráfico de barras agrupadas muestra 4 grupos (uno por producto) con 3 barras cada uno (una por método).

#### Verificación
- El color verde debe coincidir con los valores numéricamente más bajos en el rango G4:I7.
- El gráfico debe tener exactamente 12 barras (4 productos × 3 métodos).
- El título y la leyenda del gráfico deben ser legibles.

---

### Paso 5: Construir la Tabla de Ranking con Fórmulas Automáticas

**Objetivo:** Usar las funciones `MIN`, `INDICE` y `COINCIDIR` para identificar automáticamente el mejor método para cada producto según MAE y MAPE.

#### Instrucciones

1. Ve a la **Tabla de Ranking** (filas 10-16 de la hoja `Practica20_Trabajo`).

2. Define un rango con nombre para los encabezados de métodos. En una celda auxiliar temporal (por ejemplo, **K3:M3**), escribe:
   ```
   K3: PMS-3    L3: PMP    M3: SES
   ```
   Estos serán los valores que devolverá `INDICE/COINCIDIR`.

3. En la celda **C13** (Valor_MAE del mejor método para SKU-A), ingresa:
   ```excel
   =MIN(D4:F4)
   ```
   Esto devuelve el valor más bajo de MAE entre los tres métodos para SKU-A.

4. En la celda **B13** (nombre del Mejor_MAE para SKU-A), ingresa:
   ```excel
   =INDICE($K$3:$M$3,COINCIDIR(MIN(D4:F4),D4:F4,0))
   ```
   Esta fórmula:
   - `MIN(D4:F4)` → encuentra el MAE mínimo.
   - `COINCIDIR(...)` → encuentra la posición (1, 2 o 3) de ese mínimo en el rango D4:F4.
   - `INDICE($K$3:$M$3,...)` → devuelve el nombre del método en esa posición.

5. En la celda **E13** (Valor_MAPE del mejor método para SKU-A):
   ```excel
   =MIN(G4:I4)
   ```

6. En la celda **D13** (nombre del Mejor_MAPE para SKU-A):
   ```excel
   =INDICE($K$3:$M$3,COINCIDIR(MIN(G4:I4),G4:I4,0))
   ```

7. Copia las fórmulas de las celdas B13:E13 hacia abajo para las filas 14, 15 y 16 (SKU-B, SKU-C, SKU-D). Asegúrate de que las referencias de fila cambien correctamente (D4→D5→D6→D7, etc.).

8. Aplica **negrita** a los valores en las columnas C y E (valores mínimos de MAE y MAPE) para destacarlos visualmente.

> **Interpretación logística:** Si para un producto el mejor método según MAE es PMS pero según MAPE es SES, existe un **conflicto de criterios**. En ese caso, la regla general es: si la demanda del producto es alta en volumen (como SKU-A con ~520 unidades), prioriza MAE (el error en unidades impacta más el inventario). Si la demanda es baja en volumen (como SKU-D con ~100 unidades), prioriza MAPE (el porcentaje de error es más informativo).

#### Resultado esperado
La Tabla de Ranking muestra automáticamente el nombre del mejor método y su valor de error para cada producto, tanto en MAE como en MAPE. Las fórmulas `INDICE/COINCIDIR` devuelven "PMS-3", "PMP" o "SES" según corresponda.

#### Verificación
- Cambia manualmente un valor en la Matriz de Comparación (por ejemplo, modifica G4) y verifica que la Tabla de Ranking se actualiza automáticamente. Luego deshaz el cambio (Ctrl+Z).
- Los valores en las columnas C y E de la Tabla de Ranking deben coincidir exactamente con el valor mínimo visible en la Matriz de Comparación para esa fila.

---

### Paso 6: Completar la Matriz de Decisión Final

**Objetivo:** Integrar el perfil de demanda, el mejor método y la recomendación operativa logística en una tabla de decisión ejecutiva.

#### Instrucciones

1. Ve a la **Matriz de Decisión Final** (filas 19-25 de la hoja `Practica20_Trabajo`).

2. Completa las columnas A, B y C con referencias a los datos ya calculados:
   ```excel
   A22: =A4          (nombre del producto: SKU-A)
   B22: =C4          (CV de SKU-A)
   C22: =B4          (Perfil: Estable)
   ```
   Repite para las filas 23, 24 y 25 (SKU-B, SKU-C, SKU-D).

3. En la columna **D** (Mejor_Método), referencia el resultado de la Tabla de Ranking. Si el mejor método según MAPE y MAE coincide, usa ese. Si difieren, usa el criterio de volumen descrito en el paso anterior:
   ```excel
   D22: =D13    (mejor método según MAPE para SKU-A)
   ```
   Ajusta manualmente si hay conflicto de criterios (en ese caso, escribe el método elegido como texto y agrega una nota en la celda con Shift+F2 explicando el criterio usado).

4. En la columna **E** (MAPE_Mejor):
   ```excel
   E22: =E13    (valor de MAPE del mejor método para SKU-A)
   ```

5. En la columna **F** (Recomendación_Operativa), escribe manualmente la recomendación logística para cada producto. Usa las siguientes guías como referencia y adáptalas a los resultados reales que obtuviste:

   | Producto | Guía de recomendación |
   |---|---|
   | **SKU-A** (Estable) | "Cualquier método es adecuado. Se recomienda PMS-3 por simplicidad operativa. Stock de seguridad bajo (1-2 semanas). Picking planificable con alta confianza." |
   | **SKU-B** (Tendencia) | "Evitar PMS-3 (subestima tendencia). Preferir [método con menor MAPE]. Revisar pronóstico mensualmente. Ajustar punto de reorden al alza cada trimestre." |
   | **SKU-C** (Estacional) | "Incorporar índices estacionales al método seleccionado. El método base [X] tiene menor error promedio, pero requiere ajuste estacional para picos. Planificar picking estacional con 3 semanas de anticipación." |
   | **SKU-D** (Errático) | "Ningún método tiene MAPE aceptable (<20%). Considerar revisión periódica del inventario. Stock de seguridad alto (3-4 semanas). Evaluar si el SKU justifica seguimiento individual o si debe gestionarse por excepción." |

6. Aplica **bordes exteriores gruesos** a toda la Matriz de Decisión Final (rango A21:F25) y bordes interiores finos para mejorar la presentación.

7. Aplica el siguiente código de color a la columna E (MAPE_Mejor) con formato condicional manual:
   - MAPE < 10%: fondo verde claro
   - MAPE 10-20%: fondo amarillo claro
   - MAPE > 20%: fondo rojo claro

#### Resultado esperado
La Matriz de Decisión Final es una tabla ejecutiva completa que un gerente de logística podría leer en 30 segundos para tomar decisiones de inventario. Integra datos cuantitativos (CV, MAPE) con recomendaciones cualitativas operativas.

#### Verificación
- La tabla debe tener exactamente 4 filas de datos (una por SKU) y 6 columnas.
- La columna F no debe estar vacía: cada producto debe tener una recomendación operativa específica.
- El código de color de la columna E debe reflejar la calidad del pronóstico: SKU-A en verde, SKU-D probablemente en rojo o amarillo.

---

### Paso 7: Validar Recomendaciones con IA Generativa

**Objetivo:** Usar una herramienta de IA Generativa para contrastar las conclusiones del análisis cuantitativo con el razonamiento de la IA, y documentar convergencias o divergencias.

#### Instrucciones

1. Abre tu herramienta de IA Generativa (Copilot, ChatGPT o Gemini) en el navegador.

2. Construye el siguiente **prompt estructurado**. Reemplaza los valores entre corchetes con los resultados reales de tu matriz:

   ```
   Actúa como un experto en planificación de demanda y logística de inventarios.
   
   Tengo los siguientes resultados de comparación de métodos de pronóstico para 4 productos:
   
   - SKU-A (Demanda Estable, CV=[valor]): PMS-3 MAPE=[X]%, PMP MAPE=[Y]%, SES MAPE=[Z]%
   - SKU-B (Demanda con Tendencia, CV=[valor]): PMS-3 MAPE=[X]%, PMP MAPE=[Y]%, SES MAPE=[Z]%
   - SKU-C (Demanda Estacional, CV=[valor]): PMS-3 MAPE=[X]%, PMP MAPE=[Y]%, SES MAPE=[Z]%
   - SKU-D (Demanda Errática, CV=[valor]): PMS-3 MAPE=[X]%, PMP MAPE=[Y]%, SES MAPE=[Z]%
   
   Mis conclusiones son:
   1. Para SKU-A recomiendo [tu método] porque [tu razón].
   2. Para SKU-B recomiendo [tu método] porque [tu razón].
   3. Para SKU-C recomiendo [tu método] porque [tu razón].
   4. Para SKU-D ningún método tiene MAPE aceptable; recomiendo gestión por excepción.
   
   Por favor:
   a) Evalúa si mis recomendaciones son correctas desde una perspectiva de gestión logística.
   b) Indica si hay alguna recomendación con la que no estés de acuerdo y explica por qué.
   c) Señala las limitaciones de usar solo MAE y MAPE para seleccionar un método de pronóstico.
   d) ¿Cuándo debería el juicio experto del planificador complementar o incluso contradecir el análisis cuantitativo?
   ```

3. Envía el prompt y espera la respuesta de la IA.

4. Lee la respuesta y en la hoja `Practica20_Trabajo`, crea una pequeña tabla de validación a partir de la celda **A38**:

   ```
   A38: VALIDACIÓN CON IA GENERATIVA
   A40: Herramienta utilizada:    [nombre de la IA]
   A41: Fecha/hora:               [fecha y hora]
   A43: CONVERGENCIAS (IA coincide con mis conclusiones):
   A44: [escribe aquí los puntos donde la IA estuvo de acuerdo]
   A47: DIVERGENCIAS (IA difiere de mis conclusiones):
   A48: [escribe aquí los puntos donde la IA discrepó y la razón]
   A51: LIMITACIONES IDENTIFICADAS POR LA IA:
   A52: [lista las limitaciones mencionadas]
   A55: REFLEXIÓN PERSONAL (2-3 líneas):
   A56: [escribe cuándo usarías el juicio experto sobre el modelo cuantitativo]
   ```

5. Completa la tabla con base en la respuesta real que recibiste de la IA.

> **Si la IA no está disponible:** El instructor tiene capturas de pantalla de respuestas de referencia. Solicítalas y úsalas para completar la tabla de validación de todas formas. El ejercicio de reflexión es igualmente válido con una respuesta de referencia.

#### Resultado esperado
La hoja `Practica20_Trabajo` tiene una sección de validación con IA que documenta convergencias, divergencias y limitaciones identificadas. La reflexión personal conecta el análisis cuantitativo con el juicio experto en contexto logístico.

#### Verificación
- La tabla de validación no debe estar vacía. Debe tener al menos 2 convergencias y 1 limitación identificada.
- La reflexión personal debe mencionar al menos un escenario logístico concreto (por ejemplo: "usaría juicio experto cuando hay una promoción no planificada que distorsiona el historial").

---

## 7. Validación y Pruebas del Laboratorio

Al finalizar todos los pasos, realiza las siguientes verificaciones de integridad:

### Lista de verificación final

| # | Verificación | ✓/✗ |
|---|---|---|
| 1 | La Matriz de Comparación tiene 24 valores numéricos (4×3×2 métricas) sin errores de fórmula | |
| 2 | El formato condicional verde-amarillo-rojo está aplicado correctamente al rango G4:I7 (verde = menor MAPE) | |
| 3 | El gráfico de barras agrupadas tiene 12 barras (4 productos × 3 métodos) con título y leyenda | |
| 4 | Las fórmulas `INDICE/COINCIDIR` en la Tabla de Ranking devuelven "PMS-3", "PMP" o "SES" (no números) | |
| 5 | La Tabla de Ranking se actualiza automáticamente al modificar un valor de la Matriz de Comparación | |
| 6 | La Matriz de Decisión Final tiene recomendaciones operativas en la columna F para los 4 productos | |
| 7 | La sección de validación con IA está completa con convergencias, divergencias y reflexión personal | |
| 8 | El archivo está guardado como `Practica20_Matriz_Decision_[TuNombre].xlsx` | |

### Prueba de coherencia lógica

Responde mentalmente las siguientes preguntas. Si no puedes responderlas, revisa el paso correspondiente:

1. **¿Para qué producto tiene menor diferencia de MAPE entre los tres métodos?** (Debería ser SKU-A, la demanda estable. Si no es así, revisa las fórmulas de MAPE.)
2. **¿El método con menor MAPE para SKU-D tiene un MAPE aceptable (<20%)?** (Probablemente no. Esto confirma que para demanda errática ningún método simple es suficiente.)
3. **¿La IA coincidió con tu recomendación para SKU-B (tendencia)?** (Si la IA discrepó, ¿tiene sentido su argumento desde una perspectiva logística?)

---

## 8. Solución de Problemas

### Problema 1: La fórmula MAPE devuelve `#¡DIV/0!` o `#¡VALOR!`

**Síntoma:** Una o más celdas en el rango G4:I7 muestran `#¡DIV/0!` o `#¡VALOR!` en lugar de un valor numérico.

**Causa probable:** El dataset contiene al menos un período con demanda real igual a cero (división por cero en el cálculo del MAPE), o las referencias de celda en la fórmula apuntan a celdas vacías o con texto.

**Solución:**
1. Ve a la hoja del método correspondiente (`Calculos_PMS`, `Calculos_PMP` o `Calculos_SES`) y busca filas donde la demanda real sea 0 o esté vacía.
2. Si existen ceros, reemplaza la fórmula de MAPE por la versión con `SI()`:
   ```excel
   =PROMEDIO(SI(Calculos_PMS!B4:B24<>0,
     ABS((Calculos_PMS!B4:B24-Calculos_PMS!C4:C24)/Calculos_PMS!B4:B24)*100,
     ""))
   ```
   Confirma con **Ctrl+Shift+Enter** en Excel 2016, o Enter normal en Excel 365.
3. Si el error es `#¡VALOR!`, verifica que las columnas referenciadas contengan solo números. Usa `Ctrl+Fin` en la hoja de cálculos para verificar que no hay filas extra con texto debajo del rango de datos.

---

### Problema 2: La fórmula `INDICE/COINCIDIR` devuelve un número en lugar del nombre del método

**Síntoma:** Las celdas B13:B16 o D13:D16 de la Tabla de Ranking muestran "1", "2" o "3" en lugar de "PMS-3", "PMP" o "SES".

**Causa probable:** El rango de nombres de métodos en K3:M3 está vacío, mal escrito, o la fórmula `INDICE` está referenciando el rango de valores numéricos en lugar del rango de texto con los nombres.

**Solución:**
1. Verifica que las celdas **K3**, **L3** y **M3** contengan exactamente el texto: `PMS-3`, `PMP`, `SES` (sin espacios adicionales, sin tildes, en mayúsculas como los definiste).
2. Revisa la fórmula en B13. Debe ser:
   ```excel
   =INDICE($K$3:$M$3, COINCIDIR(MIN(D4:F4), D4:F4, 0))
   ```
   El primer argumento de `INDICE` debe apuntar a **$K$3:$M$3** (rango de texto con nombres), **no** a D4:F4 (rango de valores numéricos).
3. Si el problema persiste, verifica que el rango `D4:F4` tiene exactamente 3 valores y que `COINCIDIR` está usando el tercer argumento `0` (coincidencia exacta).
4. Como alternativa, puedes usar `ELEGIR` en lugar de `INDICE/COINCIDIR`:
   ```excel
   =ELEGIR(COINCIDIR(MIN(D4:F4),D4:F4,0),"PMS-3","PMP","SES")
   ```

---

## 9. Limpieza y Cierre

### Acciones de cierre del laboratorio

1. **Guarda el archivo final** con el nombre correcto:
   ```
   Practica20_Matriz_Decision_[TuNombre].xlsx
   ```
   Usa **Ctrl+S** o **Archivo → Guardar**.

2. **Elimina las celdas auxiliares** K3:M3 (nombres de métodos usados para `INDICE/COINCIDIR`) **solo si** las fórmulas de la Tabla de Ranking siguen funcionando después de eliminarlas. Si las fórmulas usan referencias directas a esas celdas, no las elimines; en cambio, cambia el color de fuente a blanco para ocultarlas visualmente.

3. **Entrega al instructor** el archivo guardado a través del canal indicado (plataforma del curso, correo electrónico o carpeta compartida).

4. **Cierra la sesión** de la herramienta de IA Generativa si usaste una cuenta personal.

5. **No elimines** el archivo de origen `Practica20_Comparacion_Metodos.xlsx`; podría ser necesario para la Práctica 21.

---

## 10. Resumen y Recursos Adicionales

### Lo que aprendiste en esta práctica

En esta práctica integraste el ciclo completo de evaluación de modelos de pronóstico logístico:

1. **Construiste una matriz de comparación 4×3** que aplica el principio del error absoluto (Lección 4.1) de forma sistemática: para cada combinación producto-método calculaste el MAE (error en unidades, mismo concepto que `|Demanda Real - Pronóstico|`) y el MAPE (error porcentual), obteniendo 24 métricas de evaluación.

2. **Visualizaste los resultados** con formato condicional de escala de colores y un gráfico de barras agrupadas, transformando una tabla de números en información accionable para la toma de decisiones.

3. **Automatizaste la selección del mejor método** usando `INDICE/COINCIDIR`, eliminando la subjetividad en la identificación del modelo óptimo por producto.

4. **Tradujiste el análisis cuantitativo en decisiones operativas** a través de la Matriz de Decisión Final, conectando el CV, el MAPE y el perfil de demanda con recomendaciones concretas de inventario, nivel de servicio y planificación de picking.

5. **Validaste y cuestionaste tus conclusiones** con IA Generativa, practicando el pensamiento crítico sobre las limitaciones de las métricas de error y el rol del juicio experto en forecasting logístico.

### Conclusión metodológica clave

> El método de pronóstico con menor MAPE no siempre es el "mejor" en sentido logístico. Un MAPE de 8% en un producto de 10.000 unidades mensuales representa 800 unidades de error, con un impacto de inventario significativo. Un MAPE de 15% en un producto de 50 unidades mensuales representa solo 7.5 unidades. **La métrica de error es una herramienta, no una sentencia.** El planificador logístico debe interpretarla siempre en el contexto del volumen, el costo de quiebre de stock y el costo de sobreinventario del producto específico.

### Conexión con el ciclo de pronóstico (Baseline + Estacionalidad + Iniciativas)

Los métodos PMS, PMP y SES que evaluaste en esta práctica capturan principalmente el **componente Baseline** de la demanda. Para SKU-C (estacional), el MAPE relativamente alto de todos los métodos confirma que se necesita incorporar el **componente de Estacionalidad** (índices estacionales) para mejorar la precisión. Esta limitación identificada cuantitativamente en esta práctica justifica metodológicamente el uso de modelos más avanzados en las siguientes prácticas del curso.

### Recursos adicionales

| Recurso | Descripción | Acceso |
|---|---|---|
| Hyndman & Athanasopoulos — *Forecasting: Principles and Practice* (3ª ed.) | Secciones 5.8 y 5.9: Evaluación de precisión de pronósticos, MAE y MAPE | [otexts.com/fpp3/accuracy.html](https://otexts.com/fpp3/accuracy.html) |
| Microsoft Support — Función INDICE | Referencia oficial de sintaxis y ejemplos | [support.microsoft.com](https://support.microsoft.com/es-es/office/indice-funci%C3%B3n-indice) |
| Microsoft Support — Función COINCIDIR | Referencia oficial de sintaxis y ejemplos | [support.microsoft.com](https://support.microsoft.com/es-es/office/coincidir-funci%C3%B3n-coincidir) |
| APICS/ASCM — Diccionario de cadena de suministro | Definiciones de MAE, MAPE, Forecast Accuracy | [ascm.org/dictionary](https://www.ascm.org/learning-development/certifications-designations/dictionary/) |
| Checkpoint de la práctica | Archivo con todos los pasos completados (para recuperación) | Solicitar al instructor: `Practica20_CHECKPOINT_Completo.xlsx` |

---

*Práctica 20 completada. La siguiente práctica del módulo aplicará los resultados de esta matriz de decisión para generar recomendaciones automáticas de reabastecimiento asistidas por IA Generativa.*

---

# Práctica 21 — Selección del mejor modelo con menor error y justificación operativa

## Metadatos

| Campo              | Detalle                                                                 |
|--------------------|-------------------------------------------------------------------------|
| **Duración**       | 12 minutos                                                              |
| **Complejidad**    | Media                                                                   |
| **Nivel Bloom**    | Crear (*Create*) — sintetizar métricas de error y producir una decisión operativa documentada |
| **Práctica N°**    | 21 (Módulo 4 — Evaluación de modelos de pronóstico)                     |
| **Archivo base**   | `Practica21_Errores_Modelos.xlsx` (proporcionado por el instructor)     |

---

## Descripción General

En esta práctica el participante calculará las métricas de error **MAE** (*Mean Absolute Error*) y **MAPE** (*Mean Absolute Percentage Error*) para tres modelos de pronóstico previamente construidos —promedio móvil de 3 períodos, suavización exponencial (α = 0.3) y regresión lineal— aplicados a tres SKUs representativos de un almacén de distribución. A partir de los resultados, identificará el modelo con menor error, lo resaltará mediante formato condicional y redactará una justificación operativa que vincule la decisión estadística con impactos reales en inventario, nivel de servicio y costos logísticos. El ejercicio consolida el ciclo completo de evaluación de modelos de forecasting iniciado en las Prácticas 18–20.

---

## Objetivos de Aprendizaje

Al completar esta práctica, el participante será capaz de:

- [ ] Calcular el error absoluto período a período (`|Real − Pronóstico|`) y el error porcentual absoluto (`|Real − Pronóstico| / Real × 100`) para tres modelos de pronóstico usando las funciones `ABS()` y `PROMEDIO()` de Excel.
- [ ] Construir una tabla resumen de métricas MAE y MAPE que permita comparar el desempeño de los tres modelos de forma objetiva.
- [ ] Seleccionar el modelo de pronóstico con menor error y documentar una justificación operativa de 3–5 líneas que considere el tipo de demanda, la magnitud del error aceptable y las consecuencias logísticas de una mala predicción.
- [ ] Aplicar formato condicional en Excel para resaltar visualmente el modelo óptimo en la tabla comparativa.
- [ ] Interpretar el MAE y el MAPE en términos de impacto sobre inventario de seguridad, costos de ruptura de stock y sobreinventario.

---

## Prerrequisitos

### Conocimiento previo
- Haber completado las Prácticas 18, 19 y 20 (construcción de modelos de pronóstico en Excel) **o** contar con el archivo `Practica21_Errores_Modelos.xlsx` pre-cargado con los tres modelos calculados.
- Comprensión del concepto de error absoluto: la diferencia sin signo entre demanda real y pronóstico (`|Real − Pronóstico|`), tal como se estudió en la Lección 4.1.
- Manejo básico de fórmulas Excel: `ABS()`, `PROMEDIO()`, operadores aritméticos y referencias de celda.
- Comprensión conceptual de por qué los errores con signo opuesto se cancelan entre sí y por qué el error absoluto entrega una visión más honesta del desempeño del modelo.

### Acceso y archivos requeridos
- Computadora con **Microsoft Excel 2016 o superior** (Microsoft 365 recomendado).
- Archivo `Practica21_Errores_Modelos.xlsx` abierto y disponible (distribuido por el instructor antes de iniciar).
- No se requiere conexión a internet en esta práctica (las herramientas de IA se utilizan en Prácticas 22–23).

---

## Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente         | Mínimo                        | Recomendado                   |
|--------------------|-------------------------------|-------------------------------|
| Procesador         | Intel Core i5 / AMD Ryzen 5   | Intel Core i7 / AMD Ryzen 7   |
| RAM                | 4 GB                          | 8 GB                          |
| Almacenamiento     | 2 GB libres en disco          | 5 GB libres en disco          |
| Resolución pantalla| 1366 × 768                    | 1920 × 1080                   |

### Software requerido

| Software                  | Versión mínima     | Notas                                              |
|---------------------------|--------------------|----------------------------------------------------|
| Microsoft Excel           | 2016               | Microsoft 365 recomendado; `ABS()` y `PROMEDIO()` disponibles desde Excel 2007 |
| Archivo de práctica       | Versión instructor | `Practica21_Errores_Modelos.xlsx`                  |

### Estructura del archivo de práctica

El archivo `Practica21_Errores_Modelos.xlsx` contiene las siguientes hojas:

| Hoja                    | Contenido                                                                                   |
|-------------------------|---------------------------------------------------------------------------------------------|
| `Datos_SKU`             | Historial de demanda real mensual (12 períodos) para SKU-A, SKU-B y SKU-C                  |
| `Modelo_PromedioMovil`  | Pronósticos del modelo de promedio móvil 3 períodos (columnas por SKU)                      |
| `Modelo_SuavizacionExp` | Pronósticos del modelo de suavización exponencial α = 0.3 (columnas por SKU)               |
| `Modelo_RegresionLineal`| Pronósticos del modelo de regresión lineal (columnas por SKU)                               |
| `Practica21_Trabajo`    | **Hoja de trabajo principal** — aquí el participante construirá todas las tablas de errores  |
| `Referencia_Completada` | Hoja de referencia con resultados esperados (visible solo al finalizar o para el instructor) |

> **Nota de configuración:** Antes de iniciar, verifique que la hoja `Practica21_Trabajo` esté activa y que las columnas A–D contengan los encabezados pre-cargados: `Período`, `Demanda_Real`, `PM3_Pronostico`, `SE_Pronostico`, `RL_Pronostico`.

---

## Procedimiento Paso a Paso

---

### Paso 1 — Reconocer la estructura de datos y revisar los pronósticos pre-cargados

**Objetivo:** Familiarizarse con los datos disponibles antes de construir las columnas de error, confirmando que los tres modelos de pronóstico están correctamente cargados en la hoja de trabajo.

#### Instrucciones

1. Abra el archivo `Practica21_Errores_Modelos.xlsx` y navegue a la hoja **`Practica21_Trabajo`**.

2. Identifique la tabla principal que debe lucir similar a la siguiente estructura (filas 1–14, columnas A–E):

   | Fila | A (Período) | B (Demanda_Real) | C (PM3_Pronostico) | D (SE_Pronostico) | E (RL_Pronostico) |
   |------|-------------|-----------------|--------------------|--------------------|-------------------|
   | 1    | **Período** | **Demanda_Real**| **PM3_Pronostico** | **SE_Pronostico**  | **RL_Pronostico** |
   | 2    | Mes 1       | 540             | —                  | 510                | 498               |
   | 3    | Mes 2       | 480             | —                  | 513                | 512               |
   | 4    | Mes 3       | 610             | —                  | 504                | 526               |
   | 5    | Mes 4       | 495             | 543                | 540                | 540               |
   | ...  | ...         | ...             | ...                | ...                | ...               |
   | 14   | Mes 13      | —               | *(pronóstico)*     | *(pronóstico)*     | *(pronóstico)*    |

   > **Importante:** El modelo de promedio móvil de 3 períodos no tiene pronóstico para los meses 1–3 (se necesitan 3 períodos previos para calcularlo). Las celdas correspondientes deben aparecer vacías o con guión. Esto es correcto y esperado; las métricas de error para PM3 se calcularán solo sobre los meses 4–12 (9 períodos).

3. Verifique que el SKU de trabajo en esta hoja corresponde a **SKU-A**. Las hojas de SKU-B y SKU-C tienen estructura idéntica y se trabajarán de forma análoga si el tiempo lo permite.

4. Confirme con el instructor si los valores de pronóstico difieren de los suyos propios (provenientes de las Prácticas 18–20). En caso de discrepancia, use los valores del archivo pre-cargado para mantener coherencia en el grupo.

#### Resultado esperado
La hoja `Practica21_Trabajo` muestra 12 filas de datos (Mes 1 a Mes 12) con demanda real y los tres pronósticos correctamente posicionados. El participante comprende que PM3 inicia en Mes 4 y que los modelos SE y RL tienen valores desde Mes 1 o Mes 2.

#### Verificación
- Cuente las celdas con valor en la columna C (PM3): deben ser **9** (Mes 4 a Mes 12).
- Cuente las celdas con valor en la columna D (SE): deben ser **12** (Mes 1 a Mes 12).
- Cuente las celdas con valor en la columna E (RL): deben ser **12** (Mes 1 a Mes 12).

---

### Paso 2 — Calcular el error absoluto período a período para los tres modelos

**Objetivo:** Construir tres columnas de error absoluto (una por modelo) usando la función `ABS()`, aplicando el concepto central de la Lección 4.1: eliminar el signo para obtener la magnitud real del error sin cancelaciones.

#### Instrucciones

1. En la fila 1 (encabezados), escriba los siguientes títulos en las columnas F, G y H:
   - **F1:** `EA_PM3` *(Error Absoluto — Promedio Móvil 3 períodos)*
   - **G1:** `EA_SE` *(Error Absoluto — Suavización Exponencial)*
   - **H1:** `EA_RL` *(Error Absoluto — Regresión Lineal)*

2. En la celda **F5** (primer período con pronóstico disponible para PM3, Mes 4), ingrese la siguiente fórmula:
   ```
   =ABS(B5-C5)
   ```
   Esta fórmula calcula `|Demanda_Real − Pronóstico_PM3|` para el Mes 4.

3. Copie la fórmula de F5 hacia abajo hasta **F13** (Mes 12). Las celdas F2, F3 y F4 deben permanecer **vacías** (no hay pronóstico PM3 para los primeros 3 meses).

4. En la celda **G2** (Mes 1, primer período con pronóstico SE), ingrese:
   ```
   =ABS(B2-D2)
   ```
   Copie hacia abajo hasta **G13** (Mes 12).

5. En la celda **H2** (Mes 1, primer período con pronóstico RL), ingrese:
   ```
   =ABS(B2-E2)
   ```
   Copie hacia abajo hasta **H13** (Mes 12).

6. Revise que ninguna celda muestre `#¡VALOR!` o `#DIV/0!`. Si aparece algún error, verifique que las celdas de pronóstico referenciadas no estén vacías o contengan texto.

   > **Recordatorio conceptual (Lección 4.1):** Si en un período el error con signo fue +30 (subestimamos) y en otro fue −30 (sobreestimamos), la suma de errores con signo sería 0 — aparentando un modelo perfecto. El error absoluto devuelve 30 en ambos casos, revelando la verdad: nos equivocamos 30 unidades en cada período.

#### Resultado esperado
Las columnas F, G y H muestran valores numéricos positivos (o cero) para cada período. Los valores típicos deben oscilar entre 10 y 120 unidades para este dataset. La columna F tiene celdas vacías en F2:F4.

#### Verificación
- Seleccione cualquier celda de la columna F (ej. F7) y confirme que la fórmula en la barra de fórmulas sea `=ABS(B7-C7)`.
- Compruebe manualmente el valor de G5: si `B5 = 495` y `D5 = 540`, entonces `G5` debe mostrar **45**.

---

### Paso 3 — Calcular el error porcentual absoluto período a período

**Objetivo:** Construir tres columnas de error porcentual absoluto (EPA) que expresen el error como porcentaje de la demanda real, preparando los insumos para el cálculo del MAPE.

#### Instrucciones

1. En la fila 1, escriba los siguientes títulos en las columnas I, J y K:
   - **I1:** `EPA_PM3` *(Error Porcentual Absoluto — PM3)*
   - **J1:** `EPA_SE` *(Error Porcentual Absoluto — SE)*
   - **K1:** `EPA_RL` *(Error Porcentual Absoluto — RL)*

2. En la celda **I5** (Mes 4, primer período válido para PM3), ingrese:
   ```
   =ABS(B5-C5)/B5*100
   ```
   Esta fórmula calcula `|Real − Pronóstico| / Real × 100`, expresando el error como porcentaje de la demanda real.

3. Copie I5 hacia abajo hasta **I13**. Deje I2:I4 vacías.

4. En **J2**, ingrese:
   ```
   =ABS(B2-D2)/B2*100
   ```
   Copie hacia abajo hasta **J13**.

5. En **K2**, ingrese:
   ```
   =ABS(B2-E2)/B2*100
   ```
   Copie hacia abajo hasta **K13**.

6. Aplique formato de número con **1 decimal y símbolo %** a las columnas I, J y K:
   - Seleccione el rango **I2:K13**.
   - Presione `Ctrl + 1` para abrir *Formato de celdas*.
   - En la pestaña *Número*, seleccione **Número** con 1 decimal (el símbolo `%` se agregará manualmente en el encabezado para claridad; no use el formato porcentaje de Excel ya que los valores ya están multiplicados por 100).

   > **Alternativa rápida:** Si prefiere no usar el cuadro de diálogo, seleccione I2:K13 y haga clic en el botón de formato de número en la cinta → *Número* → ajuste decimales a 1.

#### Resultado esperado
Las columnas I, J y K muestran valores entre 0.0 y ~25.0 (representando 0% a ~25% de error porcentual). Un valor de `9.1` en la celda J5 significaría que el modelo SE se equivocó en un 9.1% respecto a la demanda real del Mes 4.

#### Verificación
- Verifique manualmente J5: si `B5 = 495` y `D5 = 540`, entonces `J5 = ABS(495−540)/495×100 = 45/495×100 ≈ 9.1`.
- Confirme que ningún valor supere 50% (valores superiores indicarían un error en los datos de pronóstico).

---

### Paso 4 — Construir la tabla resumen de MAE y MAPE

**Objetivo:** Sintetizar los errores período a período en dos métricas agregadas (MAE y MAPE) para cada modelo, creando la tabla comparativa que permitirá tomar la decisión de selección.

#### Instrucciones

1. Navegue a la zona inferior de la hoja (a partir de la fila **17**) o a una zona en blanco a la derecha de los datos. Se recomienda usar el rango **A17:E21** para la tabla resumen.

2. Construya la siguiente estructura de tabla resumen (ingrese manualmente los encabezados y las fórmulas):

   | Celda | Contenido a ingresar |
   |-------|----------------------|
   | A17   | `Tabla Resumen — Comparación de Modelos (SKU-A)` |
   | A18   | `Métrica` |
   | B18   | `PM3 (Promedio Móvil 3P)` |
   | C18   | `SE (Suavización Exp. α=0.3)` |
   | D18   | `RL (Regresión Lineal)` |
   | A19   | `MAE (unidades)` |
   | A20   | `MAPE (%)` |
   | A21   | `Períodos evaluados` |

3. En la celda **B19** (MAE del modelo PM3), ingrese:
   ```
   =PROMEDIO(F5:F13)
   ```
   Esta fórmula calcula el promedio de los 9 errores absolutos del modelo PM3.

4. En la celda **C19** (MAE del modelo SE), ingrese:
   ```
   =PROMEDIO(G2:G13)
   ```
   Promedio de los 12 errores absolutos del modelo SE.

5. En la celda **D19** (MAE del modelo RL), ingrese:
   ```
   =PROMEDIO(H2:H13)
   ```
   Promedio de los 12 errores absolutos del modelo RL.

6. En la celda **B20** (MAPE del modelo PM3), ingrese:
   ```
   =PROMEDIO(I5:I13)
   ```

7. En la celda **C20** (MAPE del modelo SE), ingrese:
   ```
   =PROMEDIO(J2:J13)
   ```

8. En la celda **D20** (MAPE del modelo RL), ingrese:
   ```
   =PROMEDIO(K2:K13)
   ```

9. En las celdas **B21, C21 y D21**, ingrese respectivamente los valores de períodos evaluados:
   ```
   B21: 9
   C21: 12
   D21: 12
   ```
   > Esto es importante para la interpretación: PM3 tiene menos períodos de evaluación, lo que puede influir en la comparación.

10. Aplique formato de número con **2 decimales** al rango B19:D20.

#### Resultado esperado
La tabla resumen muestra los valores de MAE (en unidades de demanda) y MAPE (en porcentaje) para los tres modelos. Ejemplo de valores esperados con el dataset del instructor:

| Métrica             | PM3    | SE     | RL     |
|---------------------|--------|--------|--------|
| MAE (unidades)      | ~42.3  | ~38.7  | ~51.2  |
| MAPE (%)            | ~8.4%  | ~7.6%  | ~10.1% |
| Períodos evaluados  | 9      | 12     | 12     |

> *Los valores exactos dependen del dataset. El patrón típico es que SE presente el menor error para demanda con tendencia moderada.*

#### Verificación
- Confirme que las fórmulas en B19:D20 muestran la función `PROMEDIO` en la barra de fórmulas.
- Sume manualmente los valores de la columna G2:G13 y divida entre 12; el resultado debe coincidir con C19.
- Verifique que B20 sea siempre mayor que cero (un MAPE de 0 indicaría pronóstico perfecto, lo cual no es realista).

---

### Paso 5 — Identificar el modelo ganador y aplicar formato condicional

**Objetivo:** Resaltar visualmente el modelo con menor MAE y MAPE en la tabla resumen utilizando formato condicional, facilitando la comunicación de la decisión a cualquier interlocutor.

#### Instrucciones

1. **Identificar el modelo ganador:** Observe la fila 19 (MAE) y la fila 20 (MAPE) de la tabla resumen. El modelo ganador es aquel con el **valor más bajo** en ambas métricas. Si los modelos difieren en el ranking (uno tiene menor MAE pero otro tiene menor MAPE), se privilegiará el **MAPE** como criterio principal, ya que permite comparación relativa independiente del volumen.

2. Seleccione el rango **B19:D19** (fila de MAE).

3. En la cinta, vaya a **Inicio → Formato condicional → Reglas superiores/inferiores → 1 valor inferior**.
   - Establezca el formato: **Relleno verde claro con texto verde oscuro** (opción predeterminada de Excel).
   - Haga clic en **Aceptar**.

4. Repita el mismo procedimiento para el rango **B20:D20** (fila de MAPE):
   - **Inicio → Formato condicional → Reglas superiores/inferiores → 1 valor inferior**.
   - Use el mismo formato verde.

5. Adicionalmente, aplique un borde grueso a la columna del modelo ganador (toda la columna B, C o D de la tabla resumen, según corresponda):
   - Seleccione el rango de la columna ganadora en la tabla (ej. C18:C21 si el ganador es SE).
   - Presione `Ctrl + 1` → pestaña *Borde* → seleccione borde exterior grueso en color azul o naranja.

6. En la celda **E18**, escriba el encabezado: `Modelo Ganador`.
7. En la celda **E19**, escriba el nombre del modelo ganador (ej. `SE — Suavización Exponencial`).

#### Resultado esperado
La tabla resumen muestra en verde las celdas con los valores mínimos de MAE y MAPE. La columna del modelo ganador tiene un borde visual distintivo. Cualquier persona que abra el archivo puede identificar inmediatamente cuál modelo tuvo mejor desempeño.

#### Verificación
- Haga clic en la celda B19 y verifique en **Inicio → Formato condicional → Administrar reglas** que existe una regla activa del tipo "1 valor inferior" para el rango B19:D19.
- Cambie temporalmente un valor en B19 a un número muy bajo (ej. 1) y confirme que el verde se mueve a esa celda. Luego deshaga el cambio con `Ctrl + Z`.

---

### Paso 6 — Redactar la justificación operativa en una celda de comentario

**Objetivo:** Producir una justificación operativa documentada de 3–5 líneas que traduzca la decisión estadística (modelo con menor MAE/MAPE) en lenguaje logístico, vinculando el error con impactos reales en inventario, nivel de servicio y costos.

#### Instrucciones

1. Haga clic en la celda **A23** y escriba el encabezado:
   ```
   JUSTIFICACIÓN OPERATIVA — SELECCIÓN DEL MODELO DE PRONÓSTICO (SKU-A)
   ```
   Aplique **negrita** y color de fuente azul oscuro.

2. En la celda **A24**, redacte su justificación operativa. Active el ajuste de texto (`Ctrl + 1` → Alineación → Ajustar texto) y expanda la altura de la fila para visualizar el texto completo.

   La justificación debe responder las siguientes preguntas guía (no como lista, sino integradas en un párrafo coherente):
   - ¿Qué modelo fue seleccionado y por qué métrica?
   - ¿Qué tipo de demanda tiene el SKU-A que hace adecuado este modelo?
   - ¿Qué consecuencia operativa tendría usar el modelo con mayor error? (quiebre de stock, sobreinventario, costos)
   - ¿Cuál es el MAE en términos de días de inventario o unidades de seguridad?

   **Ejemplo de justificación operativa de referencia** (el participante debe redactar la suya propia con los valores reales obtenidos):

   > *"Se selecciona el modelo de Suavización Exponencial (α = 0.3) como modelo de pronóstico para SKU-A, dado que presenta el menor MAE (38.7 unidades) y el menor MAPE (7.6%) entre los tres modelos evaluados. SKU-A exhibe una demanda con tendencia moderada y variabilidad media, características para las que la suavización exponencial se adapta mejor que el promedio móvil simple, ya que pondera más los períodos recientes. Utilizar el modelo de regresión lineal (MAPE = 10.1%) implicaría un error adicional de ~12 unidades por período, lo que en términos de inventario de seguridad representaría un costo de oportunidad mensual estimado de $240 por sobreinventario o un riesgo de quiebre de stock equivalente al 2.5% de la demanda mensual promedio. El modelo SE es, por tanto, el más adecuado para dimensionar el stock de seguridad y planificar las órdenes de reabastecimiento de este SKU."*

3. Ajuste el texto de la celda A24 según los valores reales que obtuvo en su práctica. No copie el ejemplo textualmente; adáptelo con sus números.

4. Aplique un borde exterior a las celdas **A23:A24** con línea gruesa de color verde oscuro para destacar la sección de justificación.

#### Resultado esperado
La celda A24 contiene un párrafo de 3–5 líneas que menciona explícitamente: (a) el nombre del modelo seleccionado, (b) los valores de MAE y MAPE obtenidos, (c) al menos una consecuencia operativa logística de usar el modelo incorrecto, y (d) una referencia al impacto en inventario o nivel de servicio.

#### Verificación
- Lea su justificación en voz alta y confirme que menciona al menos: nombre del modelo, valor de MAE o MAPE, y una consecuencia en términos de inventario o nivel de servicio.
- Comparta brevemente su justificación con el compañero de al lado y pida retroalimentación sobre si es comprensible para alguien sin conocimiento estadístico.

---

### Paso 7 — Guardar y nombrar el archivo con convención estándar

**Objetivo:** Preservar el trabajo realizado con una convención de nombre que facilite la trazabilidad para las prácticas siguientes.

#### Instrucciones

1. Presione `Ctrl + Mayús + S` (o **Archivo → Guardar como**) para guardar con un nuevo nombre.

2. Guarde el archivo con el siguiente nombre estándar:
   ```
   Practica21_[SusIniciales]_Completada.xlsx
   ```
   Por ejemplo: `Practica21_JGM_Completada.xlsx`

3. Guarde en la carpeta de trabajo del curso (`C:\CursoForecasting\Practica21\` o la carpeta indicada por el instructor).

4. Confirme que el archivo se guardó correctamente verificando que el nombre en la barra de título de Excel haya cambiado.

#### Resultado esperado
El archivo queda guardado con nombre personalizado y sin errores. La hoja `Practica21_Trabajo` contiene todas las columnas de error, la tabla resumen con formato condicional y la justificación operativa.

#### Verificación
- Cierre y vuelva a abrir el archivo para confirmar que todos los datos y formatos persisten correctamente.
- Confirme que la hoja `Practica21_Trabajo` es la primera hoja visible al abrir el archivo.

---

## Validación y Pruebas

Una vez completados todos los pasos, realice las siguientes verificaciones finales para confirmar que la práctica está correctamente ejecutada:

### Lista de verificación final

| N° | Elemento a verificar | Criterio de aceptación | ✓ |
|----|----------------------|------------------------|---|
| 1 | Columnas EA_PM3, EA_SE, EA_RL (F, G, H) | Todas las celdas con pronóstico disponible muestran valores ≥ 0; ningún error `#¡VALOR!` | ☐ |
| 2 | Columnas EPA_PM3, EPA_SE, EPA_RL (I, J, K) | Valores entre 0% y 30%; formato numérico con 1 decimal | ☐ |
| 3 | Tabla resumen (A17:E21) | MAE y MAPE calculados con `PROMEDIO()`; períodos evaluados correctos (9 para PM3, 12 para SE y RL) | ☐ |
| 4 | Formato condicional | Celdas con menor MAE y menor MAPE resaltadas en verde | ☐ |
| 5 | Justificación operativa (A24) | Párrafo de 3–5 líneas con nombre del modelo, valores numéricos y consecuencia logística | ☐ |
| 6 | Archivo guardado | Nombre con convención `Practica21_[Iniciales]_Completada.xlsx` | ☐ |

### Prueba de coherencia cruzada

Realice esta prueba rápida para validar la coherencia de sus cálculos:

1. En una celda temporal (ej. B26), calcule manualmente el MAE del modelo SE sumando los valores de G2:G13 y dividiéndolos entre 12:
   ```
   =SUMA(G2:G13)/12
   ```
2. Compare el resultado con el valor en C19 (MAE_SE calculado con `PROMEDIO()`). Ambos deben ser **idénticos** (diferencia = 0).
3. Si hay diferencia, revise si alguna celda en G2:G13 contiene texto o está vacía inadvertidamente.
4. Elimine la fórmula temporal de B26 una vez confirmada la coherencia.

---

## Solución de Problemas

### Problema 1: La función `PROMEDIO()` devuelve un valor incorrecto o ignora celdas

**Síntoma:** El MAE calculado en la tabla resumen parece demasiado bajo o no coincide con una suma manual dividida entre el número de períodos. Por ejemplo, `PROMEDIO(F5:F13)` devuelve un valor diferente a `SUMA(F5:F13)/9`.

**Causa probable:** Algunas celdas en el rango F5:F13 contienen texto, espacios en blanco o el carácter de guión `"—"` en lugar de estar vacías. La función `PROMEDIO()` en Excel ignora las celdas verdaderamente vacías, pero **incluye en el denominador** solo las celdas con valores numéricos. Si hay texto, puede generar un error o un promedio sobre un subconjunto incorrecto de datos.

**Solución:**
1. Seleccione el rango F5:F13 y presione `Ctrl + G` → **Especial → Celdas en blanco** para identificar celdas vacías.
2. Si encuentra celdas con texto o guiones, selecciónelas y presione `Supr` para limpiarlas.
3. Verifique que la fórmula `ABS()` en cada celda referencia correctamente las columnas B y C (no B y D o B y E por error de copia).
4. Recalcule con `F9` y compare nuevamente con la suma manual.

---

### Problema 2: El formato condicional no resalta el valor mínimo correcto

**Síntoma:** El formato condicional verde aparece en una celda que no contiene el valor más bajo, o aparece en más de una celda cuando los valores son distintos. Por ejemplo, tanto B19 como C19 aparecen en verde aunque sus valores son diferentes.

**Causa probable:** Se aplicaron reglas de formato condicional superpuestas en pasos anteriores (ej. al probar la verificación del Paso 5), o la regla se configuró como "1 valor superior" en lugar de "1 valor inferior". Excel puede acumular reglas contradictorias que producen resultados inesperados.

**Solución:**
1. Seleccione el rango B19:D19.
2. Vaya a **Inicio → Formato condicional → Administrar reglas**.
3. En el cuadro de diálogo, elimine **todas las reglas existentes** para ese rango haciendo clic en cada una y presionando **Eliminar regla**.
4. Cierre el cuadro de diálogo.
5. Vuelva a aplicar el formato condicional desde cero siguiendo las instrucciones del Paso 5: **Inicio → Formato condicional → Reglas superiores/inferiores → 1 valor inferior**.
6. Repita el mismo proceso para el rango B20:D20.
7. Confirme que ahora solo una celda por fila aparece resaltada en verde (la que contiene el valor mínimo).

---

## Limpieza y Cierre

Al finalizar la práctica, realice las siguientes acciones de cierre:

1. **Guarde el archivo final** con `Ctrl + S` para asegurar que todos los cambios de la verificación final quedaron guardados.

2. **Entregue o comparta el archivo** según las instrucciones del instructor (correo electrónico, carpeta compartida de red, plataforma del curso).

3. **No elimine** ninguna hoja del archivo, especialmente `Practica21_Trabajo`, ya que puede ser insumo de referencia para las Prácticas 22 y 23.

4. Si utilizó una hoja de trabajo adicional para cálculos temporales, puede eliminarla o dejarla con el nombre `Borrador_Temporal` para no confundirla con los resultados finales.

5. Cierre el archivo Excel. Si el sistema solicita guardar cambios nuevamente, confirme con **Guardar**.

---

## Resumen

En esta práctica construyó el ciclo completo de evaluación de modelos de pronóstico aplicado al contexto logístico:

| Etapa | Lo que hizo | Concepto clave |
|-------|-------------|----------------|
| **1** | Revisó la estructura de datos y pronósticos pre-cargados | Comprensión del alcance de cada modelo |
| **2** | Calculó el error absoluto período a período con `ABS()` | `|Real − Pronóstico|` sin cancelación de signos (Lección 4.1) |
| **3** | Calculó el error porcentual absoluto por período | `|Real − Pronóstico| / Real × 100` |
| **4** | Construyó la tabla resumen de MAE y MAPE con `PROMEDIO()` | Síntesis de la precisión global de cada modelo |
| **5** | Aplicó formato condicional para resaltar el modelo ganador | Comunicación visual de la decisión |
| **6** | Redactó la justificación operativa | Vinculación estadística → decisión logística |
| **7** | Guardó con convención de nombre estándar | Trazabilidad para prácticas siguientes |

### Conceptos consolidados

- **MAE (Mean Absolute Error):** Promedio de los errores absolutos. Expresa el error en las mismas unidades que la demanda (ej. cajas, pallets). Un MAE de 38.7 unidades significa que, en promedio, el modelo se equivocó en 38.7 unidades por período.

- **MAPE (Mean Absolute Percentage Error):** Promedio de los errores porcentuales absolutos. Permite comparar modelos entre SKUs con escalas de demanda muy diferentes. Un MAPE del 7.6% significa que el modelo se equivocó, en promedio, en el 7.6% de la demanda real de cada período.

- **Impacto logístico del error:** Cada unidad de MAE representa un riesgo de sobreinventario (si el modelo sobreestima) o de quiebre de stock (si subestima). Seleccionar el modelo con menor error no es un ejercicio estadístico abstracto: es una decisión que afecta directamente el nivel de servicio al cliente, el capital inmovilizado en inventario y los costos operativos de picking y reabastecimiento.

### Conexión con la siguiente práctica

Los resultados de esta práctica (modelo seleccionado, valores de MAE y MAPE) serán el punto de partida de la **Práctica 22**, donde utilizará herramientas de IA Generativa (Microsoft Copilot) para generar escenarios predictivos basados en el modelo de mejor desempeño y producir recomendaciones automatizadas de reabastecimiento para los SKUs analizados.

### Recursos de referencia

- Hyndman, R.J. & Athanasopoulos, G. — *Forecasting: Principles and Practice* (3rd ed.), Sección 5.8: [https://otexts.com/fpp3/accuracy.html](https://otexts.com/fpp3/accuracy.html)
- Microsoft Support — Función `ABS()` en Excel: [https://support.microsoft.com/es-es/office/abs-función-abs](https://support.microsoft.com/es-es/office/abs-funci%C3%B3n-abs-3420200f-5628-4e8c-99da-c99d7c87713c)
- Microsoft Support — Función `PROMEDIO()` en Excel: [https://support.microsoft.com/es-es/office/promedio-función-promedio](https://support.microsoft.com/es-es/office/promedio-funci%C3%B3n-promedio-047bac88-d466-426c-a32b-8f33eb960cf6)
- APICS / ASCM — Diccionario de términos de cadena de suministro: [https://www.ascm.org/learning-development/certifications-designations/dictionary/](https://www.ascm.org/learning-development/certifications-designations/dictionary/)

---
*Práctica 21 — Módulo 4: Evaluación de Modelos de Pronóstico | Duración estimada: 12 minutos*
