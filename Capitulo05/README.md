# Práctica 22 — Construcción de un dashboard básico de demanda en Power BI Desktop

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Crear (*Create*) |
| **Módulo** | 5 — Visualización y pronóstico en Power BI |
| **Práctica número** | 22 |
| **Archivo de datos requerido** | `Demanda_Dashboard_P22.xlsx` |

---

## 2. Descripción general

En esta práctica construirás un **dashboard básico de demanda histórica** en Power BI Desktop a partir de un dataset Excel con 18 meses de información para 5 productos distribuidos en 2 categorías logísticas. Partirás desde la importación del archivo, verificarás los tipos de datos en Power Query, y levantarás tres visualizaciones complementarias: un gráfico de línea de evolución histórica, un gráfico de barras apiladas por producto/categoría y una tarjeta KPI con la demanda total del período. Finalmente, activarás el **Analytics Pane** de Power BI para agregar una línea de tendencia con ecuación visible y un pronóstico nativo de 3 períodos con intervalo de confianza del 95%, conectando directamente el análisis estadístico con decisiones operativas de inventario y abastecimiento.

---

## 3. Objetivos de aprendizaje

- [ ] Importar correctamente un dataset de demanda histórica desde Excel a Power BI Desktop, verificando que la columna `Fecha` sea tipo *Date* y `Unidades_Demandadas` sea tipo *Número entero*.
- [ ] Construir un dashboard de una sola página con al menos tres visualizaciones: gráfico de línea, gráfico de barras apiladas y tarjeta KPI de demanda total.
- [ ] Aplicar el **Analytics Pane** de Power BI para agregar una línea de tendencia (con ecuación visible) y activar el pronóstico nativo para 3 períodos adicionales con intervalo de confianza del 95%.
- [ ] Configurar títulos descriptivos y colores diferenciados por categoría para comunicar el análisis de forma profesional.

---

## 4. Prerrequisitos

### Conocimiento previo
- Haber revisado el material conceptual de la **Lección 5.1**: importación de datos desde Excel a Power BI y uso básico de Power Query.
- Comprensión básica de los conceptos de demanda histórica, tendencia y estacionalidad (Módulos 1–4 del curso).
- Conocimiento básico de navegación en Windows: abrir archivos, crear carpetas, copiar rutas.

### Acceso y software
- **Power BI Desktop** instalado y funcionando (versión 2024 o superior recomendada; descarga gratuita en [microsoft.com/power-bi](https://www.microsoft.com/es-es/download/details.aspx?id=58494)).
- **Microsoft Excel 2016 o superior** para visualizar el archivo fuente si es necesario.
- Archivo `Demanda_Dashboard_P22.xlsx` descargado en la carpeta de trabajo (proporcionado por el instructor).
- Espacio disponible en disco: mínimo 500 MB libres.
- Resolución de pantalla mínima: 1366×768 (se recomienda 1920×1080).

---

## 5. Entorno de laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1366×768 | 1920×1080 |
| Conexión a internet | No requerida para esta práctica | — |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Power BI Desktop | 2024 (gratuito) | Descargar desde microsoft.com si no está instalado |
| Microsoft Excel | 2016 o superior | Solo para inspección del archivo fuente |

### Configuración inicial del entorno

Antes de iniciar el cronómetro de 12 minutos, ejecuta los siguientes pasos de preparación (pueden realizarse antes de la sesión):

1. Crea la carpeta de trabajo en tu equipo:
```
C:\LabPBI\Practica22\
```

2. Copia el archivo proporcionado por el instructor en esa carpeta:
```
Demanda_Dashboard_P22.xlsx  →  C:\LabPBI\Practica22\
```

3. Abre Power BI Desktop. Si aparece la pantalla de bienvenida (*Getting Started*), ciérrala haciendo clic en la **X** de la esquina superior derecha del panel emergente.

4. Verifica que Power BI Desktop esté en **idioma español**. Si aparece en inglés, ve a *File → Options and Settings → Options → Regional Settings* y selecciona **Spanish (Spain)** o **Spanish (Latin America)** según corresponda.

> **Nota para el instructor:** El archivo `Demanda_Dashboard_P22.xlsx` debe contener una hoja llamada `Demanda_Historica` con las columnas: `Fecha` (formato dd/mm/aaaa, 18 registros mensuales desde ene-2023 a jun-2024), `Producto` (5 productos: Leche UHT, Yogur Natural, Arroz 5kg, Aceite Vegetal, Avena), `Categoría` (Perecederos / No Perecederos), `Unidades_Demandadas` (valores enteros entre 80 y 950 por registro). La hoja debe estar convertida en **Tabla de Excel** (Ctrl+T) con nombre `Demanda_Historica` para facilitar la detección automática en Power BI.

---

## 6. Procedimiento paso a paso

---

### Paso 1: Importar el archivo Excel a Power BI Desktop

**Objetivo:** Establecer la conexión entre Power BI Desktop y el archivo `Demanda_Dashboard_P22.xlsx`, seleccionando la tabla correcta y abriendo Power Query para verificar tipos de datos.

**Instrucciones:**

1. Con Power BI Desktop abierto y en pantalla en blanco, ve a la cinta superior y haz clic en la pestaña **Inicio**.

2. Haz clic en el botón **"Obtener datos"**. En el menú desplegable que aparece, selecciona **"Excel"** (aparece cerca de la parte superior de la lista con el ícono verde de Excel).

3. Se abrirá el explorador de archivos de Windows. Navega hasta:
```
C:\LabPBI\Practica22\Demanda_Dashboard_P22.xlsx
```
Selecciona el archivo y haz clic en **"Abrir"**.

4. Aparecerá el **Navegador de Power BI** (*Navigator*). En el panel izquierdo verás el nombre del archivo y debajo la tabla **`Demanda_Historica`**. Haz clic sobre el nombre de la tabla para seleccionarla (aparecerá una marca de verificación).

5. En el panel derecho del Navegador observarás una vista previa de los datos con las columnas: `Fecha`, `Producto`, `Categoría`, `Unidades_Demandadas`. Verifica visualmente que los datos se ven correctos.

6. **No hagas clic en "Cargar"**. En su lugar, haz clic en el botón **"Transformar datos"** (parte inferior del Navegador). Esto abrirá el editor de Power Query.

**Resultado esperado:** Se abre la ventana de Power Query con la tabla `Demanda_Historica` cargada en el panel central. En el panel derecho (*Pasos aplicados*) verás al menos los pasos: `Origen`, `Navegación` y `Tipo cambiado`.

**Verificación:** Confirma que en el panel central de Power Query puedes ver las 4 columnas (`Fecha`, `Producto`, `Categoría`, `Unidades_Demandadas`) y al menos 5 filas de datos visibles en la vista previa.

---

### Paso 2: Verificar y corregir tipos de datos en Power Query

**Objetivo:** Asegurar que la columna `Fecha` sea reconocida como tipo *Fecha* y `Unidades_Demandadas` como tipo *Número entero*, condición indispensable para que el gráfico de línea y el pronóstico funcionen correctamente.

**Instrucciones:**

1. En Power Query, observa la fila de encabezados de la tabla. A la **izquierda de cada nombre de columna** hay un ícono pequeño que indica el tipo de dato:
   - Ícono de calendario = Fecha ✓
   - Ícono `123` = Número entero ✓
   - Ícono `ABC` = Texto
   - Ícono `1.2` = Número decimal

2. **Verificar columna `Fecha`:** Haz clic en el ícono a la izquierda del encabezado `Fecha`. Si ya muestra el ícono de calendario, el tipo es correcto. Si muestra `ABC` o `1.2`, haz clic sobre ese ícono y selecciona **"Fecha"** del menú desplegable.

3. **Verificar columna `Unidades_Demandadas`:** Haz clic en el ícono a la izquierda del encabezado `Unidades_Demandadas`. Si muestra `123`, el tipo es correcto. Si muestra `1.2` (decimal), haz clic sobre ese ícono y selecciona **"Número entero"**.

4. **Verificar columnas de texto:** Confirma que `Producto` y `Categoría` muestren el ícono `ABC` (texto). No requieren cambio.

5. Una vez verificados los tipos, en la esquina superior izquierda de Power Query, haz clic en **"Cerrar y aplicar"** (*Close & Apply*). Power BI cargará los datos al modelo.

**Resultado esperado:** Power BI Desktop regresa a la vista de reporte. En el panel derecho **Campos** (*Fields*), verás la tabla `Demanda_Historica` con sus 4 columnas. La columna `Fecha` mostrará un ícono de calendario y `Unidades_Demandadas` mostrará un ícono de suma (∑), indicando que es un campo numérico agregable.

**Verificación:** En el panel **Campos** (derecha), expande la tabla `Demanda_Historica`. Confirma que:
- `Fecha` tiene ícono de calendario 📅
- `Unidades_Demandadas` tiene ícono ∑ (suma)
- `Producto` y `Categoría` tienen ícono de texto (sin símbolo especial)

---

### Paso 3: Crear el gráfico de línea de demanda histórica

**Objetivo:** Construir el gráfico principal del dashboard que muestra la evolución mensual de la demanda total a lo largo de los 18 meses históricos.

**Instrucciones:**

1. Asegúrate de estar en la **Vista de Informe** (ícono de gráfico de barras en el panel izquierdo de Power BI Desktop). La página del informe debe estar en blanco.

2. En el panel **Visualizaciones** (*Visualizations*, panel central derecho), haz clic en el ícono de **Gráfico de líneas** (*Line chart*). Es el ícono que muestra una línea ondulada. Se insertará un contenedor vacío en el lienzo.

3. Con el gráfico seleccionado (borde azul visible), ve al panel **Campos** y arrastra los siguientes campos a las secciones correspondientes del panel Visualizaciones:
   - Arrastra **`Fecha`** → sección **"Eje X"** (*X-axis*)
   - Arrastra **`Unidades_Demandadas`** → sección **"Eje Y"** (*Y-axis*)

4. Power BI puede agrupar automáticamente las fechas en jerarquía (Año → Trimestre → Mes). Para mostrar los meses individuales, haz clic en la flecha doble hacia abajo (▽▽) que aparece en la esquina superior derecha del gráfico cuando pasas el cursor sobre él. Esto expandirá la jerarquía hasta el nivel de **Mes**.

   > **Alternativa:** Si prefieres ver los meses directamente sin jerarquía, en la sección "Eje X" del panel Visualizaciones, haz clic en la flecha junto a `Fecha` y selecciona **"Fecha"** (en lugar de la jerarquía de fecha automática).

5. Redimensiona el gráfico arrastrando sus esquinas para que ocupe aproximadamente la **mitad superior izquierda** del lienzo (deja espacio para los otros dos elementos).

6. Con el gráfico aún seleccionado, ve al panel **Formato** (ícono de pincel en el panel Visualizaciones). Expande la sección **"Título"** y escribe como título:
```
Evolución Histórica de Demanda (18 meses)
```

**Resultado esperado:** Un gráfico de línea que muestra 18 puntos de datos mensuales con la demanda total (suma de todos los productos) en el eje Y y los meses en el eje X. La línea debe mostrar variaciones visibles entre meses.

**Verificación:** Pasa el cursor sobre cualquier punto de la línea. Debe aparecer un tooltip con el mes específico y el valor de `Suma de Unidades_Demandadas`. Confirma que el eje X muestre etiquetas de mes/año.

---

### Paso 4: Crear el gráfico de barras apiladas por producto y categoría

**Objetivo:** Construir una visualización comparativa que muestre la demanda acumulada por producto, diferenciando visualmente las dos categorías (Perecederos / No Perecederos) mediante colores.

**Instrucciones:**

1. Haz clic en un área **vacía** del lienzo para deseleccionar el gráfico de línea.

2. En el panel **Visualizaciones**, haz clic en el ícono de **Gráfico de barras apiladas** (*Stacked bar chart*). Es el ícono que muestra barras horizontales apiladas en diferentes colores. Se insertará un nuevo contenedor vacío.

3. Con el nuevo gráfico seleccionado, arrastra los campos desde el panel **Campos**:
   - Arrastra **`Producto`** → sección **"Eje Y"** (*Y-axis*) — esto pondrá los 5 productos en el eje vertical
   - Arrastra **`Unidades_Demandadas`** → sección **"Eje X"** (*X-axis*) — valores numéricos en el eje horizontal
   - Arrastra **`Categoría`** → sección **"Leyenda"** (*Legend*) — esto diferenciará los colores por categoría

4. Power BI asignará automáticamente colores distintos a *Perecederos* y *No Perecederos*. Cada barra de producto quedará segmentada en dos colores según la proporción de demanda de cada categoría.

   > **Nota:** Como cada producto pertenece a una sola categoría (no hay productos en ambas), las barras aparecerán de un solo color por producto, pero la leyenda de colores permitirá identificar visualmente qué productos son perecederos y cuáles no.

5. Redimensiona y posiciona este gráfico en la **mitad superior derecha** del lienzo, junto al gráfico de línea.

6. Ve al panel **Formato** → sección **"Título"** y escribe:
```
Demanda Total por Producto y Categoría
```

7. Opcional (si el tiempo lo permite): En el panel Formato, expande **"Colores de datos"** y verifica que los colores asignados a cada categoría sean claramente distinguibles. Si deseas cambiarlos, haz clic sobre el color de cada categoría y selecciona uno nuevo.

**Resultado esperado:** Un gráfico de barras horizontales apiladas con los 5 productos en el eje Y (Leche UHT, Yogur Natural, Arroz 5kg, Aceite Vegetal, Avena). Cada barra muestra la demanda total acumulada de los 18 meses, coloreada según su categoría. Una leyenda visible identifica los colores de Perecederos y No Perecederos.

**Verificación:** Pasa el cursor sobre una barra. El tooltip debe mostrar el nombre del producto, la categoría y el valor de `Suma de Unidades_Demandadas`. Confirma que aparecen exactamente 5 barras (una por producto).

---

### Paso 5: Crear la tarjeta KPI de demanda total

**Objetivo:** Agregar una tarjeta de indicador clave (KPI Card) que muestre de forma prominente la demanda total acumulada del período completo de 18 meses, funcionando como métrica de referencia rápida en el dashboard.

**Instrucciones:**

1. Haz clic en un área **vacía** del lienzo para deseleccionar el gráfico anterior.

2. En el panel **Visualizaciones**, haz clic en el ícono de **Tarjeta** (*Card*). Es el ícono que muestra el número "123" en un recuadro. Se insertará un contenedor pequeño.

3. Con la tarjeta seleccionada, arrastra el campo **`Unidades_Demandadas`** desde el panel **Campos** → sección **"Campos"** (*Fields*) del panel Visualizaciones.

4. Power BI mostrará automáticamente la **suma total** de todas las unidades demandadas en los 18 meses. El número aparecerá grande en el centro de la tarjeta.

5. Posiciona la tarjeta en la **parte inferior central** del lienzo, debajo de los dos gráficos anteriores, para que actúe como resumen global.

6. Ve al panel **Formato** → sección **"Etiqueta de categoría"** (*Category label*) y activa la opción. Esto mostrará debajo del número la etiqueta `Suma de Unidades_Demandadas`.

7. En el panel **Formato** → sección **"Título"**, activa el título y escribe:
```
Demanda Total del Período (18 meses)
```

8. Opcional: En el panel **Formato** → sección **"Valor de datos"**, aumenta el tamaño de fuente a **28 pt** para que el número sea claramente visible.

**Resultado esperado:** Una tarjeta que muestra un número grande (la suma total de todas las unidades demandadas por todos los productos en los 18 meses). El valor debe ser consistente con la suma visible en los gráficos anteriores.

**Verificación:** Anota el valor que muestra la tarjeta. Luego, en el gráfico de barras, suma visualmente los valores de las 5 barras (o pasa el cursor sobre cada una). La suma debe coincidir con el valor de la tarjeta. Si hay discrepancia, verifica que no haya filtros activos en el lienzo.

---

### Paso 6: Agregar línea de tendencia mediante el Analytics Pane

**Objetivo:** Utilizar el panel Analytics de Power BI para superponer una línea de tendencia matemática sobre el gráfico de línea histórico, con la ecuación visible, permitiendo interpretar la dirección general del comportamiento de la demanda.

**Instrucciones:**

1. Haz clic sobre el **gráfico de línea** (el que muestra la evolución histórica de 18 meses) para seleccionarlo. Debe aparecer el borde azul de selección.

2. En el panel de la derecha, localiza los tres íconos de panel en la parte superior:
   - Ícono de pincel = Formato visual
   - Ícono de lupa con gráfico = **Analytics Pane** (Panel de análisis)
   - Ícono de campos = Campos

   Haz clic en el **ícono de lupa con gráfico** para abrir el **Panel de análisis** (*Analytics Pane*).

3. En el Panel de análisis, verás una lista de opciones analíticas disponibles para el gráfico de línea. Localiza la opción **"Línea de tendencia"** (*Trend line*).

4. Haz clic en el botón **"Agregar"** (*Add*) junto a "Línea de tendencia". Aparecerá una línea recta superpuesta sobre la línea de demanda histórica.

5. Expande la sección **"Línea de tendencia"** haciendo clic sobre su nombre. Aparecerán las opciones de configuración:
   - **Color:** Cambia el color a **rojo** (`#FF0000`) para que sea claramente distinguible de la línea de demanda azul.
   - **Estilo:** Selecciona **"Guion"** (*Dashed*) o **"Sólido"** según preferencia.
   - **Mostrar etiqueta de datos** (*Show data label*): Activa esta opción — **ON**.
   - **Mostrar ecuación** (*Show equation*) — si aparece esta opción, actívala en **ON** para visualizar la ecuación de la línea de tendencia.

   > **Nota de versión:** En algunas versiones de Power BI Desktop, la opción "Mostrar ecuación" puede no estar disponible directamente en el Analytics Pane. Si no aparece, la línea de tendencia seguirá siendo válida visualmente; la ecuación exacta puede consultarse pasando el cursor sobre la línea.

6. Verifica que la línea de tendencia sea visible sobre el gráfico, con una dirección clara (ascendente, descendente o plana) que refleje el comportamiento general de la demanda en los 18 meses.

**Resultado esperado:** El gráfico de línea ahora muestra dos elementos: (1) la línea de demanda histórica con los 18 puntos de datos mensuales, y (2) una línea recta de tendencia en color rojo (o el color elegido) que atraviesa el gráfico de izquierda a derecha, indicando la dirección general de la demanda. Si la tendencia es ascendente, la línea roja subirá de izquierda a derecha.

**Verificación:** La línea de tendencia debe ser visualmente distinguible de la línea de datos históricos. Pasa el cursor sobre la línea de tendencia para confirmar que Power BI la identifica como "Línea de tendencia" en el tooltip.

---

### Paso 7: Activar el pronóstico nativo de Power BI (3 períodos, IC 95%)

**Objetivo:** Habilitar la función de pronóstico nativo del Analytics Pane para proyectar la demanda 3 meses hacia el futuro con un intervalo de confianza del 95%, conectando el análisis histórico con la planificación de inventario.

**Instrucciones:**

1. Asegúrate de que el **gráfico de línea** siga seleccionado y el **Panel de análisis** (*Analytics Pane*) siga abierto (si lo cerraste, haz clic nuevamente en el ícono de lupa con gráfico).

2. En el Panel de análisis, localiza la sección **"Pronóstico"** (*Forecast*). Haz clic en el botón **"Agregar"** (*Add*).

   > **Importante:** Si el botón "Agregar" aparece desactivado (gris), verifica que:
   > - El campo en el eje X sea de tipo **Fecha** (no texto ni número).
   > - El gráfico sea de tipo **Línea** (no área ni dispersión).
   > - No haya una jerarquía de fechas activa — si el eje X muestra "Año → Trimestre → Mes", debes cambiar el campo de fecha a un campo de fecha simple (ver nota en Paso 3).

3. Una vez agregado el pronóstico, expande la sección **"Pronóstico"** para configurar los parámetros:

   - **Unidades de pronóstico** (*Forecast length*): Escribe **`3`** y selecciona **"Meses"** (*Months*) en el selector de unidades.
   - **Intervalo de confianza** (*Confidence interval*): Ajusta el valor a **`95`** (porcentaje). Esto activará un área sombreada alrededor de la línea de pronóstico que representa el rango de incertidumbre.
   - **Estacionalidad** (*Seasonality*): Deja en **"Automático"** (*Auto detect*) para que Power BI detecte patrones estacionales automáticamente.
   - **Ignorar último** (*Ignore last*): Deja en `0` (no ignora ningún período).

4. Haz clic en **"Aplicar"** si aparece ese botón, o los cambios se aplicarán automáticamente según la versión de Power BI.

5. Observa el gráfico: a partir del último punto histórico (mes 18), aparecerá una **línea punteada** que se extiende 3 meses hacia la derecha (pronóstico), rodeada por un **área sombreada** que representa el intervalo de confianza del 95%.

6. Pasa el cursor sobre la zona de pronóstico para ver los valores proyectados mes a mes.

**Resultado esperado:** El gráfico de línea ahora muestra tres elementos diferenciados:
- **Línea sólida azul:** demanda histórica de los 18 meses.
- **Línea roja:** tendencia general calculada.
- **Línea punteada con área sombreada:** pronóstico de 3 meses futuros con intervalo de confianza del 95%.

El área sombreada (IC 95%) indica que Power BI estima con un 95% de confianza que la demanda real futura caerá dentro de ese rango. Un área más ancha indica mayor incertidumbre en la proyección.

**Verificación:** Confirma visualmente que:
- La línea punteada de pronóstico comienza exactamente donde termina la línea histórica (mes 18).
- El área sombreada es visible y se expande ligeramente hacia los meses más lejanos (mayor incertidumbre a mayor distancia).
- El eje X muestra las etiquetas de los 3 meses proyectados adicionales.

---

### Paso 8: Configuración final del dashboard — títulos y presentación

**Objetivo:** Aplicar configuración básica de presentación al dashboard completo para que sea legible, profesional y comunicativamente efectivo en un contexto logístico.

**Instrucciones:**

1. Haz clic en un área vacía del lienzo para deseleccionar todos los elementos.

2. **Título de la página del informe:** En la parte inferior del lienzo verás la pestaña de la página (llamada "Página 1" por defecto). Haz doble clic sobre esa pestaña y renómbrala:
```
Dashboard Demanda Histórica
```

3. **Agregar un cuadro de texto con título general:** En la cinta superior, ve a **Insertar → Cuadro de texto**. Haz clic en la parte superior del lienzo y escribe:
```
Análisis de Demanda Histórica — 5 Productos | Ene 2023 – Jun 2024
```
Formatea el texto con tamaño **16 pt**, **negrita**, color de fuente **azul oscuro** (`#1F3864` o similar). Posiciona este cuadro de texto como encabezado en la parte superior del lienzo.

4. **Verificar coherencia de colores:** Recorre los tres elementos visuales y confirma que:
   - Los colores de las barras del gráfico de barras apiladas son consistentes con la leyenda de categorías.
   - La línea de tendencia (roja) es claramente distinguible de la línea de datos (azul).
   - El área de pronóstico tiene un color suave (gris o azul claro) que no interfiere con la lectura de datos históricos.

5. **Guardar el archivo:** Presiona `Ctrl + S`. En el cuadro de diálogo, navega a:
```
C:\LabPBI\Practica22\
```
Guarda el archivo con el nombre:
```
Dashboard_Demanda_P22.pbix
```

**Resultado esperado:** Un dashboard de una sola página con:
- Título general en la parte superior.
- Gráfico de línea (superior izquierda) con línea de tendencia y pronóstico visible.
- Gráfico de barras apiladas (superior derecha) con colores por categoría.
- Tarjeta KPI (parte inferior) con la demanda total del período.
- Archivo guardado como `Dashboard_Demanda_P22.pbix` en la carpeta de trabajo.

**Verificación:** Cierra Power BI Desktop y vuelve a abrir el archivo `Dashboard_Demanda_P22.pbix` desde la carpeta. Confirma que el dashboard se carga correctamente con todas las visualizaciones y configuraciones aplicadas.

---

## 7. Validación y pruebas

Ejecuta las siguientes verificaciones para confirmar que la práctica fue completada correctamente:

| # | Verificación | Cómo validar | Resultado esperado |
|---|---|---|---|
| V1 | Tipos de datos correctos | Panel Campos → tabla `Demanda_Historica` | `Fecha` con ícono 📅, `Unidades_Demandadas` con ícono ∑ |
| V2 | Gráfico de línea funcional | Pasar cursor sobre puntos de la línea | Tooltip muestra mes/año y valor numérico de demanda |
| V3 | Gráfico de barras con categorías | Verificar leyenda del gráfico de barras | Dos colores en leyenda: Perecederos y No Perecederos |
| V4 | Tarjeta KPI coherente | Comparar valor de tarjeta con suma de barras | Los valores deben ser iguales (sin filtros activos) |
| V5 | Línea de tendencia visible | Observar gráfico de línea | Línea recta de color distinto superpuesta sobre datos históricos |
| V6 | Pronóstico activo | Observar extensión del gráfico de línea | Línea punteada + área sombreada extendiéndose 3 meses más allá del mes 18 |
| V7 | Archivo guardado | Explorador de Windows → carpeta de trabajo | Archivo `Dashboard_Demanda_P22.pbix` presente en `C:\LabPBI\Practica22\` |

### Pregunta de reflexión logística

> **¿Cómo impacta este análisis en una decisión operativa de logística?**
>
> Observa la dirección de la línea de tendencia en el gráfico:
> - Si la tendencia es **ascendente**, el área de compras/abastecimiento debería anticipar incrementos en los niveles de inventario de seguridad para los próximos 3 meses.
> - Si el intervalo de confianza del pronóstico es **muy amplio**, indica alta variabilidad en la demanda, lo que sugiere la necesidad de un stock de seguridad mayor o una política de reabastecimiento más frecuente.
> - Los productos de la categoría **Perecederos** con demanda creciente requieren especial atención en la planificación de picking para evitar quiebres de stock con productos de vida útil corta.

---

## 8. Solución de problemas

### Problema 1: El pronóstico no se activa — botón "Agregar" aparece desactivado en el Analytics Pane

**Síntoma:** En el Panel de análisis (*Analytics Pane*), la sección "Pronóstico" muestra el botón "Agregar" en color gris y no responde al hacer clic. No se puede activar el pronóstico.

**Causa:** Este problema ocurre cuando el campo en el eje X del gráfico de línea no es un campo de tipo **Fecha simple**, sino una **jerarquía de fechas automática** que Power BI genera al detectar campos de fecha (Año → Trimestre → Mes → Día). El pronóstico nativo de Power BI requiere un campo de fecha continuo, no una jerarquía.

**Solución paso a paso:**

1. Selecciona el gráfico de línea.
2. En el panel **Visualizaciones**, localiza la sección **"Eje X"** (*X-axis*).
3. Observa el campo que está asignado al eje X. Si muestra `Fecha (jerarquía)` o tiene una flecha de expansión, es una jerarquía.
4. Haz clic en la **flecha hacia abajo** junto al campo en el eje X para ver las opciones.
5. Selecciona **"Fecha"** en lugar de la jerarquía automática. El eje X mostrará ahora las fechas como valores continuos.
6. Regresa al **Panel de análisis** — el botón "Agregar" de Pronóstico debería estar ahora activo (azul).

> **Alternativa:** Si el problema persiste, en Power Query agrega una columna calculada de tipo Fecha con la fórmula `Date.From([Fecha])` y usa esa nueva columna en el eje X del gráfico.

---

### Problema 2: Los datos importados muestran errores o la columna `Fecha` aparece como texto

**Síntoma:** Después de importar el archivo Excel, en Power Query la columna `Fecha` muestra el ícono `ABC` (texto) en lugar del ícono de calendario. Los valores de fecha aparecen como texto (por ejemplo, `"01/01/2023"` entre comillas) y al intentar cambiar el tipo a *Fecha* aparece un error de conversión con valores `Error` en las celdas.

**Causa:** El archivo Excel tiene las fechas almacenadas como **texto** en lugar de como valores de fecha reales, o el separador de fecha del sistema operativo del equipo (punto o barra) no coincide con el formato del archivo Excel. Esto es común cuando el archivo fue creado en un sistema con configuración regional diferente (por ejemplo, fechas en formato `MM/DD/YYYY` en un sistema configurado para `DD/MM/YYYY`).

**Solución paso a paso:**

1. En Power Query, haz clic en el ícono `ABC` de la columna `Fecha` para intentar cambiar el tipo.
2. Selecciona **"Fecha"** del menú desplegable.
3. Si aparece un cuadro de diálogo preguntando sobre la configuración regional, selecciona **"Usar configuración regional"** y elige **"Español (España)"** o **"Español (México)"** según el formato del archivo.
4. Si aún aparecen errores, en Power Query ve a **"Agregar columna" → "Columna personalizada"** y escribe la siguiente fórmula para convertir texto a fecha:
```
= Date.FromText([Fecha], "es-ES")
```
5. Renombra la nueva columna como `Fecha_Corregida`, elimina la columna original `Fecha` y renombra `Fecha_Corregida` como `Fecha`.
6. Verifica que el ícono ahora sea de calendario y haz clic en **"Cerrar y aplicar"**.

> **Prevención:** Para evitar este problema en el futuro, al preparar el archivo Excel fuente, selecciona la columna de fechas, ve a *Formato de celdas → Fecha* y elige un formato estándar ISO (`AAAA-MM-DD`) que Power BI interpreta sin ambigüedad en cualquier configuración regional.

---

## 9. Limpieza del entorno

Al finalizar la práctica, realiza los siguientes pasos para dejar el entorno ordenado y listo para la siguiente práctica:

1. **Guardar el archivo final** (si no lo hiciste en el Paso 8):
   - Presiona `Ctrl + S` en Power BI Desktop.
   - Confirma que el archivo `Dashboard_Demanda_P22.pbix` está guardado en `C:\LabPBI\Practica22\`.

2. **Cerrar Power BI Desktop** normalmente (no forzar cierre). Si aparece un mensaje de "¿Desea guardar los cambios?", selecciona **"Guardar"**.

3. **Verificar la carpeta de trabajo:** Abre el explorador de Windows y navega a `C:\LabPBI\Practica22\`. Deberías ver los siguientes archivos:
   ```
   C:\LabPBI\Practica22\
   ├── Demanda_Dashboard_P22.xlsx     (archivo fuente — no modificar)
   └── Dashboard_Demanda_P22.pbix    (archivo de Power BI generado)
   ```

4. **No eliminar ningún archivo:** El archivo `.pbix` generado en esta práctica será el punto de partida para prácticas posteriores del Módulo 5, donde se agregarán medidas DAX, filtros interactivos y conexión con análisis de IA generativa.

5. **Compartir con el instructor** (si se requiere): Si el instructor solicita entrega, comprime la carpeta `Practica22` en un archivo `.zip` y compártela por el medio indicado:
   ```
   Practica22_[TuNombre].zip
   ```

---

## 10. Resumen y recursos adicionales

### Resumen de lo aprendido

En esta práctica completaste el ciclo completo de construcción de un dashboard básico de demanda en Power BI Desktop:

| Etapa | Acción realizada | Impacto logístico |
|---|---|---|
| **Importación** | Conexión a Excel via *Obtener datos*, verificación de tipos en Power Query | Garantiza que los datos de demanda sean interpretados correctamente para cálculos y visualizaciones |
| **Gráfico de línea** | Evolución mensual de demanda total (18 meses) | Permite identificar visualmente tendencias y estacionalidades que afectan el plan de compras |
| **Gráfico de barras** | Demanda comparativa por producto y categoría | Facilita la priorización de SKUs críticos en la planificación de picking y reabastecimiento |
| **Tarjeta KPI** | Demanda total del período | Referencia rápida para comparar con capacidad de almacenamiento y objetivos de nivel de servicio |
| **Línea de tendencia** | Dirección general de la demanda con ecuación | Confirma si la demanda está creciendo o decreciendo para ajustar políticas de inventario |
| **Pronóstico nativo** | Proyección 3 meses + IC 95% | Proporciona una estimación inicial de demanda futura para dimensionar órdenes de compra |

### Conceptos clave reforzados

- **Power Query como puerta de entrada:** Ningún dato debe cargarse al modelo sin pasar por Power Query. La verificación de tipos de datos es el primer control de calidad del análisis.
- **El Analytics Pane no es magia:** La línea de tendencia y el pronóstico nativo de Power BI son herramientas exploratorias basadas en regresión lineal simple. Para pronósticos logísticos complejos con estacionalidad marcada, se requieren modelos más sofisticados (como los desarrollados en las prácticas anteriores del curso).
- **El intervalo de confianza comunica incertidumbre:** Un IC 95% amplio en el pronóstico es una señal operativa: indica que la variabilidad de la demanda es alta y que el stock de seguridad debe ser mayor.
- **El dashboard es una herramienta de decisión, no solo de visualización:** Cada elemento construido en esta práctica responde a una pregunta logística concreta: ¿cuánto se demandó?, ¿qué productos lideran?, ¿hacia dónde va la demanda?

### Próximos pasos

En las prácticas siguientes del Módulo 5 agregarás:
- **Medidas DAX** para calcular indicadores como el coeficiente de variación directamente en Power BI.
- **Segmentadores de datos** (*Slicers*) para filtrar el dashboard por categoría, producto o rango de fechas.
- **Integración con análisis de IA generativa** (Copilot) para generar recomendaciones de reabastecimiento basadas en los patrones visualizados en este dashboard.

### Recursos adicionales

| Recurso | Descripción | Enlace |
|---|---|---|
| Documentación oficial Microsoft | Conectar Power BI a Excel | [learn.microsoft.com](https://learn.microsoft.com/es-es/power-bi/connect-data/desktop-connect-excel) |
| Documentación oficial Microsoft | Panel Analytics en Power BI | [learn.microsoft.com](https://learn.microsoft.com/es-es/power-bi/transform-model/desktop-analytics-pane) |
| Documentación oficial Microsoft | Pronóstico en gráficos de línea | [learn.microsoft.com](https://learn.microsoft.com/es-es/power-bi/visuals/power-bi-visualization-forecasting) |
| Canal YouTube oficial | Microsoft Power BI — tutoriales en español | [youtube.com/@MicrosoftPowerBI](https://www.youtube.com/@MicrosoftPowerBI) |
| Descarga gratuita | Power BI Desktop | [microsoft.com/power-bi](https://www.microsoft.com/es-es/download/details.aspx?id=58494) |

---

> **Nota para el instructor:** Si algún participante no logra activar el pronóstico nativo por limitaciones de versión de Power BI o problemas con el tipo de dato de fecha, puede usar el archivo *checkpoint* `Dashboard_Demanda_P22_checkpoint.pbix` que incluye el dashboard completo hasta el Paso 7 con el pronóstico ya configurado, permitiendo que el participante continúe desde el Paso 8 (configuración final) sin perder el hilo de la práctica.

---

# Identificación de productos de alta frecuencia de picking mediante análisis de demanda

## 1. Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Práctica número** | 23 |
| **Práctica previa requerida** | Práctica 22 (archivo `.pbix` base) |
| **Herramienta principal** | Power BI Desktop |

---

## 2. Descripción General

Esta práctica amplía el dashboard construido en la Práctica 22 para enfocarse en el análisis operativo del proceso de **picking** en almacén. El participante incorporará segmentadores dinámicos, un gráfico de ranking de productos y una línea de referencia estadística que permiten identificar qué productos concentran el mayor volumen de unidades demandadas y en qué períodos la demanda supera el umbral operativo normal.

Al finalizar, el dashboard contará con una capa analítica orientada a decisiones concretas de planificación de personal, asignación de recursos y priorización de rutas de picking, conectando directamente los resultados visuales con la gestión operativa del almacén.

---

## 3. Objetivos de Aprendizaje

Al completar esta práctica, el participante será capaz de:

- [ ] Agregar y configurar segmentadores (slicers) de **Categoría** y **Producto** conectados a todas las visualizaciones del dashboard en Power BI Desktop.
- [ ] Crear un gráfico de barras horizontales ordenado de mayor a menor para identificar visualmente el **top 3 de productos** con mayor volumen acumulado de unidades demandadas (alta frecuencia de picking).
- [ ] Activar una **línea de referencia constante** en el Analytics Pane del gráfico de línea existente, usando el valor `Promedio + 1 Desviación Estándar` como umbral para detectar picos de demanda.
- [ ] Documentar los hallazgos operativos en un **cuadro de texto** dentro del dashboard, identificando los 3 productos de alta frecuencia y los períodos de pico con su interpretación logística.
- [ ] Exportar la página del dashboard como **imagen** desde Power BI Desktop.

---

## 4. Prerrequisitos

### Conocimiento previo requerido

| Área | Nivel requerido |
|---|---|
| Proceso de picking en operaciones de almacén | Comprensión básica (qué es, por qué impacta la operación) |
| Promedio y desviación estándar aplicados a demanda | Familiaridad conceptual (calculados en Prácticas anteriores) |
| Navegación básica en Power BI Desktop | Completar Práctica 22 |
| Concepto de segmentador (slicer) en Power BI | Deseable, no obligatorio |

### Acceso y archivos requeridos

| Recurso | Estado requerido |
|---|---|
| Power BI Desktop (versión 2024 o superior) | Instalado y funcional |
| Archivo `.pbix` guardado de la Práctica 22 | Disponible en disco local |
| Dataset base de 8 SKUs / 24 meses (ya cargado en el `.pbix`) | Cargado en el modelo de Power BI |
| Valores de Promedio y Desviación Estándar de la demanda total | Calculados en prácticas previas (se usarán como referencia numérica) |

> **Nota para el instructor:** Si algún participante no cuenta con el archivo `.pbix` de la Práctica 22, distribuir el archivo de *checkpoint* correspondiente antes de iniciar esta práctica.

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Resolución de pantalla | 1280 × 768 | 1920 × 1080 |
| Espacio en disco libre | 2 GB | 2 GB+ |
| Conexión a internet | 10 Mbps | 20 Mbps+ |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Power BI Desktop | Más reciente (2024) | Descarga gratuita en microsoft.com/power-bi |
| Microsoft Excel | 2016 o superior | Para consultar valores estadísticos de referencia |
| Navegador web | Edge / Chrome actualizado | Para exportar o compartir si se requiere |

### Configuración previa al inicio de la práctica

Antes de comenzar, verificar los siguientes puntos:

1. Abrir Power BI Desktop.
2. Ir a **Archivo → Abrir** y cargar el archivo `.pbix` guardado en la Práctica 22.
3. Confirmar que la página del dashboard muestra el gráfico de línea de demanda construido previamente.
4. Tener a mano (en papel o en Excel abierto) los siguientes valores estadísticos del dataset:
   - **Promedio mensual de unidades demandadas** (calculado en práctica previa).
   - **Desviación estándar** de la demanda mensual.
   - **Valor umbral = Promedio + 1 Desviación Estándar** (suma de ambos valores).

> **Ejemplo de referencia:** Si el promedio es 320 unidades y la desviación estándar es 85 unidades, el umbral a ingresar en la línea de referencia será **405 unidades**. Ajustar con los valores reales del dataset de la práctica.

---

## 6. Procedimiento Paso a Paso

---

### Paso 1: Verificar el estado del archivo .pbix y preparar el lienzo

**Objetivo:** Confirmar que el dashboard de la Práctica 22 está correctamente cargado y dejar espacio en el lienzo para los nuevos elementos visuales.

**Instrucciones:**

1. Con el archivo `.pbix` abierto en Power BI Desktop, verificar que la **página del dashboard** (creada en la Práctica 22) está activa y visible en la parte inferior de la pantalla (pestaña de páginas).
2. Hacer clic en la página del dashboard para activarla.
3. Revisar que el gráfico de línea de demanda histórica esté presente y funcional.
4. **Reorganizar el espacio del lienzo** para dejar zonas libres:
   - Mover el gráfico de línea existente hacia la **mitad inferior del lienzo** (arrastrarlo desde su borde superior).
   - Dejar la **zona superior izquierda** libre (aproximadamente 30% del ancho del lienzo) para los segmentadores.
   - Dejar la **zona superior derecha** libre para el gráfico de barras de ranking.
5. Guardar el archivo: **Archivo → Guardar** (o `Ctrl + S`).

**Resultado esperado:** El lienzo muestra el gráfico de línea reposicionado en la parte inferior y dos zonas claramente libres en la parte superior para los nuevos elementos.

**Verificación:** El gráfico de línea sigue mostrando datos correctamente después de ser movido. No hay mensajes de error en el lienzo.

---

### Paso 2: Agregar el segmentador (slicer) de Categoría

**Objetivo:** Insertar un segmentador que permita filtrar todas las visualizaciones del dashboard por la columna `Categoría` del dataset.

**Instrucciones:**

1. En el panel **Visualizaciones** (lado derecho), hacer clic en el ícono de **Segmentador** (Slicer). Su ícono representa un embudo con líneas horizontales. Si no lo ubicas fácilmente, posicionar el cursor sobre los íconos hasta encontrar el que dice *"Segmentador"*.
2. Se creará un segmentador vacío en el lienzo. Arrastrarlo a la **zona superior izquierda** del lienzo.
3. Con el segmentador seleccionado, ir al panel **Campos** (lado derecho) y expandir la tabla de datos del dataset.
4. Arrastrar el campo **`Categoría`** al área **Campo** del segmentador (o hacer clic en la casilla junto al campo `Categoría` en el panel de campos).
5. El segmentador mostrará las categorías disponibles del dataset (por ejemplo: *Alimentos*, *Bebidas*, *Limpieza*, etc., según los datos del instructor).
6. Ajustar el tamaño del segmentador para que sea legible: aproximadamente **200 px de ancho × 150 px de alto**.
7. **Cambiar el formato del segmentador** a lista vertical:
   - Con el segmentador seleccionado, ir al panel **Formato** (ícono de rodillo de pintura).
   - Expandir la sección **Configuración del segmentador → Opciones**.
   - En **Estilo**, seleccionar **"Lista"** (si no está ya seleccionado).
8. Agregar un **título** al segmentador:
   - En el panel Formato, expandir **Título del segmentador**.
   - Activar el título y escribir: `Filtrar por Categoría`.

**Resultado esperado:** El segmentador muestra una lista con todas las categorías del dataset. Al hacer clic en una categoría, el gráfico de línea existente se filtra automáticamente (esto confirma que la conexión entre visual y datos es correcta).

**Verificación:** Hacer clic en una categoría del segmentador y observar que el gráfico de línea cambia para mostrar solo los datos de esa categoría. Hacer clic en el ícono de borrar filtro (borrador o "×" en el segmentador) para restablecer la vista completa.

---

### Paso 3: Agregar el segmentador (slicer) de Producto

**Objetivo:** Insertar un segundo segmentador para filtrar por `Producto` (o `SKU`), conectado a todas las visualizaciones.

**Instrucciones:**

1. Hacer clic en un área vacía del lienzo para **deseleccionar** el segmentador anterior.
2. En el panel **Visualizaciones**, hacer clic nuevamente en el ícono de **Segmentador**.
3. Se creará un nuevo segmentador vacío. Arrastrarlo justo **debajo del segmentador de Categoría** (zona izquierda del lienzo).
4. Con el nuevo segmentador seleccionado, arrastrar el campo **`Producto`** (o `SKU` / `Descripción`, según el nombre en el dataset) al área **Campo** del segmentador.
5. El segmentador mostrará los 8 productos/SKUs del dataset.
6. Ajustar el tamaño a aproximadamente **200 px de ancho × 200 px de alto** para que todos los productos sean visibles sin desplazamiento.
7. Configurar el formato:
   - **Estilo:** Lista.
   - **Título:** `Filtrar por Producto`.
   - Activar **"Selección múltiple"** si se desea permitir filtrar por más de un producto simultáneamente:
     - Panel Formato → **Configuración del segmentador → Selección** → activar *"Selección múltiple con CTRL"*.
8. Guardar el archivo: `Ctrl + S`.

**Resultado esperado:** El lienzo ahora muestra dos segmentadores en la zona izquierda superior: uno para Categoría y otro para Producto. Ambos están conectados al modelo de datos y filtran el gráfico de línea al interactuar con ellos.

**Verificación:** Seleccionar un producto específico en el segmentador de Producto. El gráfico de línea debe mostrar únicamente la demanda de ese producto. Verificar también que al seleccionar una Categoría, el segmentador de Producto muestre solo los productos pertenecientes a esa categoría (filtrado en cascada automático de Power BI).

---

### Paso 4: Crear el gráfico de barras horizontales para ranking de productos

**Objetivo:** Construir una visualización de barras horizontales que muestre el volumen acumulado de unidades demandadas por producto, ordenado de mayor a menor, para identificar el top 3 de productos con alta frecuencia de picking.

**Instrucciones:**

1. Hacer clic en un área vacía del lienzo (zona superior derecha).
2. En el panel **Visualizaciones**, seleccionar el ícono de **Gráfico de barras agrupadas** (Clustered Bar Chart). Este ícono muestra barras horizontales. Asegurarse de elegir el de barras **horizontales** (no columnas verticales).

   > **Distinción importante:** En Power BI, el *"Gráfico de barras agrupadas"* (Clustered Bar Chart) tiene barras horizontales. El *"Gráfico de columnas agrupadas"* tiene barras verticales. Para ranking de productos, las barras horizontales son más legibles cuando los nombres de los SKUs son largos.

3. Con el gráfico seleccionado, configurar los campos en el panel **Campos del visual**:
   - **Eje Y (categorías):** Arrastrar el campo `Producto` (o `SKU` / `Descripción`).
   - **Eje X (valores):** Arrastrar el campo `Unidades_Vendidas` (o `Unidades_Demandadas`, según el nombre en el dataset). Power BI aplicará automáticamente la función **SUMA** para acumular el total por producto.
   - **Leyenda:** Dejar vacío por ahora.

4. **Ordenar el gráfico de mayor a menor:**
   - Hacer clic en los **tres puntos (...)** en la esquina superior derecha del gráfico.
   - Seleccionar **"Ordenar por"** → elegir `Suma de Unidades_Vendidas`.
   - Hacer clic nuevamente en los tres puntos → **"Orden descendente"** (de mayor a menor).
   - El producto con mayor demanda acumulada aparecerá en la parte superior.

5. **Formatear el gráfico** para hacerlo más legible:
   - Panel Formato → **Título del gráfico:** Escribir `Ranking de Productos por Volumen de Picking`.
   - Panel Formato → **Etiquetas de datos:** Activar para mostrar el valor numérico al final de cada barra.
   - Panel Formato → **Colores de datos:** Seleccionar un color que contraste bien (azul corporativo o el color del tema del dashboard).
   - Ajustar el tamaño del gráfico para que los 8 productos sean visibles sin scroll: aproximadamente **450 px de ancho × 280 px de alto**.

6. Guardar el archivo: `Ctrl + S`.

**Resultado esperado:** El gráfico muestra 8 barras horizontales ordenadas de mayor a menor demanda acumulada. Las 3 barras superiores representan los productos con mayor frecuencia de picking. Los valores numéricos son visibles al final de cada barra.

**Verificación:** Identificar y anotar mentalmente (o en papel) los nombres de los **3 productos en la parte superior** del gráfico. Estos son los productos de alta frecuencia de picking que se documentarán en el Paso 6. Probar el filtrado: seleccionar una categoría en el segmentador de Categoría y verificar que el gráfico de barras también se actualiza.

---

### Paso 5: Agregar la línea de referencia de umbral en el gráfico de línea

**Objetivo:** Activar una línea de referencia constante en el gráfico de línea existente, posicionada en el valor `Promedio + 1 Desviación Estándar`, para identificar visualmente los períodos donde la demanda supera el umbral normal (picos de demanda).

**Instrucciones:**

1. Hacer clic sobre el **gráfico de línea** (el que muestra la demanda histórica mensual, construido en la Práctica 22) para seleccionarlo.

2. En el panel derecho, localizar el **panel de Analytics** (ícono de lupa con línea de tendencia, generalmente el tercer ícono en la barra de paneles después de Visualizaciones y Campos). Hacer clic en él.

   > **Nota:** El panel de Analytics solo está disponible cuando se tiene seleccionado un visual compatible (gráfico de línea, columnas, área). Si el ícono aparece atenuado, verificar que el gráfico de línea esté seleccionado.

3. En el panel de Analytics, localizar la opción **"Línea constante"** (Constant Line) y hacer clic en el botón **"+ Agregar"** o **"Agregar línea"**.

4. Se desplegará una sección de configuración para la línea constante. Configurar los siguientes parámetros:

   | Parámetro | Valor a ingresar |
   |---|---|
   | **Valor** | Ingresar el valor numérico de `Promedio + 1 Desviación Estándar` (calculado en prácticas previas). Ejemplo: si Promedio = 320 y Desv. Est. = 85, ingresar **405**. |
   | **Color de línea** | Rojo o naranja (para destacar el umbral de alerta). |
   | **Estilo de línea** | Discontinua (punteada) para diferenciarla de las líneas de datos. |
   | **Posición** | Detrás de los datos (Behind). |
   | **Etiqueta de datos** | Activar y escribir: `Umbral: Promedio + 1σ` |
   | **Transparencia** | 0% (línea completamente visible). |

5. Verificar visualmente que la línea aparece en el gráfico como una línea horizontal discontinua de color rojo/naranja atravesando el gráfico.

6. **Identificar los períodos de pico:** Observar en qué puntos del eje X (meses) la línea de demanda **supera** la línea de referencia roja. Estos son los meses de pico de demanda. Anotar los 2-3 períodos más evidentes (por ejemplo: *Diciembre 2022, Marzo 2023, Noviembre 2023*).

7. Guardar el archivo: `Ctrl + S`.

**Resultado esperado:** El gráfico de línea muestra la evolución mensual de la demanda con una línea horizontal discontinua roja que representa el umbral estadístico. Los meses donde la línea de demanda cruza hacia arriba de la línea roja son claramente identificables como períodos de pico.

**Verificación:** La línea de referencia debe ser visible en el gráfico y estar etiquetada. Deben identificarse al menos 2 períodos donde la demanda supera el umbral. Si la línea de referencia no aparece, verificar que el gráfico de línea esté seleccionado antes de abrir el panel de Analytics.

---

### Paso 6: Documentar los hallazgos en un cuadro de texto dentro del dashboard

**Objetivo:** Insertar un cuadro de texto en el dashboard que documente los 3 productos de alta frecuencia de picking y los períodos de pico identificados, con una interpretación operativa concisa para la planificación del área de picking.

**Instrucciones:**

1. Hacer clic en un área vacía del lienzo para deseleccionar todos los visuales.

2. En la cinta superior de Power BI Desktop, ir a la pestaña **"Insertar"** → hacer clic en **"Cuadro de texto"** (Text Box).

3. Se creará un cuadro de texto en el lienzo. Arrastrarlo a la **zona inferior derecha** del dashboard (debajo del gráfico de barras o en el espacio disponible).

4. Ajustar el tamaño del cuadro de texto a aproximadamente **400 px de ancho × 200 px de alto**.

5. Hacer clic dentro del cuadro de texto y escribir el siguiente contenido, **adaptando los datos reales** identificados en los pasos anteriores:

```
HALLAZGOS OPERATIVOS – ANÁLISIS DE PICKING

📦 TOP 3 PRODUCTOS – ALTA FRECUENCIA DE PICKING:
1. [Nombre Producto 1] – [X,XXX] unidades acumuladas
2. [Nombre Producto 2] – [X,XXX] unidades acumuladas
3. [Nombre Producto 3] – [X,XXX] unidades acumuladas

⚠️ PERÍODOS CON PICOS DE DEMANDA (superan umbral Prom. + 1σ = [VALOR]):
• [Mes/Año 1]: Demanda estimada [X,XXX] unidades
• [Mes/Año 2]: Demanda estimada [X,XXX] unidades
• [Mes/Año 3]: Demanda estimada [X,XXX] unidades (si aplica)

📋 IMPLICACIÓN OPERATIVA:
Los productos del top 3 requieren posicionamiento prioritario
en zona de picking de alta rotación. Durante los períodos de
pico, se recomienda reforzar personal de picking en al menos
20-30% y pre-posicionar stock en ubicaciones de fácil acceso
para reducir tiempos de desplazamiento y evitar quiebres.
```

6. **Formatear el texto** para mayor legibilidad:
   - Seleccionar la línea de título y aplicar **negrita** y tamaño de fuente 11.
   - Seleccionar los nombres de productos y aplicar negrita.
   - Cambiar el color del encabezado a azul oscuro o el color corporativo del dashboard.

7. Agregar un **borde al cuadro de texto**:
   - Con el cuadro de texto seleccionado, ir al panel **Formato**.
   - Activar **Borde** y seleccionar color gris oscuro, grosor 1 pt.
   - Activar **Fondo** con color blanco o gris muy claro para que contraste con el lienzo.

8. Guardar el archivo: `Ctrl + S`.

**Resultado esperado:** El dashboard muestra un cuadro de texto bien formateado con los 3 productos de alta frecuencia de picking, los períodos de pico identificados y una interpretación operativa clara. El contenido es específico con los datos reales del dataset (no genérico).

**Verificación:** Leer el cuadro de texto completo y confirmar que los nombres de productos y valores numéricos corresponden exactamente a lo que muestra el gráfico de barras horizontales y el gráfico de línea con la línea de referencia.

---

### Paso 7: Revisar la interacción entre todos los elementos del dashboard

**Objetivo:** Verificar que los segmentadores controlan correctamente todas las visualizaciones del dashboard y que el análisis integrado es coherente.

**Instrucciones:**

1. Probar el **filtrado por Categoría:**
   - Hacer clic en una categoría del segmentador de Categoría.
   - Verificar que el gráfico de línea, el gráfico de barras horizontales **y** el cuadro de texto (que es estático) se comportan correctamente.
   - El cuadro de texto no se filtra automáticamente (es texto fijo), lo cual es esperado.

2. Probar el **filtrado por Producto:**
   - Hacer clic en el producto que aparece en la posición #1 del ranking en el gráfico de barras.
   - El gráfico de línea debe mostrar únicamente la demanda histórica de ese producto.
   - Verificar que la línea de referencia permanece visible (el umbral no cambia al filtrar, ya que es una constante).

3. **Ajustar interacciones si es necesario:**
   - Si algún visual no responde a los segmentadores, hacer clic en el segmentador problemático.
   - Ir a la cinta **"Formato"** → **"Editar interacciones"**.
   - Verificar que el ícono de filtro (embudo) esté activado para todos los visuales que deben responder al segmentador.
   - Hacer clic en **"Editar interacciones"** nuevamente para salir del modo de edición.

4. **Verificación final del dashboard completo:** El dashboard debe mostrar de forma integrada:
   - Segmentador de Categoría (izquierda superior).
   - Segmentador de Producto (izquierda, debajo del anterior).
   - Gráfico de barras horizontales de ranking (derecha superior).
   - Gráfico de línea de demanda histórica con línea de referencia (parte inferior).
   - Cuadro de texto con hallazgos operativos (esquina inferior derecha o zona disponible).

5. Guardar el archivo: `Ctrl + S`.

**Resultado esperado:** Todos los visuales responden coherentemente a los segmentadores. El dashboard presenta una vista integrada y operativa del análisis de picking.

**Verificación:** Seleccionar el producto #1 del ranking en el segmentador de Producto y confirmar que el gráfico de línea muestra su demanda individual. Limpiar todos los filtros (hacer clic en el ícono de borrar en ambos segmentadores) y confirmar que se vuelve a la vista completa.

---

### Paso 8: Exportar el dashboard como imagen

**Objetivo:** Exportar la página del dashboard como imagen para documentación, entrega al instructor o uso en reportes externos.

**Instrucciones:**

1. Asegurarse de que **ningún filtro esté activo** en los segmentadores (vista completa del dashboard sin filtros aplicados). Esto garantiza que la imagen exportada muestre el análisis completo.

2. En la cinta superior de Power BI Desktop, ir a la pestaña **"Archivo"**.

3. Seleccionar **"Exportar"** → **"Exportar como PowerPoint"** o usar el método alternativo:

   > **Método recomendado para exportar como imagen:**
   > - En la cinta superior, ir a **"Archivo" → "Imprimir"** (o usar `Ctrl + P`).
   > - En el cuadro de diálogo de impresión, cambiar la impresora a **"Microsoft Print to PDF"** o **"Guardar como PDF"**.
   > - Esto generará un PDF de alta calidad del dashboard.
   >
   > **Método alternativo (captura de pantalla directa):**
   > - Usar la tecla `Windows + Shift + S` (Recorte y anotación de Windows).
   > - Seleccionar el área del lienzo del dashboard.
   > - Guardar la captura con el nombre: `Lab05-02_Dashboard_Picking_[TuNombre].png`.

4. **Método oficial de Power BI (si está disponible):**
   - Ir a la cinta **"Archivo"** → **"Exportar"** → **"Exportar a PDF"**.
   - Seleccionar la página del dashboard.
   - Guardar el archivo con el nombre: `Lab05-02_Dashboard_Picking_[TuNombre].pdf`.

5. Verificar que el archivo exportado (imagen o PDF) muestra todos los elementos del dashboard correctamente: ambos segmentadores, el gráfico de barras, el gráfico de línea con la línea de referencia y el cuadro de texto con los hallazgos.

6. Guardar el archivo `.pbix` por última vez: `Ctrl + S`.

**Resultado esperado:** Un archivo de imagen (`.png`) o PDF que muestra el dashboard completo de análisis de picking, listo para ser entregado al instructor o incluido en un reporte.

**Verificación:** Abrir el archivo de imagen o PDF exportado y confirmar que todos los elementos son legibles y que la línea de referencia roja es claramente visible en el gráfico de línea.

---

## 7. Validación y Pruebas

Al finalizar la práctica, el participante debe poder responder afirmativamente a cada uno de los siguientes puntos de control:

| # | Punto de validación | ¿Cumplido? |
|---|---|---|
| 1 | El dashboard contiene exactamente **2 segmentadores** (Categoría y Producto) visibles y funcionales. | ☐ |
| 2 | El gráfico de barras horizontales muestra **8 productos ordenados de mayor a menor** demanda acumulada. | ☐ |
| 3 | El **top 3 de productos** (mayor frecuencia de picking) es claramente identificable en la parte superior del gráfico de barras. | ☐ |
| 4 | El gráfico de línea muestra una **línea de referencia horizontal discontinua** en el valor `Promedio + 1σ` con etiqueta visible. | ☐ |
| 5 | Se identificaron al menos **2 períodos donde la demanda supera el umbral** (picos de demanda). | ☐ |
| 6 | El cuadro de texto documenta los **3 productos de alta frecuencia** y los **períodos de pico** con interpretación operativa. | ☐ |
| 7 | Todos los visuales **responden a los segmentadores** (interacciones correctamente configuradas). | ☐ |
| 8 | El archivo `.pbix` está **guardado** con todos los cambios. | ☐ |
| 9 | Se generó y guardó la **exportación del dashboard** (imagen o PDF). | ☐ |

**Prueba de integración final:**
Seleccionar en el segmentador de Categoría la categoría que contiene el producto #1 del ranking. Verificar que:
- El gráfico de barras muestra solo los productos de esa categoría.
- El gráfico de línea muestra la demanda filtrada por esa categoría.
- La línea de referencia permanece en el mismo valor constante (no se modifica con el filtro).

Esto confirma que la línea de referencia es correctamente una **constante estadística** y no un valor dinámico que cambia con los filtros.

---

## 8. Resolución de Problemas

### Problema 1: El panel de Analytics no muestra la opción "Línea constante" o aparece atenuado

**Síntoma:** Al hacer clic en el ícono de Analytics (lupa), las opciones de líneas de referencia aparecen en gris y no se pueden seleccionar, o la opción "Línea constante" directamente no aparece en el panel.

**Causa:** El panel de Analytics con la opción de línea constante solo está disponible para ciertos tipos de visualizaciones. Si el gráfico de línea fue creado como un **gráfico de área**, **gráfico combinado** o si tiene configuraciones especiales de ejes, es posible que algunas opciones del Analytics Pane no estén disponibles. También puede ocurrir si el visual no tiene un campo numérico continuo en el eje X.

**Solución:**
1. Seleccionar el gráfico de línea y verificar en el panel de Visualizaciones que el tipo de visual sea específicamente **"Gráfico de líneas"** (Line Chart), no un gráfico de área ni combinado.
2. Si el visual es de otro tipo, hacer clic en el ícono de **"Gráfico de líneas"** en el panel de Visualizaciones para convertirlo (los datos se conservan).
3. Verificar que el eje X del gráfico tenga el campo `Fecha` o `Mes` correctamente configurado como tipo de dato **Fecha** (no texto).
4. Si el problema persiste, intentar crear un nuevo gráfico de línea desde cero: insertar un nuevo visual de tipo línea, asignar los mismos campos y luego intentar agregar la línea constante desde el Analytics Pane.
5. Como alternativa temporal, agregar una **línea de promedio** (Average Line) desde el Analytics Pane, que sí está disponible para más tipos de visuales, y ajustar manualmente el valor mostrado.

---

### Problema 2: Los segmentadores no filtran el gráfico de barras horizontales (o algún visual no responde)

**Síntoma:** Al seleccionar una categoría o producto en los segmentadores, el gráfico de línea se filtra correctamente, pero el gráfico de barras horizontales no cambia y sigue mostrando todos los productos con los mismos valores.

**Causa:** Las interacciones entre visuales en Power BI son configurables. Es posible que la interacción entre el segmentador y el gráfico de barras esté configurada en modo **"Ninguno"** (sin filtrado) en lugar del modo **"Filtro"**. Esto puede ocurrir si el gráfico de barras fue creado después de que se configuraron las interacciones, o si accidentalmente se desactivó la interacción.

**Solución:**
1. Hacer clic en el **segmentador de Categoría** para seleccionarlo.
2. En la cinta superior, ir a la pestaña **"Formato"** (que aparece en la cinta cuando un visual está seleccionado).
3. Hacer clic en **"Editar interacciones"**. Aparecerán íconos pequeños en la esquina superior derecha de cada visual del dashboard.
4. Localizar el gráfico de barras horizontales. Si el ícono activo es el círculo con línea diagonal (🚫, modo "Ninguno"), hacer clic en el ícono de **embudo/filtro** (🔽) para activar el filtrado.
5. Repetir el mismo proceso seleccionando el **segmentador de Producto** y verificando que el gráfico de barras también tenga el modo filtro activado.
6. Hacer clic en **"Editar interacciones"** nuevamente para salir del modo de edición.
7. Probar los segmentadores nuevamente para confirmar que todos los visuales responden correctamente.

---

## 9. Limpieza del Entorno

Al finalizar la práctica, realizar las siguientes acciones de cierre:

1. **Guardar el archivo `.pbix` final:**
   - Ir a **Archivo → Guardar como**.
   - Nombrar el archivo: `Lab05-02_Practica23_Dashboard_Picking_[TuNombre].pbix`.
   - Guardar en la carpeta de trabajo del curso (por ejemplo: `C:\LabsCurso\Practica23\`).

2. **Verificar el archivo exportado:**
   - Confirmar que el archivo de imagen o PDF del dashboard existe en la carpeta de trabajo.
   - Nombrar el archivo exportado: `Lab05-02_Dashboard_Picking_[TuNombre].pdf` (o `.png`).

3. **Entregar al instructor** (según indicaciones del curso):
   - El archivo `.pbix` guardado.
   - La imagen o PDF del dashboard exportado.
   - Los valores numéricos utilizados: Promedio, Desviación Estándar y Umbral calculado.

4. **No cerrar Power BI Desktop** si la práctica siguiente se realizará de forma inmediata. El archivo `.pbix` de esta práctica será el punto de partida para prácticas posteriores en la secuencia del curso.

> **Recordatorio de secuencialidad:** El archivo `.pbix` generado en esta práctica (Práctica 23) será utilizado como insumo en las prácticas siguientes. Asegurarse de que el archivo esté correctamente guardado y accesible antes de continuar.

---

## 10. Resumen

### Conceptos aplicados en esta práctica

En esta práctica se aplicaron cuatro capacidades clave de Power BI Desktop en un contexto operativo de análisis de picking:

| Herramienta | Aplicación logística |
|---|---|
| **Segmentadores (Slicers)** | Permiten al analista o supervisor de almacén filtrar el dashboard por categoría o producto específico, respondiendo preguntas operativas en tiempo real sin modificar el modelo de datos. |
| **Gráfico de barras horizontales ordenado** | Proporciona un ranking visual inmediato de los productos que concentran mayor volumen de picking, facilitando decisiones de posicionamiento de inventario y asignación de rutas de picking. |
| **Línea de referencia constante (Analytics Pane)** | Establece un umbral estadístico basado en `Promedio + 1σ` que permite identificar períodos de demanda excepcional, esencial para la planificación anticipada de recursos humanos y materiales en el área de picking. |
| **Cuadro de texto con hallazgos operativos** | Convierte el análisis cuantitativo en recomendaciones accionables documentadas dentro del mismo dashboard, cerrando el ciclo entre análisis de datos y decisión operativa. |

### Conexión con el ciclo de forecasting logístico

El análisis de alta frecuencia de picking realizado en esta práctica está directamente conectado con la estructura metodológica del pronóstico (Baseline + Estacionalidad + Iniciativas) trabajada en el curso:

- Los **productos del top 3** identificados corresponden al componente **Baseline** del pronóstico: son los SKUs que sistemáticamente generan mayor volumen y deben estar en el núcleo de cualquier modelo de demanda.
- Los **períodos de pico** detectados con la línea de referencia corresponden al componente de **Estacionalidad** o **Eventos Extraordinarios**: momentos donde la demanda supera el patrón normal y requieren planificación diferenciada.
- La **implicación operativa** documentada en el cuadro de texto traduce el análisis estadístico en una **decisión de inventario y nivel de servicio**: pre-posicionamiento de stock, refuerzo de personal y ajuste de capacidad de picking.

### Recursos adicionales

- [Documentación oficial: Segmentadores en Power BI](https://learn.microsoft.com/es-es/power-bi/visuals/power-bi-visualization-slicers)
- [Documentación oficial: Panel de análisis en Power BI Desktop](https://learn.microsoft.com/es-es/power-bi/transform-model/desktop-analytics-pane)
- [Documentación oficial: Exportar reportes de Power BI](https://learn.microsoft.com/es-es/power-bi/consumer/end-user-export)
- [Documentación oficial: Editar interacciones entre visualizaciones](https://learn.microsoft.com/es-es/power-bi/create-reports/service-reports-visual-interactions)

---

---

# Práctica 24 — Visualización comparativa entre demanda histórica y pronóstico seleccionado

## Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 12 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Práctica número** | 24 |
| **Módulo** | Capítulo 5 — Visualización de pronósticos en Power BI Desktop |
| **Archivo de entrada** | `Practica21_Pronostico_Seleccionado.xlsx` (generado en Práctica 21) + archivo `.pbix` base (generado en Práctica 22) |
| **Archivo de salida** | `Practica24_Visual_Comparativa.pbix` + `Dataset_Comparativo_Largo.xlsx` |

---

## Descripción General

En esta práctica construirás una visualización comparativa en Power BI Desktop que muestra simultáneamente la demanda histórica real y los valores del modelo de pronóstico seleccionado en la Práctica 21, ambas series sobre el mismo eje temporal. Para lograrlo, primero estructurarás los datos en Excel en **formato largo** (*long format*), con una columna `Tipo` que diferencia entre registros `"Real"` y `"Pronóstico"`, y luego importarás ese dataset a Power BI siguiendo el proceso de conexión y transformación aprendido en la Lección 5.1. El resultado final es un gráfico de líneas dual con colores diferenciados, una tabla comparativa con formato condicional y un cuadro de texto con observaciones logísticas sobre los períodos de mayor brecha.

---

## Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] **Preparar** en Excel un dataset en formato largo que combine 18 meses de demanda histórica y 6 períodos de pronóstico, con columnas `Periodo`, `Tipo` y `Valor`, listo para importación a Power BI.
- [ ] **Importar** el dataset combinado a Power BI Desktop usando el flujo *Obtener datos → Excel → Transformar datos → Cerrar y aplicar*, verificando tipos de datos en Power Query.
- [ ] **Construir** un gráfico de líneas con dos series diferenciadas (`Real` en azul, `Pronóstico` en naranja) sobre el mismo eje temporal dentro de una nueva página del archivo `.pbix`.
- [ ] **Crear** una tabla comparativa en Power BI con formato condicional que resalte en rojo las diferencias absolutas superiores al 15%.
- [ ] **Documentar** en un cuadro de texto dentro de Power BI al menos dos observaciones sobre períodos de sobreestimación o subestimación y su impacto en decisiones de reabastecimiento.

---

## Prerrequisitos

### Conocimiento previo requerido

| Área | Detalle |
|---|---|
| **Práctica 21 completada** | Contar con el modelo de pronóstico seleccionado (Promedio Móvil, Suavización Exponencial o Regresión Lineal) y sus 6 valores proyectados calculados en Excel |
| **Práctica 22 completada** | Contar con el archivo `.pbix` base con al menos una página de visualizaciones previas |
| **Lección 5.1 revisada** | Comprender el flujo: Preparar Excel → Conectar Power BI → Transformar en Power Query → Cerrar y aplicar |
| **Conceptos de sobreestimación/subestimación** | Entender que pronóstico > real genera exceso de inventario; pronóstico < real genera quiebre de stock |
| **Formato largo (long format)** | Saber que en formato largo cada observación ocupa una fila, en contraste con el formato ancho donde cada serie es una columna |

### Acceso y archivos necesarios

| Recurso | Estado requerido |
|---|---|
| Power BI Desktop (versión 2024 o superior) | Instalado y funcional |
| Microsoft Excel 2016 o superior | Abierto con archivo de Práctica 21 disponible |
| `Practica21_Pronostico_Seleccionado.xlsx` | Disponible en carpeta de trabajo local |
| `Practica22_Dashboard_Base.pbix` | Disponible en carpeta de trabajo local |
| Carpeta de trabajo local | `C:\LabLogistica\Practica24\` (o equivalente) |

> **Nota para el instructor:** Si algún participante no completó la Práctica 21, entregar el archivo de *checkpoint* `Checkpoint_P21_Pronostico.xlsx` que contiene los valores de pronóstico de referencia para el SKU de práctica.

---

## Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 2 GB | 5 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Conexión a internet | No requerida para esta práctica | — |

### Software requerido

| Software | Versión | Uso en esta práctica |
|---|---|---|
| Microsoft Excel | 2016 o superior (Microsoft 365 recomendado) | Preparar dataset en formato largo |
| Power BI Desktop | Versión más reciente (gratuita) | Importar datos, construir visualizaciones |

### Configuración inicial del entorno

Antes de comenzar, verificar que la carpeta de trabajo esté organizada correctamente:

```
C:\LabLogistica\Practica24\
├── Practica21_Pronostico_Seleccionado.xlsx   ← archivo de entrada
├── Practica22_Dashboard_Base.pbix            ← archivo .pbix de entrada
```

Si la carpeta no existe, crearla con el siguiente comando en el símbolo del sistema (CMD):

```cmd
mkdir C:\LabLogistica\Practica24
```

Luego copiar los archivos de entrada a esa carpeta antes de comenzar los pasos.

---

## Pasos del Laboratorio

---

### Paso 1 — Revisar los datos de origen en Excel

**Objetivo:** Identificar los 18 valores históricos reales y los 6 valores de pronóstico disponibles en el archivo de la Práctica 21, antes de restructurarlos.

#### Instrucciones

1. Abre el archivo `Practica21_Pronostico_Seleccionado.xlsx` en Microsoft Excel.
2. Localiza la hoja que contiene el historial de demanda. Típicamente se llama `Historial` o `Datos_Base`.
3. Identifica la columna de **períodos** (etiquetada como `Periodo`, `Mes` o `Fecha`) y la columna de **demanda real** (etiquetada como `Demanda_Real` o `Unidades`). Deberías ver **18 filas de datos históricos** (meses 1 al 18).
4. Localiza ahora la hoja o sección que contiene el **pronóstico seleccionado**. Identifica las columnas `Periodo` y el nombre del modelo ganador (por ejemplo, `Pronostico_SES`, `Pronostico_PM3` o `Pronostico_RegLineal`). Deberías ver **6 filas de valores proyectados** (meses 19 al 24).
5. Anota en papel o en una celda auxiliar:
   - El **nombre del SKU** analizado (por ejemplo, `SKU-003`).
   - El **nombre del modelo de pronóstico seleccionado** (el de menor MAPE calculado en la Práctica 21).
   - El **rango de períodos**: histórico (meses 1–18) y proyectado (meses 19–24).

> **Referencia de la Lección 5.1:** Antes de importar a Power BI, es fundamental entender la estructura de los datos de origen. Una importación exitosa comienza con datos bien comprendidos y organizados.

#### Resultado esperado

Tienes claridad sobre los 24 registros que conformarán el dataset comparativo: 18 reales + 6 pronosticados, con sus respectivos períodos y valores numéricos identificados.

#### Verificación

✅ Puedes responder: ¿Cuántas filas de datos históricos existen? (Respuesta: 18)
✅ Puedes responder: ¿Cuántas filas de pronóstico existen? (Respuesta: 6)
✅ Sabes cuál es el modelo de pronóstico seleccionado y su columna de valores.

---

### Paso 2 — Construir el dataset en formato largo en Excel

**Objetivo:** Crear un nuevo archivo Excel con los 24 registros estructurados en **formato largo**, con tres columnas: `Periodo`, `Tipo` y `Valor`, listos para importar a Power BI.

#### Instrucciones

1. Abre un **nuevo libro de Excel** (Ctrl+N) o inserta una nueva hoja en el archivo existente.
2. En la **celda A1**, escribe el encabezado: `Periodo`
3. En la **celda B1**, escribe el encabezado: `Tipo`
4. En la **celda C1**, escribe el encabezado: `Valor`

5. A partir de la fila 2, ingresa los **18 registros históricos** con la etiqueta `Real` en la columna `Tipo`. La estructura debe verse así:

   | Periodo | Tipo | Valor |
   |---|---|---|
   | Mes 1 | Real | 320 |
   | Mes 2 | Real | 345 |
   | Mes 3 | Real | 298 |
   | … | Real | … |
   | Mes 18 | Real | 412 |

   > **Importante:** Usa los valores reales de tu archivo de Práctica 21. Los números de la tabla anterior son solo de referencia.

6. A continuación, en las filas 20 a 25, ingresa los **6 registros de pronóstico** con la etiqueta `Pronóstico` en la columna `Tipo`:

   | Periodo | Tipo | Valor |
   |---|---|---|
   | Mes 19 | Pronóstico | 398 |
   | Mes 20 | Pronóstico | 415 |
   | … | Pronóstico | … |
   | Mes 24 | Pronóstico | 441 |

   > **Importante:** Usa los valores del modelo de pronóstico seleccionado en tu Práctica 21.

7. Verifica que la tabla completa tenga exactamente **24 filas de datos** (filas 2 a 25) más 1 fila de encabezado.

8. **Convierte el rango en Tabla de Excel:**
   - Selecciona el rango `A1:C25`.
   - Presiona **Ctrl+T**.
   - Confirma que la casilla *"La tabla tiene encabezados"* esté marcada.
   - Haz clic en **Aceptar**.
   - En el campo *Nombre de tabla* (esquina superior izquierda de la cinta cuando la tabla está seleccionada), escribe: `Comparativo_Demanda_Pronostico`

9. Guarda el archivo como:
   - Nombre: `Dataset_Comparativo_Largo.xlsx`
   - Ubicación: `C:\LabLogistica\Practica24\`
   - Formato: Libro de Excel (.xlsx)

#### Resultado esperado

Un archivo `Dataset_Comparativo_Largo.xlsx` con una Tabla de Excel llamada `Comparativo_Demanda_Pronostico`, con 24 filas de datos y tres columnas bien definidas: `Periodo` (texto o número de mes), `Tipo` (valores `"Real"` o `"Pronóstico"`) y `Valor` (número entero).

#### Verificación

✅ La tabla tiene exactamente 24 filas de datos (sin contar encabezado).
✅ La columna `Tipo` contiene únicamente dos valores distintos: `Real` y `Pronóstico`.
✅ No hay celdas vacías en ninguna de las tres columnas.
✅ El nombre de la tabla en Excel es `Comparativo_Demanda_Pronostico`.

> **Consejo de tiempo:** Este paso debería completarse en aproximadamente 3 minutos. Si los datos de la Práctica 21 están bien organizados, simplemente copias y pegas los valores en el nuevo formato.

---

### Paso 3 — Importar el dataset a Power BI Desktop con Power Query

**Objetivo:** Conectar Power BI Desktop al archivo `Dataset_Comparativo_Largo.xlsx`, verificar los tipos de datos en Power Query y cargar la tabla al modelo.

#### Instrucciones

1. Abre Power BI Desktop y luego abre el archivo `.pbix` base generado en la Práctica 22:
   - Ve a *Archivo → Abrir informe → Examinar este dispositivo*
   - Navega hasta `C:\LabLogistica\Practica24\`
   - Selecciona `Practica22_Dashboard_Base.pbix` y haz clic en **Abrir**.

2. Con el archivo `.pbix` abierto, importa el nuevo dataset:
   - En la cinta superior, ve a **Inicio → Obtener datos**.
   - En el menú desplegable, selecciona **Excel**.

3. En el cuadro de diálogo de exploración de archivos:
   - Navega hasta `C:\LabLogistica\Practica24\`
   - Selecciona `Dataset_Comparativo_Largo.xlsx`
   - Haz clic en **Abrir**.

4. En el **Navegador de Power BI** que aparece:
   - Verás la tabla `Comparativo_Demanda_Pronostico` listada (gracias a que la convertiste en Tabla de Excel en el Paso 2).
   - Haz clic sobre el nombre de la tabla para ver la **vista previa** en el panel derecho.
   - Confirma que se muestran las columnas `Periodo`, `Tipo` y `Valor` con datos correctos.
   - **NO hagas clic en "Cargar"**. En su lugar, haz clic en **"Transformar datos"**.

5. Se abre **Power Query Editor**. Realiza las siguientes verificaciones y transformaciones:

   **a) Verificar tipo de dato de la columna `Periodo`:**
   - Observa el ícono a la izquierda del encabezado `Periodo`. Si muestra `ABC` (texto) y los valores son como "Mes 1", "Mes 2", etc., déjalo como texto. Si son números enteros, verifica que el ícono muestre `123`.
   - Si necesitas cambiar el tipo: haz clic en el ícono del tipo → selecciona el tipo correcto.

   **b) Verificar tipo de dato de la columna `Tipo`:**
   - Debe mostrar el ícono `ABC` (Texto). Si no es así, haz clic en el ícono → selecciona **Texto**.

   **c) Verificar tipo de dato de la columna `Valor`:**
   - Debe mostrar el ícono `1.2` (Número decimal) o `123` (Número entero). Si los valores son unidades enteras, selecciona **Número entero**.
   - Si el ícono no es correcto: haz clic en el ícono del tipo → selecciona **Número entero**.

   **d) Verificar que no haya filas en blanco:**
   - En el menú de Power Query, ve a **Inicio → Quitar filas → Quitar filas en blanco**.
   - Confirma que el conteo de filas en la barra inferior sigue mostrando **24 filas**.

6. Revisa el panel **"Pasos aplicados"** en el lado derecho de Power Query. Deberías ver al menos:
   - `Origen`
   - `Navegación`
   - `Tipo cambiado` (si modificaste algún tipo)

7. Haz clic en **"Cerrar y aplicar"** (botón en la esquina superior izquierda de Power Query).

8. Power BI cargará la tabla al modelo. En el panel **Datos** (lado derecho de Power BI), verifica que aparece la tabla `Comparativo_Demanda_Pronostico` con sus tres campos.

#### Resultado esperado

La tabla `Comparativo_Demanda_Pronostico` está disponible en el modelo de datos de Power BI con 24 filas, tipos de datos correctos y sin errores de carga. El panel de Datos muestra los campos `Periodo`, `Tipo` y `Valor`.

#### Verificación

✅ En el panel Datos de Power BI aparece `Comparativo_Demanda_Pronostico`.
✅ Al expandir la tabla en el panel Datos, se ven los tres campos: `Periodo`, `Tipo`, `Valor`.
✅ No aparece ningún triángulo amarillo de advertencia junto al nombre de la tabla.
✅ Power Query muestra 24 filas en el paso final antes de cerrar.

---

### Paso 4 — Crear una nueva página en el archivo .pbix y construir el gráfico de líneas dual

**Objetivo:** Agregar una nueva página al archivo `.pbix` y construir el gráfico de líneas con dos series diferenciadas (Real en azul, Pronóstico en naranja) sobre el mismo eje temporal.

#### Instrucciones

**Parte A: Crear nueva página**

1. En la barra de páginas en la parte inferior de Power BI Desktop, haz clic en el **ícono "+"** para agregar una nueva página.
2. Haz doble clic sobre el nombre de la nueva página (aparece como "Página 2" o similar) y renómbrala como: `Comparativa Real vs Pronóstico`
3. Presiona **Enter** para confirmar.

**Parte B: Insertar el gráfico de líneas**

4. Con la nueva página activa, en el panel **Visualizaciones** (lado derecho), haz clic en el ícono de **Gráfico de líneas** (Line Chart). Es el ícono que muestra una línea ascendente sobre un área.

5. Aparecerá un contenedor de visualización vacío en el lienzo. Arrástralo para que ocupe aproximadamente el **70% superior** del lienzo (dejando espacio abajo para la tabla que crearás en el Paso 5).

6. Con el gráfico seleccionado, configura los campos en el panel de Visualizaciones:

   | Campo del gráfico | Columna a asignar | Cómo asignar |
   |---|---|---|
   | **Eje X** (Axis) | `Periodo` | Arrastra desde el panel Datos |
   | **Eje Y** (Values) | `Valor` | Arrastra desde el panel Datos |
   | **Leyenda** (Legend) | `Tipo` | Arrastra desde el panel Datos |

7. El gráfico debería mostrar automáticamente dos líneas: una para los registros con `Tipo = "Real"` y otra para `Tipo = "Pronóstico"`.

**Parte C: Personalizar colores de las series**

8. Con el gráfico seleccionado, haz clic en el **ícono de pincel** (Formato) en el panel de Visualizaciones.

9. Expande la sección **"Líneas"** o **"Colores de datos"** (el nombre puede variar según la versión de Power BI).

10. Localiza la serie **"Real"** y cambia su color a **azul** (código hexadecimal recomendado: `#0070C0` o el azul estándar de la paleta).

11. Localiza la serie **"Pronóstico"** y cambia su color a **naranja** (código hexadecimal recomendado: `#FF7A00` o el naranja estándar de la paleta).

12. Expande la sección **"Título"** del panel de formato y escribe como título del gráfico:
    ```
    Demanda Real vs. Pronóstico — SKU [nombre de tu SKU]
    ```

13. Activa la **leyenda** si no está visible: en el panel de formato, expande **"Leyenda"** y activa el toggle a **"Activado"**. Configura la posición como **"Inferior central"** o **"Superior derecha"**.

**Parte D: Agregar línea de referencia con el Analytics Pane**

14. Con el gráfico de líneas aún seleccionado, haz clic en el **ícono de lupa con gráfico** (Analytics Pane / Panel de análisis) en el panel de Visualizaciones. Es el tercer ícono de la fila superior del panel.

15. En el Panel de análisis, busca la opción **"Línea constante"** (Constant Line) y haz clic en **"+ Agregar"**.

16. Configura la línea de referencia:
    - **Valor:** Ingresa el **promedio de la demanda histórica real** (puedes calcularlo rápidamente en Excel: promedio de los 18 valores reales).
    - **Color:** Gris claro (`#AAAAAA`).
    - **Estilo:** Línea punteada.
    - **Etiqueta de datos:** Activar y escribir `Promedio histórico`.

17. Haz clic fuera del gráfico para deseleccionarlo y observa el resultado visual.

#### Resultado esperado

Un gráfico de líneas que muestra:
- Línea **azul** con los 18 puntos de demanda real (Meses 1–18).
- Línea **naranja** con los 6 puntos de pronóstico (Meses 19–24).
- Una línea de referencia gris punteada en el valor del promedio histórico.
- Título descriptivo y leyenda visible con etiquetas "Real" y "Pronóstico".

#### Verificación

✅ El gráfico muestra exactamente dos líneas con colores diferenciados.
✅ La leyenda identifica correctamente "Real" y "Pronóstico".
✅ La línea azul tiene 18 puntos y la línea naranja tiene 6 puntos.
✅ La línea de referencia del promedio histórico es visible y está etiquetada.
✅ El eje X muestra los períodos en orden cronológico (Mes 1 → Mes 24).

---

### Paso 5 — Crear la tabla comparativa con formato condicional

**Objetivo:** Construir una visualización tipo tabla en Power BI que muestre período a período el valor real, el valor pronosticado y la diferencia absoluta, con formato condicional que resalte en rojo las diferencias superiores al 15%.

#### Instrucciones

**Parte A: Preparar las medidas necesarias**

Antes de construir la tabla, necesitas crear una medida calculada para la diferencia porcentual. En Power BI, esto se hace con DAX.

1. En el panel **Datos** (lado derecho), haz clic derecho sobre la tabla `Comparativo_Demanda_Pronostico` y selecciona **"Nueva medida"**.

2. En la barra de fórmulas que aparece en la parte superior, escribe la siguiente medida DAX:

   ```dax
   Diferencia_Abs =
   VAR ValorReal =
       CALCULATE(
           SUM(Comparativo_Demanda_Pronostico[Valor]),
           Comparativo_Demanda_Pronostico[Tipo] = "Real"
       )
   VAR ValorPronostico =
       CALCULATE(
           SUM(Comparativo_Demanda_Pronostico[Valor]),
           Comparativo_Demanda_Pronostico[Tipo] = "Pronóstico"
       )
   RETURN
       ABS(ValorReal - ValorPronostico)
   ```

3. Presiona **Enter** o haz clic en el ícono de confirmación (✓). La medida `Diferencia_Abs` aparecerá en el panel Datos.

4. Crea una segunda medida para la diferencia porcentual:

   ```dax
   Diferencia_Pct =
   VAR ValorReal =
       CALCULATE(
           SUM(Comparativo_Demanda_Pronostico[Valor]),
           Comparativo_Demanda_Pronostico[Tipo] = "Real"
       )
   VAR ValorPronostico =
       CALCULATE(
           SUM(Comparativo_Demanda_Pronostico[Valor]),
           Comparativo_Demanda_Pronostico[Tipo] = "Pronóstico"
       )
   RETURN
       DIVIDE(ABS(ValorReal - ValorPronostico), ValorReal, 0)
   ```

   > **Nota:** Esta medida devuelve un valor decimal (por ejemplo, 0.18 para 18%). El formato condicional se configurará con el umbral 0.15 (equivalente al 15%).

**Parte B: Insertar la visualización de tabla**

5. En el lienzo de la página `Comparativa Real vs Pronóstico`, haz clic en un área vacía debajo del gráfico de líneas.

6. En el panel **Visualizaciones**, haz clic en el ícono de **Tabla** (Table). Aparece como una cuadrícula de filas y columnas.

7. Ajusta el tamaño del contenedor de la tabla para que ocupe el **30% inferior** del lienzo.

8. Con la tabla seleccionada, configura las columnas en el panel de Visualizaciones (sección **"Columnas"**):

   Arrastra los siguientes campos desde el panel Datos hacia la sección Columnas:
   - `Periodo` (de la tabla `Comparativo_Demanda_Pronostico`)
   - `Valor` (de la tabla) — esto mostrará los valores; más adelante lo ajustarás
   - `Tipo` (de la tabla)
   - `Diferencia_Abs` (medida creada)
   - `Diferencia_Pct` (medida creada)

   > **Nota:** La tabla mostrará inicialmente filas duplicadas por tipo. Esto es normal dado el formato largo. En un entorno de práctica de 12 minutos, esta visualización sirve para identificar visualmente los períodos de mayor brecha. Para una vista más limpia, se puede pivotar en Power Query (opcional, fuera del alcance de esta práctica).

9. Formatea la columna `Diferencia_Pct` como porcentaje:
   - Selecciona la medida `Diferencia_Pct` en el panel Datos.
   - En la cinta superior (pestaña *Herramientas de medición*), cambia el formato a **Porcentaje** con 1 decimal.

**Parte C: Aplicar formato condicional**

10. Con la tabla seleccionada en el lienzo, en el panel de Visualizaciones haz clic en el **ícono de pincel** (Formato).

11. Expande la sección **"Formato condicional"** (Conditional formatting) o busca la opción dentro de la configuración de columnas específicas.

12. Localiza la columna `Diferencia_Pct` en la lista de columnas de la tabla y activa el **formato de color de fondo** (Background color):
    - Haz clic en el botón **"fx"** o **"Formato condicional"** junto a la columna.
    - Se abre el cuadro de diálogo de formato condicional.

13. Configura la regla de formato condicional:
    - **Basado en:** `Valor del campo`
    - **Campo:** `Diferencia_Pct`
    - **Tipo de formato:** Reglas
    - Haz clic en **"+ Nueva regla"** y configura:
      - Si el valor es **mayor que** `0.15` → Color de fondo: **Rojo** (`#FF0000`) con texto **Blanco** (`#FFFFFF`)
    - Haz clic en **Aceptar**.

14. Agrega un título a la tabla:
    - En el panel de formato, expande **"Título"** y escribe: `Comparativo Período a Período — Diferencias de Pronóstico`

#### Resultado esperado

Una tabla en Power BI que muestra los registros del dataset con la columna `Diferencia_Pct` resaltada en **rojo** para todos los períodos donde la diferencia entre real y pronóstico supera el **15%**, facilitando la identificación visual inmediata de los períodos de mayor error de pronóstico.

#### Verificación

✅ Las medidas `Diferencia_Abs` y `Diferencia_Pct` aparecen en el panel Datos sin errores.
✅ La tabla muestra la columna `Diferencia_Pct` en formato porcentual.
✅ Al menos una celda de `Diferencia_Pct` aparece con fondo rojo (si el modelo de pronóstico tiene errores > 15% en algún período).
✅ El título de la tabla es visible y descriptivo.

---

### Paso 6 — Agregar cuadro de texto con observaciones logísticas

**Objetivo:** Documentar dentro del lienzo de Power BI al menos dos observaciones sobre los períodos de mayor brecha entre demanda real y pronóstico, conectando el análisis visual con decisiones operativas de reabastecimiento.

#### Instrucciones

1. En la página `Comparativa Real vs Pronóstico`, ve a la cinta superior y haz clic en **Insertar → Cuadro de texto** (Text Box).

2. Dibuja el cuadro de texto en un espacio libre del lienzo (puede ser a la derecha del gráfico o debajo de la tabla, según el espacio disponible).

3. En el cuadro de texto, escribe tus observaciones siguiendo esta estructura:

   ```
   ANÁLISIS DE BRECHAS — [SKU analizado] — [Fecha de análisis]

   Observación 1: [Período con mayor sobreestimación]
   En el Mes [X], el pronóstico fue de [valor] unidades pero la demanda real
   fue de [valor] unidades (diferencia: [%]%). Esto representa una
   SOBREESTIMACIÓN que en términos logísticos implica:
   → Riesgo de exceso de inventario en [número] unidades
   → Capital inmovilizado estimado en bodega
   → Posible necesidad de ajustar orden de compra a la baja

   Observación 2: [Período con mayor subestimación]
   En el Mes [X], el pronóstico fue de [valor] unidades pero la demanda real
   fue de [valor] unidades (diferencia: [%]%). Esto representa una
   SUBESTIMACIÓN que en términos logísticos implica:
   → Riesgo de quiebre de stock de [número] unidades
   → Potencial pérdida de ventas o penalización de nivel de servicio
   → Necesidad de activar reabastecimiento de emergencia o buffer adicional

   Observación 3 (opcional): Patrón general
   El modelo [nombre del modelo] tiende a [sobreestimar/subestimar] en los
   meses de [descripción del patrón, ej. alta demanda estacional]. Se
   recomienda revisar el factor de ajuste estacional para los períodos
   proyectados Mes 19–24.
   ```

   > **Importante:** Reemplaza todos los valores entre corchetes `[...]` con los datos reales de tu análisis. Las observaciones deben estar basadas en lo que ves en el gráfico de líneas y la tabla comparativa.

4. Aplica formato básico al cuadro de texto:
   - Título principal en **negrita**, tamaño 11pt.
   - Observaciones en tamaño 9pt.
   - Usa colores coherentes con el gráfico: menciona la línea **azul** para real y **naranja** para pronóstico.

5. Haz clic fuera del cuadro de texto para finalizar la edición.

#### Resultado esperado

Un cuadro de texto visible en el lienzo de Power BI con al menos dos observaciones estructuradas que conectan la brecha visual del gráfico con consecuencias operativas concretas en inventario y reabastecimiento.

#### Verificación

✅ El cuadro de texto contiene al menos 2 observaciones diferenciadas.
✅ Cada observación identifica un período específico (número de mes).
✅ Cada observación menciona si es sobreestimación o subestimación.
✅ Cada observación incluye al menos una consecuencia logística operativa.
✅ Los valores numéricos en el texto coinciden con los datos del gráfico.

---

### Paso 7 — Guardar el archivo .pbix final

**Objetivo:** Guardar el trabajo completado como archivo `.pbix` con nombre correcto para uso en prácticas posteriores.

#### Instrucciones

1. En Power BI Desktop, ve a **Archivo → Guardar como**.
2. Navega hasta `C:\LabLogistica\Practica24\`.
3. Escribe el nombre del archivo: `Practica24_Visual_Comparativa`
4. Haz clic en **Guardar**.
5. Verifica que el archivo `Practica24_Visual_Comparativa.pbix` aparece en la carpeta `C:\LabLogistica\Practica24\`.
6. Confirma que el archivo `.pbix` tiene un tamaño mayor a 0 KB (normalmente entre 500 KB y 2 MB para datasets de este tamaño).

#### Resultado esperado

El archivo `Practica24_Visual_Comparativa.pbix` guardado correctamente en la carpeta de trabajo, listo para ser utilizado como insumo en prácticas posteriores del curso.

#### Verificación

✅ El archivo `.pbix` existe en `C:\LabLogistica\Practica24\`.
✅ Al abrir el archivo, la página `Comparativa Real vs Pronóstico` muestra el gráfico de líneas, la tabla comparativa y el cuadro de texto con observaciones.
✅ El archivo también conserva las páginas previas del `.pbix` de la Práctica 22.

---

## Validación y Pruebas del Laboratorio

Una vez completados todos los pasos, realiza las siguientes verificaciones de integridad:

### Lista de verificación final

| # | Elemento a verificar | Criterio de éxito |
|---|---|---|
| 1 | Dataset en formato largo | 24 filas exactas en `Dataset_Comparativo_Largo.xlsx` con columnas `Periodo`, `Tipo`, `Valor` |
| 2 | Tabla importada en Power BI | `Comparativo_Demanda_Pronostico` visible en panel Datos sin errores |
| 3 | Página nueva en .pbix | Página `Comparativa Real vs Pronóstico` existe y está activa |
| 4 | Gráfico de líneas | Dos líneas visibles: azul (Real, 18 puntos) y naranja (Pronóstico, 6 puntos) |
| 5 | Colores diferenciados | Línea Real = azul (`#0070C0`), línea Pronóstico = naranja (`#FF7A00`) |
| 6 | Línea de referencia | Línea punteada gris en el valor del promedio histórico con etiqueta |
| 7 | Tabla comparativa | Tabla con columnas `Periodo`, `Tipo`, `Valor`, `Diferencia_Abs`, `Diferencia_Pct` |
| 8 | Formato condicional | Celdas con `Diferencia_Pct > 15%` resaltadas en rojo |
| 9 | Cuadro de texto | Mínimo 2 observaciones con período específico e impacto logístico |
| 10 | Archivo guardado | `Practica24_Visual_Comparativa.pbix` en `C:\LabLogistica\Practica24\` |

### Pregunta de reflexión logística

> **¿Cómo impacta este análisis en una decisión operativa de logística?**
>
> Los períodos donde el pronóstico **sobreestima** la demanda real son señales de alerta para revisar los niveles de stock de seguridad hacia abajo, evitando costos de almacenamiento innecesarios. Los períodos donde el pronóstico **subestima** la demanda real son señales para aumentar el stock de seguridad o reducir el tiempo de ciclo de reabastecimiento, protegiendo el nivel de servicio. La visualización comparativa permite al planificador logístico tomar estas decisiones de forma informada y justificada con datos, en lugar de basarse en criterios subjetivos.

---

## Solución de Problemas

### Problema 1: El gráfico de líneas muestra una sola línea o no diferencia las series

**Síntoma:** Al construir el gráfico de líneas con los campos `Periodo` (eje X), `Valor` (eje Y) y `Tipo` (leyenda), el gráfico muestra solo una línea o ambas series aparecen superpuestas con el mismo color.

**Causa probable:** La columna `Tipo` no fue asignada al campo **"Leyenda"** del gráfico, sino que fue colocada en otro campo (como "Información sobre herramientas" o "Eje secundario"). Otra causa frecuente es que los valores en la columna `Tipo` no son exactamente `"Real"` y `"Pronóstico"` (por ejemplo, tienen espacios adicionales, mayúsculas inconsistentes o caracteres especiales como tildes que difieren entre filas).

**Solución:**
1. Selecciona el gráfico de líneas en el lienzo.
2. En el panel de Visualizaciones, verifica que la columna `Tipo` esté ubicada específicamente en el campo **"Leyenda"** (Legend), no en otro campo.
3. Si el campo está correcto pero las series no se diferencian, regresa a Excel y verifica que los valores en la columna `Tipo` sean exactamente `Real` y `Pronóstico` (sin espacios al inicio o al final).
4. En Power Query, puedes limpiar los valores de texto con: *Transformar → Formato → Recortar* sobre la columna `Tipo`, luego *Cerrar y aplicar*.
5. Si el problema persiste, en Power Query verifica el paso *Tipo cambiado* y asegúrate de que la columna `Tipo` tiene tipo de dato **Texto**.

---

### Problema 2: El formato condicional rojo no aparece en la tabla aunque hay diferencias mayores al 15%

**Síntoma:** La tabla comparativa muestra valores en la columna `Diferencia_Pct`, algunos claramente superiores a 0.15 (o 15%), pero ninguna celda aparece con fondo rojo.

**Causa probable:** La regla de formato condicional fue configurada con el umbral en `15` en lugar de `0.15`. Como la medida `Diferencia_Pct` devuelve un valor decimal (por ejemplo, 0.18 para 18%), la comparación debe hacerse contra `0.15`, no contra `15`. Otra causa es que el formato condicional se aplicó a la columna incorrecta (por ejemplo, a `Diferencia_Abs` en lugar de `Diferencia_Pct`).

**Solución:**
1. Selecciona la tabla en el lienzo de Power BI.
2. En el panel de Visualizaciones, haz clic en el ícono de **Formato** (pincel).
3. Localiza la configuración de formato condicional para la columna `Diferencia_Pct`.
4. Abre la regla existente y verifica el valor umbral. Si dice `15`, cámbialo a `0.15`.
5. Confirma que el campo seleccionado en la regla es `Diferencia_Pct` y no otra medida.
6. Haz clic en **Aceptar** y verifica que las celdas correspondientes ahora aparecen en rojo.
7. Si la medida `Diferencia_Pct` está formateada como porcentaje en Power BI (mostrando "18%" en lugar de "0.18"), entonces el umbral en la regla de formato condicional debe ser `15` (no `0.15`), ya que Power BI aplica la regla sobre el valor mostrado. Ajusta según corresponda.

---

## Limpieza del Entorno

Al finalizar la práctica, realiza las siguientes acciones para mantener el entorno de trabajo organizado:

1. **Guardar el archivo .pbix** una última vez (Ctrl+S) para asegurar que todos los cambios están persistidos.

2. **Cerrar Power BI Desktop** normalmente (no forzar cierre).

3. **Verificar la carpeta de salida** — debe contener los siguientes archivos:

   ```
   C:\LabLogistica\Practica24\
   ├── Practica21_Pronostico_Seleccionado.xlsx   ← archivo de entrada (no modificar)
   ├── Practica22_Dashboard_Base.pbix            ← archivo de entrada (no modificar)
   ├── Dataset_Comparativo_Largo.xlsx            ← NUEVO: dataset en formato largo
   └── Practica24_Visual_Comparativa.pbix        ← NUEVO: archivo de salida principal
   ```

4. **Compartir el archivo de salida** con el instructor si así se indica: el archivo `Practica24_Visual_Comparativa.pbix` es el entregable de esta práctica.

5. **No eliminar** el archivo `Dataset_Comparativo_Largo.xlsx`, ya que la conexión del archivo `.pbix` depende de que este archivo permanezca en la misma ruta. Si mueves el `.xlsx`, la conexión en Power BI se romperá y deberás restablecerla desde *Inicio → Transformar datos → Configuración de origen de datos*.

---

## Resumen

En esta práctica aplicaste el flujo completo de importación de datos desde Excel a Power BI Desktop para construir una visualización comparativa entre demanda histórica real y pronóstico seleccionado. Los pasos clave fueron:

| Paso | Acción | Herramienta |
|---|---|---|
| 1 | Identificar datos de origen (18 reales + 6 pronóstico) | Excel |
| 2 | Estructurar dataset en formato largo con columnas `Periodo`, `Tipo`, `Valor` | Excel |
| 3 | Importar con Power Query, verificar tipos y cargar al modelo | Power BI / Power Query |
| 4 | Construir gráfico de líneas dual con colores diferenciados y línea de referencia | Power BI |
| 5 | Crear tabla comparativa con formato condicional (rojo > 15%) | Power BI |
| 6 | Documentar observaciones logísticas sobre brechas en cuadro de texto | Power BI |
| 7 | Guardar archivo `.pbix` final | Power BI |

### Conceptos clave consolidados

- **Formato largo (*long format*):** Estructura de datos donde cada observación ocupa una fila, con una columna de tipo/categoría. Es el formato requerido por Power BI para crear gráficos de múltiples series usando la leyenda.
- **Sobreestimación vs. subestimación:** La sobreestimación (pronóstico > real) genera riesgo de exceso de inventario; la subestimación (pronóstico < real) genera riesgo de quiebre de stock. Ambos tienen costos logísticos directos.
- **Formato condicional en Power BI:** Permite identificar visualmente de forma inmediata los períodos problemáticos sin necesidad de revisar manualmente cada valor.
- **Analytics Pane:** Herramienta de Power BI que agrega líneas de referencia estadísticas (promedio, percentil, constante) sobre gráficos existentes para enriquecer el análisis visual.

### Conexión con prácticas siguientes

El archivo `Practica24_Visual_Comparativa.pbix` generado en esta práctica es el insumo base para las prácticas 25 y siguientes, donde se agregarán visualizaciones adicionales de indicadores de nivel de servicio y recomendaciones de reabastecimiento. Mantén el archivo en la carpeta `C:\LabLogistica\Practica24\` y no cambies su nombre ni su ubicación.

### Recursos de referencia

- [Documentación oficial: Conectar a un libro de Excel en Power BI Desktop](https://learn.microsoft.com/es-es/power-bi/connect-data/desktop-connect-excel)
- [Documentación oficial: Gráficos de líneas en Power BI](https://learn.microsoft.com/es-es/power-bi/visuals/power-bi-line-chart)
- [Documentación oficial: Formato condicional en tablas de Power BI](https://learn.microsoft.com/es-es/power-bi/create-reports/desktop-conditional-table-formatting)
- [Documentación oficial: Panel de análisis en Power BI Desktop](https://learn.microsoft.com/es-es/power-bi/transform-model/desktop-analytics-pane)
- [Documentación oficial: Introducción a DAX en Power BI](https://learn.microsoft.com/es-es/dax/dax-overview)

---

# Análisis de comportamiento por categoría de producto

## 1. Metadatos

| Campo            | Detalle                                      |
|------------------|----------------------------------------------|
| **Duración**     | 12 minutos                                   |
| **Complejidad**  | Alta                                         |
| **Nivel Bloom**  | Aplicar (Apply)                              |
| **Módulo**       | Módulo 5 — Visualización y análisis en Power BI Desktop |
| **Práctica**     | 25 de 25 (práctica final del Módulo 5)       |

---

## 2. Descripción General

En esta práctica final del Módulo 5 el participante construirá una tercera página de análisis en el archivo `.pbix` existente, dedicada exclusivamente a comparar el comportamiento de demanda entre categorías de producto (perecederos vs. no perecederos). Se crearán tres visualizaciones complementarias —gráfico de líneas con tendencias individuales, gráfico de área apilada y matriz con heat map— y se cerrará el módulo configurando la sincronización de segmentadores entre las tres páginas del reporte, consolidando así un dashboard integrado y navegable como producto final del Módulo 5.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica, el participante será capaz de:

- [ ] Construir un gráfico de líneas con múltiples series por categoría y aplicar líneas de tendencia independientes mediante el Analytics Pane de Power BI Desktop.
- [ ] Crear un gráfico de área apilada para analizar la composición del mix de productos a lo largo del tiempo e identificar cambios en períodos de pico.
- [ ] Configurar una visualización Matrix con formato condicional de escala de colores (heat map) para detectar meses de alta y baja demanda por categoría.
- [ ] Sincronizar segmentadores entre las tres páginas del reporte usando la función Sync Slicers, garantizando coherencia de filtros en todo el dashboard.
- [ ] Integrar y revisar el reporte completo del Módulo 5 como un producto analítico coherente con navegación entre páginas.

---

## 4. Prerrequisitos

### Conocimiento previo
- Haber completado las **Prácticas 22, 23 y 24** del Módulo 5 y contar con el archivo `.pbix` con las dos páginas previas ya construidas y guardadas.
- Comprensión básica de la diferencia entre categorías de producto (perecederos / no perecederos) y su impacto diferenciado en la cadena logística (caducidad, rotación, planificación de inventario).
- Familiaridad con el concepto de mix de productos y su relevancia en la planificación de operaciones de picking y abastecimiento.
- Conocimiento del proceso de importación de datos desde Excel a Power BI Desktop (Lección 5.1) y de la estructura del dataset de práctica con 24 meses de historial para 8 SKUs.

### Acceso y archivos requeridos
- Archivo `.pbix` generado en las Prácticas 22–24 con las páginas **"Tendencias de Demanda"** y **"Análisis de Variabilidad"** ya construidas.
- Dataset de práctica `Demanda_Historica_8SKUs.xlsx` (proporcionado por el instructor) con las columnas: `Fecha`, `SKU`, `Descripción`, `Categoría`, `Unidades_Demandadas`, `Canal`.
- La columna `Categoría` debe contener exactamente dos valores: `Perecedero` y `No Perecedero`.
- Power BI Desktop instalado y actualizado (versión 2024 recomendada).

---

## 5. Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente         | Mínimo                          | Recomendado                     |
|--------------------|---------------------------------|---------------------------------|
| Procesador         | Intel Core i5 / AMD Ryzen 5     | Intel Core i7 / AMD Ryzen 7     |
| RAM                | 4 GB                            | 8 GB                            |
| Almacenamiento     | 2 GB libres en disco            | 5 GB libres en disco            |
| Resolución         | 1280 × 768                      | 1920 × 1080                     |
| Conexión a internet| 10 Mbps                         | 20 Mbps o superior              |

### Software requerido

| Software              | Versión mínima          | Notas                                              |
|-----------------------|-------------------------|----------------------------------------------------|
| Power BI Desktop      | Versión 2024 (gratuita) | Descarga: microsoft.com/es-es/download (gratuito)  |
| Microsoft Excel       | 2016 o superior         | Para verificar el dataset fuente si es necesario   |
| Microsoft Edge / Chrome | Versión actualizada   | Para acceso a recursos de referencia               |

### Verificación del entorno antes de comenzar

Antes de iniciar los pasos de la práctica, confirmar lo siguiente:

1. Abrir Power BI Desktop.
2. Abrir el archivo `.pbix` de las prácticas anteriores.
3. Verificar que existen exactamente **dos páginas** en la barra inferior: `Tendencias de Demanda` y `Análisis de Variabilidad`.
4. Hacer clic en cualquier visual de la página 1 y confirmar que los datos se cargan sin errores (sin triángulo amarillo de advertencia en las visualizaciones).
5. En el panel **Datos** (derecha), expandir la tabla importada y confirmar que la columna `Categoría` existe y está visible.

> ⚠️ **Si el archivo `.pbix` no está disponible:** Solicitar al instructor el archivo de checkpoint `Lab_Modulo5_P24_checkpoint.pbix` antes de continuar.

---

## 6. Procedimiento Paso a Paso

---

### Paso 1: Crear la tercera página del reporte

**Objetivo:** Agregar una nueva página dedicada al análisis por categoría dentro del archivo `.pbix` existente.

**Instrucciones:**

1. Con el archivo `.pbix` abierto en Power BI Desktop, localizar la barra de pestañas de páginas en la parte inferior de la pantalla.
2. Hacer clic en el ícono **"+"** (signo más) ubicado a la derecha de la última pestaña de página para agregar una nueva página.
3. Hacer doble clic sobre la nueva pestaña (que aparece como `Página 3` o similar) para renombrarla.
4. Escribir el nombre: `Análisis por Categoría` y presionar **Enter**.
5. Confirmar que la nueva página está en blanco y activa.

**Resultado esperado:** La barra inferior muestra tres pestañas: `Tendencias de Demanda`, `Análisis de Variabilidad` y `Análisis por Categoría`. La tercera pestaña está seleccionada y el lienzo está vacío.

**Verificación:** Las tres pestañas son visibles simultáneamente en la barra inferior sin necesidad de desplazarse horizontalmente.

---

### Paso 2: Construir el gráfico de líneas con múltiples series por categoría

**Objetivo:** Crear un gráfico de líneas donde cada serie representa una categoría de producto, con Fecha en el eje X y suma de Unidades_Demandadas en el eje Y, usando la columna Categoría como leyenda para comparar visualmente el comportamiento de demanda entre perecederos y no perecederos.

**Instrucciones:**

1. En el panel **Visualizaciones** (derecha), hacer clic en el ícono de **Gráfico de líneas** (Line chart). Aparecerá un visual vacío en el lienzo.
2. Arrastrar el visual hacia la mitad superior izquierda del lienzo. Redimensionarlo para que ocupe aproximadamente el **50% del ancho** y el **45% del alto** del lienzo.
3. Con el visual seleccionado, ir al panel **Datos** (extremo derecho) y configurar los campos de la siguiente manera:
   - **Eje X:** Arrastrar la columna `Fecha` al campo **"Eje X"** (o "Axis"). Power BI puede agrupar automáticamente por jerarquía de fecha; si esto ocurre, hacer clic en la flecha desplegable del campo y seleccionar **"Fecha"** (nivel de mes o el nivel disponible más granular).
   - **Eje Y:** Arrastrar la columna `Unidades_Demandadas` al campo **"Eje Y"** (o "Values"). Verificar que la agregación sea **Suma** (si aparece "Recuento", hacer clic en la flecha del campo → seleccionar "Suma").
   - **Leyenda:** Arrastrar la columna `Categoría` al campo **"Leyenda"** (o "Legend").
4. Verificar que el gráfico muestra **dos líneas de colores distintos**: una para `Perecedero` y otra para `No Perecedero`, con el tiempo en el eje horizontal.
5. Hacer clic en el ícono de **pincel de formato** (Format visual) en el panel Visualizaciones para abrir las opciones de formato.
6. Expandir la sección **"Título"** y escribir: `Demanda por Categoría — Serie Temporal (24 meses)`.
7. Expandir la sección **"Leyenda"** y asegurarse de que esté **activada** (toggle en On) y posicionada en la parte superior del gráfico.

**Resultado esperado:** Un gráfico de líneas con dos series claramente diferenciadas por color, con etiqueta de título visible y leyenda que identifica cada categoría. Las líneas deben mostrar la evolución mensual de la demanda durante los 24 meses del dataset.

**Verificación:** Al pasar el cursor sobre cualquier punto de las líneas, el tooltip muestra la `Fecha`, la `Categoría` y el valor de `Suma de Unidades_Demandadas` correspondiente.

---

### Paso 3: Agregar líneas de tendencia individuales por categoría mediante el Analytics Pane

**Objetivo:** Activar el panel de análisis (Analytics Pane) para agregar líneas de tendencia independientes a cada serie del gráfico de líneas, permitiendo identificar visualmente si las categorías presentan comportamientos divergentes (creciente, decreciente o estable).

**Instrucciones:**

1. Asegurarse de que el gráfico de líneas del Paso 2 esté seleccionado (borde azul visible alrededor del visual).
2. En el panel **Visualizaciones**, localizar los tres íconos en la parte superior del panel:
   - Ícono de **campos** (tabla con columnas)
   - Ícono de **pincel de formato** (pincel)
   - Ícono de **Analytics** (lupa con gráfico o símbolo de análisis `∿`)
3. Hacer clic en el ícono de **Analytics** (tercer ícono). Se abrirá el panel de Analytics con opciones de líneas de referencia.
4. Localizar la opción **"Línea de tendencia"** (Trend line) dentro del panel de Analytics.
5. Hacer clic en el botón **"Agregar"** o activar el toggle junto a "Línea de tendencia".

   > **Nota importante:** En versiones recientes de Power BI Desktop, el Analytics Pane agrega automáticamente una línea de tendencia **por cada serie** cuando la leyenda está configurada. Si solo aparece una línea, verificar que la columna `Categoría` esté correctamente asignada al campo Leyenda (Paso 2).

6. Expandir la configuración de la línea de tendencia recién agregada:
   - **Color:** Mantener el color automático por serie (Power BI asigna el mismo color de cada línea de datos a su respectiva tendencia, con trazo discontinuo).
   - **Estilo de línea:** Seleccionar **"Discontinuo"** (Dashed) para diferenciarla visualmente de la línea de datos.
   - **Transparencia:** Establecer en **20%** para mayor visibilidad.
7. Verificar en el gráfico que ahora aparecen **dos líneas de tendencia adicionales** (una por categoría), representadas con trazo discontinuo, superpuestas sobre las líneas de datos originales.
8. Observar la dirección de cada línea de tendencia:
   - ¿La tendencia de `Perecedero` es creciente, decreciente o estable?
   - ¿La tendencia de `No Perecedero` sigue la misma dirección o diverge?
   - Anotar mentalmente esta observación para la discusión final de la práctica.

**Resultado esperado:** El gráfico de líneas muestra cuatro líneas en total: dos líneas sólidas de datos (una por categoría) y dos líneas de tendencia discontinuas superpuestas. Las pendientes de las líneas de tendencia permiten comparar visualmente la dirección de crecimiento de cada categoría.

**Verificación:** Al hacer clic en el Analytics Pane con el visual seleccionado, la sección "Línea de tendencia" muestra un contador indicando **"2 líneas"** o similar, confirmando que hay una tendencia por serie.

---

### Paso 4: Crear el gráfico de área apilada para análisis de composición del mix

**Objetivo:** Construir un gráfico de área apilada (Stacked Area Chart) que muestre la proporción de cada categoría sobre la demanda total a lo largo del tiempo, permitiendo identificar si la composición del mix de productos cambia durante los períodos de pico de demanda.

**Instrucciones:**

1. Hacer clic en un área vacía del lienzo para deseleccionar el visual anterior.
2. En el panel **Visualizaciones**, hacer clic en el ícono de **Gráfico de área apilada** (Stacked Area Chart). Es el ícono que muestra áreas de colores apiladas con línea en la parte superior.

   > **Tip de identificación:** Si no lo encuentras directamente, hacer clic en los tres puntos `...` al final del panel de visualizaciones para expandir todas las opciones y buscarlo. También puedes pasar el cursor sobre los íconos para ver el tooltip con el nombre.

3. Arrastrar el nuevo visual hacia la mitad superior derecha del lienzo. Redimensionarlo para que ocupe aproximadamente el **50% del ancho** y el **45% del alto** del lienzo (simétrico al gráfico de líneas).
4. Configurar los campos del gráfico de área apilada:
   - **Eje X:** Arrastrar la columna `Fecha` al campo **"Eje X"**.
   - **Eje Y:** Arrastrar la columna `Unidades_Demandadas` al campo **"Eje Y"**. Verificar que la agregación sea **Suma**.
   - **Leyenda:** Arrastrar la columna `Categoría` al campo **"Leyenda"**.
5. Observar que el gráfico muestra áreas de colores apiladas: la suma de ambas categorías forma la demanda total en cada período, y cada color representa la contribución de una categoría.
6. Hacer clic en el ícono de **formato** (pincel) y configurar:
   - **Título:** `Composición del Mix de Demanda por Categoría`
   - **Leyenda:** Activada, posición superior.
   - **Etiquetas de datos:** Desactivar (para mantener el visual limpio con 24 meses de datos).
7. Analizar visualmente el gráfico: ¿La proporción de perecederos aumenta o disminuye en los meses de mayor demanda total? ¿La composición del mix es estable o variable a lo largo del tiempo?

**Resultado esperado:** Un gráfico de área apilada donde la altura total de las áreas representa la demanda total mensual, y los dos colores apilados representan la contribución de `Perecedero` y `No Perecedero`. Los cambios en la proporción relativa de los colores indican variaciones en el mix de productos.

**Verificación:** La suma de las dos áreas en cualquier mes debe coincidir aproximadamente con el total de demanda visible en el gráfico de líneas del Paso 2 para ese mismo mes. Al pasar el cursor sobre el gráfico, el tooltip muestra los valores individuales por categoría y el total.

---

### Paso 5: Construir la matriz con formato condicional heat map

**Objetivo:** Crear una visualización Matrix (tabla dinámica visual) con Categoría en filas y Mes en columnas, aplicando formato condicional de escala de colores para identificar visualmente los meses de mayor y menor demanda por categoría a modo de mapa de calor (heat map).

**Instrucciones:**

1. Hacer clic en un área vacía del lienzo (preferiblemente en la mitad inferior).
2. En el panel **Visualizaciones**, hacer clic en el ícono de **Matriz** (Matrix). Es el ícono que muestra una tabla con filas y columnas anidadas.
3. Arrastrar y redimensionar el visual para que ocupe el **100% del ancho** y el **40% del alto** inferior del lienzo, de modo que quede debajo de los dos gráficos superiores.
4. Configurar los campos de la Matriz:
   - **Filas (Rows):** Arrastrar la columna `Categoría`.
   - **Columnas (Columns):** Para mostrar los meses, necesitamos una columna de mes. Arrastrar la columna `Fecha` al campo **"Columnas"**. Power BI puede mostrar una jerarquía; hacer clic en la flecha del campo y seleccionar el nivel **"Mes"** o **"Nombre del mes"** según esté disponible.

   > **Si no existe una columna de Mes separada:** En el panel Datos, expandir la tabla y buscar si existe una columna `Mes` o `Nombre_Mes` creada en prácticas anteriores. Si no existe, usar `Fecha` y Power BI agrupará automáticamente por mes al colocarla en columnas.

   - **Valores (Values):** Arrastrar la columna `Unidades_Demandadas`. Verificar que la agregación sea **Suma**.
5. La matriz debe mostrar una cuadrícula con `Perecedero` y `No Perecedero` en las filas, los 24 meses en las columnas, y los valores de demanda en las celdas.
6. **Aplicar formato condicional (Heat Map):**
   a. Con la Matriz seleccionada, hacer clic en el ícono de **formato** (pincel) en el panel Visualizaciones.
   b. Expandir la sección **"Valores de celda"** (Cell elements) o **"Formato condicional"** según la versión de Power BI.
   c. Localizar la opción **"Color de fondo"** (Background color) y hacer clic en el botón **"fx"** o en el toggle para activarla.
   d. Se abrirá el diálogo de formato condicional. Configurar:
      - **Estilo de formato:** Seleccionar **"Escala de colores"** (Color scale / Gradient).
      - **Basado en campo:** `Suma de Unidades_Demandadas`.
      - **Mínimo:** Color **blanco** o **azul claro** (valores bajos de demanda).
      - **Máximo:** Color **azul oscuro** o **naranja intenso** (valores altos de demanda). Usar una escala que sea intuitiva: colores fríos para demanda baja, colores cálidos o intensos para demanda alta.
      - Hacer clic en **"Aceptar"** para aplicar.
7. Verificar que las celdas de la matriz ahora muestran un gradiente de colores: los meses con mayor demanda aparecen con el color más intenso, y los de menor demanda con el color más claro.
8. Formatear el título de la Matriz: hacer clic en **Título** en el panel de formato y escribir: `Heat Map de Demanda por Categoría y Mes`.
9. Ajustar el ancho de columnas si es necesario: hacer clic derecho sobre cualquier encabezado de columna en la Matriz → **"Ajustar automáticamente"** (Auto-fit).

**Resultado esperado:** Una matriz visualmente similar a un mapa de calor donde los meses de alta demanda destacan en colores intensos y los de baja demanda en colores claros. La comparación visual entre las filas de `Perecedero` y `No Perecedero` permite identificar si ambas categorías comparten los mismos meses pico o si sus picos ocurren en momentos distintos.

**Verificación:** Los valores numéricos en las celdas son coherentes con los totales visibles en los gráficos de líneas y área apilada construidos en los pasos anteriores. La fila de totales (si está activada) debe mostrar la suma correcta de ambas categorías por mes.

---

### Paso 6: Agregar un segmentador de Categoría a la página 3

**Objetivo:** Agregar un segmentador (Slicer) de Categoría en la página actual antes de configurar la sincronización, para que la función Sync Slicers tenga un segmentador de referencia en esta página.

**Instrucciones:**

1. Hacer clic en un área vacía del lienzo (puede ser una esquina superior o un espacio lateral).
2. En el panel **Visualizaciones**, hacer clic en el ícono de **Segmentador** (Slicer) — el ícono con un embudo o filtro.
3. Arrastrar la columna `Categoría` al campo **"Campo"** (Field) del segmentador.
4. Redimensionar el segmentador para que sea compacto (aproximadamente 15% del ancho del lienzo) y posicionarlo en la esquina superior derecha del lienzo, sin tapar los gráficos principales.
5. En el panel de formato del segmentador, configurar:
   - **Título:** Activado, texto: `Filtrar por Categoría`.
   - **Estilo:** Seleccionar **"Lista"** (List) para que muestre las opciones `Perecedero` y `No Perecedero` como casillas de verificación.
6. Verificar que al hacer clic en `Perecedero` en el segmentador, los tres visuales de la página (gráfico de líneas, área apilada y matriz) se filtran mostrando solo esa categoría. Luego hacer clic en el ícono de borrar filtro (X o eraser) del segmentador para restablecer la vista completa.

**Resultado esperado:** El segmentador de Categoría funciona correctamente en la página 3, filtrando todos los visuales de la página de forma interactiva.

**Verificación:** Al seleccionar `No Perecedero` en el segmentador, el gráfico de líneas muestra solo una línea, el área apilada muestra solo un color y la matriz muestra solo la fila de `No Perecedero`. Al deseleccionar, vuelven a aparecer ambas categorías.

---

### Paso 7: Configurar la sincronización de segmentadores entre páginas (Sync Slicers)

**Objetivo:** Activar la función Sync Slicers de Power BI Desktop para sincronizar el segmentador de Categoría entre las tres páginas del reporte, de modo que al filtrar en cualquier página, todas las demás se actualicen con el mismo filtro.

**Instrucciones:**

1. Asegurarse de estar en la página `Análisis por Categoría` (página 3) con el segmentador de Categoría seleccionado (clic sobre él para que tenga el borde de selección azul).
2. En la cinta superior de Power BI Desktop, ir a la pestaña **"Ver"** (View).
3. En el grupo de opciones, hacer clic en **"Sincronizar segmentadores"** (Sync Slicers). Se abrirá un panel flotante en el lado izquierdo o derecho de la pantalla llamado **"Sincronizar segmentadores"**.
4. El panel mostrará una lista de todas las páginas del reporte con dos columnas de íconos:
   - **Columna de sincronización** (ícono de dos flechas circulares): activa la sincronización bidireccional.
   - **Columna de visibilidad** (ícono de ojo): controla si el segmentador es visible en esa página.
5. Configurar la sincronización de la siguiente manera:

   | Página                    | Sincronización | Visibilidad |
   |---------------------------|:--------------:|:-----------:|
   | Tendencias de Demanda     | ✅ Activado    | ✅ Activado |
   | Análisis de Variabilidad  | ✅ Activado    | ✅ Activado |
   | Análisis por Categoría    | ✅ Activado    | ✅ Activado |

   > Para activar cada opción, hacer clic en el ícono correspondiente en cada fila. Los íconos activados se mostrarán resaltados o con fondo de color.

6. Cerrar el panel de Sync Slicers haciendo clic en la **"X"** del panel o volviendo a hacer clic en "Sincronizar segmentadores" en la cinta.
7. **Verificar la sincronización:**
   a. En el segmentador de la página 3, seleccionar `Perecedero`.
   b. Navegar a la página 1 (`Tendencias de Demanda`) haciendo clic en su pestaña.
   c. Verificar que el segmentador en la página 1 también muestra `Perecedero` seleccionado y que los visuales de esa página están filtrados.
   d. Navegar a la página 2 (`Análisis de Variabilidad`) y verificar el mismo comportamiento.
   e. Regresar a la página 3 y hacer clic en el ícono de borrar filtro del segmentador para restablecer la selección completa. Verificar que las páginas 1 y 2 también se restablecen.

**Resultado esperado:** Al seleccionar cualquier valor en el segmentador de Categoría en cualquiera de las tres páginas, las otras dos páginas se actualizan automáticamente con el mismo filtro. El reporte se comporta como una unidad coherente e integrada.

**Verificación:** Navegar entre las tres páginas con `No Perecedero` seleccionado en el segmentador. En todas las páginas, los visuales deben mostrar exclusivamente datos de la categoría `No Perecedero`. Al borrar el filtro en cualquier página, todas las páginas muestran datos de ambas categorías.

> ⚠️ **Nota sobre compatibilidad de segmentadores:** Para que Sync Slicers funcione correctamente, los segmentadores de las diferentes páginas deben estar basados en el **mismo campo** de la **misma tabla** del modelo de datos. Si los segmentadores de las páginas 1 y 2 fueron creados con un campo diferente o una tabla duplicada, la sincronización no funcionará. En ese caso, verificar la fuente del campo `Categoría` en cada segmentador.

---

### Paso 8: Revisión final e integración del dashboard del Módulo 5

**Objetivo:** Revisar el reporte completo del Módulo 5 como producto integrado, verificar la coherencia visual y narrativa entre las tres páginas, y guardar el archivo final.

**Instrucciones:**

1. **Revisión de la página 3 — Análisis por Categoría:**
   - Verificar que los tres visuales (gráfico de líneas, área apilada y matriz heat map) están correctamente posicionados sin superponerse.
   - Confirmar que todos los títulos de los visuales son descriptivos y legibles.
   - Asegurarse de que el segmentador de Categoría es visible y accesible.

2. **Revisión de coherencia entre páginas:**
   - Navegar por las tres páginas en orden y verificar que la experiencia de navegación es fluida.
   - Confirmar que los colores usados para `Perecedero` y `No Perecedero` son **consistentes** entre el gráfico de líneas de la página 3 y cualquier visual de páginas anteriores que use la misma categoría. Si hay inconsistencia de colores, esto es normal y aceptable en esta práctica.

3. **Agregar navegación entre páginas (opcional si el tiempo lo permite):**
   - En la página 3, ir a la cinta **"Insertar"** → **"Botones"** → **"Atrás"** (Back button).
   - Posicionar el botón en la esquina inferior izquierda del lienzo.
   - Repetir en las páginas 1 y 2 si aún no tienen botones de navegación de prácticas anteriores.

4. **Reflexión analítica — responder mentalmente (o anotar) las siguientes preguntas:**
   - ¿Las tendencias de demanda de perecederos y no perecederos son divergentes o paralelas? ¿Qué implicación tiene esto para la planificación de inventario?
   - ¿La composición del mix cambia en los meses de mayor demanda total? ¿Qué categoría domina en los picos?
   - ¿Los meses de mayor demanda de perecederos coinciden con los de no perecederos según el heat map? ¿Cómo impacta esto en la operación de picking y en las necesidades de almacenamiento diferenciado?

5. **Guardar el archivo:**
   - Presionar **Ctrl + S** para guardar el archivo `.pbix`.
   - Confirmar que el nombre del archivo incluye una referencia al módulo y práctica, por ejemplo: `Modulo5_Dashboard_Completo_P25.pbix`.
   - Si el instructor lo requiere, exportar el reporte como PDF: ir a **Archivo → Exportar → Exportar a PDF** para generar una versión estática del dashboard completo.

**Resultado esperado:** Un archivo `.pbix` guardado con tres páginas completas, visualizaciones coherentes, segmentadores sincronizados y títulos descriptivos. El reporte representa el producto integrado del Módulo 5 y puede ser presentado como un dashboard analítico de demanda por categoría.

**Verificación:** Al abrir el archivo `.pbix` desde cero (cerrar y volver a abrir), las tres páginas cargan correctamente, los datos se muestran sin errores y los segmentadores funcionan de forma sincronizada.

---

## 7. Validación y Pruebas

Completar la siguiente lista de verificación antes de dar por finalizada la práctica:

| # | Criterio de validación | Estado |
|---|------------------------|--------|
| 1 | El archivo `.pbix` tiene exactamente **tres páginas** nombradas correctamente | ☐ |
| 2 | El gráfico de líneas muestra **dos series** diferenciadas por color (una por categoría) | ☐ |
| 3 | El Analytics Pane muestra **dos líneas de tendencia** (una por categoría) con trazo discontinuo | ☐ |
| 4 | El gráfico de área apilada muestra la **composición proporcional** de ambas categorías en el tiempo | ☐ |
| 5 | La Matriz tiene **formato condicional de escala de colores** aplicado correctamente | ☐ |
| 6 | Al seleccionar una categoría en el segmentador de la página 3, **las páginas 1 y 2 también se filtran** | ☐ |
| 7 | Los valores de la Matriz son **coherentes** con los totales del gráfico de líneas para el mismo período | ☐ |
| 8 | El archivo `.pbix` ha sido **guardado** con nombre descriptivo | ☐ |

**Prueba de integración final:**

Ejecutar la siguiente secuencia de acciones para validar el reporte completo:

1. Ir a la página `Análisis por Categoría`.
2. Seleccionar `Perecedero` en el segmentador.
3. Navegar a `Tendencias de Demanda` → verificar que el filtro está activo.
4. Navegar a `Análisis de Variabilidad` → verificar que el filtro está activo.
5. Regresar a `Análisis por Categoría` → seleccionar `No Perecedero`.
6. Verificar que las tres páginas muestran solo datos de `No Perecedero`.
7. Borrar el filtro → verificar que las tres páginas muestran datos completos.

Si todos los pasos de esta secuencia funcionan correctamente, el dashboard está completamente integrado y la práctica está finalizada con éxito.

---

## 8. Solución de Problemas

### Problema 1: Las líneas de tendencia del Analytics Pane no aparecen separadas por categoría (solo aparece una línea de tendencia general)

**Síntoma:** Al activar la Línea de tendencia en el Analytics Pane, solo aparece una línea de tendencia que representa el total combinado de ambas categorías, en lugar de dos líneas separadas (una por categoría).

**Causa:** El campo `Categoría` no está correctamente asignado al campo **"Leyenda"** del gráfico de líneas, o está asignado a un campo diferente como "Detalles" o "Información sobre herramientas". El Analytics Pane solo genera líneas de tendencia independientes por serie cuando la segmentación se realiza mediante el campo **Leyenda** específicamente.

**Solución:**
1. Seleccionar el gráfico de líneas.
2. En el panel de campos del visual, verificar que `Categoría` está en el campo **"Leyenda"** (no en "Detalles" ni otro campo).
3. Si está en el lugar incorrecto, arrastrar `Categoría` desde el campo actual al campo **"Leyenda"**.
4. Volver al Analytics Pane y verificar que ahora aparecen dos líneas de tendencia.
5. Si el problema persiste, eliminar la línea de tendencia existente haciendo clic en el ícono de papelera junto a ella, y luego volver a agregarla con el botón "Agregar".

---

### Problema 2: La sincronización de segmentadores (Sync Slicers) no funciona entre páginas — los filtros no se propagan

**Síntoma:** Al seleccionar `Perecedero` en el segmentador de la página 3, las páginas 1 y 2 no reflejan el filtro. Al navegar entre páginas, cada segmentador mantiene su selección independiente.

**Causa:** Los segmentadores de las diferentes páginas están basados en campos de **tablas diferentes** o en **copias duplicadas** de la columna `Categoría` en el modelo de datos. Sync Slicers solo sincroniza segmentadores que comparten exactamente el mismo campo de la misma tabla en el modelo. También puede ocurrir que las columnas de sincronización y visibilidad en el panel Sync Slicers no estén activadas para todas las páginas.

**Solución:**
1. Verificar en el panel **Sync Slicers** que tanto la columna de sincronización (flechas) como la de visibilidad (ojo) están activadas para las **tres páginas**. Si alguna fila no tiene ambos íconos activos, hacer clic para activarlos.
2. Si el problema persiste, verificar la fuente del campo `Categoría` en cada segmentador:
   - Hacer clic en el segmentador de la página 1.
   - En el panel de campos, verificar el nombre completo del campo: debe mostrar `NombreTabla[Categoría]`.
   - Repetir para los segmentadores de las páginas 2 y 3.
   - Si los segmentadores usan tablas diferentes (por ejemplo, `Ventas[Categoría]` vs. `Categorias[Categoría]`), eliminar los segmentadores de las páginas 1 y 2, y crear nuevos segmentadores usando exactamente el mismo campo que el segmentador de la página 3.
3. Después de corregir, seleccionar el segmentador de la página 3 y volver a configurar Sync Slicers desde la cinta **Ver → Sincronizar segmentadores**.

---

## 9. Limpieza del Entorno

Al finalizar la práctica, realizar las siguientes acciones de cierre:

1. **Guardar el archivo final:** Presionar **Ctrl + S** una última vez para asegurar que todos los cambios están guardados.
2. **Verificar el nombre del archivo:** Confirmar que el archivo se llama `Modulo5_Dashboard_Completo_P25.pbix` (o el nombre indicado por el instructor).
3. **Cerrar el panel Sync Slicers** si está abierto: ir a **Ver → Sincronizar segmentadores** para cerrarlo y liberar espacio en pantalla.
4. **Restablecer todos los filtros de segmentadores:** Antes de cerrar, asegurarse de que ningún segmentador tiene una selección activa (todos los valores deben estar seleccionados o el segmentador debe estar en estado "sin filtro"), para que el archivo se abra sin filtros preseleccionados la próxima vez.
5. **Exportar a PDF (si el instructor lo solicita):** Ir a **Archivo → Exportar → Exportar a PDF**, guardar el archivo PDF con el mismo nombre base del `.pbix` en la misma carpeta.
6. **Cerrar Power BI Desktop** o dejarlo abierto según las instrucciones del instructor para la sesión de discusión final del módulo.

> **Nota de secuencialidad:** Este archivo `.pbix` es el producto final del Módulo 5. Si el curso continúa con módulos posteriores que requieran este dashboard como insumo, asegurarse de tener una copia de respaldo en una ubicación segura antes de continuar.

---

## 10. Resumen y Recursos Adicionales

### Lo que construiste en esta práctica

En esta práctica final del Módulo 5 completaste los siguientes entregables:

| Entregable | Descripción | Relevancia logística |
|------------|-------------|----------------------|
| **Gráfico de líneas con tendencias** | Dos series (una por categoría) con líneas de tendencia discontinuas individuales | Permite identificar si perecederos y no perecederos requieren estrategias de abastecimiento divergentes a largo plazo |
| **Gráfico de área apilada** | Visualización de la composición del mix de demanda a lo largo del tiempo | Revela si los picos de demanda son impulsados por una categoría específica, afectando la planificación de espacio y temperatura en almacén |
| **Matriz con heat map** | Tabla dinámica con escala de colores por mes y categoría | Herramienta de diagnóstico rápido para identificar meses críticos de alta demanda que requieren mayor stock de seguridad |
| **Sync Slicers configurado** | Segmentadores sincronizados entre las tres páginas | Permite al analista filtrar el reporte completo desde cualquier página, manteniendo coherencia en el análisis multi-perspectiva |
| **Dashboard integrado del Módulo 5** | Tres páginas coherentes como producto analítico completo | Reporte listo para ser presentado a tomadores de decisiones en operaciones logísticas |

### Conexión con decisiones operativas logísticas

La pregunta central de esta práctica es: **¿cómo impacta el análisis por categoría en una decisión operativa de logística?**

- **Tendencias divergentes** entre categorías sugieren que los modelos de pronóstico deben ser independientes por categoría, no un único modelo agregado.
- **Cambios en el mix** durante picos de demanda implican ajustes en la asignación de zonas de picking y en los recursos de almacenamiento diferenciado (temperatura controlada para perecederos).
- **El heat map** permite identificar meses donde el stock de seguridad debe incrementarse selectivamente por categoría, optimizando el capital inmovilizado en inventario.
- **Los segmentadores sincronizados** facilitan el análisis de nivel de servicio segmentado: ¿el nivel de servicio objetivo es el mismo para perecederos y no perecederos? Generalmente no lo es, y este dashboard permite analizar ambas dimensiones de forma integrada.

### Recursos adicionales de referencia

- [Documentación oficial de Microsoft: Analytics Pane en Power BI Desktop](https://learn.microsoft.com/es-es/power-bi/transform-model/desktop-analytics-pane)
- [Documentación oficial de Microsoft: Sincronizar segmentadores en varias páginas](https://learn.microsoft.com/es-es/power-bi/visuals/power-bi-visualization-slicers#sync-and-use-slicers-on-other-pages)
- [Documentación oficial de Microsoft: Formato condicional en visualizaciones de Power BI](https://learn.microsoft.com/es-es/power-bi/create-reports/desktop-conditional-table-formatting)
- [Documentación oficial de Microsoft: Gráfico de área en Power BI](https://learn.microsoft.com/es-es/power-bi/visuals/power-bi-visualization-area-chart)
- [Microsoft Power BI Desktop – Descarga gratuita oficial](https://www.microsoft.com/es-es/download/details.aspx?id=58494)

---

> **Cierre del Módulo 5:** Has completado las cinco prácticas del Módulo 5 (Prácticas 22–25), construyendo progresivamente un dashboard de análisis de demanda en Power BI Desktop. El archivo `.pbix` resultante integra visualizaciones de tendencia, variabilidad y comportamiento por categoría, con filtros sincronizados que permiten un análisis coherente y multidimensional. Este dashboard es el fundamento visual sobre el cual se aplicarán las herramientas de pronóstico y recomendaciones de reabastecimiento en los módulos siguientes del curso.
