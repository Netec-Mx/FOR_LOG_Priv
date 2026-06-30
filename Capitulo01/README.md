# Práctica 1. Análisis de demanda y construcción conceptual de pronóstico para kits de instalación FTTH

## Objetivos
Analizar una base histórica de demanda logística en Excel para identificar patrones de comportamiento y construir un pronóstico conceptual basado en tres componentes principales:

```text
Baseline + Estacionalidad + Iniciativas del negocio
```

Al finalizar la práctica, el participante podrá explicar qué productos tienen mayor impacto operativo, qué meses presentan mayor demanda, qué eventos deben excluirse del análisis normal y cómo una promoción puede modificar la planeación de bodega.

## Duración aproximada
- 144 minutos.

## Tabla de ayuda

Para realizar esta práctica, se recomienda contar con los siguientes recursos:

| Recurso | Descripción |
| --- | --- |
| Microsoft Excel | Herramienta principal para el análisis de datos, creación de tablas dinámicas, filtros y gráficos. |
| Archivos de datos [exploracion_datos.xlsx](https://docs.google.com/spreadsheets/d/1skr10xBSiUVzJrq60d05eyweJHbL-jfc/edit?usp=sharing&ouid=100656380906672528108&rtpof=true&sd=true) y [Pronostico ajustado para kits de instalacion.xlsx](https://docs.google.com/spreadsheets/d/1svt0SDRI5QbXFC41CXH6eXgugNse2Seg/edit?usp=sharing&ouid=100656380906672528108&rtpof=true&sd=true) | Bases históricas de demanda logística utilizadas durante la práctica. |


## Instrucciones
Sigue los pasos a continuación para completar cada actividad que conforma la práctica. Durante el desarrollo, analiza los resultados obtenidos y responde las preguntas planteadas en cada situación a resolver.

## Parte 1. 

### Contexto de la práctica

El área de logística de Claro Ecuador debe estimar la demanda de kits de instalación para servicios de fibra óptica durante el mes de diciembre.

Cada kit contiene los materiales necesarios para realizar una instalación residencial, por lo que una estimación incorrecta puede provocar:

- quiebres de inventario,
- retrasos en instalaciones,
- compras urgentes,
- incremento de costos logísticos.

Para apoyar la planeación de inventarios se analizarán dos años de demanda histórica y posteriormente se incorporará el efecto de una campaña comercial.

Durante esta práctica construirás un pronóstico utilizando los mismos conceptos empleados en procesos reales de planeación de la demanda.

---

### Actividad 1. Calcular el promedio histórico mensual

Antes de construir cualquier pronóstico es necesario resumir la información histórica.

Como existen dos años de datos, el primer paso consiste en obtener el promedio de demanda de cada mes.

1. Abre el archivo de Excel *Pronostico ajustado para kits de instalacion.xlsx*.
2. Observa las columnas:
    - Mes
    - Año 1
    - Año 2
    - Promedio
3. En la columna Promedio debes calcular el promedio del Año 1 y Año 2 de cada Mes.


Recuerda que el promedio se calcula sumando los valores y dividiendo entre la cantidad de valores.

$$
\text{Promedio} = \frac{x_1 + x_2 + \cdots + x_n}{n}
$$

En la celda E13 escribe la fórmula:

```text
=(C13+D13)/2
```

Copia la fórmula hacia abajo hasta diciembre.

4. ¿Por qué es mejor utilizar el promedio de dos años que únicamente un año?

### Actividad 2. Calcular el Baseline

El Baseline representa la demanda normal del producto.

Es la referencia que posteriormente será ajustada por estacionalidad y campañas comerciales.

1. Ubica la columna destinada al Baseline.
2. Calcula el Baseline como el promedio de los doce valores obtenidos anteriormente.

Puedes escribir la fórmula:

```text
=SUMA(E13:E24)/12
```

3. ¿Qué representa este valor? ¿Puede considerarse la demanda "esperada" del producto durante un mes promedio?

### Actividad 3. Calcular el índice estacional

No todos los meses presentan el mismo comportamiento.

El índice estacional mide cuánto se desvía cada mes respecto a la demanda promedio anual.

1. En la columna Índice estacional, calcula el valor correspondiente a diciembre.

$$
\text{Índice estacional} = \frac{\text{Promedio del mes}}{\text{Promedio general(Baseline)}}
$$

Utiliza:

```text
=E24/F24
```

2. Observa el resultado.

Interpreta: 

- Si el índice es mayor que 1, la demanda está por encima del promedio.

- Si es cercano a 1, el comportamiento es normal.

- Si es menor que 1, la demanda se encuentra por debajo del promedio.

3. Calcula el índice estacional para los demás meses.
4. ¿Por qué diciembre presenta un índice diferente al resto de los meses?

### Actividad 4. Medir la variabilidad de la demanda

No basta con conocer el promedio.

También es importante medir qué tan estable o variable es la demanda.

Una demanda muy variable implica mayor incertidumbre para el pronóstico.

1. Calcula la desviación estándar de toda la demanda histórica.

$$
s = \sqrt{\frac{1}{n-1} \sum_{i=1}^{n} (x_i - \bar{x})^2}
$$

Utiliza:

```text
=DESVEST.M(B13:C24)
```

2. Interpreta el resultado: ¿qué tan dispersos están los datos respecto al promedio? 
3. Calcula la variabilidad relativa dividiendo la desviación estándar entre el Baseline.

Utiliza:

```text
=DESVEST.M(B13:C24)/Baseline
```

4. ¿La demanda parece estable? ¿Qué implicaciones tendría una alta variabilidad para la planeación de inventarios?

### Actividad 5. Calcular el Uplift

Las campañas comerciales modifican temporalmente la demanda.

El Uplift mide cuánto aumentó (o disminuyó) la demanda respecto al pronóstico original.

Escenario:

Durante la campaña anterior se esperaba vender: 1,200 kits

Sin embargo, la demanda real fue de: 1,380 kits

1. Calcula el Uplift:

$$
\text{Uplift} = \frac{\text{Demanda real}}{\text{Demanda esperada}} - 1
$$

Utiliza:

```text
=(1380/1200)-1
```

También puedes expresarlo como porcentaje.

2. Analiza el resultado, ¿qué significa un Uplift positivo? ¿Y un Uplift negativo?
3. ¿Qué porcentaje aumentó la demanda? ¿Qué podría explicar este incremento?

### Actividad 6. Construcción del pronóstico final

En la práctica profesional el pronóstico no depende únicamente del promedio histórico.

Debe incorporar el efecto de la estacionalidad y de las iniciativas comerciales.

1. Calcula el pronóstico utilizando la siguiente expresión:

$$
\text{Pronóstico Final} = \text{Baseline} \times \text{Índice estacional} \times (1 + \text{Uplift})
$$

En Excel puede escribirse como:

```text
=Baseline*Indice_Estacional*(1+Uplift)
```

2. ¿El resultado es mayor que el Baseline?
3. ¿Qué componente tuvo mayor impacto sobre el pronóstico?
4. ¿Cómo utilizaría este resultado el área de logística?


El pronóstico final estima la demanda de kits para el mes de diciembre en el siguiente periodo, incorporando el nivel promedio histórico (baseline), la estacionalidad específica del mes y un ajuste incremental (uplift) derivado de campañas comerciales, resultando en un valor de aproximadamente 1633 unidades.


## Parte 2.

### Contexto de la práctica
Claro Ecuador desea mejorar la planeación de inventarios para productos utilizados en instalaciones de fibra óptica, como:

- ONT
- Routers WiFi
- Kits de instalación FTTH
- Cable de fibra
- Conectores
- Splitters
- Accesorios técnicos

El equipo de Bodega debe analizar datos históricos de demanda para anticipar necesidades futuras, reducir quiebres de stock, organizar mejor el picking y prepararse ante campañas comerciales.

Esta práctica integra análisis exploratorio, depuración de datos, identificación de tendencias, estacionalidad, productos de alta rotación, eventos extraordinarios, construcción conceptual del pronóstico y simulación del impacto de una promoción.

---

### Actividad 1. Exploración inicial de la base de datos

Antes de construir un pronóstico, primero se debe entender qué información contiene la base de datos. Un mal análisis suele iniciar cuando no se comprende qué representa cada columna o cuándo se usan datos sin revisar su significado operativo.

En esta actividad revisarás la estructura general de la base y crearás una primera tabla dinámica para identificar qué categorías concentran mayor demanda histórica.

1. Abre el archivo de Excel *exploracion_datos.xlsx*.
2. Ubícate en la hoja `Datos_Demanda`.
3. Observa la primera fila de la tabla. Esa fila contiene los nombres de las columnas.
4. Identifica las siguientes columnas:
    - `Fecha`
    - `Año`
    - `Mes`
    - `Ciudad`
    - `Bodega`
    - `SKU`
    - `SKU_Descripcion`
    - `Categoría`
    - `Unidades_Demandadas`
    - `Unidades_Despachadas`
    - `Lineas_Picking`
    - `Tipo_Evento`
5. Haz clic en cualquier celda dentro de la tabla.
6. En la parte superior de Excel, haz clic en **Insertar**.
7. Selecciona **Tabla dinámica**.
8. En la ventana que aparece, elige **Nueva hoja de cálculo**.
9. Haz clic en **Aceptar**.
10. En el panel derecho llamado **Campos de tabla dinámica**, configura la tabla de la siguiente manera:
    - Arrastra `Categoría` al área **Filas**.
    - Arrastra `Unidades_Demandadas` al área **Valores**.
11. Verifica que el campo `Unidades_Demandadas` se esté calculando como **Suma** y no como **Recuento**.
12. Revisa qué categorías tienen mayor demanda total.
13. Ordena los resultados de mayor a menor:
    - Haz clic derecho sobre cualquier número de la columna de valores.
    - Selecciona **Ordenar > Ordenar de mayor a menor**.

14. ¿Cuáles son las tres categorías con mayor volumen de demanda?
15. ¿Qué implicaciones podrían tener esas categorías para la operación de bodega?
16. ¿Estas categorías deberían recibir mayor atención en la planeación de inventarios? Justifica tu respuesta.

### Actividad 2. Depuración y organización de datos

No todos los datos sirven igual para analizar demanda normal. Algunos registros pueden contener errores, datos incompletos o eventos especiales que alteran el comportamiento esperado de la demanda.

Antes de calcular un promedio histórico o construir un pronóstico, es necesario revisar qué datos representan una operación normal y cuáles deben analizarse con cuidado.

1. Regresa a la hoja `Datos_Demanda`.
2. Haz clic en cualquier celda de la fila de encabezados.
3. Ve a la pestaña **Datos**.
4. Haz clic en **Filtro**.
5. Verifica que en cada encabezado aparezca una flecha desplegable.
6. Haz clic en la flecha de la columna `Registro_Valido`.
7. Revisa si existen registros marcados como:
    - `Sí`
    - `Revisar`
    - `No`
8. Filtra únicamente los registros marcados como `Revisar`.
9. Observa las columnas principales de esos registros e identifica cuál podría ser el motivo de revisión, por ejemplo:
    - demanda inusualmente alta,
    - demanda inusualmente baja,
    - unidades no atendidas,
    - quiebre de stock,
    - evento comercial,
    - registro incompleto
10. Ahora haz clic en la flecha de la columna `Tipo_Evento`.
11. Identifica si existen eventos como:
    - `Promoción comercial`
    - `Quiebre de stock`
    - `Retraso de importación`
    - `Campaña nacional`
    - `Cambio de proveedor`
    - `Sin evento`
12. Observa qué tipo de eventos aparecen con mayor frecuencia en los registros marcados como `Revisar`.
13. Quita el filtro de `Registro_Valido` y `Tipo_Evento`:
    - Haz clic en la flecha de cada columna filtrada.
    - Selecciona **Borrar filtro**.

14. ¿Qué tipo de registros deberían revisarse antes de usarlos para construir un baseline?
15. ¿Qué registros sí podrían usarse directamente para calcular demanda normal?
16. ¿Por qué no conviene mezclar registros normales con registros afectados por eventos extraordinarios?

Para construir un baseline, se recomienda considerar principalmente los registros donde:

- `Registro_Valido` sea igual a `Sí`.
- `Tipo_Evento` sea igual a `Sin evento` o equivalente.
- No existan quiebres de stock que hayan limitado el despacho.
- No existan eventos comerciales que hayan elevado artificialmente la demanda.

Los registros marcados como `Revisar` no necesariamente deben eliminarse, pero sí deben analizarse antes de incluirlos en el cálculo.


### Actividad 3. Identificación visual de tendencias

La tendencia permite reconocer la dirección general de la demanda. Una categoría puede mostrar crecimiento sostenido, disminución gradual o estabilidad. Esta información ayuda a tomar decisiones sobre inventario, espacio de almacenamiento, personal operativo y prioridades de abastecimiento.

En esta actividad construirás una tabla dinámica y un gráfico de líneas para observar el comportamiento de una categoría FTTH a lo largo del tiempo.

1. Haz clic en cualquier celda dentro de la tabla `Datos_Demanda`.
2. Ve a **Insertar > Tabla dinámica**.
3. Selecciona **Nueva hoja de cálculo**.
4. Haz clic en **Aceptar**.
5. En el panel de campos, configura la tabla dinámica de la siguiente manera:
    - Arrastra `Mes_Año` al área **Filas**.
    - Arrastra `Unidades_Demandadas` al área **Valores**.
    - Arrastra `Categoría` al área **Filtros**.
6. En la parte superior de la tabla dinámica aparecerá el filtro `Categoría`.
7. Haz clic en el filtro y selecciona una categoría relacionada con fibra óptica, por ejemplo:
    - `ONT Fibra Óptica`
    - `Kit de instalación FTTH`
8. Selecciona la tabla dinámica.
9. Ve a **Insertar**.
10. Selecciona **Gráfico de líneas**.
11. Observa la línea del gráfico:
    - Si sube con el tiempo, hay una tendencia creciente.
    - Si baja con el tiempo, hay una tendencia decreciente.
    - Si se mantiene similar, hay una tendencia estable.
12. Cambia el filtro de categoría y repite el análisis con `Router WiFi` u otra categoría relevante.

La gerencia evalúa aumentar el espacio de almacenamiento para productos FTTH.
14. ¿La demanda de productos FTTH muestra una tendencia creciente, decreciente o estable?
15. ¿Se justifica preparar más capacidad para estos productos?
16. ¿Qué riesgo tendría ignorar esta tendencia en la planeación de bodega?

El gráfico de líneas te permite interpretar visualmente la tendencia de demanda de una categoría logística.

### Actividad 4. Identificación manual de estacionalidad

La estacionalidad permite detectar meses en los que la demanda sube o baja de forma repetitiva. En logística, reconocer estos patrones ayuda a preparar inventario antes de los picos, reforzar turnos de picking y reducir el riesgo de quiebres de stock.

En esta actividad compararás la demanda por mes y año para identificar si existen meses que se repiten como periodos de alta demanda.

1. Crea una nueva tabla dinámica desde la hoja `Datos_Demanda`.
2. En el panel de campos, configura la tabla de la siguiente manera:
    - Arrastra `Mes` al área **Filas**.
    - Arrastra `Año` al área **Columnas**.
    - Arrastra `Unidades_Demandadas` al área **Valores**.
    - Arrastra `Categoría` al área **Filtros**.
3. En el filtro `Categoría`, selecciona `Kit de instalación FTTH`.
4. Observa los valores por mes y por año.
5. Identifica los meses que aparecen con demanda alta en más de un año.
6. Selecciona la tabla dinámica.
7. Ve a **Insertar > Gráfico de columnas**.
8. Observa si algunos meses se repiten como picos de demanda.
9. Cambia el filtro de categoría a `Router WiFi` y revisa si el patrón se repite.

La bodega solo puede reforzar inventario en cuatro meses del año.
10. ¿Qué cuatro meses elegirías según la estacionalidad observada?
11. ¿Los picos se repiten en más de un año o parecen ser eventos aislados?
12. ¿Qué acciones debería preparar la bodega antes de esos meses?

Un mes puede considerarse estacional si muestra demanda alta de forma repetida en distintos años. En cambio, si solo aparece alto una vez, podría tratarse de un evento extraordinario y no necesariamente de una estacionalidad real.


### Actividad 5. Productos de alta rotación e impacto en picking

No todos los productos afectan igual la operación. Algunos SKU generan alto volumen de unidades demandadas, mientras que otros generan muchas líneas de picking. Ambos indicadores son importantes.

Un producto con muchas unidades demandadas impacta el inventario. Un producto con muchas líneas de picking impacta la carga operativa, porque requiere más movimientos, preparación y despacho.

1. Crea una nueva tabla dinámica desde la hoja `Datos_Demanda`.
2. En el panel derecho, configura la tabla de la siguiente manera:
    - Arrastra `SKU_Descripcion` al área **Filas**.
    - Arrastra `Unidades_Demandadas` al área **Valores**.
3. Verifica que `Unidades_Demandadas` se calcule como **Suma**.
4. Haz clic derecho sobre cualquier número de la columna de valores.
5. Selecciona **Ordenar > Ordenar de mayor a menor**.
6. Observa los primeros 10 productos de la lista.
7. Ahora arrastra `Lineas_Picking` al área **Valores**.
8. Verifica que `Lineas_Picking` se calcule como **Suma**.
9. Compara:
    - productos con más unidades demandadas,
    - productos con más líneas de picking.
10. Identifica productos que aparecen altos en ambos indicadores.
11. Selecciona la tabla dinámica.
12. Ve a **Insertar > Gráfico de barras**.

La bodega solo puede reubicar 5 productos cerca de la zona de despacho.
13. ¿Qué 5 SKU seleccionarías considerando demanda y líneas de picking?
14. ¿Elegirías únicamente los productos con más unidades demandadas? ¿Por qué?
15. ¿Qué beneficio operativo tendría reubicar esos productos cerca de la zona de despacho?

Para seleccionar los 5 SKU prioritarios, considera aquellos que combinen:
- alta demanda,
- alto número de líneas de picking,
- riesgo operativo si no están disponibles o si requieren muchos recorridos dentro de bodega


### Actividad 6. Detección de eventos extraordinarios

Un evento extraordinario puede generar picos o caídas que no representan la demanda normal. Si estos eventos se usan sin analizarlos, el pronóstico puede quedar distorsionado.

Por ejemplo, una promoción comercial puede elevar la demanda temporalmente, mientras que un quiebre de stock puede reducir las unidades despachadas aunque la demanda real siga existiendo.

1. Regresa a la hoja `Datos_Demanda`.
2. Activa los filtros si no están activos:
    - Ve a **Datos > Filtro**.
3. Haz clic en la flecha de la columna `Tipo_Evento`.
4. Revisa los eventos disponibles.
5. Selecciona un evento, por ejemplo:
    - `Promoción comercial`
    - `Quiebre de stock`
    - `Retraso de importación`
6. Observa las siguientes columnas:
    - `Unidades_Demandadas`
    - `Unidades_Despachadas`
    - `Unidades_No_Atendidas`
    - `Quiebre_Stock`
7. Quita el filtro aplicado.
8. Crea una tabla dinámica nueva.
9. En el panel de campos, configura la tabla de la siguiente manera:
    - Arrastra `Tipo_Evento` al área **Filas**.
    - Arrastra `Unidades_Demandadas` al área **Valores**.
    - Arrastra `Unidades_Despachadas` al área **Valores**.
    - Arrastra `Unidades_No_Atendidas` al área **Valores**.
10. Compara los eventos:
    - una promoción suele elevar la demanda,
    - un quiebre de stock puede reducir el despacho,
    - un retraso de importación puede generar demanda no atendida.

Se detectó una caída fuerte en unidades despachadas durante el mes de Enero.
11. ¿Fue una caída real de la demanda o pudo deberse a un evento extraordinario?
12. ¿Qué diferencia observas entre `Unidades_Demandadas` y `Unidades_Despachadas`?
13. ¿Qué pasaría si construyes un pronóstico usando únicamente las unidades despachadas durante un quiebre de stock?

Cuando existe quiebre de stock, las unidades despachadas pueden ser menores que la demanda real. En ese caso, no conviene interpretar la caída del despacho como una caída del mercado o de la necesidad operativa.


### Actividad 7. Construcción conceptual del pronóstico

El pronóstico no debe construirse únicamente considerando el promedio histórico. Primero hay que separar tres elementos:

- **Baseline:** demanda normal esperada sin eventos extraordinarios.
- **Estacionalidad:** ajuste por meses que normalmente suben o bajan.
- **Iniciativas del negocio:** campañas, promociones o acciones comerciales que pueden modificar la demanda esperada.

1. Selecciona un producto de alta rotación.
2. Crea una tabla dinámica nueva.
3. Configura la tabla dinámica: 
- Arrastra `Mes_Año` al área **Filas**.
- Arrastra `Unidades_Demandadas` al área **Valores**.
- Arrastra `SKU_Descripcion` al área **Filtros**.
- Arrastra `Tipo_Evento` al área **Filtros**.
- En el filtro `SKU_Descripcion`, selecciona el producto elegido.
- En el filtro `Tipo_Evento`, selecciona `Sin evento`.
4. Observa la demanda mensual del producto seleccionado.
5. Copia los resultados de la tabla dinámica en una zona libre de la hoja para realizar cálculos.
6. Organiza una pequeña tabla con las siguientes columnas y calcula cada valor:
    - `Mes`
    - `Promedio`
    - `Baseline`
    - `Indice_Estacional`
    - `Pronostico_teórico (sin eventos extraordinarios)`
7. Interpreta el índice estacional:
    - Un índice de `1.00` indica un mes normal.
    - Un índice mayor a `1.00` indica un mes por encima del promedio.
    - Un índice menor a `1.00` indica un mes por debajo del promedio.
8. Calcula el pronóstico conceptual multiplicando el baseline por el índice estacional.

Recuerda que el baseline debe calcularse con datos que representen demanda normal. Por lo tanto, antes de calcularlo, excluye o revisa meses afectados por:
- promociones comerciales,
- campañas nacionales,
- quiebres de stock,
- retrasos de importación,
- cambios de proveedor,
- registros inválidos o incompletos.

Si un mes no tiene datos normales porque todos sus registros estuvieron afectados por eventos extraordinarios, no lo uses directamente para calcular el baseline. En ese caso, documenta que el mes requiere revisión.

Debes preparar el pronóstico base para el siguiente año de operación.
9. ¿Qué baseline usarías para el SKU seleccionado?
10. ¿Qué meses excluirías por no representar demanda normal?
11. ¿Qué meses tienen un índice estacional mayor a `1.00`?
12. ¿Qué meses tienen un índice estacional menor a `1.00`?
13. ¿Qué pronóstico conceptual obtienes para el siguiente año? ¿Por qué?


### Actividad 8. Simulación del impacto de una promoción

Las iniciativas comerciales pueden cambiar la demanda esperada. Bodega debe anticiparse para evitar quiebres de stock, saturación de picking o reabastecimientos urgentes.

Considera el efecto de una campaña nacional de migración a fibra óptica sobre el pronóstico calculado en la actividad anterior.

Claro Ecuador lanzará una campaña nacional de migración a fibra óptica. Marketing estima que la demanda de instalaciones aumentará **35%** durante el próximo año.

1. Usa el SKU trabajado en la Actividad 7.
2. Toma el `Pronostico_Conceptual` calculado.
3. Añade la siguiente columna a la tabla anterior:
- `Incremento por promoción (Uplift)`
4. En las celdas inferiores calcula el incremento esperado por la promoción comercial.

Calcula el pronóstico ajustado con la siguiente fórmula:

```text
=Pronostico_Conceptual*1.35
```

5. Ahora calcula el impacto en líneas de picking. Revisa el promedio de `Lineas_Picking` para el SKU seleccionado integrando `Lineas_Picking`en la tabla dinámica.
6. Aplica también el incremento del 35% a las líneas de picking.

```excel
=Lineas_Picking_Promedio*1.35
```

7. Compara:
    - demanda normal,
    - demanda con promoción,
    - picking normal,
    - picking con promoción.

8. Revisa si el inventario disponible promedio sería suficiente para cubrir el nuevo volumen.
9. Si el inventario parece insuficiente, marca el SKU como riesgo de quiebre.
10. ¿Qué acción recomendarías?
    - aumentar inventario,
    - reforzar picking,
    - adelantar compras,
    - reubicar productos,
    - coordinar entregas con proveedores,
    - combinar varias acciones.
- ¿Cuál sería el riesgo operativo de no anticipar la promoción?


Al finalizar esta actividad, habrás ajustado un pronóstico conceptual por una promoción comercial y podrás proponer una decisión operativa para reducir riesgos de quiebre, sobrecarga de picking o falta de inventario.

---

## Reflexión final
Reflexiona desde el rol del equipo de Bodega:

- ¿Por qué no conviene construir un pronóstico únicamente con el promedio histórico?
- ¿Qué diferencia existe entre demanda normal y demanda afectada por eventos extraordinarios?
- ¿Cómo ayuda la estacionalidad a preparar inventario con anticipación?
- ¿Por qué las líneas de picking también deben considerarse en la planeación?
- ¿Qué información debería compartir Marketing con Bodega antes de lanzar una campaña?
- ¿Qué decisiones operativas pueden tomarse a partir de un pronóstico conceptual?

## Resultado esperado de toda la práctica
Al finalizar la práctica, el participante habrá realizado un análisis completo de demanda histórica en Excel y será capaz de:

- Reconocer la estructura de una base logística.
- Identificar categorías y SKU relevantes para la operación de bodega.
- Depurar datos básicos y distinguir registros normales de registros atípicos.
- Crear tablas dinámicas sencillas.
- Crear gráficos de líneas, columnas y barras.
- Identificar tendencias de demanda.
- Reconocer meses con posible estacionalidad.
- Detectar productos de alta rotación. 
- Diferenciar eventos extraordinarios de demanda normal.
- Construir un baseline simple.
- Ajustar un pronóstico conceptual con estacionalidad. 
- Simular el impacto de una promoción comercial.
- Proponer decisiones operativas para reducir riesgos de quiebre, sobrecarga de picking o falta de inventario.
