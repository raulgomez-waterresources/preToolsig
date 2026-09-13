# ============================================================
# preToolsig — Script de completado de series de precipitación
# por regresión lineal simple entre estaciones, con componente
# SIG (tabla de sumatorios por año hidrológico, visualización,
# selección de estación auxiliar por criterios geográficos,
# cascada de modelos para maximizar el completado y mapa de
# estaciones)
# Versión v4 — sept. 2026
# ============================================================
#
# NOTA DE PROCEDENCIA: este script parte del motor de cálculo de
# preTools v9 (proyecto anterior, archivo preTools.R, sin
# modificaciones en su lógica de regresión/completado). La v1 de
# preToolsig añadió, sin tocar esa lógica, la tabla de sumatorios
# mensuales por año hidrológico y el gráfico interactivo de la serie
# (idea 5). La v2 añadió el filtro geográfico opcional de estación
# auxiliar por distancia y diferencia de altitud (idea 2, un único
# modelo ganador). Esta v3 sustituye ese único modelo, CUANDO SE USA
# TABLA DE ESTACIONES, por una cascada que prueba todas las auxiliares
# geográficamente válidas por orden de R² antes de dejar una fecha sin
# completar, con una reserva opcional (desactivada por defecto) para
# recurrir, como último recurso y de forma señalada, a auxiliares
# fuera del filtro geográfico. Sin tabla de estaciones, el
# comportamiento sigue siendo idéntico al de antes de la idea 2. Esta
# v4 añade el mapa interactivo de estaciones (idea 4): reutiliza la
# misma tabla de estaciones y la misma tabla de modelos ya calculadas,
# sin repetir ningún cálculo de distancia/altitud/selección, solo los
# traduce a colores y texto sobre un mapa. El proyecto preTools
# original (preTools.R, v9) se mantiene intacto como referencia.
#
# ------------------------------------------------------------
# INSTRUCCIONES DE USO (léelas antes de tocar nada más)
# ------------------------------------------------------------
#
# Este archivo NO ejecuta nada por sí solo: solo DEFINE funciones.
# Para usarlo necesitas dos pasos siempre, en este orden:
#
#   PASO 1 — Cargar las funciones en tu sesión de R
#   ------------------------------------------------
#   Con este archivo abierto y activo en el editor de RStudio,
#   pulsa Ctrl+Shift+S (Windows/Linux) o Cmd+Shift+S (macOS).
#   Esto define todas las funciones del script en tu entorno, pero
#   no procesa ningún Excel todavía.
#
#   PASO 2 — Ejecutar el proceso sobre tu archivo
#   ------------------------------------------------
#   Elige UNA de estas dos formas (no hace falta hacer las dos):
#
#   Opción A (recomendada mientras estés haciendo pruebas):
#   escribe esto directamente en la CONSOLA de RStudio (no en este
#   archivo) y pulsa Enter:
#
#       resultado <- completar_primera_estacion(
#         ruta_excel = "RUTA/A/TU/ARCHIVO.xlsx",
#         ruta_salida = "NOMBRE_SALIDA.xlsx"
#       )
#
#   Opción B (si quieres dejarlo guardado en el script): ve al
#   final de este archivo, a la sección "Ejemplo de uso", quita
#   las almohadillas (#) de esas líneas y sustituye la ruta de
#   ejemplo por la tuya. Guarda el archivo (Ctrl+S) y repite el
#   PASO 1 (Ctrl+Shift+S): al no estar ya comentada, esa llamada
#   se ejecutará automáticamente al cargar el script.
#
#   Parámetros opcionales de completar_primera_estacion() (si no
#   los indicas, se usan estos valores por defecto):
#     hoja                       = 1     (número de hoja del Excel a leer)
#     min_n                      = 100   (mínimo de datos coincidentes)
#     min_r2                     = 0.5   (R² mínimo para aceptar un modelo)
#     limitar_negativos          = TRUE  (recorta a 0 estimaciones negativas)
#     ruta_gpkg                  = NULL  (tabla maestra de estaciones, idea 2;
#                                          sin ella, no hay filtro geográfico)
#     max_distancia_m            = NULL  (idea 2; requiere ruta_gpkg)
#     max_diferencia_altitud_m   = NULL  (idea 2; requiere ruta_gpkg)
#     permitir_reserva_fuera_filtro = FALSE  (idea 2, opción 3; requiere
#                                          ruta_gpkg; ver sección 7bis)
#
#   Qué esperar en la consola al ejecutar: mensajes informativos
#   (message()) sobre qué estación auxiliar se ha seleccionado y
#   cuántos valores se han completado, y avisos (warning()) si se
#   detectan filas con fecha inválida, columnas vacías, valores
#   negativos en el Excel de origen, o estaciones sin coordenadas en
#   la tabla de estaciones (idea 2). Ningún warning detiene el
#   proceso; un error (stop()) sí lo detendría, y ocurre si el Excel
#   no tiene la estructura mínima esperada (al menos 3 columnas:
#   fecha, objetivo, auxiliar), o si el .gpkg no tiene las columnas
#   obligatorias ('Estacion', 'Z_m') o no tiene CRS definido.
#
#   El resultado se guarda en el archivo indicado en ruta_salida,
#   con las hojas Datos (datos originales, sin completar), Datos_
#   completados, Registro_completados, Modelos_regresion (incluye
#   distancia_m, diferencia_altitud_m y aviso_coordenadas cuando se
#   usa ruta_gpkg), una hoja por cada par estación objetivo -
#   estación auxiliar (DEP_<objetivo>_<auxiliar>), y dos hojas de
#   sumatorios mensuales por año hidrológico: Resumen_mensual_AH_
#   original y Resumen_mensual_AH_completado (ver sección 10 más
#   abajo).
#
#   En Datos_completados, junto a la columna de la estación objetivo,
#   se añade una columna "<objetivo>_estado" con una de estas
#   etiquetas por fila: "Original", "Completado (regresión)",
#   "Completado (auxiliar=0)" o "SIN COMPLETAR". Esta última señala
#   las filas que no se han podido rellenar (la estación auxiliar
#   seleccionada tampoco tenía dato esa fecha, o no hubo ningún
#   modelo válido). Se usa esta columna de texto en lugar de un
#   valor numérico centinela (p. ej. 999) para no introducir en la
#   columna de precipitación un número que pueda confundirse con un
#   dato real en cálculos posteriores.
# ------------------------------------------------------------
#
# NOTA METODOLÓGICA (vigente desde v9 de preTools, heredada aquí):
# Para cada par (estación objetivo, estación auxiliar) se utilizan
# TODOS los registros coincidentes disponibles entre ambas, sin
# exigir que el resto de estaciones auxiliares tengan dato ese día
# (filtrado POR PARES). Esto sustituye al criterio de "excel
# depurado global" usado en versiones anteriores.
#
# CONSECUENCIA A TENER EN CUENTA (no se corrige automáticamente,
# por instrucción expresa de no complicar la metodología): el
# número de datos coincidentes (n) es distinto para cada estación
# auxiliar evaluada, ya que depende únicamente de su solape temporal
# con la estación objetivo. La selección del mejor modelo compara
# R² entre auxiliares con n potencialmente muy distintos entre sí.
# Esto es una práctica habitual y aceptada, pero conviene tenerlo
# presente: un R² ligeramente mayor con un n mucho menor no es
# necesariamente "mejor" en sentido estadístico estricto que un R²
# ligeramente menor con un n mucho mayor. No se introduce ningún
# criterio adicional (ponderación por n, RMSE, etc.) para corregir
# esto, tal como se ha solicitado explícitamente.

# Paquetes necesarios
library(readxl)
library(dplyr)
library(purrr)
library(writexl)
library(stringr)
library(tibble)
library(tidyr)
library(plotly)
library(sf)
library(leaflet)

# ============================================================
# 1. LECTURA
# ============================================================
#
# Lee el Excel de origen. Se asume:
#   - primera columna = fecha
#   - segunda columna = estación objetivo (a completar)
#   - resto de columnas = estaciones auxiliares
#   - primera fila = encabezados
#
# Se fuerza explícitamente el tipo de cada columna (fecha / numeric)
# en lugar de dejar que read_excel() lo adivine. Se detectó en
# versiones anteriores que, para determinados archivos, readxl puede
# malinterpretar el formato interno de una columna numérica y
# devolver solo su parte fraccionaria (p. ej. 3.9 -> 0.9), sin
# lanzar error ni aviso. Forzar col_types evita ese problema.

leer_datos_precipitacion <- function(ruta_excel,
                                     hoja = 1,
                                     col_fecha = 1) {

  cabecera <- read_excel(ruta_excel, sheet = hoja, n_max = 0)
  n_cols <- ncol(cabecera)

  if (n_cols < 3) {
    stop("El archivo debe tener al menos 3 columnas: fecha, estación objetivo y al menos una estación auxiliar.")
  }

  tipos_columnas <- rep("numeric", n_cols)
  tipos_columnas[col_fecha] <- "date"

  df <- read_excel(ruta_excel, sheet = hoja, col_types = tipos_columnas)

  # --- Fechas no interpretables: se eliminan solo esas filas ---
  idx_fecha_invalida <- which(is.na(df[[col_fecha]]))
  if (length(idx_fecha_invalida) > 0) {
    warning(
      length(idx_fecha_invalida),
      " fila(s) con fecha vacía o no interpretable han sido eliminadas antes del análisis ",
      "(filas de datos: ", paste(head(idx_fecha_invalida, 10), collapse = ", "),
      if (length(idx_fecha_invalida) > 10) ", ..." else "",
      "). Revisa si son filas espurias (totales, fórmulas residuales, celdas vacías, etc.) en el Excel de origen."
    )
    df <- df[-idx_fecha_invalida, ]
  }

  return(df)
}

# ============================================================
# 2. VALIDACIÓN
# ============================================================
#
# Controles objetivamente detectables sobre los datos ya leídos:
#   - columnas completamente vacías
#   - valores negativos de precipitación (físicamente inválidos)
#
# Deliberadamente NO se incluyen heurísticas basadas en el valor
# máximo de una columna (p. ej. "máximo sospechosamente bajo"):
# una estación con precipitaciones reales bajas es perfectamente
# posible y no debe marcarse como anómala solo por eso.
#
# Esta función solo informa mediante warning(); no modifica ni
# elimina datos.

validar_datos_precipitacion <- function(df, col_fecha = 1) {
  cols_datos <- setdiff(colnames(df), colnames(df)[col_fecha])

  for (cn in cols_datos) {
    valores <- df[[cn]]

    if (all(is.na(valores))) {
      warning("La columna '", cn, "' está completamente vacía (todos los valores son NA).")
      next
    }

    idx_negativos <- which(valores < 0)
    if (length(idx_negativos) > 0) {
      warning(
        "La columna '", cn, "' contiene ", length(idx_negativos),
        " valor(es) negativo(s), físicamente inválidos para precipitación ",
        "(filas de datos: ", paste(head(idx_negativos, 10), collapse = ", "),
        if (length(idx_negativos) > 10) ", ..." else "",
        "). Revisa el Excel de origen; el script NO modifica estos valores automáticamente."
      )
    }
  }

  invisible(df)
}

# ============================================================
# 3. PREPARACIÓN DE DATOS POR PARES
# ============================================================
#
# Para un par (estación objetivo, estación auxiliar), conserva
# únicamente las filas en las que AMBAS estaciones tienen dato.
# La ausencia de dato en otras estaciones auxiliares no afecta
# a este subconjunto.

preparar_datos_par <- function(df, col_fecha, col_target, col_aux) {
  df_par <- df %>%
    select(all_of(c(col_fecha, col_target, col_aux))) %>%
    filter(!is.na(.data[[col_target]]),
           !is.na(.data[[col_aux]]))

  return(df_par)
}

# ============================================================
# 4. AJUSTE DE REGRESIONES (con control de validez)
# ============================================================
#
# Ajusta Y = a + bX (estación objetivo ~ estación auxiliar) sobre
# el subconjunto por pares. Antes de aceptar el modelo comprueba:
#   - número mínimo de observaciones para poder ajustar una recta
#   - variabilidad en la auxiliar y en la objetivo
#   - que el ajuste no produzca error
#   - que los coeficientes sean finitos (no NA/NaN/Inf)
#   - que R² sea calculable y finito
#
# Si cualquiera de estos controles falla, el modelo se marca como
# NO VÁLIDO (valido = FALSE) con un motivo textual, y la función
# NUNCA detiene la ejecución: simplemente informa mediante el
# campo 'motivo' para que quede reflejado en la tabla resumen y se
# continúe con el resto de estaciones auxiliares.
#
# NOTA: el mínimo técnico de observaciones para ajustar una recta
# se fija internamente en 3 (para tener al menos 1 grado de libertad
# tras estimar 2 coeficientes). Esto es independiente del parámetro
# de usuario `min_n`, que se aplica más adelante, en la selección
# del mejor modelo — un modelo puede ser técnicamente VÁLIDO aquí y
# aun así no ser SELECCIONADO por no alcanzar min_n o min_r2.

N_MINIMO_TECNICO <- 3

ajustar_modelo_par <- function(df_par, col_target, col_aux) {

  n_datos <- nrow(df_par)

  resultado_base <- list(
    target = col_target,
    aux = col_aux,
    intercept = NA_real_,
    slope = NA_real_,
    r2 = NA_real_,
    n = n_datos,
    valido = FALSE,
    motivo = NA_character_,
    modelo = NULL
  )

  if (n_datos < N_MINIMO_TECNICO) {
    resultado_base$motivo <- paste0("n insuficiente para ajustar (n=", n_datos,
                                     ", mínimo técnico=", N_MINIMO_TECNICO, ")")
    return(resultado_base)
  }

  x <- df_par[[col_aux]]
  y <- df_par[[col_target]]

  if (stats::var(x, na.rm = TRUE) == 0) {
    resultado_base$motivo <- "la estación auxiliar no tiene variabilidad (varianza = 0)"
    return(resultado_base)
  }

  if (stats::var(y, na.rm = TRUE) == 0) {
    resultado_base$motivo <- "la estación objetivo no tiene variabilidad (varianza = 0)"
    return(resultado_base)
  }

  ajuste <- tryCatch({
    # reformulate() evita fallos cuando los nombres de columna
    # contienen espacios, tildes u otros caracteres no válidos en
    # sintaxis de fórmula.
    formula_modelo <- reformulate(col_aux, response = col_target)
    modelo <- lm(formula_modelo, data = df_par)
    resumen <- summary(modelo)
    list(
      intercept = unname(coef(modelo)[1]),
      slope     = unname(coef(modelo)[2]),
      r2        = resumen$r.squared,
      modelo    = modelo,
      error     = NULL
    )
  }, error = function(e) {
    list(intercept = NA_real_, slope = NA_real_, r2 = NA_real_,
         modelo = NULL, error = conditionMessage(e))
  })

  if (!is.null(ajuste$error)) {
    resultado_base$motivo <- paste0("error al ajustar el modelo: ", ajuste$error)
    return(resultado_base)
  }

  coeficientes_validos <- is.finite(ajuste$intercept) && is.finite(ajuste$slope)
  r2_valido <- is.finite(ajuste$r2)

  if (!coeficientes_validos) {
    resultado_base$motivo <- "coeficientes no finitos (NA/NaN/Inf)"
    return(resultado_base)
  }

  if (!r2_valido) {
    resultado_base$motivo <- "R² no calculable (NA/NaN/Inf)"
    return(resultado_base)
  }

  list(
    target = col_target,
    aux = col_aux,
    intercept = ajuste$intercept,
    slope = ajuste$slope,
    r2 = ajuste$r2,
    n = n_datos,
    valido = TRUE,
    motivo = NA_character_,
    modelo = ajuste$modelo
  )
}

# ============================================================
# 5. EVALUACIÓN DE MODELOS PARA UNA ESTACIÓN OBJETIVO
# ============================================================
#
# Construye el subconjunto por pares y ajusta el modelo para CADA
# estación auxiliar disponible. Devuelve:
#   - tabla_modelos: una fila por auxiliar evaluada (válida o no)
#   - datos_pares: lista con el data.frame usado en cada par
#     (necesario para exportar la hoja individual de cada par)

evaluar_modelos <- function(df, col_fecha, col_target, cols_aux) {

  datos_pares <- map(cols_aux, ~ preparar_datos_par(df, col_fecha, col_target, .x))
  names(datos_pares) <- cols_aux

  resultados <- map2(datos_pares, cols_aux,
                      ~ ajustar_modelo_par(.x, col_target, .y))

  tabla_modelos <- map_dfr(resultados, ~ tibble(
    estacion_objetivo = .x$target,
    estacion_auxiliar = .x$aux,
    intercept = .x$intercept,
    slope = .x$slope,
    r2 = .x$r2,
    n = .x$n,
    valido = .x$valido,
    motivo = .x$motivo
  ))

  message("Estación '", col_target, "': ", nrow(tabla_modelos),
          " estación(es) auxiliar(es) evaluada(s), ",
          sum(tabla_modelos$valido), " modelo(s) válido(s).")

  list(
    tabla_modelos = tabla_modelos,
    datos_pares = datos_pares
  )
}

# ============================================================
# 6. SELECCIÓN DEL MEJOR MODELO
# ============================================================
#
# Entre los modelos VÁLIDOS que cumplen n >= min_n y R² >= min_r2, y
# que además superan el filtro geográfico opcional (idea 2, ver más
# abajo), selecciona el de mayor R². Si ninguno cumple, no se
# completa esa estación (se informa mediante message(), no se
# detiene el script).
#
# FILTRO GEOGRÁFICO (NUEVO — idea 2): max_distancia_m y
# max_diferencia_altitud_m son opcionales (por defecto NULL = sin
# filtro, para no alterar el comportamiento de quien no use tabla de
# estaciones). Cuando se indican, se exige distancia_m <=
# max_distancia_m Y diferencia_altitud_m <= max_diferencia_altitud_m
# (ambos a la vez) ANTES de comparar por R². Es decir: el criterio de
# prioridad acordado es primero la coherencia geográfica/orográfica,
# y solo entre las auxiliares que la cumplen se elige la de mayor R².
# Una auxiliar con mejor R² pero fuera de los umbrales queda
# descartada aunque una alternativa geográficamente más razonable
# tenga peor R².
#
# Si una estación (objetivo o auxiliar) no tiene coordenadas
# conocidas, distancia_m/diferencia_altitud_m llegan como NA desde
# calcular_distancia_altitud(). El filtro se escribe deliberadamente
# como "is.na(...) | ... <= umbral" para que esas auxiliares NO se
# descarten por falta de coordenadas (se evalúan sin restricción
# geográfica, tal como se decidió), en vez de perderlas por el
# comportamiento por defecto de filter() con NA.

seleccionar_mejor_modelo <- function(tabla_modelos, min_n = 100, min_r2 = 0.5,
                                     max_distancia_m = NULL,
                                     max_diferencia_altitud_m = NULL) {

  tabla_filtrada <- tabla_modelos %>%
    filter(valido,
           n >= min_n,
           r2 >= min_r2)

  if (!is.null(max_distancia_m)) {
    tabla_filtrada <- tabla_filtrada %>%
      filter(is.na(distancia_m) | distancia_m <= max_distancia_m)
  }

  if (!is.null(max_diferencia_altitud_m)) {
    tabla_filtrada <- tabla_filtrada %>%
      filter(is.na(diferencia_altitud_m) | diferencia_altitud_m <= max_diferencia_altitud_m)
  }

  tabla_filtrada <- tabla_filtrada %>% arrange(desc(r2))

  if (nrow(tabla_filtrada) == 0) {
    return(NULL)
  }

  tabla_filtrada[1, ]
}

# ============================================================
# 7. COMPLETADO DE LA ESTACIÓN OBJETIVO
# ============================================================
#
# Aplica el modelo seleccionado sobre el DATASET ORIGINAL COMPLETO
# (no sobre los subconjuntos por pares usados para el ajuste).
#
# Reglas:
#   - auxiliar == 0        -> completado = 0 (no se aplica regresión)
#   - auxiliar != 0 y != NA -> completado = a + b * auxiliar
#   - auxiliar == NA        -> no se puede completar; permanece NA
#
# `limitar_negativos` (por defecto TRUE) controla si un valor
# estimado por regresión que resulte negativo se recorta a 0.
# La precipitación no puede ser físicamente negativa, pero se deja
# como parámetro explícito para que la decisión sea visible y
# configurable, en lugar de una regla oculta en el código.
#
# Devuelve una lista con:
#   - datos: el data.frame completado
#   - registro: tibble con una fila por cada valor efectivamente
#     completado (trazabilidad), o un tibble vacío si no hubo
#     completado.

completar_estacion <- function(df,
                               col_fecha,
                               col_target,
                               mejor_modelo,
                               limitar_negativos = TRUE) {

  # NOTA: en lugar de rellenar las filas sin dato y sin posibilidad de
  # completado con un valor centinela (p. ej. 999), se añade una
  # columna de trazabilidad "<objetivo>_estado" con las etiquetas
  # "Original", "Completado (regresión)", "Completado (auxiliar=0)" y
  # "SIN COMPLETAR". Esto permite identificar de un vistazo (filtro de
  # Excel) las filas sin completar, sin introducir en la columna de
  # precipitación un valor numérico que pueda confundirse con un dato
  # real en cálculos posteriores.
  col_estado <- paste0(col_target, "_estado")

  columnas_registro <- c("fecha", "estacion_objetivo", "valor_original",
                          "valor_completado", "estacion_auxiliar_utilizada",
                          "valor_auxiliar", "metodo_completado",
                          "intercepto", "pendiente", "r2_modelo", "n_modelo")

  registro_vacio <- setNames(
    as.list(rep(list(character(0)), length(columnas_registro))),
    columnas_registro
  ) %>% as_tibble()

  df_out <- df
  idx_na_objetivo <- which(is.na(df_out[[col_target]]))

  estado <- rep("Original", nrow(df_out))
  estado[idx_na_objetivo] <- "SIN COMPLETAR"

  if (is.null(mejor_modelo)) {
    message("No se completa la estación '", col_target,
            "': ningún modelo cumple los criterios mínimos (n, R²).")
    df_out[[col_estado]] <- estado
    df_out <- df_out %>% relocate(all_of(col_estado), .after = all_of(col_target))
    return(list(datos = df_out, registro = registro_vacio))
  }

  aux_col <- mejor_modelo$estacion_auxiliar
  a       <- mejor_modelo$intercept
  b       <- mejor_modelo$slope

  message("Estación '", col_target, "' completada con auxiliar '", aux_col,
          "' (R2 = ", round(mejor_modelo$r2, 3), ", n = ", mejor_modelo$n, ").")

  # Caso 3 (auxiliar NA): se excluyen de idx_completable y permanecen NA
  # (quedarán marcadas como "SIN COMPLETAR" en la columna de estado)
  idx_completable <- idx_na_objetivo[!is.na(df_out[[aux_col]][idx_na_objetivo])]

  registro <- vector("list", length(idx_completable))

  for (k in seq_along(idx_completable)) {
    i <- idx_completable[k]
    x_aux <- df_out[[aux_col]][i]

    if (x_aux == 0) {
      valor_completado <- 0
      metodo <- "Auxiliar = 0"
      estado[i] <- "Completado (auxiliar=0)"
    } else {
      valor_estimado <- a + b * x_aux
      if (limitar_negativos && valor_estimado < 0) {
        valor_completado <- 0
      } else {
        valor_completado <- valor_estimado
      }
      metodo <- "Regresión lineal"
      estado[i] <- "Completado (regresión)"
    }

    df_out[[col_target]][i] <- valor_completado

    registro[[k]] <- tibble(
      fecha = df_out[[col_fecha]][i],
      estacion_objetivo = col_target,
      valor_original = NA_real_,
      valor_completado = valor_completado,
      estacion_auxiliar_utilizada = aux_col,
      valor_auxiliar = x_aux,
      metodo_completado = metodo,
      intercepto = a,
      pendiente = b,
      r2_modelo = mejor_modelo$r2,
      n_modelo = mejor_modelo$n
    )
  }

  registro_df <- if (length(registro) > 0) bind_rows(registro) else registro_vacio

  df_out[[col_estado]] <- estado
  df_out <- df_out %>% relocate(all_of(col_estado), .after = all_of(col_target))

  n_na_restantes <- sum(is.na(df_out[[col_target]]))
  message("  -> ", length(idx_completable), " valores completados. ",
          n_na_restantes, " permanecen NA / SIN COMPLETAR (auxiliar también NA en esas fechas).")

  list(datos = df_out, registro = registro_df)
}

# ============================================================
# 7bis. CASCADA DE MODELOS CANDIDATOS (NUEVO — idea 2, opciones 2 y 3)
# ============================================================
#
# CONTEXTO: con un único modelo "ganador" (completar_estacion(), ya
# usada cuando NO hay tabla de estaciones), una fecha sin dato queda
# SIN COMPLETAR si la auxiliar elegida no tiene dato ESE día, aunque
# otra auxiliar igualmente válida sí lo tuviera. Al añadir el filtro
# geográfico esto se notó más: forzar la auxiliar geográficamente más
# coherente puede dejar mucha menos cobertura si esa auxiliar tiene
# poco solape temporal con la objetivo.
#
# SOLUCIÓN ACORDADA (opción 2, comportamiento por defecto en cuanto se
# usa tabla de estaciones): en vez de un único modelo, se prueban TODOS
# los modelos que superan min_n/min_r2 y el filtro geográfico
# (candidatos_filtro), ordenados de mayor a menor R². Para cada fecha
# sin dato se intenta primero el de mayor R²; si esa auxiliar tampoco
# tiene dato ese día, se prueba la siguiente; y así sucesivamente. Un
# día solo queda SIN COMPLETAR si NINGUNA auxiliar geográficamente
# válida tiene dato ese día. Nunca se sacrifica rigor geográfico por
# cobertura: todas las auxiliares de esta cascada cumplen el filtro.
#
# OPCIÓN 3 (interruptor opcional, permitir_reserva_fuera_filtro =
# FALSE por defecto): si tras agotar candidatos_filtro aún quedan
# fechas sin completar, y el responsable del proyecto ha activado
# explícitamente este parámetro, se prueban además candidatos_reserva
# (modelos válidos que el filtro geográfico había descartado),
# también de mayor a menor R². Cada valor completado así queda
# etiquetado de forma DISTINTA y VISIBLE ("..., fuera de filtro
# geográfico") en la columna de estado, en Registro_completados y en
# el gráfico, para que quede claro cuáles son "limpios" y cuáles son
# un compromiso consciente. Con el interruptor en FALSE (por defecto),
# candidatos_reserva se calcula igualmente (para dejarlo documentado
# en Modelos_regresion) pero NUNCA se usa para completar nada.

seleccionar_modelos_candidatos <- function(tabla_modelos, min_n = 100, min_r2 = 0.5,
                                           max_distancia_m = NULL,
                                           max_diferencia_altitud_m = NULL) {

  base <- tabla_modelos %>% filter(valido, n >= min_n, r2 >= min_r2)

  if (nrow(base) == 0) {
    return(list(candidatos_filtro = base, candidatos_reserva = base))
  }

  cumple <- rep(TRUE, nrow(base))
  if (!is.null(max_distancia_m)) {
    cumple <- cumple & (is.na(base$distancia_m) | base$distancia_m <= max_distancia_m)
  }
  if (!is.null(max_diferencia_altitud_m)) {
    cumple <- cumple & (is.na(base$diferencia_altitud_m) | base$diferencia_altitud_m <= max_diferencia_altitud_m)
  }

  list(
    candidatos_filtro  = base[cumple, , drop = FALSE] %>% arrange(desc(r2)),
    candidatos_reserva = base[!cumple, , drop = FALSE] %>% arrange(desc(r2))
  )
}

completar_estacion_cascada <- function(df,
                                       col_fecha,
                                       col_target,
                                       candidatos_filtro,
                                       candidatos_reserva = NULL,
                                       limitar_negativos = TRUE,
                                       permitir_reserva_fuera_filtro = FALSE) {

  col_estado <- paste0(col_target, "_estado")

  columnas_registro <- c("fecha", "estacion_objetivo", "valor_original",
                          "valor_completado", "estacion_auxiliar_utilizada",
                          "valor_auxiliar", "metodo_completado",
                          "intercepto", "pendiente", "r2_modelo", "n_modelo")

  registro_vacio <- setNames(
    as.list(rep(list(character(0)), length(columnas_registro))),
    columnas_registro
  ) %>% as_tibble()

  df_out <- df
  idx_na_objetivo <- which(is.na(df_out[[col_target]]))

  estado <- rep("Original", nrow(df_out))
  estado[idx_na_objetivo] <- "SIN COMPLETAR"

  hay_filtro  <- !is.null(candidatos_filtro)  && nrow(candidatos_filtro)  > 0
  hay_reserva <- !is.null(candidatos_reserva) && nrow(candidatos_reserva) > 0

  if (!hay_filtro && !(permitir_reserva_fuera_filtro && hay_reserva)) {
    message("No se completa la estación '", col_target,
            "': ningún modelo cumple los criterios mínimos (n, R², y filtro geográfico si aplica).")
    df_out[[col_estado]] <- estado
    df_out <- df_out %>% relocate(all_of(col_estado), .after = all_of(col_target))
    return(list(datos = df_out, registro = registro_vacio))
  }

  # Lista de intentos, en orden: primero TODOS los candidatos dentro
  # del filtro (mayor a menor R²); solo si se permite la reserva, se
  # añaden al final los candidatos fuera de filtro (también por R²).
  lista_intentos <- list()
  if (hay_filtro) {
    for (k in seq_len(nrow(candidatos_filtro))) {
      lista_intentos[[length(lista_intentos) + 1]] <- list(modelo = candidatos_filtro[k, ], reserva = FALSE)
    }
  }
  if (permitir_reserva_fuera_filtro && hay_reserva) {
    for (k in seq_len(nrow(candidatos_reserva))) {
      lista_intentos[[length(lista_intentos) + 1]] <- list(modelo = candidatos_reserva[k, ], reserva = TRUE)
    }
  }

  registro <- list()
  pendientes <- idx_na_objetivo

  for (intento in lista_intentos) {
    if (length(pendientes) == 0) break

    modelo  <- intento$modelo
    aux_col <- modelo$estacion_auxiliar
    a       <- modelo$intercept
    b       <- modelo$slope
    reserva <- intento$reserva

    # De las fechas aún pendientes, nos quedamos solo con las que
    # ESTA auxiliar sí tiene dato.
    idx_completable <- pendientes[!is.na(df_out[[aux_col]][pendientes])]
    if (length(idx_completable) == 0) next

    message("Estación '", col_target, "': completando ", length(idx_completable),
            " fecha(s) pendiente(s) con auxiliar '", aux_col, "' (R2 = ", round(modelo$r2, 3), ")",
            if (reserva) " [fuera de filtro geográfico]" else "", ".")

    for (i in idx_completable) {
      x_aux <- df_out[[aux_col]][i]

      if (x_aux == 0) {
        valor_completado <- 0
        metodo <- if (reserva) "Auxiliar = 0 (fuera de filtro geográfico)" else "Auxiliar = 0"
        estado[i] <- if (reserva) "Completado (auxiliar=0, fuera de filtro geográfico)" else "Completado (auxiliar=0)"
      } else {
        valor_estimado <- a + b * x_aux
        valor_completado <- if (limitar_negativos && valor_estimado < 0) 0 else valor_estimado
        metodo <- if (reserva) "Regresión lineal (fuera de filtro geográfico)" else "Regresión lineal"
        estado[i] <- if (reserva) "Completado (regresión, fuera de filtro geográfico)" else "Completado (regresión)"
      }

      df_out[[col_target]][i] <- valor_completado

      registro[[length(registro) + 1]] <- tibble(
        fecha = df_out[[col_fecha]][i],
        estacion_objetivo = col_target,
        valor_original = NA_real_,
        valor_completado = valor_completado,
        estacion_auxiliar_utilizada = aux_col,
        valor_auxiliar = x_aux,
        metodo_completado = metodo,
        intercepto = a,
        pendiente = b,
        r2_modelo = modelo$r2,
        n_modelo = modelo$n
      )
    }

    pendientes <- setdiff(pendientes, idx_completable)
  }

  registro_df <- if (length(registro) > 0) bind_rows(registro) else registro_vacio

  df_out[[col_estado]] <- estado
  df_out <- df_out %>% relocate(all_of(col_estado), .after = all_of(col_target))

  n_completados   <- nrow(registro_df)
  n_na_restantes  <- length(pendientes)
  message("Estación '", col_target, "' completada en cascada: ", n_completados,
          " valores completados en total. ", n_na_restantes,
          " permanecen SIN COMPLETAR (ninguna auxiliar disponible tenía dato esas fechas).")

  list(datos = df_out, registro = registro_df)
}

# ============================================================
# 8. NOMBRES DE HOJA VÁLIDOS Y ÚNICOS
# ============================================================
#
# Excel exige nombres de hoja de máximo 31 caracteres y prohíbe los
# caracteres: [ ] : * ? / \
# Esta función sanea el nombre y garantiza unicidad frente a los
# nombres ya usados, reservando espacio para un sufijo numérico si
# hiciera falta.

sanear_nombre_hoja <- function(nombre) {
  nombre <- gsub("[\\[\\]:*?/\\\\]", "_", nombre)
  substr(nombre, 1, 31)
}

generar_nombre_hoja_par <- function(prefijo, col_target, col_aux, nombres_usados) {
  base <- sanear_nombre_hoja(paste0(prefijo, "_", col_target, "_", col_aux))

  if (!(base %in% nombres_usados)) {
    return(base)
  }

  sufijo <- 2
  repeat {
    sufijo_str <- paste0("_", sufijo)
    corte <- 31 - nchar(sufijo_str)
    candidato <- paste0(substr(base, 1, corte), sufijo_str)
    if (!(candidato %in% nombres_usados)) {
      return(candidato)
    }
    sufijo <- sufijo + 1
  }
}

# ============================================================
# 8bis. TABLA MAESTRA DE ESTACIONES — SIG (NUEVO — idea 2)
# ============================================================
#
# La tabla maestra de estaciones es un archivo GeoPackage (.gpkg),
# separado del Excel de precipitación y reutilizable entre proyectos,
# con al menos estas columnas:
#   - Estacion    (texto; debe coincidir EXACTO con el nombre de
#                  columna usado en el Excel de precipitación)
#   - Z_m         (numérico, metros — altitud de la estación)
#   - geometría de punto (X, Y) en el CRS del proyecto
#
# CRS del proyecto: EPSG:25830 (ETRS89 / UTM huso 30N), que es el que
# usa Raúl en QGIS para las estaciones de esta cuenca. Se fija como
# constante en vez de dejarlo implícito, para que quede documentado
# en un único sitio y sea fácil de cambiar si algún día se trabaja
# con estaciones de otro huso.

CRS_ESTACIONES <- 25830  # ETRS89 / UTM huso 30N

# cargar_tabla_estaciones()
#
# Lee el .gpkg y comprueba que tenga las columnas mínimas necesarias.
# Si el archivo viene en un CRS distinto a CRS_ESTACIONES, se
# reproyecta automáticamente (en vez de asumir silenciosamente que ya
# está en el CRS correcto, lo que desplazaría todas las distancias
# calculadas sin que se note a simple vista). Si no tiene CRS
# definido, se detiene con stop(): sin esa información no se puede
# garantizar que las distancias sean correctas.

cargar_tabla_estaciones <- function(ruta_gpkg, crs_esperado = CRS_ESTACIONES) {

  tabla <- sf::st_read(ruta_gpkg, quiet = TRUE)

  columnas_necesarias <- c("Estacion", "Z_m")
  faltan <- setdiff(columnas_necesarias, colnames(tabla))
  if (length(faltan) > 0) {
    stop("La tabla de estaciones (", ruta_gpkg, ") no tiene las columnas obligatorias: ",
         paste(faltan, collapse = ", "),
         ". Se esperan al menos 'Estacion', 'Z_m' y una geometría de punto.")
  }

  crs_actual <- sf::st_crs(tabla)
  if (is.na(crs_actual)) {
    stop("La tabla de estaciones (", ruta_gpkg, ") no tiene un CRS definido. ",
         "Asigna un CRS válido (se esperaba EPSG:", crs_esperado, ") antes de continuar.")
  }

  epsg_actual <- crs_actual$epsg
  if (is.na(epsg_actual) || epsg_actual != crs_esperado) {
    message("La tabla de estaciones está en un CRS distinto al esperado (EPSG:", crs_esperado,
            "); se reproyecta automáticamente.")
    tabla <- sf::st_transform(tabla, crs_esperado)
  }

  if (any(duplicated(tabla$Estacion))) {
    stop("La tabla de estaciones tiene nombres de 'Estacion' duplicados: ",
         paste(unique(tabla$Estacion[duplicated(tabla$Estacion)]), collapse = ", "),
         ". Cada estación debe aparecer una única vez.")
  }

  tabla
}

# calcular_distancia_altitud()
#
# Para la estación objetivo y cada estación auxiliar, calcula la
# distancia (metros, euclídea directa sobre el CRS proyectado
# CRS_ESTACIONES) y la diferencia de altitud (metros, valor
# absoluto). Si la objetivo o una auxiliar concreta no aparece en la
# tabla de estaciones, esa fila queda con distancia_m/
# diferencia_altitud_m = NA y un aviso en la columna
# aviso_coordenadas, y se lanza un warning() — pero NO se descarta:
# se decidió explícitamente evaluarla sin filtro geográfico en ese
# caso, por si tuviera un buen ajuste estadístico aunque falte su
# ubicación exacta.

calcular_distancia_altitud <- function(tabla_estaciones, col_target, cols_aux) {

  resultado <- tibble(
    estacion_auxiliar = cols_aux,
    distancia_m = NA_real_,
    diferencia_altitud_m = NA_real_,
    aviso_coordenadas = NA_character_
  )

  fila_objetivo <- tabla_estaciones[tabla_estaciones$Estacion == col_target, ]

  if (nrow(fila_objetivo) == 0) {
    aviso <- paste0("Sin coordenadas para la estación objetivo '", col_target,
                     "': ninguna auxiliar de esta ejecución llevará filtro geográfico.")
    warning(aviso)
    resultado$aviso_coordenadas <- aviso
    return(resultado)
  }

  for (i in seq_along(cols_aux)) {
    aux <- cols_aux[i]
    fila_aux <- tabla_estaciones[tabla_estaciones$Estacion == aux, ]

    if (nrow(fila_aux) == 0) {
      aviso <- paste0("Sin coordenadas para la estación auxiliar '", aux,
                       "': se evalúa sin filtro geográfico (no se descarta por este motivo).")
      warning(aviso)
      resultado$aviso_coordenadas[i] <- aviso
      next
    }

    resultado$distancia_m[i] <- as.numeric(sf::st_distance(fila_objetivo, fila_aux))
    resultado$diferencia_altitud_m[i] <- abs(fila_objetivo$Z_m[1] - fila_aux$Z_m[1])
  }

  resultado
}

# ============================================================
# 9. PIPELINE PARA UNA ESTACIÓN OBJETIVO (núcleo reutilizable)
# ============================================================
#
# Encapsula: evaluación de modelos -> (filtro geográfico) -> selección
# -> completado. No lee ni escribe ficheros: recibe un data.frame ya
# leído y devuelve resultados en memoria. Esta separación es
# intencionada para permitir, en el futuro, iterar esta misma función
# sobre varias columnas objetivo sin duplicar lógica.
#
# tabla_estaciones, max_distancia_m y max_diferencia_altitud_m son
# opcionales (NULL por defecto): sin ellos, el comportamiento es
# idéntico al de antes de la idea 2.

completar_estacion_objetivo <- function(df,
                                        col_fecha,
                                        col_target,
                                        cols_aux,
                                        min_n = 100,
                                        min_r2 = 0.5,
                                        limitar_negativos = TRUE,
                                        tabla_estaciones = NULL,
                                        max_distancia_m = NULL,
                                        max_diferencia_altitud_m = NULL,
                                        permitir_reserva_fuera_filtro = FALSE) {

  evaluacion <- evaluar_modelos(df, col_fecha, col_target, cols_aux)

  # NUEVO (idea 2): si se ha proporcionado tabla de estaciones, se
  # calcula distancia y diferencia de altitud de cada auxiliar
  # respecto a la objetivo, y se incorporan a la tabla de modelos
  # ANTES de seleccionar el mejor, para que la selección pueda
  # aplicarlas como filtro. Sin tabla de estaciones, estas columnas
  # quedan a NA (comportamiento idéntico al de antes de la idea 2).
  if (!is.null(tabla_estaciones)) {
    info_geo <- calcular_distancia_altitud(tabla_estaciones, col_target, cols_aux)
    evaluacion$tabla_modelos <- evaluacion$tabla_modelos %>%
      left_join(info_geo, by = "estacion_auxiliar")
  } else {
    evaluacion$tabla_modelos <- evaluacion$tabla_modelos %>%
      mutate(distancia_m = NA_real_,
             diferencia_altitud_m = NA_real_,
             aviso_coordenadas = NA_character_)
  }

  if (!is.null(tabla_estaciones)) {

    # NUEVO (idea 2, opción 2 por defecto + opción 3 opcional):
    # cascada de modelos en vez de un único ganador. Ver sección 7bis.
    candidatos <- seleccionar_modelos_candidatos(evaluacion$tabla_modelos, min_n, min_r2,
                                                 max_distancia_m, max_diferencia_altitud_m)

    resultado_completado <- completar_estacion_cascada(
      df, col_fecha, col_target,
      candidatos_filtro = candidatos$candidatos_filtro,
      candidatos_reserva = candidatos$candidatos_reserva,
      limitar_negativos = limitar_negativos,
      permitir_reserva_fuera_filtro = permitir_reserva_fuera_filtro
    )

    # Transparencia en Modelos_regresion: orden en que se intentó cada
    # auxiliar dentro de la cascada (1 = primera intentada), si supera
    # o no el filtro geográfico, y cuántos valores completó REALMENTE
    # esa auxiliar en concreto (puede ser 0 si una auxiliar de mayor
    # prioridad ya cubrió todas sus fechas coincidentes).
    orden_filtro <- candidatos$candidatos_filtro %>%
      mutate(cumple_filtro_geografico = TRUE, orden_prioridad = row_number()) %>%
      select(estacion_auxiliar, cumple_filtro_geografico, orden_prioridad)

    orden_reserva <- candidatos$candidatos_reserva %>%
      mutate(cumple_filtro_geografico = FALSE,
             orden_prioridad = nrow(candidatos$candidatos_filtro) + row_number()) %>%
      select(estacion_auxiliar, cumple_filtro_geografico, orden_prioridad)

    orden_todos <- bind_rows(orden_filtro, orden_reserva)

    n_por_aux <- if (nrow(resultado_completado$registro) > 0) {
      resultado_completado$registro %>%
        count(estacion_auxiliar_utilizada, name = "n_valores_completados") %>%
        rename(estacion_auxiliar = estacion_auxiliar_utilizada)
    } else {
      tibble(estacion_auxiliar = character(0), n_valores_completados = integer(0))
    }

    evaluacion$tabla_modelos <- evaluacion$tabla_modelos %>%
      left_join(orden_todos, by = "estacion_auxiliar") %>%
      left_join(n_por_aux, by = "estacion_auxiliar") %>%
      mutate(n_valores_completados = tidyr::replace_na(n_valores_completados, 0L))

    # "seleccionado" se mantiene por compatibilidad con la interfaz de
    # Shiny (que muestra "una" estación auxiliar destacada): marca la
    # primera auxiliar realmente intentada en la cascada (orden 1),
    # solo si esa auxiliar llegó a intentarse de verdad (dentro de
    # filtro siempre; fuera de filtro solo si el interruptor de
    # reserva estaba activado).
    mejor_modelo <- NULL
    if (nrow(candidatos$candidatos_filtro) > 0) {
      mejor_modelo <- candidatos$candidatos_filtro[1, ]
    } else if (permitir_reserva_fuera_filtro && nrow(candidatos$candidatos_reserva) > 0) {
      mejor_modelo <- candidatos$candidatos_reserva[1, ]
    }

  } else {
    # Comportamiento idéntico al de antes de la idea 2: un único
    # modelo ganador para toda la serie.
    mejor_modelo <- seleccionar_mejor_modelo(evaluacion$tabla_modelos, min_n, min_r2)
    resultado_completado <- completar_estacion(df, col_fecha, col_target,
                                               mejor_modelo, limitar_negativos)

    evaluacion$tabla_modelos <- evaluacion$tabla_modelos %>%
      mutate(cumple_filtro_geografico = NA,
             orden_prioridad = NA_integer_,
             n_valores_completados = NA_integer_)
  }

  tabla_modelos_final <- evaluacion$tabla_modelos %>%
    mutate(
      ecuacion = if_else(
        valido,
        paste0(estacion_objetivo, " = ", round(intercept, 4),
               " + ", round(slope, 4), " * ", estacion_auxiliar),
        NA_character_
      ),
      seleccionado = if (is.null(mejor_modelo)) {
        FALSE
      } else {
        valido &
          (estacion_objetivo == mejor_modelo$estacion_objetivo) &
          (estacion_auxiliar == mejor_modelo$estacion_auxiliar)
      }
    ) %>%
    select(estacion_objetivo, estacion_auxiliar, intercept, slope,
           ecuacion, r2, n, distancia_m, diferencia_altitud_m,
           cumple_filtro_geografico, orden_prioridad, n_valores_completados,
           valido, motivo, aviso_coordenadas, seleccionado) %>%
    rename(intercepto = intercept, pendiente = slope) %>%
    arrange(desc(seleccionado), orden_prioridad, desc(valido), desc(r2))

  list(
    datos_completados = resultado_completado$datos,
    registro_completados = resultado_completado$registro,
    tabla_modelos = tabla_modelos_final,
    datos_pares = evaluacion$datos_pares,
    mejor_modelo = mejor_modelo
  )
}

# ============================================================
# 10. AÑO HIDROLÓGICO Y TABLA DE SUMATORIOS MENSUALES (NUEVO — idea 5)
# ============================================================
#
# El año hidrológico (AH) en la Península Ibérica va de octubre a
# septiembre. Regla de asignación: si el mes de la fecha es >= 10
# (oct/nov/dic), el AH "empieza" ese mismo año civil; si el mes es
# <= 9 (ene...sep), el AH empezó el año civil anterior. Es decir,
# oct-2010 a sep-2011 pertenecen al mismo AH, cuyo "año de inicio"
# es 2010.
#
# Esta función NO decide etiquetas AH1/AH2/...: solo calcula el año
# de inicio de cada fecha, para que generar_tabla_resumen_ah()
# construya las etiquetas en orden cronológico sobre el conjunto de
# datos que le llegue (necesario para que "AH1" sea siempre el primer
# año hidrológico presente en ESE dataframe, ya sea el de datos
# originales o el de datos completados).

calcular_ah_inicio <- function(fecha) {
  mes  <- as.integer(format(fecha, "%m"))
  anio <- as.integer(format(fecha, "%Y"))
  if_else(mes >= 10, anio, anio - 1L)
}

MESES_AH <- c("Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo",
              "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre")

# generar_tabla_resumen_ah()
#
# Construye la tabla de sumatorios mensuales por año hidrológico
# para UNA columna de valores (se llama una vez sobre los datos
# originales y otra sobre los datos completados; ver sección 11).
#
# Formato de cada celda: "<suma redondeada a 1 decimal> (n=<n>)",
# donde n es el número de días con dato (no NA) que han contribuido
# a esa celda concreta (mes x AH). Se deja como texto, igual que la
# columna "<objetivo>_estado", para que el n vaya siempre pegado a
# su sumatorio y no se pueda leer un valor sin saber cuántos días lo
# respaldan.
#
# Devuelve una lista con:
#   - tabla: la tabla ancha (filas = meses AH + fila "Total AH",
#     columnas = AH1, AH2, ...)
#   - leyenda: tibble con el rango de calendario real de cada AH
#     (columnas: etiqueta_ah, ah_inicio, rango), para incluirla como
#     hoja de referencia común a ambas tablas (original/completada).

generar_tabla_resumen_ah <- function(df, col_fecha, col_valor) {

  fechas  <- df[[col_fecha]]
  valores <- df[[col_valor]]

  ah_inicio <- calcular_ah_inicio(fechas)
  mes_num   <- as.integer(format(fechas, "%m"))
  nombre_mes <- MESES_AH[match(mes_num, c(10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9))]

  base <- tibble(
    ah_inicio = ah_inicio,
    mes = factor(nombre_mes, levels = MESES_AH),
    valor = valores
  )

  ah_ordenados <- sort(unique(ah_inicio))
  etiquetas_ah <- paste0("AH", seq_along(ah_ordenados))
  names(etiquetas_ah) <- as.character(ah_ordenados)

  leyenda <- tibble(
    etiqueta_ah = etiquetas_ah,
    ah_inicio = ah_ordenados,
    rango = paste0("oct-", ah_ordenados, " a sep-", ah_ordenados + 1L)
  )

  # --- sumatorio mensual (mes x AH) ---
  resumen_mensual <- base %>%
    group_by(ah_inicio, mes) %>%
    summarise(suma = sum(valor, na.rm = TRUE), n = sum(!is.na(valor)), .groups = "drop") %>%
    mutate(
      celda = paste0(round(suma, 1), " (n=", n, ")"),
      etiqueta_ah = etiquetas_ah[as.character(ah_inicio)]
    )

  tabla_mensual <- resumen_mensual %>%
    select(mes, etiqueta_ah, celda) %>%
    tidyr::pivot_wider(names_from = etiqueta_ah, values_from = celda) %>%
    arrange(match(mes, MESES_AH)) %>%
    rename(Mes = mes)

  # --- fila de totales anuales por AH ---
  resumen_anual <- base %>%
    group_by(ah_inicio) %>%
    summarise(suma = sum(valor, na.rm = TRUE), n = sum(!is.na(valor)), .groups = "drop") %>%
    mutate(
      celda = paste0(round(suma, 1), " (n=", n, ")"),
      etiqueta_ah = etiquetas_ah[as.character(ah_inicio)]
    )

  fila_total <- resumen_anual %>%
    select(etiqueta_ah, celda) %>%
    tidyr::pivot_wider(names_from = etiqueta_ah, values_from = celda) %>%
    mutate(Mes = "Total AH", .before = 1)

  # Asegura que fila_total tenga las mismas columnas AH que tabla_mensual,
  # por si algún AH no tuviera ningún mes con datos (caso límite).
  cols_ah_faltantes <- setdiff(colnames(tabla_mensual), colnames(fila_total))
  for (cc in cols_ah_faltantes) fila_total[[cc]] <- NA_character_
  fila_total <- fila_total[, colnames(tabla_mensual)]

  tabla_final <- bind_rows(tabla_mensual, fila_total)

  list(tabla = tabla_final, leyenda = leyenda)
}

# ============================================================
# 11. GRÁFICO INTERACTIVO DE LA SERIE (NUEVO — idea 5)
# ============================================================
#
# Genera un gráfico plotly de la serie ya completada: línea continua
# con el valor final de cada fecha, y puntos superpuestos coloreados
# según "<objetivo>_estado" (Original / Completado (regresión) /
# Completado (auxiliar=0) / SIN COMPLETAR). Incluye un rangeslider
# nativo de plotly en el eje X para poder hacer zoom sobre series
# largas sin necesidad de un control adicional en la interfaz.
#
# No depende de Shiny: puede llamarse directamente desde la consola
# de R sobre el resultado de completar_primera_estacion() para
# inspeccionar la serie sin necesidad de abrir la app.

generar_grafico_serie <- function(datos_completados, col_fecha, col_target, col_estado) {

  colores_estado <- c(
    "Original" = "#1f6f8b",
    "Completado (regresión)" = "#e07b39",
    "Completado (auxiliar=0)" = "#8a6d3b",
    "SIN COMPLETAR" = "#c0392b",
    # NUEVO (idea 2, opción 3): valores completados recurriendo a la
    # reserva fuera de filtro geográfico. Colores diferenciados pero
    # de la misma familia que sus equivalentes "dentro de filtro",
    # para que se distingan a simple vista sin dejar de asociarse
    # visualmente con su categoría (regresión / auxiliar=0).
    "Completado (regresión, fuera de filtro geográfico)" = "#f4a261",
    "Completado (auxiliar=0, fuera de filtro geográfico)" = "#bfa16b"
  )

  # CORRECCIÓN (tras revisión con datos de prueba): se construye una
  # secuencia diaria continua entre la primera y la última fecha del
  # dataset, y se le unen los datos reales. Cualquier fecha AUSENTE en
  # el Excel de origen (no una celda en blanco, sino una fila que
  # directamente no existe) queda como NA en esta secuencia completa.
  # Con connectgaps = FALSE, la línea de la serie se corta en esos
  # huecos en lugar de dibujar una conexión recta que sugeriría una
  # continuidad de datos que no existe. Esto NO afecta a los puntos
  # coloreados por estado, que se siguen dibujando únicamente sobre
  # las filas reales de datos_completados.
  rango_fechas <- seq(min(datos_completados[[col_fecha]]), max(datos_completados[[col_fecha]]), by = "day")
  serie_completa <- tibble(fecha_aux = rango_fechas) %>%
    left_join(datos_completados, by = c("fecha_aux" = col_fecha))

  p <- plot_ly()

  p <- add_trace(
    p,
    x = serie_completa[["fecha_aux"]],
    y = serie_completa[[col_target]],
    type = "scatter",
    mode = "lines",
    connectgaps = FALSE,
    line = list(color = "#a0a0a0", width = 1),
    name = "Serie completada",
    hoverinfo = "skip"
  )

  # NOTA: la categoría "SIN COMPLETAR" aparecerá en la leyenda pero sin
  # ningún punto visible en el gráfico. Es el comportamiento esperado:
  # esas filas no tienen valor (col_target es NA), así que no hay nada
  # que dibujar en el eje Y para ellas; su presencia en la leyenda solo
  # confirma que la categoría existe y cuántas filas tiene el dataset
  # con ese estado (visible en el resumen de texto, no en el gráfico).
  for (est in names(colores_estado)) {
    subset_est <- datos_completados[datos_completados[[col_estado]] == est, ]
    if (nrow(subset_est) == 0) next

    p <- add_trace(
      p,
      data = subset_est,
      x = subset_est[[col_fecha]],
      y = subset_est[[col_target]],
      type = "scatter",
      mode = "markers",
      marker = list(color = colores_estado[[est]], size = 5),
      name = est,
      text = paste0(
        "Fecha: ", format(subset_est[[col_fecha]], "%d/%m/%Y"),
        "<br>Valor: ", round(subset_est[[col_target]], 2),
        "<br>Estado: ", est
      ),
      hoverinfo = "text"
    )
  }

  # SEGUNDA CORRECCIÓN (tras las capturas con datos reales de COBE y
  # con el Excel de prueba): con el título del gráfico Y la leyenda
  # compartiendo la misma franja superior, ambos textos se solapaban y
  # quedaban ilegibles. Se elimina el título propio del gráfico (la
  # estación objetivo ya se muestra en el texto "Estación objetivo: "
  # justo encima, en la interfaz de Shiny) y se deja únicamente la
  # leyenda arriba, con margen superior suficiente para que no se
  # recorte.
  p <- layout(
    p,
    xaxis = list(title = "Fecha", rangeslider = list(visible = TRUE)),
    yaxis = list(title = "Precipitación (mm)"),
    legend = list(orientation = "h", x = 0.5, xanchor = "center", y = 1.15, yanchor = "bottom"),
    margin = list(t = 70, b = 60)
  )

  p
}

# ============================================================
# 11bis. MAPA DE ESTACIONES (NUEVO — idea 4)
# ============================================================
#
# Dibuja un mapa interactivo (leaflet) con TODAS las estaciones de la
# tabla maestra: la objetivo señalada de forma distinta, y cada
# auxiliar coloreada según su papel en ESTA ejecución concreta:
#   - "Objetivo": la estación que se está completando.
#   - "Auxiliar usada": ha completado al menos un valor (esté dentro
#     o fuera del filtro geográfico; el popup lo aclara).
#   - "Auxiliar evaluada (sin uso)": pasó el filtro geográfico y los
#     criterios de n/R², pero un candidato de mayor prioridad ya cubrió
#     todas sus fechas coincidentes.
#   - "Descartada por filtro geográfico": no se usó porque no cumplía
#     la distancia y/o diferencia de altitud máximas indicadas.
#   - "Descartada (n/R² insuficiente)": no llega al mínimo estadístico,
#     independientemente de dónde esté.
#
# Reutiliza tabla_modelos tal cual la genera completar_estacion_
# objetivo(): no repite ningún cálculo de distancia/altitud/selección,
# solo los traduce a colores y texto para el mapa.
#
# La tabla de estaciones se reproyecta a WGS84 (EPSG:4326) SOLO para
# la visualización — todos los cálculos de distancia/altitud ya se
# hicieron antes, en CRS_ESTACIONES (EPSG:25830), y no se repiten ni
# se ven afectados por esta reproyección.

generar_mapa_estaciones <- function(tabla_estaciones, col_target, tabla_modelos) {

  tabla_wgs84 <- sf::st_transform(tabla_estaciones, 4326)
  coords <- sf::st_coordinates(tabla_wgs84)
  tabla_wgs84$lon <- coords[, 1]
  tabla_wgs84$lat <- coords[, 2]

  info <- tabla_modelos %>%
    select(estacion_auxiliar, r2, n, distancia_m, diferencia_altitud_m,
           cumple_filtro_geografico, n_valores_completados, valido) %>%
    rename(Estacion = estacion_auxiliar)

  tabla_mapa <- tabla_wgs84 %>% left_join(info, by = "Estacion")

  tabla_mapa$categoria <- case_when(
    tabla_mapa$Estacion == col_target ~ "Objetivo",
    is.na(tabla_mapa$valido) ~ "No evaluada en esta ejecución",
    !tabla_mapa$valido ~ "Descartada (n/R² insuficiente)",
    tabla_mapa$n_valores_completados > 0 ~ "Auxiliar usada",
    tabla_mapa$cumple_filtro_geografico %in% TRUE ~ "Auxiliar evaluada (sin uso)",
    tabla_mapa$cumple_filtro_geografico %in% FALSE ~ "Descartada por filtro geográfico",
    TRUE ~ "Auxiliar evaluada (sin uso)"
  )

  colores_categoria <- c(
    "Objetivo" = "#c0392b",
    "Auxiliar usada" = "#1f8a4c",
    "Auxiliar evaluada (sin uso)" = "#1f6f8b",
    "Descartada por filtro geográfico" = "#8a6d3b",
    "Descartada (n/R² insuficiente)" = "#7f7f7f",
    "No evaluada en esta ejecución" = "#b0b0b0"
  )
  tabla_mapa$color <- colores_categoria[tabla_mapa$categoria]

  tabla_mapa$popup <- purrr::pmap_chr(
    list(tabla_mapa$Estacion, tabla_mapa$categoria, tabla_mapa$Z_m,
         tabla_mapa$r2, tabla_mapa$n, tabla_mapa$distancia_m,
         tabla_mapa$diferencia_altitud_m, tabla_mapa$n_valores_completados),
    function(nombre, categoria, z, r2, n, dist, desnivel, n_completados) {
      base <- paste0("<b>", nombre, "</b><br>Altitud: ", z, " m<br>", categoria)
      if (categoria == "Objetivo") return(base)
      extra <- paste0(
        "<br>R\u00b2: ", if (is.na(r2)) "-" else round(r2, 3),
        "<br>Datos coincidentes: ", if (is.na(n)) "-" else n,
        "<br>Distancia a la objetivo: ", if (is.na(dist)) "-" else paste0(round(dist, 0), " m"),
        "<br>Diferencia de altitud: ", if (is.na(desnivel)) "-" else paste0(round(desnivel, 0), " m"),
        "<br>Valores completados: ", if (is.na(n_completados)) 0 else n_completados
      )
      paste0(base, extra)
    }
  )

  # CORRECCIÓN (tras revisión de capturas): CartoDB.Positron está
  # pidiendo API key en las cuentas nuevas de sus proveedores de
  # teselas; se sustituye por OpenStreetMap estándar, que no requiere
  # ninguna clave y es el que trae Leaflet por defecto.
  leaflet(tabla_mapa) %>%
    addProviderTiles(providers$OpenStreetMap) %>%
    addCircleMarkers(
      lng = ~lon, lat = ~lat,
      radius = 8, color = "#333333", weight = 1,
      fillColor = ~color, fillOpacity = 0.9,
      # CORRECCIÓN: con noHide = TRUE el nombre de la estación queda
      # visible permanentemente junto al punto, en vez de solo al
      # pasar el ratón por encima. Con el número de estaciones
      # habitual en un análisis de completado (la objetivo y un
      # puñado de auxiliares cercanas) esto no satura el mapa y evita
      # tener que pasar el ratón una a una para identificarlas.
      label = ~Estacion,
      labelOptions = labelOptions(noHide = TRUE, direction = "auto", textOnly = FALSE),
      popup = ~popup
    ) %>%
    addLegend(
      position = "bottomright",
      colors = unname(colores_categoria),
      labels = names(colores_categoria),
      opacity = 0.9
    )
}

# ============================================================
# 12. PIPELINE COMPLETO: LECTURA -> ... -> EXPORTACIÓN
# ============================================================
#
# Orquesta todo el proceso para la PRIMERA estación (segunda
# columna del Excel) como objetivo, y escribe el Excel de salida.
#
# Preparado para ampliarse a varias estaciones objetivo: bastaría
# con iterar la llamada a completar_estacion_objetivo() para cada
# columna candidata y combinar los resultados antes de exportar.

completar_primera_estacion <- function(ruta_excel,
                                       hoja = 1,
                                       min_n = 100,
                                       min_r2 = 0.5,
                                       limitar_negativos = TRUE,
                                       ruta_gpkg = NULL,
                                       max_distancia_m = NULL,
                                       max_diferencia_altitud_m = NULL,
                                       permitir_reserva_fuera_filtro = FALSE,
                                       ruta_salida = "salida_completado.xlsx") {

  # --- Lectura ---
  df <- leer_datos_precipitacion(ruta_excel, hoja = hoja)

  # --- Validación ---
  validar_datos_precipitacion(df, col_fecha = 1)

  nombres <- colnames(df)
  col_fecha  <- nombres[1]
  col_target <- nombres[2]
  cols_aux   <- nombres[-c(1, 2)]

  # --- Tabla de estaciones (NUEVO — idea 2), opcional ---
  # Sin ruta_gpkg, tabla_estaciones queda NULL y el comportamiento es
  # idéntico al de antes de la idea 2 (sin filtro geográfico posible,
  # aunque se indiquen max_distancia_m/max_diferencia_altitud_m: sin
  # coordenadas no hay nada que filtrar).
  tabla_estaciones <- if (!is.null(ruta_gpkg)) cargar_tabla_estaciones(ruta_gpkg) else NULL

  # --- Preparación por pares -> Ajuste -> Evaluación -> Filtro geográfico -> Cascada/Selección -> Completado ---
  resultado <- completar_estacion_objetivo(df, col_fecha, col_target, cols_aux,
                                           min_n = min_n, min_r2 = min_r2,
                                           limitar_negativos = limitar_negativos,
                                           tabla_estaciones = tabla_estaciones,
                                           max_distancia_m = max_distancia_m,
                                           max_diferencia_altitud_m = max_diferencia_altitud_m,
                                           permitir_reserva_fuera_filtro = permitir_reserva_fuera_filtro)

  # --- Registro de completados ---
  # (ya incluido en resultado$registro_completados)

  # --- Tablas de sumatorios mensuales por año hidrológico (NUEVO) ---
  # Original: se calcula sobre la columna objetivo tal como se leyó,
  # ANTES de completar nada (df, no resultado$datos_completados).
  # Completada: se calcula sobre la serie ya completada.
  resumen_ah_original   <- generar_tabla_resumen_ah(df, col_fecha, col_target)
  resumen_ah_completado <- generar_tabla_resumen_ah(resultado$datos_completados, col_fecha, col_target)

  # La leyenda de rangos de AH puede diferir entre original y completada
  # solo si el completado alarga la serie (no debería, ya que se opera
  # sobre las mismas fechas); se exportan ambas por transparencia, con
  # nombres de hoja distintos.

  # --- Exportación ---
  # "Datos" es el data.frame original tal como se leyó y validó, ANTES
  # de completar nada. Se coloca como primera hoja para poder comparar
  # de un vistazo, dentro del mismo Excel, el dato original frente al
  # completado, sin tener que abrir el Excel de origen aparte.
  nombres_usados <- c("Datos", "Datos_completados", "Registro_completados", "Modelos_regresion",
                      "Resumen_mensual_AH_original", "Resumen_mensual_AH_completado",
                      "Leyenda_AH_original", "Leyenda_AH_completado")
  hojas_pares <- list()

  for (col_aux in cols_aux) {
    nombre_hoja <- generar_nombre_hoja_par("DEP", col_target, col_aux, nombres_usados)
    nombres_usados <- c(nombres_usados, nombre_hoja)
    hojas_pares[[nombre_hoja]] <- resultado$datos_pares[[col_aux]]
  }

  lista_hojas <- c(
    list(
      "Datos" = df,
      "Datos_completados" = resultado$datos_completados,
      "Registro_completados" = resultado$registro_completados,
      "Modelos_regresion" = resultado$tabla_modelos,
      "Resumen_mensual_AH_original" = resumen_ah_original$tabla,
      "Leyenda_AH_original" = resumen_ah_original$leyenda,
      "Resumen_mensual_AH_completado" = resumen_ah_completado$tabla,
      "Leyenda_AH_completado" = resumen_ah_completado$leyenda
    ),
    hojas_pares
  )

  write_xlsx(lista_hojas, path = ruta_salida)

  message("Archivo de salida escrito en: ", ruta_salida)
  message("Hojas generadas: ", paste(names(lista_hojas), collapse = ", "))

  invisible(list(
    datos_originales = df,
    datos_completados = resultado$datos_completados,
    registro_completados = resultado$registro_completados,
    tabla_modelos = resultado$tabla_modelos,
    datos_pares = resultado$datos_pares,
    resumen_ah_original = resumen_ah_original,
    resumen_ah_completado = resumen_ah_completado,
    # NUEVO (idea 4): se devuelve la tabla de estaciones ya cargada
    # (si se usó ruta_gpkg) para que la interfaz pueda dibujar el mapa
    # sin tener que releer el .gpkg de nuevo.
    tabla_estaciones = tabla_estaciones,
    col_target = col_target
  ))
}

# ------------------------------------------------------------
# 13. Ejemplo de uso
# ------------------------------------------------------------

# resultado <- completar_primera_estacion(
#   ruta_excel = "precipitacion.xlsx",
#   hoja = 1,
#   min_n = 100,
#   min_r2 = 0.5,
#   limitar_negativos = TRUE,
#   ruta_gpkg = "Tabla_Coordenadas.gpkg",         # opcional (idea 2)
#   max_distancia_m = 5000,                       # opcional (idea 2)
#   max_diferencia_altitud_m = 100,               # opcional (idea 2)
#   permitir_reserva_fuera_filtro = FALSE,        # opcional (idea 2, opción 3)
#   ruta_salida = "precipitacion_completada.xlsx"
# )
#
# Sin ruta_gpkg (o dejándola en NULL), el comportamiento es idéntico
# al de antes de la idea 2: selección de un único modelo por n y R².
# Con ruta_gpkg, por defecto se usa la cascada (opción 2): se prueban
# todas las auxiliares que pasan el filtro geográfico, de mayor a
# menor R², antes de dejar una fecha sin completar. Solo si
# permitir_reserva_fuera_filtro = TRUE se recurre, como último
# recurso y quedando señalado en el estado de cada valor, a
# auxiliares que el filtro geográfico había descartado.
#
# Para ver el gráfico interactivo tras ejecutar lo anterior:
# col_target <- names(resultado$datos_completados)[2]
# col_estado <- names(resultado$datos_completados)[3]
# generar_grafico_serie(resultado$datos_completados, "FECHA", col_target, col_estado)
