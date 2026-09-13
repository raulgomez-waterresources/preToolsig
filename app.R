library(shiny)
library(plotly)
library(leaflet)
source("preToolsig.R")

# ============================================================
# UI — lo que ve el usuario
# ============================================================
ui <- fluidPage(

  # Tipografía manuscrita para la firma personal (carga desde Google Fonts)
  tags$head(
    tags$link(
      rel = "stylesheet",
      href = "https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&display=swap"
    )
  ),

  fluidRow(
    column(12, align = "center",
      tags$img(src = "logo.png", height = "100px"),
      titlePanel("preToolsig"),
      p("Herramienta para el completado de series de precipitación mediante regresión lineal entre estaciones meteorológicas, con visualización, sumatorios por año hidrológico y mapa de estaciones.")
    )
  ),

  hr(),

  fluidRow(
    column(12,
      h4("Instrucciones"),
      p("El Excel debe tener la siguiente estructura, con encabezados en la primera fila:"),
      tags$ul(
        tags$li("Columna 1: Fecha"),
        tags$li("Columna 2: Estación a completar"),
        tags$li("Columnas 3 en adelante: Estaciones de apoyo")
      ),
      p("Los datos de precipitación deben ser numéricos. Los datos ausentes pueden dejarse como celdas vacías."),
      tableOutput("tabla_ejemplo")
    )
  ),

  # Aviso de privacidad y responsabilidad, en tono discreto
  fluidRow(
    column(12,
      tags$div(
        style = "font-size:13px; color:#6c757d; font-style:italic; border-left:3px solid #dee2e6; padding:8px 12px; margin-bottom:15px; background-color:#f8f9fa;",
        "Los archivos que subas se procesan únicamente para generar el resultado y no se almacenan de forma permanente. ",
        "preToolsig es una herramienta de apoyo técnico y no sustituye la validación profesional de los datos ni de los resultados obtenidos."
      )
    )
  ),

  hr(),

  sidebarLayout(
    sidebarPanel(
      h4("1. Cargar archivo"),
      fileInput("archivo", "Selecciona tu Excel (.xlsx)", accept = ".xlsx"),

      h4("2. Parámetros"),
      numericInput("min_n", "Mínimo de datos coincidentes", value = 100, min = 1),
      numericInput("min_r2", "R² mínimo", value = 0.5, min = 0, max = 1, step = 0.05),
      checkboxInput("limitar_negativos", "Limitar estimaciones negativas a 0", value = TRUE),

      h4("3. Filtro geográfico (opcional)"),
      p(style = "font-size:13px; color:#6c757d;",
        "Si subes la tabla maestra de estaciones, se prueban todas las auxiliares que cumplan la distancia y diferencia de altitud indicadas, de mayor a menor R², antes de dejar una fecha sin completar. Sin tabla de estaciones, este apartado no tiene efecto."),
      fileInput("archivo_gpkg", "Tabla maestra de estaciones (.gpkg)", accept = ".gpkg"),
      numericInput("max_distancia_m", "Distancia máxima (m) — vacío = sin límite", value = NA),
      numericInput("max_diferencia_altitud_m", "Diferencia máxima de altitud (m) — vacío = sin límite", value = NA),
      checkboxInput("permitir_reserva", "Si aun así quedan huecos, permitir completar con estaciones fuera del filtro geográfico (quedan señaladas como tal)", value = FALSE),

      actionButton("procesar", "Procesar archivo", class = "btn-primary")
    ),

    mainPanel(
      h4("4. Resultado"),
      textOutput("estado_texto"),
      uiOutput("avisos_ui"),
      uiOutput("resumen_ui"),
      uiOutput("cascada_ui"),
      plotlyOutput("grafico_serie", height = "500px"),
      h4("5. Mapa de estaciones"),
      uiOutput("mapa_ui"),
      uiOutput("descarga_ui")
    )
  ),

  hr(),

  # Firma personal / sello de marca
  fluidRow(
    column(12, align = "center",
      tags$div(
        style = "font-family:'Dancing Script', cursive; font-size:34px; color:#1f6f8b; margin-top:10px;",
        "Raúl Gómez C."
      ),
      tags$div(
        style = "font-family: Arial, sans-serif; font-size:14px; letter-spacing:1px; color:#5a8a99; margin-top:-6px; margin-bottom:20px;",
        "HIDROLOGÍA Y GESTIÓN DE RECURSOS HÍDRICOS"
      )
    )
  )
)

# ============================================================
# SERVER — la lógica
# ============================================================
server <- function(input, output, session) {

  resultado       <- reactiveVal(NULL)          # lo que devuelve completar_primera_estacion()
  ruta_salida_tmp <- reactiveVal(NULL)          # ruta temporal del Excel ya generado
  estado          <- reactiveVal(NULL)          # NULL / "ok" / texto de error
  avisos          <- reactiveVal(character(0))  # warnings capturados del motor

  output$tabla_ejemplo <- renderTable({
    data.frame(
      Fecha = c("01/01/2020", "02/01/2020", "03/01/2020"),
      ESTACION_OBJETIVO = c(12.4, NA, 4.2),
      ESTACION_APOYO_1 = c(13.1, 8.4, 3.9),
      ESTACION_APOYO_2 = c(11.8, 9.1, 4.8)
    )
  })

  # AJUSTE 1 (Apartado 4): si el usuario selecciona un archivo NUEVO sin
  # haber pulsado todavía "Procesar", se limpia el resultado anterior.
  # Así se evita que quede visible (y descargable) el resultado de un
  # archivo distinto al que hay seleccionado en ese momento.
  observeEvent(input$archivo, {
    resultado(NULL)
    ruta_salida_tmp(NULL)
    estado(NULL)
    avisos(character(0))
  })

  observeEvent(input$procesar, {

    req(input$archivo)

    ext <- tolower(tools::file_ext(input$archivo$name))
    if (ext != "xlsx") {
      estado(paste0("El archivo debe ser .xlsx (se ha subido un archivo ." , ext, ")."))
      resultado(NULL)
      avisos(character(0))
      return()
    }

    ruta_gpkg_sel <- NULL
    if (!is.null(input$archivo_gpkg)) {
      ext_gpkg <- tolower(tools::file_ext(input$archivo_gpkg$name))
      if (ext_gpkg != "gpkg") {
        estado(paste0("La tabla de estaciones debe ser .gpkg (se ha subido un archivo ." , ext_gpkg, ")."))
        resultado(NULL)
        avisos(character(0))
        return()
      }
      ruta_gpkg_sel <- input$archivo_gpkg$datapath
    }

    # Un numericInput vacío devuelve NA; se traduce a NULL (= sin
    # límite) para completar_primera_estacion().
    max_dist_sel <- if (is.na(input$max_distancia_m)) NULL else input$max_distancia_m
    max_alt_sel  <- if (is.na(input$max_diferencia_altitud_m)) NULL else input$max_diferencia_altitud_m

    salida_tmp <- tempfile(fileext = ".xlsx")
    avisos_capturados <- character(0)

    withProgress(message = "Procesando archivo...", value = 0.6, {

      tryCatch({

        res <- withCallingHandlers({
          completar_primera_estacion(
            ruta_excel                     = input$archivo$datapath,
            hoja                           = 1,
            min_n                          = input$min_n,
            min_r2                         = input$min_r2,
            limitar_negativos              = input$limitar_negativos,
            ruta_gpkg                      = ruta_gpkg_sel,
            max_distancia_m                = max_dist_sel,
            max_diferencia_altitud_m       = max_alt_sel,
            permitir_reserva_fuera_filtro  = input$permitir_reserva,
            ruta_salida                    = salida_tmp
          )
        }, warning = function(w) {
          avisos_capturados <<- c(avisos_capturados, conditionMessage(w))
          invokeRestart("muffleWarning")
        })

        resultado(res)
        ruta_salida_tmp(salida_tmp)
        estado("ok")
        avisos(avisos_capturados)

      }, error = function(e) {
        estado(paste0("No se ha podido procesar el archivo.\n\nMotivo: ", conditionMessage(e)))
        resultado(NULL)
        avisos(character(0))
      })
    })
  })

  # AJUSTE 2 (Apartado 4): mensaje inicial mientras no se ha procesado
  # nada todavía, en vez de dejar el panel vacío sin ninguna explicación.
  output$estado_texto <- renderText({
    if (is.null(estado())) {
      "Sube un archivo y pulsa \"Procesar archivo\" para ver aquí el resultado."
    } else if (identical(estado(), "ok")) {
      "Procesamiento completado correctamente."
    } else {
      estado()
    }
  })

  output$avisos_ui <- renderUI({
    if (length(avisos()) == 0) return(NULL)
    tags$div(
      style = "color:#8a6d3b; background:#fcf8e3; padding:8px; border-radius:4px; margin-top:8px;",
      tags$b("Avisos del proceso:"),
      tags$ul(lapply(avisos(), tags$li))
    )
  })

  output$resumen_ui <- renderUI({
    req(resultado())
    r <- resultado()

    datos_completados <- r$datos_completados
    col_target <- names(datos_completados)[2]
    col_estado <- names(datos_completados)[3]

    modelo_sel <- r$tabla_modelos[r$tabla_modelos$seleccionado, ]

    n_total         <- nrow(datos_completados)
    n_completados   <- nrow(r$registro_completados)
    n_sin_completar <- sum(datos_completados[[col_estado]] == "SIN COMPLETAR")
    pct_completado  <- round((n_total - n_sin_completar) / n_total * 100, 1)

    bloque_pct <- tags$p(
      tags$b("Serie completada: "),
      paste0(pct_completado, " % (", n_sin_completar, " de ", n_total, " días sin completar)")
    )

    if (nrow(modelo_sel) == 0) {
      return(tags$div(
        tags$p(tags$b("Estación objetivo: "), col_target),
        tags$p("Ningún modelo cumplió los criterios mínimos: no se ha completado ningún valor."),
        bloque_pct
      ))
    }

    # NUEVO (idea 2): si se usó tabla de estaciones, se muestra la
    # distancia y diferencia de altitud de la auxiliar seleccionada,
    # cuando estén disponibles (no serán NA salvo que faltaran
    # coordenadas para alguna de las dos estaciones).
    bloque_geo <- NULL
    if (!is.na(modelo_sel$distancia_m)) {
      bloque_geo <- tags$p(
        tags$b("Distancia / desnivel a la objetivo: "),
        paste0(round(modelo_sel$distancia_m, 0), " m / ",
               round(modelo_sel$diferencia_altitud_m, 0), " m")
      )
    }

    # NUEVO (idea 2, cascada): con tabla de estaciones, "modelo_sel" es
    # la auxiliar de mayor prioridad (primera intentada), pero puede no
    # ser la única que haya completado valores — el desglose completo
    # se muestra aparte, en cascada_ui.
    etiqueta_auxiliar <- if (is.na(modelo_sel$orden_prioridad)) {
      "Estación auxiliar seleccionada: "
    } else {
      "Estación auxiliar de mayor prioridad: "
    }

    tags$div(
      tags$p(tags$b("Archivo procesado: "), input$archivo$name),
      tags$p(tags$b("Estación objetivo: "), col_target),
      tags$p(tags$b(etiqueta_auxiliar), modelo_sel$estacion_auxiliar),
      tags$p(tags$b("R²: "), round(modelo_sel$r2, 3)),
      tags$p(tags$b("Datos utilizados: "), modelo_sel$n),
      bloque_geo,
      tags$p(tags$b("Valores completados: "), n_completados),
      tags$p(tags$b("Valores sin completar: "), n_sin_completar),
      bloque_pct
    )
  })

  # NUEVO (idea 5): gráfico interactivo de la serie completada,
  # coloreado por "<objetivo>_estado". Usa generar_grafico_serie(),
  # definida en preToolsig.R, para no duplicar lógica de trazado.
  # NUEVO (idea 2, cascada): cuando se ha usado tabla de estaciones,
  # muestra qué auxiliares han completado cuántos valores cada una,
  # no solo la "seleccionada" principal — para que quede visible el
  # reparto real, ya que en cascada pueden intervenir varias.
  output$cascada_ui <- renderUI({
    req(resultado())
    r <- resultado()

    desglose <- r$tabla_modelos[!is.na(r$tabla_modelos$orden_prioridad) &
                                 r$tabla_modelos$n_valores_completados > 0, ]
    if (nrow(desglose) == 0) return(NULL)

    desglose <- desglose[order(desglose$orden_prioridad), ]

    filas <- lapply(seq_len(nrow(desglose)), function(i) {
      fila <- desglose[i, ]
      etiqueta <- if (isTRUE(fila$cumple_filtro_geografico)) "" else " (fuera de filtro geográfico)"
      tags$li(
        tags$b(fila$estacion_auxiliar), etiqueta, ": ",
        fila$n_valores_completados, " valores completados (R² = ", round(fila$r2, 3), ")"
      )
    })

    tags$div(
      tags$p(tags$b("Desglose de la cascada (auxiliares que han completado algún valor):")),
      tags$ul(filas)
    )
  })

  output$grafico_serie <- renderPlotly({
    req(resultado())
    r <- resultado()

    datos_completados <- r$datos_completados
    col_fecha  <- names(datos_completados)[1]
    col_target <- names(datos_completados)[2]
    col_estado <- names(datos_completados)[3]

    generar_grafico_serie(datos_completados, col_fecha, col_target, col_estado)
  })

  # NUEVO (idea 4): mapa de estaciones. Solo tiene sentido si se usó
  # una tabla maestra de estaciones (ruta_gpkg); sin ella, se muestra
  # un texto explicativo en su lugar en vez de un mapa vacío.
  output$mapa_ui <- renderUI({
    req(resultado())
    r <- resultado()
    if (is.null(r$tabla_estaciones)) {
      return(tags$p(style = "color:#6c757d; font-size:14px;",
                    "Sube una tabla maestra de estaciones (.gpkg) en el apartado 3 para ver aquí el mapa."))
    }
    leafletOutput("mapa_estaciones", height = "420px")
  })

  output$mapa_estaciones <- renderLeaflet({
    req(resultado())
    r <- resultado()
    req(r$tabla_estaciones)
    generar_mapa_estaciones(r$tabla_estaciones, r$col_target, r$tabla_modelos)
  })

  output$descarga_ui <- renderUI({
    req(resultado())
    downloadButton("descargar", "Descargar Excel generado", class = "btn-success")
  })

  output$descargar <- downloadHandler(
    filename = function() {
      paste0("preToolsig_completado_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      req(ruta_salida_tmp())
      file.copy(ruta_salida_tmp(), file)
    }
  )
}

shinyApp(ui = ui, server = server)
