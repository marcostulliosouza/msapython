# Carregar pacotes necessários
library(dplyr)
library(SixSigma)
library(openxlsx)
library(ggplot2)
library(janitor)

# Definir argumentos de teste
input_file <- "/Users/marcostullio/Documents/Projeto MSA/Dados/MSA 912701.xlsx"
device_col <- "Basement ID"
part_col <- "Serial Number"
appr_col <- "Facility"
is_one_way <- FALSE
columns_to_analyze <- c("Current in Idle Mode (mA)") # Exemplo de coluna para análise

# Limpar os nomes das colunas fornecidas como argumentos
device_col <- janitor::make_clean_names(device_col)
part_col <- janitor::make_clean_names(part_col)
appr_col <- janitor::make_clean_names(appr_col)
columns_to_analyze <- sapply(columns_to_analyze, janitor::make_clean_names)

# Definir a pasta de resultados e criar se não existir
output_dir <- "Resultados"
if (!dir.exists(output_dir)) {
  dir.create(output_dir)
}

# Gerar o nome do arquivo de saída com base no nome do arquivo de entrada
file_name <- tools::file_path_sans_ext(basename(input_file))
output_file <- file.path(output_dir, paste0(file_name, "_resultado_gage_rr.xlsx"))

# Carregar os dados do arquivo Excel
data_table <- read.xlsx(input_file, sheet = 1)

# Limpar os nomes das colunas no data_table
data_table <- janitor::clean_names(data_table)

# Verificar se as colunas fornecidas existem após a limpeza dos nomes
if (!(device_col %in% colnames(data_table))) {
  stop(paste("Coluna do dispositivo", device_col, "não encontrada."))
}
if (!(part_col %in% colnames(data_table))) {
  stop(paste("Coluna do part (peça)", part_col, "não encontrada."))
}
if (!is_one_way && !(appr_col %in% colnames(data_table))) {
  stop(paste("Coluna do appr (operador)", appr_col, "não encontrada."))
}

# Filtrar os dados para as colunas relevantes
filtered_data <- data_table %>%
  rename(device = !!sym(device_col), part = !!sym(part_col))

if (!is_one_way) {
  filtered_data <- filtered_data %>%
    rename(appr = !!sym(appr_col))
}

# Criar e salvar o arquivo Excel
wb <- createWorkbook()

# Obter a lista de dispositivos únicos
devices <- unique(filtered_data$device)

for (device in devices) {
  # Filtrar os dados para o dispositivo atual
  device_data <- filtered_data %>%
    filter(device == !!device)

  # Adicionar uma aba para o dispositivo
  sheet_name <- paste("Resultados -", device)
  addWorksheet(wb, sheet_name)

  for (col in columns_to_analyze) {
    gage_data <- device_data %>%
      select(part, !!sym(col)) %>%
      rename(var = !!sym(col)) %>%
      mutate(var = as.numeric(var))

    # Verificar equilíbrio do design e balanceamento dos dados
    balance_check <- gage_data %>%
      group_by(part) %>%
      summarise(count = n(), .groups = "drop")

    min_count <- balance_check %>%
      summarise(min_count = min(count)) %>%
      pull(min_count)

    balanced_data <- gage_data %>%
      group_by(part) %>%
      sample_n(min_count, replace = TRUE) %>%
      ungroup() %>%
      mutate(part = as.factor(part))

    # Realizar a análise Gage R&R
    result <- tryCatch(
      {
        if (is_one_way) {
          ss.rr(
            var = "var",
            part = "part",
            lsl = NA,
            usl = NA,
            sigma = 6,
            data = balanced_data,
            main = paste("Six Sigma Gage R&R Study -", col),
            sub = "Análise Gage R&R",
            alphaLim = 0.05,
            digits = 4,
            method = "oneway",
            print_plot = TRUE, # Habilitado para gerar gráficos
            signifstars = FALSE
          )
        } else {
          ss.rr(
            var = "var",
            part = "part",
            appr = "appr",
            lsl = NA,
            usl = NA,
            sigma = 6,
            data = balanced_data,
            main = paste("Six Sigma Gage R&R Study -", col),
            sub = "Análise Gage R&R",
            alphaLim = 0.05,
            errorTerm = "interaction",
            digits = 4,
            method = "crossed",
            print_plot = TRUE, # Habilitado para gerar gráficos
            signifstars = FALSE
          )
        }
      },
      error = function(e) {
        message(paste("Erro ao realizar análise Gage R&R para a coluna", col, ":", e$message))
        NULL
      }
    )

    if (!is.null(result)) {
      # Adicionar resultados da análise Gage R&R à aba do dispositivo
      writeData(wb, sheet = sheet_name, x = result$anova, startCol = 1, startRow = 1, colNames = TRUE)
      writeData(wb, sheet = sheet_name, x = result$variance, startCol = 1, startRow = nrow(result$anova) + 3, colNames = TRUE)

      # Salvar gráficos gerados em PNG
      for (plot_name in names(result$plots)) {
        gage_plot <- result$plots[[plot_name]]
        plot_file <- file.path(output_dir, paste0(plot_name, "_", device, "_", col, ".png"))
        ggsave(plot_file, plot = gage_plot, width = 8, height = 6)

        # Inserir o gráfico na aba do dispositivo
        insertImage(wb, sheet = sheet_name, file = plot_file, width = 6, height = 4, startRow = nrow(result$anova) + nrow(result$variance) + 6, startCol = 1)
      }
    } else {
      message(paste("Resultado da análise Gage R&R para a coluna", col, "não disponível."))
    }
    if (is_one_way) {
      gage_data <- device_data %>%
        select(part, !!sym(col)) %>%
        rename(var = !!sym(col)) %>%
        mutate(var = as.numeric(var))
    } else {
      gage_data <- device_data %>%
        select(part, appr, !!sym(col)) %>%
        rename(var = !!sym(col)) %>%
        mutate(var = as.numeric(var))
    }

    # Verificar equilíbrio do design e balanceamento dos dados
    if (is_one_way) {
      balance_check <- gage_data %>%
        group_by(part) %>%
        summarise(count = n(), .groups = "drop")
    } else {
      balance_check <- gage_data %>%
        group_by(part, appr) %>%
        summarise(count = n(), .groups = "drop")
    }

    min_count <- balance_check %>%
      summarise(min_count = min(count)) %>%
      pull(min_count)

    if (is_one_way) {
      balanced_data <- gage_data %>%
        group_by(part) %>%
        sample_n(min_count, replace = TRUE) %>%
        ungroup() %>%
        mutate(part = as.factor(part))
    } else {
      balanced_data <- gage_data %>%
        group_by(part, appr) %>%
        sample_n(min_count, replace = TRUE) %>%
        ungroup() %>%
        mutate(part = as.factor(part), appr = as.factor(appr))
    }

    # Realizar a análise Gage R&R
    result <- tryCatch(
      {
        if (is_one_way) {
          ss.rr(
            var = "var",
            part = "part",
            appr = NA,
            lsl = NA,
            usl = NA,
            sigma = 6,
            data = balanced_data,
            main = paste("Six Sigma Gage R&R Study -", col),
            sub = "Análise Gage R&R",
            alphaLim = 0.05,
            digits = 4,
            method = "oneway",
            print_plot = TRUE, # Habilitado para gerar gráficos
            signifstars = FALSE
          )
        } else {
          ss.rr(
            var = "var",
            part = "part",
            appr = "appr",
            lsl = NA,
            usl = NA,
            sigma = 6,
            data = balanced_data,
            main = paste("Six Sigma Gage R&R Study -", col),
            sub = "Análise Gage R&R",
            alphaLim = 0.05,
            errorTerm = "interaction",
            digits = 4,
            method = "crossed",
            print_plot = TRUE, # Habilitado para gerar gráficos
            signifstars = FALSE
          )
        }
      },
      error = function(e) {
        message(paste("Erro ao realizar análise Gage R&R para a coluna", col, ":", e$message))
        NULL
      }
    )

    if (!is.null(result)) {
      # Adicionar resultados da análise Gage R&R à aba do dispositivo
      writeData(wb, sheet = sheet_name, x = result$anova, startCol = 1, startRow = 1, colNames = TRUE)
      writeData(wb, sheet = sheet_name, x = result$variance, startCol = 1, startRow = nrow(result$anova) + 3, colNames = TRUE)

      # Salvar gráficos gerados em PNG
      for (plot_name in names(result$plots)) {
        gage_plot <- result$plots[[plot_name]]
        plot_file <- file.path(output_dir, paste0(plot_name, "_", device, "_", col, ".png"))
        ggsave(plot_file, plot = gage_plot, width = 8, height = 6)

        # Inserir o gráfico na aba do dispositivo
        insertImage(wb, sheet = sheet_name, file = plot_file, width = 6, height = 4, startRow = nrow(result$anova) + nrow(result$variance) + 6, startCol = 1)
      }
    } else {
      message(paste("Resultado da análise Gage R&R para a coluna", col, "não disponível."))
    }
  }
}

# Salvar o arquivo Excel
saveWorkbook(wb, output_file, overwrite = TRUE)
