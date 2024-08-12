args <- commandArgs(trailingOnly = TRUE)

# Verificar se os argumentos essenciais foram fornecidos
if (length(args) < 6) {
    stop("Número insuficiente de argumentos. Necessário:
    arquivo, coluna_dispositivo, coluna_part,
    coluna_appr, modo_one_way, colunas_para_analisar, [lsls], [usls]")
}

# Receber os argumentos passados do Python
input_file <- args[1]
device_col <- args[2]
part_col <- args[3]
appr_col <- args[4]
is_one_way <- args[7]
columns_to_analyze <- args[8:length(args)]

# Número de colunas a serem analisadas
num_cols <- length(columns_to_analyze)

# Inicializar LSL e USL como NA por padrão
lsls <- rep(NA, num_cols)
usls <- rep(NA, num_cols)

if (is_one_way == "True") {
    is_one_way <- TRUE
} else {
    is_one_way <- FALSE
}

# Verificar se LSL e USL foram fornecidos
if (length(args) > 6 + num_cols) {
    # Se fornecidos, atribuir os valores adequados
    lsls <- as.numeric(args[(7 + num_cols):(7 + 2 * num_cols - 1)])
    usls <- as.numeric(args[(7 + 2 * num_cols):length(args)])
}

# Limpar os nomes das colunas fornecidas como argumentos
device_col <- janitor::make_clean_names(device_col)
part_col <- janitor::make_clean_names(part_col)
appr_col <- janitor::make_clean_names(appr_col)
columns_to_analyze <- sapply(columns_to_analyze, janitor::make_clean_names)

# Carregar pacotes necessários
if (!require("dplyr")) install.packages("dplyr", dependencies = TRUE)
if (!require("SixSigma")) install.packages("SixSigma", dependencies = TRUE)
if (!require("openxlsx")) install.packages("openxlsx", dependencies = TRUE)
if (!require("ggplot2")) install.packages("ggplot2", dependencies = TRUE)
if (!require("janitor")) install.packages("janitor", dependencies = TRUE)
library(janitor)
library(dplyr)
library(SixSigma)
library(openxlsx)
library(ggplot2)

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

# Verificar se as colunas para análise existem
missing_cols <- columns_to_analyze[!(columns_to_analyze %in% colnames(data_table))]
if (length(missing_cols) > 0) {
    stop(paste("As seguintes colunas para análise não foram encontradas:", paste(missing_cols, collapse = ", ")))
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

    # Inicializa a linha de início
    current_row <- 1

    # Adicionar uma aba para o dispositivo
    sheet_name <- paste("Resultados -", device)
    addWorksheet(wb, sheet_name)

    for (i in seq_along(columns_to_analyze)) {
        col <- columns_to_analyze[i]
        lsl <- lsls[i]
        usl <- usls[i]

        # Selecionar as colunas apropriadas dependendo do valor de is_one_way
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
                        lsl = lsl,
                        usl = usl,
                        sigma = 6,
                        data = balanced_data,
                        main = paste("Six Sigma Gage R&R Study -", col),
                        sub = "Análise Gage R&R",
                        alphaLim = 0.05,
                        digits = 4,
                        method = "oneway",
                        print_plot = FALSE, # Desativar a geração de gráficos
                        signifstars = FALSE
                    )
                } else {
                    ss.rr(
                        var = "var",
                        part = "part",
                        appr = "appr",
                        lsl = lsl,
                        usl = usl,
                        sigma = 6,
                        data = balanced_data,
                        main = paste("Six Sigma Gage R&R Study -", col),
                        sub = "Análise Gage R&R",
                        alphaLim = 0.05,
                        errorTerm = "interaction",
                        digits = 4,
                        method = "crossed",
                        print_plot = FALSE, # Desativar a geração de gráficos
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
            print(paste("Salvando resultados para a coluna:", col))
            writeData(wb, sheet = sheet_name, x = result$anova, startCol = 1, startRow = current_row, colNames = TRUE)
            current_row <- current_row + nrow(result$anova) + 2 # Atualiza a linha de início para evitar sobrescrita
            writeData(wb, sheet = sheet_name, x = result$variance, startCol = 1, startRow = current_row, colNames = TRUE)
            current_row <- current_row + nrow(result$variance) + 4 # Atualiza a linha de início para a próxima tabela
        } else {
            message(paste("Resultado da análise Gage R&R para a coluna", col, "não disponível."))
        }
    }
}

# Salvar o arquivo Excel
saveWorkbook(wb, output_file, overwrite = TRUE)
