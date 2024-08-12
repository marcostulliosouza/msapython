args <- commandArgs(trailingOnly = TRUE)

# Verificar se todos os argumentos foram fornecidos
if (length(args) < 5) {
    stop("Número insuficiente de argumentos. Necessário: arquivo, coluna_dispositivo, coluna_part, coluna_appr, colunas_para_analisar")
}

# Receber os argumentos passados do Python
input_file <- args[1]
device_col <- args[2]
part_col <- args[3]
appr_col <- args[4]
columns_to_analyze <- args[5:length(args)]

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

# Carregar os dados do arquivo Excel
data_table <- read.xlsx(input_file, sheet = 1)

# input_file <- "/Users/marcostullio/Documents/Projeto MSA/Dados/MSA 912701.xlsx"
# device_col <- "Chamber"
# part_col <- "Serial Number"
# appr_col <- "Facility"

# Carregar os dados do arquivo Excel
data_table <- read.xlsx(input_file, sheet = 1)

# Limpar os nomes das colunas
data_table <- janitor::clean_names(data_table)

# # Imprimir os nomes das colunas carregadas
# print("Nomes das colunas no dataframe após limpeza:")
# print(colnames(data_table))

# # Limpar os nomes das colunas: remover espaços em branco e caracteres especiais
# clean_names <- function(names) {
#     names <- gsub("\\s+", " ", names) # Substituir múltiplos espaços por um único espaço
#     names <- trimws(names) # Remover espaços em branco no início e no final
#     names <- gsub("[^[:alnum:]_ ]", "", names) # Remover caracteres especiais
#     names <- gsub(" ", "_", names) # Substituir espaços por sublinhados para evitar problemas de nomeação
#     return(names)
# }

# colnames(data_table) <- clean_names(colnames(data_table))

# # Imprimir os nomes das colunas carregadas
# print("Nomes das colunas no dataframe após limpeza:")
# print(colnames(data_table))

# Verificar se as colunas fornecidas existem após a limpeza dos nomes
if (!(device_col %in% colnames(data_table))) {
    stop(paste("Coluna do dispositivo", device_col, "não encontrada."))
}
if (!(part_col %in% colnames(data_table))) {
    stop(paste("Coluna do part (peça)", part_col, "não encontrada."))
}
if (!(appr_col %in% colnames(data_table))) {
    stop(paste("Coluna do appr (operador)", appr_col, "não encontrada."))
}

# Filtrar os dados para as colunas relevantes
filtered_data <- data_table %>%
    rename(device = !!sym(device_col), part = !!sym(part_col), appr = !!sym(appr_col))

# Criar e salvar o arquivo Excel
wb <- createWorkbook()

# Obter a lista de dispositivos únicos
devices <- unique(filtered_data$device)

for (device in devices) {
    # Filtrar os dados para o dispositivo atual
    device_data <- filtered_data %>%
        filter(device == !!device)

    # Adicionar uma aba para o dispositivo
    addWorksheet(wb, paste("Resultados", device))

    for (col in columns_to_analyze) {
        gage_data <- device_data %>%
            select(part, appr, !!sym(col)) %>%
            rename(var = !!sym(col)) %>%
            mutate(var = as.numeric(var))

        # Verificar equilíbrio do design e balanceamento dos dados
        balance_check <- gage_data %>%
            group_by(part, appr) %>%
            summarise(count = n(), .groups = "drop")

        min_count <- balance_check %>%
            summarise(min_count = min(count)) %>%
            pull(min_count)

        balanced_data <- gage_data %>%
            group_by(part, appr) %>%
            sample_n(min_count, replace = TRUE) %>%
            ungroup() %>%
            mutate(part = as.factor(part), appr = as.factor(appr))

        # Realizar a análise Gage R&R
        result <- tryCatch(
            {
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
                    print_plot = FALSE,
                    signifstars = FALSE
                )
            },
            error = function(e) {
                message(paste("Erro ao realizar análise Gage R&R para a coluna", col, ":", e$message))
                NULL
            }
        )

        if (!is.null(result)) {
            # Adicionar resultados da análise Gage R&R à aba do dispositivo
            writeData(wb, sheet = paste("Resultados", device), x = result$anova, startCol = 1, startRow = 1, colNames = TRUE)
            writeData(wb, sheet = paste("Resultados", device), x = result$variance, startCol = 1, startRow = nrow(result$anova) + 3, colNames = TRUE)

            # Gerar o gráfico de Gage R&R
            gage_plot <- result$plot
            plot_file <- tempfile(fileext = ".png")
            png(plot_file)
            print(gage_plot)
            dev.off()

            # Inserir o gráfico na aba do dispositivo
            insertImage(wb, sheet = paste("Resultados", device), file = plot_file, width = 6, height = 4, startRow = nrow(result$anova) + 10, startCol = 1)
        }
    }
}

# Salvar o arquivo Excel
saveWorkbook(wb, "resultado_gage_rr.xlsx", overwrite = TRUE)
