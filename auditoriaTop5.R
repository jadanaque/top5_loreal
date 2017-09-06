# This script gives the necessary info to answer TOP 5 Questionnaire (Audit Process).
# It creates the function that will do the work, then we just need to call the function, without any argument.

# Packages ----
library(readxl)
library(reshape2)
library(dplyr)
library(tidyr)
library(lubridate)
library(ggplot2)
library(openxlsx)
# library(RColorBrewer)
# library(ggmap)
# library(leaflet)
#library(plotly)
#library(googleVis)
library(magrittr)

# Function ----
Top5Answers <- function(){
	
    # CUSTOMERS	----
    # Reading the last Master Data in the specified folder
    files <- list.files("./data/clientes/", full.names = T)
    
    filePath <- files  %>%
        file.mtime %>%
        which.max %>%
        files[.]
    
    #data_types <- rep("text", 64)  # Removed: no longer needed
    
    clientesDF <- read_excel(filePath, sheet=1, col_types = "text")
    
    names(clientesDF) <- make.names(names(clientesDF), unique = T)
    names(clientesDF) <- iconv(names(clientesDF), to='ASCII//TRANSLIT')
    names(clientesDF) <- gsub("[^[:alnum:]]", "", names(clientesDF))
    
    ##clientesDF$Creadoel <- as.Date(as.numeric(clientesDF$Creadoel), origin = as.Date("1899-12-30"), tz = "GMT")  # Not necessary for Top5
    clientesDF <- mutate(clientesDF, Se = as.integer(Se), BqPed = as.integer(BqPed), BloqPed = as.integer(BloqPed))
    
    # Eliminating duplicates
    clientesDF <- select(clientesDF, -Noctaant, -CP, -Viaspago, -PrioE, -ImpRmto, -OrdRmto, -RecFinan,
                         -CobFlet, -NIF3, -Grupodeestadisticascliente, -Codfiscsupl, -Cl, -GP, -CoordX,
                         -CoordY, -Comentariosreferentesaladi, -Contrato, -Di, -Clasededistribucionparael,
                         -BqEnt, -BloqueodeEntrega, -Mon, -Mon1, -Mon2, -Mss, -Prt, -EP, -EP1)
    clientesDF <- distinct(clientesDF, Cliente, OrgVt, CDis, Se, .keep_all = TRUE)
    
    # Creating a new variable with the names of "Grupo de Cuentas"
    groupNames <- data.frame(Grupo = c("P002", "P003", "PP01", "PP02", "PP03", "PP04", "PP05", "PP12", "Z001"),
                             NombreGrupoCuentas = c("Empleado", "Free Delivery", "Solicitante", "Ship-to", "Pagador", "Receptor de Factura", "Boca de Ventas", "Nodo de Clientes", "Clientes IntraGrupo"),
                             stringsAsFactors = F)
    clientesDF <- left_join(clientesDF, groupNames, by = "Grupo")
		
		# Summarizing data
		uniqueClients <- dcast(clientesDF, Grupo ~ OrgVt, n_distinct, value.var = "Cliente", margins = T)  # Total Unique Codes. Combination: Grupo-OrgVt
		
		# Blocked customers: It considers global inactivation, i.e., the customer must be inactive (blocked)
		# in every "sector" to be considered BLOCKED. If it is active only in one sector (say Kerastase), it is
		# considered ACTIVE
		activos_bloqueadosDF <- clientesDF %>%
		    group_by(Cliente, CDis) %>%
		    mutate(sec_len_G = n(),
		           BqPed_len_G = sum(BqPed %in% c(10, 12)),
		           BloqPed_len_G = sum(BloqPed %in% c(10, 12))) %>%
		    mutate(status_G = ifelse(BqPed_len_G > 0 | BloqPed_len_G >= sec_len_G,
		                             "BLOQUEADO", "ACTIVO")) %>%
		    ungroup()
		
		# Dataframe with unique codes (customers) and status (active-inactive)
		unique_activos_bloqueadosDF <- activos_bloqueadosDF %>% 
		                            select(-(1:3)) %>%
		                            distinct(Cliente, .keep_all = TRUE)
		
		# Summary of ACTIVE-BLOCKED
		resumen_activos_bloqueados <- activos_bloqueadosDF %>%
		                group_by(status_G) %>%
		                summarise(cuenta = n_distinct(Cliente)) %>%  # códigos únicos
		                mutate(porcentaje = paste0(format((cuenta / sum(cuenta)) * 100, digits = 2, nsmall = 2), "%")) %>%
		                add_row(status_G = "TOTAL", cuenta = sum(.$cuenta), porcentaje = "100%")
		
		# ACTIVE-BLOCKED. This time, by sector
		status_x_sector <- clientesDF %>%
		    select(Cliente, OrgVt, CDis, Se, Grupo, BqPed, BloqPed) %>%
		    filter(CDis == "02", Grupo %in% c("PP01", "P003", "PP03")) %>%  # Sólo Solicitantes, Free Deliverys y Pagadores del canal 02
		    distinct(Cliente, OrgVt, CDis, Se, .keep_all = TRUE) %>%
		    mutate(status_se = ifelse(BqPed %in% c(10, 12) | BloqPed %in% c(10, 12), "B", "A")) %>%  # B: Bloqueado, A: Activo. Para el sector
		    select(Cliente, Se, status_se) %>%
		    spread(Se, status_se, fill = "N")  # N: No creado

		# Saving the summary tables to files
		top5excelPath <- paste0("./reportes/top5answers-", gsub("-", "", today()), ".xlsx")
		
		activos_bloqueadosDF <- select(activos_bloqueadosDF, 1:37, 41)  # Removes columns created for computation only
		
		wb <- createWorkbook()
		addWorksheet(wb, "CodigosUnicos")
		addWorksheet(wb, "status_clientesDB")
		addWorksheet(wb, "status_clientesDB-unicos")
		addWorksheet(wb, "status_clientes-resumen")
		addWorksheet(wb, "statusXsector-ActivBloqNocreado")
		writeData(wb, "CodigosUnicos", uniqueClients)
		writeData(wb, "status_clientesDB", activos_bloqueadosDF)
		writeData(wb, "status_clientesDB-unicos", unique_activos_bloqueadosDF)
		writeData(wb, "status_clientes-resumen", resumen_activos_bloqueados)
		writeData(wb, "statusXsector-ActivBloqNocreado", status_x_sector)
		
		###############################################################
		
		# PROVEEDORES
		# Loading data
		files <- list.files(paste0("./data/proveedores/", year(today())), full.names = T, recursive = T)
		
		filePath <- files  %>%
		    file.mtime %>%
		    which.max %>%
		    files[.]
		
		data_types <- rep("text", ncol(read_excel(filePath)))
		
		proveedoresDF <- read_excel(filePath, sheet=1, col_types = data_types)
		
		# Data Cleaning
		names(proveedoresDF) <- make.names(names(proveedoresDF), unique = T)
		names(proveedoresDF) <- iconv(names(proveedoresDF), to='ASCII//TRANSLIT')
		names(proveedoresDF) <- gsub("[^[:alnum:]]", "", names(proveedoresDF))
		
		# Eliminating duplicates
		proveedoresDF <- select(proveedoresDF, -CPapdo, -Concbusq, -Ramo, -RecAltPago,
		                        -Codfiscsupl, -Telefono1, -GrRef, -Tpimpto)
		proveedoresDF <- distinct(proveedoresDF, Acreedor, .keep_all = TRUE)
		
		# Activos/Bloqueados
		proveedoresDF <- mutate(proveedoresDF, status = ifelse(is.na(B) & is.na(B1) & is.na(B2), "ACTIVO", "BLOQUEADO"))
		
		resumen_proveedores <- proveedoresDF %>%
		                            group_by(status) %>%
		                            summarize(cuenta = n_distinct(Acreedor)) %>%
		                            mutate(porcentaje = paste0(format((cuenta / sum(cuenta)) * 100,
		                                                              digits = 2, nsmall = 2),
		                                                       "%")) %>%
		                            add_row(status = "TOTAL", cuenta = sum(.$cuenta), porcentaje = "100%")
		
		# Saving the summary tables to a file
		addWorksheet(wb, "statusProveedoresDB")
		addWorksheet(wb, "status_proveedores-resumen")
		writeData(wb, "statusProveedoresDB", proveedoresDF)
		writeData(wb, "status_proveedores-resumen", resumen_proveedores)
		
		saveWorkbook(wb, top5excelPath, overwrite = T)
		


}