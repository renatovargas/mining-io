# Insumo-Producto Ambientalmente Extendido
# Cuenta Satélite de Bioeconomía
# Matriz de Insumo Producto y 
# Cuentas Ambientales de Costa Rica
# Consultor CEPAL: Renato Vargas renovargas@gmail.com

# Preámbulo
library(openxlsx)
library(reshape2)
library(Matrix.utils)
library(ggplot2)
rm(  list = ls()  )

getwd()

# Consumo intermedio
Z_cruda <- as.matrix(read.xlsx("datos/MIP-AE-AE-017-CR.xlsx", 
                               sheet = "MIP 2017", 
                               rows= c(12:264), 
                               cols = c(4:256), 
                               skipEmptyRows = FALSE, 
                               colNames = FALSE, 
                               rowNames = FALSE)
                     )  # <-- Fin del paréntesis

nombres <- read.xlsx("datos/MIP-AE-AE-017-CR.xlsx", 
                               sheet = "MIP 2017", 
                               rows= c(12:264), 
                               cols = 1, 
                               skipEmptyRows = FALSE, 
                               colNames = FALSE, 
                               rowNames = FALSE
)  # <-- Fin del paréntesis


# Agregamos nuestra matriz para tener un valor por actividad

# Nombramos nuestra matriz Z
colnames(Z_cruda) <- as.vector(nombres$X1)
rownames(Z_cruda) <- as.vector(nombres$X1)

Z <- aggregate.Matrix(Z_cruda, as.factor(nombres$X1),fun = "sum")
Z <- t(aggregate.Matrix(t(Z), as.factor(nombres$X1),fun = "sum"))
Z <- as.matrix(Z)

dim(Z)


# Demanda Final
DF_cruda <- as.matrix(read.xlsx("datos/MIP-AE-AE-017-CR.xlsx", 
                               sheet = "MIP 2017", 
                               rows= c(12:264), 
                               cols = c(258:262), 
                               skipEmptyRows = FALSE, 
                               colNames = FALSE, 
                               rowNames = FALSE)
                     )  # <-- Fin del paréntesis


# Lo mismo para la demanda final
codsDF <- as.matrix(
  t(
    read.xlsx(
      "datos/MIP-AE-AE-017-CR.xlsx",
      sheet = "MIP 2017", 
      rows= c(9:10), 
      cols = c(258:262), 
      skipEmptyRows = FALSE,
      colNames = FALSE, 
      rowNames = FALSE
  ))
)# <-- Fin del paréntesis

# Y nuestra matriz de demanda final DF
colnames(DF_cruda) <- as.vector(codsDF[,1])
rownames(DF_cruda) <- as.vector(rownames(Z_cruda))

# Y agregamos las filas que se repiten
DF <- aggregate.Matrix(DF_cruda, as.factor(rownames(DF_cruda)),fun = "sum")
DF <- as.matrix(DF)

dim(DF)



# =============================================================================
# Cuenta de energía

# Importamos los datos crudos
E_cruda <- as.matrix(read.xlsx("datos/COUF-2017.xlsx", 
                               sheet = "COUF-E 2017", 
                               rows= c(127:146), 
                               cols = c(3:169), 
                               skipEmptyRows = FALSE, 
                               colNames = FALSE, 
                               rowNames = TRUE # Sí hay nombres de fila
                               )
)  # <-- Fin del paréntesis

# Extraemos los nombres de columna
nombres_e <- t(read.xlsx("datos/COUF-2017.xlsx", 
                               sheet = "COUF-E 2017", 
                               rows= c(16), 
                               cols = c(4:169), 
                               skipEmptyRows = FALSE, 
                               colNames = FALSE, 
                               rowNames = FALSE)
                     )  # <-- Fin del paréntesis

# Y nombramos las columnas de nuestra matriz de usos energéticos
colnames(E_cruda) <- c(nombres_e[,1])

# Las dimensiones de E_cruda son mayores a las de Z
# porque hay agregaciones por grupos de sectores y
# hay sectores desagregados a mayor detalle.

dim(E_cruda)

# Identificamos las posiciones que son sumas de sectores
# Nótese que dejamos dos sumas dentro que no tienen detalle
# correspondientes a AE082 (Electricidad) y AE144 (hogares como empl.)

posGruposEnergia <- c(1,31,35,78,82,94,97,106,110,113,
                      119,122,133,143,147,150,153,158)

# Extraemos solo los sectores (nótese el "-" antes de posGruposEnergia)
E_cruda <- E_cruda[ , -posGruposEnergia]

# Utilizando la función substr() extraemos los primeros 5 digitos de
# la nomenclatura para poder agregar por actividades que comparten
# esos mismos.
colnames(E_cruda) <- substr(colnames(E_cruda), start = 1, stop = 5)

# Y agregamos utilizando el mismo procedimiento que anteriormente.
E <- as.matrix(t(aggregate.Matrix(t(E_cruda), colnames(E_cruda),fun = "sum")))
E <- as.matrix(E)
# Y chequeamos que nuestras dimensiones sean iguales a las columnas
# de Z
dim(E)

# Para ser congruentes con Z, renombramos las columnas con los nombres
# completos de Z
colnames(E) <- colnames(Z)

# Hacemos limpieza
rm(Z_cruda,DF_cruda,E_cruda, nombres_e)

# =============================================================================
## Cuenta de energía 2018 (Preliminar No-Citar)

# Importamos los datos crudos
E_cruda <- as.matrix(read.xlsx("datos/COUF-2018.xlsx", 
                               sheet = "COUF-E 2018", 
                               rows= c(126:145), 
                               cols = c(3:170), 
                               skipEmptyRows = FALSE, 
                               colNames = FALSE, 
                               rowNames = TRUE # Sí hay nombres de fila
                               )
)  # <-- Fin del paréntesis

# Extraemos los nombres de columna
nombres_e <- t(read.xlsx("datos/COUF-2018.xlsx", 
                         sheet = "COUF-E 2018", 
                         rows= c(15), 
                         cols = c(4:170), 
                         skipEmptyRows = FALSE, 
                         colNames = FALSE, 
                         rowNames = FALSE)
)  # <-- Fin del paréntesis

# Y nombramos las columnas de nuestra matriz de usos energéticos
colnames(E_cruda) <- c(nombres_e[,1])

# Las dimensiones de E_cruda son mayores a las de Z
# porque hay agregaciones por grupos de sectores y
# hay sectores desagregados a mayor detalle.

dim(E_cruda)

# Identificamos las posiciones que son sumas de sectores
# Nótese que dejamos dos sumas dentro que no tienen detalle
# correspondientes a AE082 (Electricidad) y AE144 (hogares como empl.)

# Notar que respecto de 2017 lo siguiente cambia a partir de 106 (107)

posGruposEnergia <-c(1,31,35,78,82,94,97,107,111,114,
                      120,123,134,144,148,151,154,159)

# Extraemos solo los sectores (nótese el "-" antes de posGruposEnergia)
E_cruda <- E_cruda[ , -posGruposEnergia]

# Utilizando la función substr() extraemos los primeros 5 digitos de
# la nomenclatura para poder agregar por actividades que comparten
# esos mismos.
colnames(E_cruda) <- substr(colnames(E_cruda), start = 1, stop = 5)

# Y agregamos utilizando el mismo procedimiento que anteriormente.
E18 <- as.matrix(t(aggregate.Matrix(t(E_cruda), colnames(E_cruda),fun = "sum")))
E18 <- as.matrix(E18)
# Y chequeamos que nuestras dimensiones sean iguales a las columnas
# de Z
dim(E18)

# Para ser congruentes con Z, renombramos las columnas con los nombres
# completos de Z
colnames(E18) <- colnames(Z)

# Hacemos limpieza
rm(E_cruda)

moltenE18 <- as.matrix(melt(E18))

# Exportamos a Excel
write.xlsx( as.data.frame(moltenE18) , 
            "datos/CuentaEnergiaBD_2018.xlsx",
            sheetName= "datos",
            startRow = 5,
            startCol = 1,
            asTable = FALSE, 
            colNames = TRUE, 
            rowNames = TRUE, 
            overwrite = TRUE
)

# write.table(as.data.frame(moltenE18),"clipboard", sep = "\t", row.names = TRUE)

# =============================================================================
# Modelo de insumo producto

# Producción
x <- as.vector(rowSums(Z) + rowSums(DF))

# Demanda final
f <- as.vector(rowSums(DF))

# x sombrero
xhat <- diag(x)
xhat_inv <- solve(xhat)

# Matriz de coeficientes técnicos
A <- Z %*% solve( xhat )

# Matriz identidad
I <- diag( dim(A)[1])

# Matriz de Leontief
L <- solve(I - A )

# Coeficientes de uso de cada energético por unidad de producto
EC <- E %*% solve(xhat)
colnames(EC) <- colnames(Z)

# Nueva demanda final
f1 <- f
f1[80] <- f1[80] *1.20

# Cálculo de nuevas demandas de energía por energético
EC %*% L %*% f1

# Diferencias
deltaE <- cbind( rowSums(E), 
       (EC %*% L %*% f1), 
       (EC %*% L %*% f1)- rowSums(E), 
       ((EC %*% L %*% f1)- rowSums(E))*100/ rowSums(E) 
       ) # <-- fin del paréntesis
colnames(deltaE) <- c("Original", "Política", "Diferencia", "Porcentual")

# Y si queremos el detalle
E1 <- EC %*% diag(as.vector(L %*% f1))
colnames(E1) <- colnames(E)

# =============================================================================
# Bioeconomía

# Clasificación cruzada bioeconomía AECR
nombres_bio <- read.xlsx("datos/clasificacion_cruzada_bioeconomia.xlsx", 
                                            sheet = "datos", 
                                            rows= c(1:137), 
                                            cols = c(1:6), 
                                            skipEmptyRows = FALSE, 
                                            colNames =TRUE, 
                                            rowNames = FALSE
)  # <-- Fin del paréntesis

# Primero solamente emitir los multiplicadores de los sectores
rownames(L) <- colnames(L)
bio_mult <- as.data.frame(colSums(L))

# Lo copiamos para pegarlo
write.table(as.data.frame(colSums(L)),"clipboard", sep = "\t", row.names = TRUE)


# Luego agregamos nuestras matrices para considerar a todas las actividades
# características como un solo sector

Z_bio <- Z

# Renombramos con la clasificación cruzada
colnames(Z_bio) <- nombres_bio$BECR
rownames(Z_bio) <- nombres_bio$BECR

# Y Agregamos
Z_bio <- aggregate.Matrix(Z_bio, as.factor(nombres_bio$BECR),fun = "sum")
Z_bio <- t(aggregate.Matrix(t(Z_bio), as.factor(nombres_bio$BECR),fun = "sum"))
Z_bio <- as.matrix(Z_bio)

# Repetimos  para la demanda final

f_bio <- as.matrix(f)
rownames(f_bio) <- nombres_bio$BECR

f_bio <- aggregate.Matrix(f_bio, as.factor(nombres_bio$BECR),fun = "sum")
f_bio <- as.matrix(f_bio)
f_bio <- as.vector(f_bio)

# Modelo insumo producto

# Modelo de insumo producto

# Producción
x_bio <- as.vector(rowSums(Z_bio) + f_bio)

# Demanda final
f_bio <- f_bio

# x sombrero
xhat_bio <- diag(x_bio)
xhat_inv_bio <- solve(xhat_bio)

# Matriz de coeficientes técnicos
A_bio <- Z_bio %*% solve( xhat_bio )

# Matriz identidad
I_bio <- diag( dim(A_bio)[1])

# Matriz de Leontief
L_bio <- solve(I_bio - A_bio )

bio_mult2 <- as.data.frame(colSums(L_bio))

write.table(bio_mult2,"clipboard", sep = "\t", row.names = TRUE)

#


# Empleo BIO
# ==========

# Consumo intermedio
LAB_cruda <- as.matrix(read.xlsx("datos/MIP-AE-AE-017-CR.xlsx", 
                               sheet = "MIP 2017", 
                               rows= c(296:301), 
                               cols = c(4:256), 
                               skipEmptyRows = FALSE, 
                               colNames = FALSE, 
                               rowNames = FALSE)
)  # <-- Fin del paréntesis

nombres_l <- read.xlsx("datos/MIP-AE-AE-017-CR.xlsx", 
                     sheet = "MIP 2017", 
                     rows= c(296:301), 
                     cols = 1, 
                     skipEmptyRows = FALSE, 
                     colNames = FALSE, 
                     rowNames = FALSE
)  # <-- Fin del paréntesis


# Agregamos nuestra matriz para tener un valor por actividad

# Nombramos nuestra matriz Z
colnames(LAB_cruda) <- as.vector(nombres$X1)
rownames(LAB_cruda) <- as.vector(nombres_l$X1)

LAB <- aggregate.Matrix(LAB_cruda, as.factor(nombres_l$X1),fun = "sum")
LAB <- t(aggregate.Matrix(t(LAB), as.factor(nombres$X1),fun = "sum"))
LAB <- as.matrix(LAB)
rm(LAB_cruda)

dim(LAB)

# Multiplicadores de empleo
# Todos

c_total     <- colSums(LAB) %*% xhat_inv
c_total_hat <- diag(as.vector(c_total))
H_total <- c_total_hat %*% L
multip_total <- as.matrix(colSums(H_total))

write.table(multip_total,"clipboard", sep = "\t", row.names = TRUE)

# Asalariados

c_asal     <- LAB[1,] %*% xhat_inv
c_asal_hat <- diag(as.vector(c_asal))
H_asal <- c_asal_hat %*% L
multip_asal <- as.matrix(colSums(H_asal))

# Cuenta propia

c_cp     <- LAB[2,] %*% xhat_inv
c_cp_hat <- diag(as.vector(c_cp))
H_cp <- c_cp_hat %*% L
multip_cp <- as.matrix(colSums(H_cp))

# Empresarios

c_empr     <- LAB[3,] %*% xhat_inv
c_empr_hat <- diag(as.vector(c_empr))
H_empr <- c_empr_hat %*% L
multip_empr <- as.matrix(colSums(H_empr))

# Trabajadores familiares no remunerados

c_tfnr     <- LAB[4,] %*% xhat_inv
c_tfnr_hat <- diag(as.vector(c_tfnr))
H_tfnr <- c_tfnr_hat %*% L
multip_tfnr <- as.matrix(colSums(H_tfnr))

# Otros trabajadores no remunerados

c_otnr     <- LAB[5,] %*% xhat_inv
c_otnr_hat <- diag(as.vector(c_otnr))
H_otnr <- c_otnr_hat %*% L
multip_otnr <- as.matrix(colSums(H_otnr))

# Personal de otros establecimientos

c_otestab     <- LAB[6,] %*% xhat_inv
c_otestab_hat <- diag(as.vector(c_otestab))
H_otestab <- c_otestab_hat %*% L
multip_otestab <- as.matrix(colSums(H_otestab))

# Multiplicadores BIO
LAB_bio <- LAB
colnames(LAB_bio) <- nombres_bio$BECR

# Y Agregamos
LAB_bio <- aggregate.Matrix(LAB_bio, as.factor(rownames(LAB_bio)),fun = "sum")
LAB_bio <- t(aggregate.Matrix(t(LAB_bio), as.factor(nombres_bio$BECR),fun = "sum"))
LAB_bio <- as.matrix(LAB_bio)



# Multiplicadores de empleo BIO
# Todos

c_total_bio     <- colSums(LAB_bio) %*% xhat_inv_bio
c_total_hat_bio <- diag(as.vector(c_total_bio))
H_total_bio <- c_total_hat_bio %*% L_bio
multip_total_bio <- as.matrix(colSums(H_total_bio))

# Asalariados

c_asal_bio     <- LAB_bio[1,] %*% xhat_inv_bio
c_asal_hat_bio <- diag(as.vector(c_asal_bio))
H_asal_bio <- c_asal_hat_bio %*% L_bio
multip_asal_bio <- as.matrix(colSums(H_asal_bio))

# Cuenta propia

c_cp_bio     <- LAB_bio[2,] %*% xhat_inv_bio
c_cp_hat_bio <- diag(as.vector(c_cp_bio))
H_cp_bio <- c_cp_hat_bio %*% L_bio
multip_cp_bio <- as.matrix(colSums(H_cp_bio))

# Empresarios

c_empr_bio     <- LAB_bio[3,] %*% xhat_inv_bio
c_empr_hat_bio <- diag(as.vector(c_empr_bio))
H_empr_bio <- c_empr_hat_bio %*% L_bio
multip_empr_bio <- as.matrix(colSums(H_empr_bio))

# Trabajadores familiares no remunerados

c_tfnr_bio     <- LAB_bio[4,] %*% xhat_inv_bio
c_tfnr_hat_bio <- diag(as.vector(c_tfnr_bio))
H_tfnr_bio <- c_tfnr_hat_bio %*% L_bio
multip_tfnr_bio <- as.matrix(colSums(H_tfnr_bio))

# Otros trabajadores no remunerados

c_otnr_bio     <- LAB_bio[5,] %*% xhat_inv_bio
c_otnr_hat_bio <- diag(as.vector(c_otnr_bio))
H_otnr_bio <- c_otnr_hat_bio %*% L_bio
multip_otnr_bio <- as.matrix(colSums(H_otnr_bio))

# Personal de otros establecimientos

c_otestab_bio     <- LAB_bio[6,] %*% xhat_inv_bio
c_otestab_hat_bio <- diag(as.vector(c_otestab_bio))
H_otestab_bio <- c_otestab_hat_bio %*% L_bio
multip_otestab_bio <- as.matrix(colSums(H_otestab_bio))


# =============================================================================
# Excel

write.xlsx(as.data.frame(colnames(Z)),
                         "salidas/nombres_Z.xlsx",
                         sheetName = "datos",
                         startRow = 5,
                         startCol = 1,
                         asTable = FALSE,
                         colNames = TRUE,
                         rowNames = TRUE,
                         overwrite = TRUE)

# Cambios en datos ambientales
write.xlsx( as.data.frame(deltaE) , 
            "salidas/datos_ambientales.xlsx",
            sheetName= "datos",
            startRow = 5,
            startCol = 1,
            asTable = FALSE, 
            colNames = TRUE, 
            rowNames = TRUE, 
            overwrite = TRUE
            )

# Gráfico
heatmap(LAB, Colv = NA, Rowv = NA)


