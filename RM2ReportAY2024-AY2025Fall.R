rm(list=ls())

detach(ts)

pacman::p_load(dplyr, tidyr, stringr, tibble, readxl, writexl, psych)
#install.packages("openxlsx")
library(writexl)

library(readxl)
library(tidyverse)
library(car)

ts <- readxl::read_xlsx("ts_raw.xlsx")

#delete incompletes
ts$filter <- ifelse(is.na(ts$DG),0,1)
ts<- ts[ts$filter==1,]

attach(ts)

A1_1 <- rowMeans( tibble(A1.1_1, A2.1_1, A3.1_1) , na.rm = TRUE )
A1_2 <- rowMeans( tibble(A1.1_2, A2.1_2, A3.1_2) , na.rm = TRUE )
A1_3 <- rowMeans( tibble(A1.1_3, A2.1_3, A3.1_3) , na.rm = TRUE )
A1_4 <- rowMeans( tibble(A1.1_4, A2.1_4, A3.1_4) , na.rm = TRUE )
A1_5 <- rowMeans( tibble(A1.1_5, A2.1_5, A3.1_5) , na.rm = TRUE )
A2_1 <- rowMeans( tibble(A1.2_1, A2.2_1, A3.2_1) , na.rm = TRUE )
A2_2 <- rowMeans( tibble(A1.2_2, A2.2_2, A3.2_2) , na.rm = TRUE )
A2_3 <- rowMeans( tibble(A1.2_3, A2.2_3, A3.2_3) , na.rm = TRUE )
A2_4 <- rowMeans( tibble(A1.2_4, A2.2_4, A3.2_4) , na.rm = TRUE )
A2_5 <- rowMeans( tibble(A1.2_5, A2.2_5, A3.2_5) , na.rm = TRUE )
I1_1 <- rowMeans( tibble(I1.1_1, I2.1_1, I3.1_1) , na.rm = TRUE )
I1_2 <- rowMeans( tibble(I1.1_2, I2.1_2, I3.1_2) , na.rm = TRUE )
I1_3 <- rowMeans( tibble(I1.1_3, I2.1_3, I3.1_3) , na.rm = TRUE )
I1_4 <- rowMeans( tibble(I1.1_4, I2.1_4, I3.1_4) , na.rm = TRUE )
I1_5 <- rowMeans( tibble(I1.1_5, I2.1_5, I3.1_5) , na.rm = TRUE )
I2_1 <- rowMeans( tibble(I1.2_1, I2.2_1, I3.2_1) , na.rm = TRUE )
I2_2 <- rowMeans( tibble(I1.2_2, I2.2_2, I3.2_2) , na.rm = TRUE )
I2_3 <- rowMeans( tibble(I1.2_3, I2.2_3, I3.2_3) , na.rm = TRUE )
I2_4 <- rowMeans( tibble(I1.2_4, I2.2_4, I3.2_4) , na.rm = TRUE )
I2_5 <- rowMeans( tibble(I1.2_5, I2.2_5, I3.2_5) , na.rm = TRUE )
N1_1 <- rowMeans( tibble(N1_1, N2.1_1, N3.1_1) , na.rm = TRUE )
N1_2 <- rowMeans( tibble(N1_2, N2.1_2, N3.1_2) , na.rm = TRUE )
N1_3 <- rowMeans( tibble(N1_3, N2.1_3, N3.1_3) , na.rm = TRUE )
N1_4 <- rowMeans( tibble(N1_4, N2.1_4, N3.1_4) , na.rm = TRUE )
N1_5 <- rowMeans( tibble(N1_5, N2.1_5, N3.1_5) , na.rm = TRUE )
N2_1 <- rowMeans( tibble(N1.2_1, N2.2_1, N3.2_1) , na.rm = TRUE )
N2_2 <- rowMeans( tibble(N1.2_2, N2.2_2, N3.2_2) , na.rm = TRUE )
N2_3 <- rowMeans( tibble(N1.2_3, N2.2_3, N3.2_3) , na.rm = TRUE )
N2_4 <- rowMeans( tibble(N1.2_4, N2.2_4, N3.2_4) , na.rm = TRUE )
N2_5 <- rowMeans( tibble(N1.2_5, N2.2_5, N3.2_5) , na.rm = TRUE )


ts_clean <- tibble(ticket, Condition, DA, DG, A1_1, A1_2, A1_3, A1_4, A1_5, A2_1, A2_2,
                   A2_3, A2_4, A2_5, I1_1, I1_2, I1_3, I1_4, I1_5, I2_1, I2_2, I2_3, I2_4,
                   I2_5, N1_1, N1_2, N1_3, N1_4, N1_5, N2_1, N2_2, N2_3, N2_4, N2_5,
                   Appropriateness_1, Appropriateness_2, Appropriateness_3)


ts_clean$Y_A_Overall <- rowMeans(tibble(A1_1, A1_2, A1_3, A1_4, A1_5, A2_1, A2_2, A2_3, A2_4, A2_5), na.rm=TRUE)
ts_clean$Y_I_Overall <- rowMeans(tibble(I1_1, I1_2, I1_3, I1_4, I1_5, I2_1, I2_2, I2_3, I2_4, I2_5), na.rm=TRUE)
ts_clean$Y_N_Overall <- rowMeans(tibble(N1_1, N1_2, N1_3, N1_4, N1_5, N2_1, N2_2, N2_3, N2_4, N2_5), na.rm=TRUE)

ts_clean$Y_A_Furious <- rowMeans(tibble(A1_1, A2_1), na.rm=TRUE)
ts_clean$Y_A_Angry <- rowMeans(tibble(A1_2, A2_2), na.rm=TRUE)
ts_clean$Y_A_Upset <- rowMeans(tibble(A1_3, A2_3), na.rm=TRUE)
ts_clean$Y_A_Annoyed <- rowMeans(tibble(A1_4, A2_4), na.rm=TRUE)
ts_clean$Y_A_Frustrated <- rowMeans(tibble(A1_5, A2_5), na.rm=TRUE)


ts_clean$Y_I_Furious <- rowMeans(tibble(I1_1, I2_1), na.rm=TRUE)
ts_clean$Y_I_Angry <- rowMeans(tibble(I1_2, I2_2), na.rm=TRUE)
ts_clean$Y_I_Upset <- rowMeans(tibble(I1_3, I2_3), na.rm=TRUE)
ts_clean$Y_I_Annoyed <- rowMeans(tibble(I1_4, I2_4), na.rm=TRUE)
ts_clean$Y_I_Frustrated <- rowMeans(tibble(I1_5, I2_5), na.rm=TRUE)

ts_clean$Y_N_Furious <- rowMeans(tibble(N1_1, N2_1), na.rm=TRUE)
ts_clean$Y_N_Angry <- rowMeans(tibble(N1_2, N2_2), na.rm=TRUE)
ts_clean$Y_N_Upset <- rowMeans(tibble(N1_3, N2_3), na.rm=TRUE)
ts_clean$Y_N_Annoyed <- rowMeans(tibble(N1_4, N2_4), na.rm=TRUE)
ts_clean$Y_N_Frustrated <- rowMeans(tibble(N1_5, N2_5), na.rm=TRUE)

ts_clean$A1M <- rowMeans(tibble(A1_1, A1_2, A1_3, A1_4, A1_5), na.rm=TRUE)
ts_clean$A2M <- rowMeans(tibble(A2_1, A2_2, A2_3, A2_4, A2_5), na.rm=TRUE)

ts_clean$I1M <- rowMeans(tibble(I1_1, I1_2, I1_3, I1_4, I1_5), na.rm=TRUE)
ts_clean$I2M <- rowMeans(tibble(I2_1, I2_2, I2_3, I2_4, I2_5), na.rm=TRUE)

ts_clean$N1M <- rowMeans(tibble(N1_1, N1_2, N1_3, N1_4, N1_5), na.rm=TRUE)
ts_clean$N2M <- rowMeans(tibble(N2_1, N2_2, N2_3, N2_4, N2_5), na.rm=TRUE)


ts_clean$missing <- rowSums(is.na(ts_clean))
ts_clean <- ts_clean[ts_clean$missing<5,]

library(writexl)

write_xlsx(ts_clean, path = "ts_cleanFINAL.xlsx")


data <- read_excel("ts_cleanFINAL.xlsx")
cor_A <- cor.test(data$Appropriateness_A, data$Y_A_Overall, method = "pearson")
cor_I <- cor.test(data$Appropriateness_I, data$Y_I_Overall, method = "pearson")
cor_N <- cor.test(data$Appropriateness_N, data$Y_N_Overall, method = "pearson")
cor_A
cor_I
cor_N


