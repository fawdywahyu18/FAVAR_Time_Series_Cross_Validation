library(readxl)
library(boot)
library(tsDyn)
library(vars)
library(repr)
library(dplyr)
library(rlist)
library(zoo)
library(factoextra)
library(forecast)

main_wd = ''
backup_wd = ''
setwd(backup_wd)

nama_file = 'ihk_prepared.xlsx'
data_all = read_excel(nama_file, sheet = 'multivariate_monthly')
# data_sub = data_all[541:nrow(data_all),] # subsetting data from january 2005
data_sub = data_all[-1]

if (nama_file=='ihk_prepared.xlsx') {
  indeks_awal_notrans = 37
  indeks_akhir_notrans = 53
  indeks_slow_vars_manufacture = 5:7
  indeks_slow_vars_jub = 32:36
  indeks_slow_vars_upah = 54:62
} else if (nama_file=='ihk_prepared new.xlsx') {
  indeks_awal_notrans = 37+22
  indeks_akhir_notrans = 53+22
  indeks_slow_vars_manufacture = 5:7
  indeks_slow_vars_jub = 54:58
  indeks_slow_vars_upah = 76:84
} else {
  print('Nama File excel belum dikenal')
}

# interpolate data using spline
col_names = colnames(data_sub)
list_int = list()
list_no_int = list()
for (c in col_names) {
  is_na = is.na(data_sub[,c])
  check = sum(is_na)
  if (check==0) {
    list_no_int[[c]] = c
  }
  if (check==0) next
  list_int[[c]] = spline(data_sub[,c], method = 'fmm', n = nrow(data_sub))$y
}
data_int = list.cbind(list_int)
col_noint = unlist(list_no_int)
col_cpi = c("ID: Consumer Price Index", "Consumer Price Index: Core", "Consumer Price Index: Administered",
            "Consumer Price Index: Volatile")
data_int = cbind(data_int, data_sub[col_noint])
col_names_all = colnames(data_int)

percentage_index = function(input_data, lag_n) {
  persen = (input_data - lag(input_data, n =lag_n))/lag(input_data, n =lag_n)
  return(persen)
}
percentage_rate = function(input_data, lag_n) {
  persen = input_data - lag(input_data, n =lag_n)
  return(persen)
}

list_pct = list()
list_log = list()
no_transform = c(indeks_awal_notrans:indeks_akhir_notrans) # apabila mengugnakan ihk_prepared new = 37+22=59 : 53+22=75
for (c in col_names) {
  check = c %in% col_names[no_transform]
  if (check==TRUE){
    list_pct[[c]] = na.locf(percentage_rate(data_int[,c], 1), fromLast = TRUE)
    list_log[[c]] = data_int[,c]
  } else {
    list_pct[[c]] = na.locf(percentage_index(data_int[,c], 1), fromLast = TRUE)
    list_log[[c]] = na.locf(log(data_int[,c]), fromLast = TRUE)
  }
}
data_pct = list.cbind(list_pct)
data_log = list.cbind(list_log)
head(data_pct, 10)
head(data_log, 10)

# Create my own function to evaluate the model based on time series cross validation in R (Rob J. Hyndman)
# Function to forecast using the Time Invariant FAVAR
favar_forecast = function(data_input, nhor) {
  # data_input: dataframe, terdapat 2 pilihan: data_pct dan data_log
  # nhor: int, merupakan number of horizon untuk ke depan
  
  # Standardizing data = all variables with mean 0 and standard deviation 1. 
  # This step is crucial in PC analysis
  data_input_scaled = scale(data_input[,1:ncol(data_input)], center = TRUE, scale = TRUE)
  
  # mean and sd of inflation
  m_input = mean(data_input[,"ID: Consumer Price Index"])
  s_input = sd(data_input[,"ID: Consumer Price Index"])
  
  data_s = data_input_scaled
  s = s_input
  m = m_input
  # Step 1: Extract principal componentes of all X (including Y)
  pc_all = prcomp(data_s, center=FALSE, scale.=FALSE, rank. = 3) 
  C = pc_all$x # saving the principal components
  
  # Step 2: Extract principal componentes of Slow Variables
  # Slow Variables terdiri dari variabel manufaktur, jumlah uang beredar, dan upah
  # Kalau pakai ihk_prepared new : 
  slow_vars = c(indeks_slow_vars_manufacture, indeks_slow_vars_jub, indeks_slow_vars_upah)
  data_slow = data_s[, slow_vars]
  pc_slow = prcomp(data_slow, center=FALSE, scale.=FALSE, rank. = 3)
  F_slow = pc_slow$x
  
  # Step 3: Clean the PC from the effect of observed Y
  # Next clean the PC of all space from the observed Y
  reg = lm(C ~ F_slow + data_s[,"ID: Consumer Price Index"])
  summary(reg)
  F_hat = C - data.matrix(data_s[,"ID: Consumer Price Index"])%*%reg$coefficients[5,] # cleaning and saving F_hat
  
  # Step 4: Estimate FAVAR
  data_var = data.frame(F_hat, "ID: Consumer Price Index" = data_s[,"ID: Consumer Price Index"])
  nlag = VARselect(data_var, lag.max = 6, type = 'const')$selection[2] # lag based on HQ criteria
  favar = VAR(data_var, p = nlag)
  
  forecast_data = predict(favar, n.ahead=nhor, level=0.99)$fcst
  inf_fcst = forecast_data[[length(forecast_data)]]
  unscaled = inf_fcst * s + m
  tail_unscaled = exp(tail(unscaled, 1)[,1])
  
  return(tail_unscaled)
}


tsCV_fawdy = function(data_acuan, data_fcst, forecast_func=favar_forecast, nhor=1, nrep=30) {
  # data_acuan : dataframe sebagai acuan evaluasi (univariate only!!)
  # data_fcst : dataframe untuk melakukan forecast, bisa bentuknya matrix atau univariate
  # forecast_func : function untuk melakukan forecast, defaultnya adalah favar_forecast
  # nhor : int, merupakan forecast horizon, default = 1
  # nrep : int, merupakan jumlah repetisi untuk melakukan loop
  
  y = as.ts(data_acuan)
  index_adj = nhor - 1
  len_split = length(y) - index_adj
  
  list_error = list()
  for (i in 1:nrep) {
    # i = 1
    index_split = len_split - i
    
    split_y = data_fcst[1:index_split,]
    hasil_forecast = forecast_func(split_y, nhor=nhor)
    
    # define data acuan mana yang sesuai index
    index_train = index_split + nhor
    hasil_error = hasil_forecast - y[index_train]
    list_error[[i]] = hasil_error
  }
  vec_error = unlist(list_error)
  
  return(vec_error)
}

rmse_list = list()
mae_list = list()
nahead_vec = c(1, 3, 6, 9, 12)

for (n in nahead_vec) {
  n_ahead = n
  error_nahead = tsCV_fawdy(data_int[,'ID: Consumer Price Index'], data_log,
                            forecast_func = favar_forecast,
                            nhor = n_ahead, nrep = 30)
  rmse_list[[n]] = sqrt(mean(error_nahead^2)) # root mean squared error (RMSE)
  mae_list[[n]] = mean(abs(error_nahead)) # mean absolute error (MAE)
}
rmse_list_84var = rmse_list[c(1, 3, 6, 9, 12)]
mae_list_84var = mae_list[c(1, 3, 6, 9, 12)]

rmse_list_62var = rmse_list[c(1, 3, 6, 9, 12)]
mae_list_62var = mae_list[c(1, 3, 6, 9, 12)]


