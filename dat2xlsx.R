
#!/usr/bin/env Rscript
args = commandArgs(trailingOnly = TRUE)

library(openxlsx)

# ******************************************************************************
# Two arguments
# 1. input file path\name
# 2. output file path\name
# ******************************************************************************
if (length(args) < 2) {
  
  cat('\nPlease provide input file name followed by output file name.',
      '\n\nFor example:\n\n\t Rscript dat2xlsx XXX.DAT XXX.xslx\n\n')
  
} else {
  
  con = file(args[1], 'r')
  text = readLines(con, n = -1, warn = F)
  close(con)
  
  Encoding(text) = 'UTF-8'
  text = iconv(text, "UTF-8", "UTF-8", sub = '')
  
  df = do.call(rbind.data.frame, lapply(text, function(x) {
    list(textvar1 = substr(x, 1, 20),
         textvar2 = substr(x, 21, 40),
         datevar1 = substr(x, 41, 50),
         datevar2 = substr(x, 51, 60),
         numvar1 = substr(x, 61, 65),
         numvar2 = substr(x, 66, 70))
  }))
  
  # rename variables to their original names
  names(df) = c('First text variable name', 'Second text variable name', 'First date variable name',
                'Second date variable name', 'First numeric variable name', 'Second numeric variable name')
  
  # trim white space
  for (col in names(df)) {
    df[,col] = trimws(df[,col])
  }
  
  # convert numeric variables from text to numbers
  for (col in c('First numeric variable name', 'Second numeric variable name')) {
    df[,col] = as.numeric(df[,col])
  }
  
  # convert date/variables from text to date/time
  for (col in c('First date variable name', 'Second date variable name')) {
    df[,col] = as.Date(df[,col], '%m/%d/%Y')
  }
  df = df[order(df$`Delinquency Date`), ]
  
  cat('\nThe input file has ', nrow(df), 'rows and ', ncol(df), 'columns...',
      '\nNow writing these records into ', args[2], 'in Excel format...\n\n')
  
  # write to xlsx
  write.xlsx(df, args[2])
  
}
