library(docxtractr)
library(dplyr)
library(purrr)
library(tidyr)
library(stringr)

rdx <- qdapTools::read_docx

check_empty_rows <- function(rw){
  row_2_vector <- unlist(rw) #Convert a row data frame to vector
  empty_cols <- sum(row_2_vector == "", na.rm =T) #Check empty value
  na_cols <- sum(is.na(row_2_vector)) #Check NA
  total_invalid_columns <- empty_cols + na_cols #Calculate total columns having invalid values
  # print(total_invalid_columns)
  ncols = length(rw) #Numbers of columns

  if (total_invalid_columns == ncols)
    .rs <- TRUE
  else if (ncols == 4){
    if ((row_2_vector[1]=="") & (row_2_vector[4])=="" | is.na(row_2_vector[4])) {
      .rs <- TRUE
      # print("it is")
    }
    else {
      .rs <- FALSE
    }
  }
  else {
    .rs <- FALSE
  }
  return(.rs)
}

pre_remove_emtpy_rows <- function(df){
  .df <- df[complete.cases(df[,2]),]
  empty_xetnghiem <- which(.df %>% .[,2]=="")
  if (length(empty_xetnghiem) > 0){
    .df <- .df[-empty_xetnghiem,]
  }
  return(.df)
}

remove_empty_rows <- function(df){
  emp_rows <- apply(df, 1, check_empty_rows)
  .new.df <- df[!emp_rows,]
  .new.df <- .new.df[!.new.df[,1]=="NƯỚC TIỂU - PHÂN - DỊCH SINH HỌC",]
  return(.new.df)
}


read_and_check <- function(file_path){
  docx <- read_docx(file_path)
  tbls <- docx_tbl_count(docx)
  file_name <- gsub("data//|.docx","", file_path)
  return(list(file_name = file_name, table_count = tbls))
}


to_extract <- function(tbl){
  rs <- FALSE
  if (ncol(tbl) >=3){
    if (any(grepl("V1", names(tbl))))
      #Todo: Check the logic below. So funny :-)
      if (any(grepl("kết quả | KQ | quả", tolower(tbl[2,]))))
        rs <- TRUE
      else
        rs <- TRUE
  }
  else {
    if (any(grepl("kết quả | KQ | quả", tolower(tbl[1,]))))
      rs <- TRUE
  }
  return(rs)
}

extract_table <- function(docx){
  tbls <- docx_extract_all_tbls(docx) #Initially get all tables from a give docx
  # Start filtering
  new.tables <- list()
  j <- 0
  for(i in seq_len(length(tbls)))
    if(to_extract(tbls[[i]])){
      names(tbls[[i]]) <- tbls[[i]][1,]
      tbls[[i]] <- tbls[[i]][-1,]
      .m <- names(tbls[[i]])
      if (sum(!.m=="") != 1){ #extract if column name not empty > 1
        j <- j+ 1
        new.tables[[j]] <- tbls[[i]]
        }
    }
  return(new.tables)
}

transform_table <- function(tbl){
  if(ncol(tbl)==8){ #Process table 1
    .tbl1 <- tbl[,c(1:4)]
    .tbl2 <- tbl[,c(5:8)]
    names(.tbl1) <- c(paste0("V",c(1:4)))
    names(.tbl2) <- c(paste0("V",c(1:4)))
    new.table <- dplyr::bind_rows(.tbl1,.tbl2) %>% pre_remove_emtpy_rows(.)
  }
  else { #Process other tables
    columns_to_select <- grep("xn|xét nghiệm|quả|kết quả", tolower(names(tbl)))
    if(length(columns_to_select)==2){
      new.table <- tbl[,columns_to_select, drop = F]
    }
    else{
      new.table <- tbl[,c(1,2)]
    }
  }
  return(new.table)
}

fill_na_value_table <- function(tbl){
  ncols <- ncol(tbl)
  if(ncols==4){
    tbl <- apply(tbl, 1, function(x){
      k <- x
      if (all(x[1:3] != "", na.rm = T) & (is.na(x[4]) || x[4]=="")){
        x[4] <- x[3]
        k <- x
      }
      return(k)
    })
    .tbl <- as.data.frame(t(tbl))
  }
  else {
    .tbl <- tbl
  }
  return(.tbl)
}

process_each_table <- function(tbl){
  # stopifnot(exists(tbl))
  .tbl <- tbl %>%
    transform_table(.) %>%
    fill_na_value_table(.) %>%
    remove_empty_rows(.)
  if (ncol(.tbl) == 4){
    .tbl <- .tbl[,c(2,4)]
  }
  names(.tbl) <- c("XN","KQ")

  check_rubbish <- grepl(":|kết quả",tolower(.tbl$KQ))
  if(sum(check_rubbish) > 2){
    .tbl <- .tbl[!check_rubbish,]
  }
  # .tbl$KQ <- ifelse(grepl("\\/L", .tbl$KQ),"", .tbl$KQ)
  return(.tbl)
}

long_to_wide <- function(tbl){
  if (nrow(tbl) > 0){
    if(grepl("000|in trang này", tolower(tbl$XN))){
      .tbl <- NULL
    } else {
      if(sum(duplicated(tbl[,1])) == 0){
        .tbl <- tbl %>% spread(XN,KQ)
      }
      .tbl <- tbl[!duplicated(tbl),] %>% spread(XN, KQ)
    }
  }
  else {
    .tbl <- NULL
  }
  # if (any(grepl("bác sĩ chỉ định",tolower(names(.tbl))))){
  #   .tbl <- NULL
  # }
  return(.tbl)
}

process_all_tables <- function(tbl){
.tbl <- tbl %>%
  process_each_table(.) %>%
  long_to_wide(.)
return(.tbl)
}

extract_date <- function(txt){
  jj <- rdx(txt)
  if (length(jj) > 0){
    jj  %>%
      grepl(pattern = "bác sỹ|ngày", x = tolower(.)) %>%
      jj[.] %>%
      str_extract_all(., "\\d+") %>% .[[1]] %>%
      paste0(.,collapse = "-") %>%
      gsub("-09-00-00","",.) %>%
      gsub("-09-15-00","",.)
  }
  else {
    NA
  }
}

