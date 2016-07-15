rm(list = ls())

source("functions.R")

doc_files <- list.files("data/",".docx", full.names = T)

rs <- list() #Create a list for storing tables
#Run a loop of number of docx files scanned
for (i in seq_len(length(doc_files))){
  print(i)
  smp_docx <- read_docx(doc_files[[i]])
  if (docx_tbl_count(smp_docx) > 0){
    smp_tbls <- extract_table(smp_docx)
    date_value <- doc_files[[i]] %>% extract_date(.)
    .tbl <- bind_cols(map(smp_tbls, process_all_tables)) %>%
      data.frame(stt = i, file_name = gsub("data//|.docx","",smp_docx$path),date_collect = date_value, .)
    names(.tbl) <- names(.tbl) %>%
      gsub("\\.{1,9}","_",.) %>%
      gsub("\\_$","",.) %>%
      gsub("_1$","",.) %>%
      tolower(.) %>%
      gsub("_test$|_viêm_gan_b$","",.) %>%
      gsub("tổng_pt","tổng_phân_tích",.) %>%
      gsub("feto_protein","foeto_protein",.) %>%
      gsub("_viêm_dạ_dày$","",.) %>%
      gsub("kt_kháng_lao$","kháng_thể_kháng_lao",.) %>%
      gsub("_viêm_gan_b$","",.) %>%
      gsub("creatinin$","creatinine",.) %>%
      gsub("cholesterol_mỡ_máu","cholesterol",.) %>%
      gsub("anti_hbsag","hbsag",.) %>%
      gsub("kháng_thể_hbsag","hbsag",.) %>%
      gsub("anti_hcv","hvc",.)

    rs[[i]] <- .tbl
  }
}
rs1 <- bind_rows(rs) #Combind all docx's tables into a data frame.
#Remove some dirty letters
rs2  <- apply(rs1, 2, function(x) ifelse(grepl("\\/L",x),"",x)) %>% as.data.frame((.))

library(xlsx) #For writing out the xls file

write.xlsx(rs2, "thongke_xetnghiem.xls", sheetName = "ThongKe", row.names = F)
saveRDS(rs2, "report.rds")
