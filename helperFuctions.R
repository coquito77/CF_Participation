theme_CF <- function () {
  theme_bw(base_size=12, base_family="Avenir") %+replace%
    hrbrthemes::theme_ipsum_rc(plot_title_size = 20, subtitle_size = 15, caption_size = 10) +
    theme(panel.border = element_blank(),
          strip.background = element_rect(fill="blue"), # element_blank(),
          strip.text.x = element_text(size = 10,
                                      colour = "white",
                                      angle = 00,
                                      hjust = 0 #, lineheight = .5
          ),
          axis.title.y = element_text(face = "bold", size = 15, angle = 90)
          ,axis.title.x = element_text(face = "bold", size = 15, angle = 00, hjust = 0),
          axis.text.x = element_text(face = "bold", size = 10, angle = 00)
          ,axis.text.y = element_text(face = "bold", size = 10, angle = 00),
          legend.position = "top")
}


CW_Clean1 <- function(data) {

  x <- readxl::read_excel(Exl_CA237_paths[i]
                          ,sheet = Exl_CA237_shts[i]
                          #,range = "A1:M66"
                          ,col_types = "text")

  FirstRowPosition <- unique(which(Vectorize(function(x) x %in% c("Data Cell", "County Name"))(x),
                                   arr.ind=TRUE)[,1]) %>% tail(1)

  LastRowPosition <- unique(which(Vectorize(function(x) x %in% "Yuba")(x),
                                  arr.ind=TRUE)[,1]) %>% tail(1)

  col_CasesOpen <-  which(apply(x, 2, function(x) any(grepl("Cases open during the month", x))))

  x <- x %>%
    setNames(.[FirstRowPosition,]%>% unlist() %>% as.character()) %>%
    setNames(gsub("[^[:alnum:]]", perl = TRUE, "", names(.))) %>%
    setNames(make.names(names(.), unique = TRUE)) %>%
    select(County = 1 , TwoParentFamilies_A = col_CasesOpen
           , ZeroParentFamilies_B =  col_CasesOpen + 1
           , AllOtherFamilies_C =  col_CasesOpen + 2
           , TANFTimedOutCases_D = col_CasesOpen + 3) %>%
    mutate_at(vars(TwoParentFamilies_A:TANFTimedOutCases_D),
              funs(as.numeric(as.character(.)))) %>%
    mutate(total = rowSums(.[2:5], na.rm = TRUE)) %>%
    filter(total >0) %>%
    mutate(path = basename(Exl_CA237_paths[i]),
           sheet = Exl_CA237_shts[i],
           ReportMonth = paste0("01",sheet),
           ReportMonth = as.Date(ReportMonth, format = "%d%b%y"))
  all_tables[[i]] <-  x

}


CW_Clean2 <- function(data) {

  x <- readxl::read_excel(Exl_CA237_paths[i]
                          ,sheet = Exl_CA237_shts[i]
                          #,range = "A1:M66"
                          ,col_types = "text")

  FirstRowPosition <- unique(which(Vectorize(function(x) x %in% c("Data Cell", "County Name"))(x),
                                   arr.ind=TRUE)[,1]) %>% tail(1)

  LastRowPosition <- unique(which(Vectorize(function(x) x %in% "Yuba")(x),
                                  arr.ind=TRUE)[,1]) %>% tail(1)

  col_CasesOpen <-  which(apply(x, 2, function(x) any(grepl("Cases open during the month", x))))

  x <- x %>%
    setNames(.[FirstRowPosition,]%>% unlist() %>% as.character()) %>%
    setNames(gsub("[^[:alnum:]]", perl = TRUE, "", names(.))) %>%
    setNames(make.names(names(.), unique = TRUE)) %>%
    select(County = 1 , TwoParentFamilies_A = col_CasesOpen
           , ZeroParentFamilies_B =  col_CasesOpen + 1
           , AllOtherFamilies_C =  col_CasesOpen + 2
           , TANFTimedOutCases_D = col_CasesOpen + 3
           , SftN_FlnF_LTS = col_CasesOpen + 4) %>%
    mutate_at(vars(TwoParentFamilies_A:SftN_FlnF_LTS),
              funs(as.numeric(as.character(.)))) %>%
    mutate(total = rowSums(.[2:6], na.rm = TRUE)) %>%
    filter(total >0) %>%
    mutate(path = basename(Exl_CA237_paths[i]),
           sheet = Exl_CA237_shts[i],
           ReportMonth = paste0("01",sheet),
           ReportMonth = as.Date(ReportMonth, format = "%d%b%y"))

  all_tables[[i]] <-  x

}

CW_Clean3 <- function(data) {

  x <- readxl::read_excel(Exl_CA237_paths[i]
                          ,sheet = Exl_CA237_shts[i]
                          #,range = "A1:M66"
                          ,col_types = "text")

  FirstRowPosition <- unique(which(Vectorize(function(x) x %in% c("Data Cell", "County Name"))(x),
                                   arr.ind=TRUE)[,1]) %>% tail(1)

  LastRowPosition <- unique(which(Vectorize(function(x) x %in% "Yuba")(x),
                                  arr.ind=TRUE)[,1]) %>% tail(1)

  col_CasesOpen <-  which(apply(x, 2, function(x) any(grepl("Cases open during the month", x))))

  x <- x %>%
    setNames(.[FirstRowPosition,]%>% unlist() %>% as.character()) %>%
    setNames(gsub("[^[:alnum:]]", perl = TRUE, "", names(.))) %>%
    setNames(make.names(names(.), unique = TRUE)) %>%
    select(County = 2, ReportMonth
           , TwoParentFamilies_A = col_CasesOpen
           , ZeroParentFamilies_B =  col_CasesOpen + 1
           , AllOtherFamilies_C =  col_CasesOpen + 2
           , TANFTimedOutCases_D = col_CasesOpen + 3
           , SftN_FlnF_LTS = col_CasesOpen + 4) %>%
    mutate_at(vars(TwoParentFamilies_A:SftN_FlnF_LTS),
              funs(as.numeric(as.character(.)))) %>%
    mutate(ReportMonth = as.Date(as.numeric(as.character(ReportMonth)), origin = "1899-12-30"),
           total = rowSums(.[3:7], na.rm = TRUE)) %>%
    filter(total >0)

  all_tables[[i]] <-  x

}


GR_Clean1 <- function(data) {

  x <- readxl::read_excel(Exl_GR237_paths[i]
                          ,sheet = Exl_GR237_shts[i]
                          #,range = "A1:M66"
                          ,col_types = "text")

  FirstRowPosition <- unique(which(Vectorize(function(x) x %in% c("Data Cell", "County Name"))(x),
                                   arr.ind=TRUE)[,1]) %>% tail(1)

  LastRowPosition <- unique(which(Vectorize(function(x) x %in% "Yuba")(x),
                                  arr.ind=TRUE)[,1]) %>% tail(1)

  col_CasesOpen <-  which(apply(x, 2, function(x) any(grepl("Total cases$", x)))) %>% head(1)

  x <- x %>%
    setNames(.[FirstRowPosition,]%>% unlist() %>% as.character()) %>%
    setNames(gsub("[^[:alnum:]]", perl = TRUE, "", names(.))) %>%
    setNames(make.names(names(.), unique = TRUE)) %>%
    select(County = 1 ,
           TotalOpenCases = col_CasesOpen) %>%
    mutate_at(vars(TotalOpenCases),
              funs(as.numeric(as.character(.)))) %>%
    #mutate(total = rowSums(.[2:5], na.rm = TRUE)) %>%
    filter(TotalOpenCases > 0) %>%
    mutate(path = basename(Exl_GR237_paths[i]),
           sheet = Exl_GR237_shts[i],
           ReportMonth = paste0("01",sheet),
           ReportMonth = as.Date(ReportMonth, format = "%d%b%y"))
  all_tables[[i]] <-  x

}

GR_Clean2 <- function(data) {

  x <- readxl::read_excel(Exl_GR237_paths[i]
                          ,sheet = Exl_GR237_shts[i]
                          #,range = "A1:M66"
                          ,col_types = "text")

  FirstRowPosition <- unique(which(Vectorize(function(x) x %in% c("Data Cell", "County Name"))(x),
                                   arr.ind=TRUE)[,1]) %>% tail(1)

  LastRowPosition <- unique(which(Vectorize(function(x) x %in% "Yuba")(x),
                                  arr.ind=TRUE)[,1]) %>% tail(1)

  col_CasesOpen <-  which(apply(x, 2, function(x) any(grepl("Total cases$", x)))) %>% min()

  x <- x %>%
    setNames(.[FirstRowPosition,]%>% unlist() %>% as.character()) %>%
    setNames(gsub("[^[:alnum:]]", perl = TRUE, "", names(.))) %>%
    setNames(make.names(names(.), unique = TRUE)) %>%
    select(County = 2, ReportMonth
           ,  TotalOpenCases = col_CasesOpen) %>%
    mutate_at(vars(TotalOpenCases),
              funs(as.numeric(as.character(.)))) %>%
    mutate(ReportMonth = as.Date(as.numeric(as.character(ReportMonth)), origin = "1899-12-30")) %>%
    filter(TotalOpenCases >0)

  all_tables[[i]] <-  x

}



CF_DFA256_Clean1 <- function(data) {

  x <- readxl::read_excel(Exl_dfa256_paths[i]
                          ,sheet = Exl_dfa256_shts[i]
                          #,range = "A1:M66"
                          ,col_types = "text")

  FirstRowPosition <- unique(which(Vectorize(function(x) x %in% c("Data Cell", "County"))(x),
                                   arr.ind=TRUE)[,1]) %>% tail(1)

  LastRowPosition <- unique(which(Vectorize(function(x) x %in% "Yuba")(x),
                                  arr.ind=TRUE)[,1]) %>% tail(1)

  col_CasesOpen <- which(apply(x, 2, function(x) any(grepl("Total", x)))) %>% head(1)

  x <- x %>%
    setNames(.[FirstRowPosition,]%>% unlist() %>% as.character()) %>%
    setNames(gsub("[^[:alnum:]]", perl = TRUE, "", names(.))) %>%
    setNames(make.names(names(.), unique = TRUE)) %>%
    select(County = 1 ,
           HH_Fed = col_CasesOpen,
           HH_Fed_State = col_CasesOpen +1 ,
           HH_State = col_CasesOpen + 2 ) %>%
    mutate_at(vars(HH_Fed:HH_State),
              funs(as.numeric(as.character(.)))) %>%
    mutate(total = rowSums(.[2:4], na.rm = TRUE)) %>%
    filter(total > 0) %>%
    mutate(path = basename(Exl_dfa256_paths[i]),
           sheet = Exl_dfa256_shts[i],
           ReportMonth = paste0("01",sheet),
           ReportMonth = as.Date(ReportMonth, format = "%d%b%y"))
  all_tables[[i]] <-  x

}

CF_DFA256_Clean2 <- function(data) {

  x <- readxl::read_excel(Exl_dfa256_paths[i]
                          ,sheet = Exl_dfa256_shts[i]
                          #,range = "A1:M66"
                          ,col_types = "text")

  FirstRowPosition <- unique(which(Vectorize(function(x) x %in% c("Data Cell", "County"))(x),
                                   arr.ind=TRUE)[,1]) %>% tail(1)

  LastRowPosition <- unique(which(Vectorize(function(x) x %in% "Yuba")(x),
                                  arr.ind=TRUE)[,1]) %>% tail(1)

  col_CasesOpen <- which(apply(x, 2, function(x) any(grepl("Total", x)))) %>% head(1)

  x <- x %>%
    setNames(.[FirstRowPosition,]%>% unlist() %>% as.character()) %>%
    setNames(gsub("[^[:alnum:]]", perl = TRUE, "", names(.))) %>%
    setNames(make.names(names(.), unique = TRUE)) %>%
    select(County = 2, ReportMonth = 7,
           HH_Fed = col_CasesOpen,
           HH_Fed_State = col_CasesOpen +1 ,
           HH_State = col_CasesOpen + 2 ) %>%
    mutate_at(vars(HH_Fed:HH_State),
              funs(as.numeric(as.character(.)))) %>%
    mutate(total = rowSums(.[3:5], na.rm = TRUE)) %>%
    filter(total > 0) %>%
    mutate(ReportMonth = as.Date(as.numeric(as.character(ReportMonth)), origin = "1899-12-30"),
           path = basename(Exl_dfa256_paths[i]),
           sheet = Exl_dfa256_shts[i])

  all_tables[[i]] <-  x

}
