


getSkipValue <- function(x, destfile, sheets, FUN){
    
    Foo <- try({
        
        if (x %in% c('gbtenderhistory', 'gbonissuehistory', 'gbsyndication', 'gbrepurchases', 'tbtenderhistory', 'ecponissue', 'iibfactors', 'nmihistory')){
            
            out <- sapply(seq_along(sheets), function(ii){
                
                tmp <- suppressWarnings(FUN(destfile, sheet = sheets[ii],
                                            .name_repair = 'minimal'))
                tmp <- as.data.frame(tmp)
                
                skip <- grep('^\\d{5}$', tmp[, 1])[1] - 1

                if (is.na(skip))
                    skip <- 0

                if (x == 'nmihistory')
                    skip <- skip - 1
                
                skip
                
            })
            
        } else if (x %in% 'gbonissue'){

            out <- sapply(seq_along(sheets), function(ii){

                tmp <- suppressWarnings(FUN(destfile, sheet = sheets[ii],
                                            .name_repair = 'minimal'))
                tmp <- as.data.frame(tmp)
                
                skip <- grep('^\\d+/\\d{4}$', tmp[, 1])[1] - 1

                if (is.na(skip))
                    skip <- 0

                skip
                
            })
            
        } else if (x %in% c('tbonissue') ){

            out <- sapply(seq_along(sheets), function(ii){

                tmp <- suppressWarnings(FUN(destfile, sheet = sheets[ii],
                                            .name_repair = 'minimal'))
                tmp <- as.data.frame(tmp)
                
                skip <- which(tmp[, 2] %in% 'Maturity')[1] - 1

                skip
                
            })
            
        } else if (x %in% 'tbonissuehistory'){

            out <- sapply(seq_along(sheets), function(ii){

                tmp <- suppressWarnings(FUN(destfile, sheet = sheets[ii],
                                            .name_repair = 'minimal'))
                tmp <- as.data.frame(tmp)

                if (ii == 1){
                    skip <- grep('^\\d+-[[:alpha:]]+-\\d{4}$', tmp[, 1])[1] - 1
                } else if (ii == 2){
                    skip <- grep('^\\d{5}$', tmp[, 1])[1] - 1
                }
                
                if (is.na(skip))
                    skip <- 0

                skip
                
            })
            
        }
        
    }, silent = TRUE)
    
    if (inherits(Foo, 'try-error') ){
        stop('readxl could not read ', destfile)
    }
    
    out
    
}
