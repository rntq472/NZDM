##' Download Data From the NZDM Website.
##' 
##' Retrieves data from
##' \url{https://debtmanagement.treasury.govt.nz/investor-resources/data}. This
##' involves downloading spreadsheets and reading them into R.
##' 
##' @param x Name of dataset to retrieve. See details below.
##' @param destDir Directory into which the spreadsheet should be downloaded into.
##'                Defaults to NULL in which case a temp directory is used.
##' @param quiet Logical passed to download.file.
##' @param webAddress URL where the spreadsheets are kept. Users don't need to change
##'                   this.
##' 
##' @return A data frame (or a list of them if the spreadsheet contained multiple
##'         sheets) containing the requested data.
##' 
##' @details There are eleven datasets that can be retrieved whose names are listed
##'          below:
##' 
##' \itemize{
##'     \item Government bonds - tender issuance history
##'     \item Government Bonds - on issue (history at month end)
##'     \item Government Bonds - on issue (at month end)
##'     \item Government Bonds - Syndication
##'     \item Government bonds - tender history (repurchases)
##'     \item Treasury bills - tender history
##'     \item Treasury Bills - on issue (at month end)
##'     \item Treasury Bills - on issue (history at month end)
##'     \item European Commercial Paper -  on issue (at month end)
##'     \item Inflation-indexed bonds - factors spreadsheet
##'     \item Non market issuance - issuance history
##' }
##' 
##' There are a number of short-cuts used so that they are easier for the user to
##' specify. Check the code for exact details, but the gist of it is that spaces are
##' removed, letters are converted to lowercase, and certain words are truncated to
##' their initials. This means you can enter the required names in a variety of
##' formats (e.g. "Treasury Bill tender issuance history" is the same as
##' "tbtenderhistory" and hopefully they will be recognised. If in doubt you can
##' enter one of the names listed below.
##' 
##' \itemize{
##'     \item GB tender history
##'     \item GB onissue history
##'     \item GB onissue
##'     \item GB syndication
##'     \item GB repurchases
##'     \item TB tender history
##'     \item TB onissue
##'     \item TB onissue history
##'     \item ECP onissue
##'     \item IIB factors
##'     \item NMI history
##' }
##' 
##' @note The author has no affiliation with the NZDM office.
##' 
##' @examples
##' \dontrun{
##' dat <- getNZDM('GB tender history')
##' plot(dat$Tender_Date, dat[, 'Wtd.Avg._successful_Yield_(%)'], type = 'l')
##' }
##' 
##' @author Jasper Watson
##' 
##' @import httr
##' @import readxl
##' @import rvest
##' @import xml2
##' @importFrom utils download.file
##' 
##' @export
##' 
##
getNZDM <- function(x, destDir = NULL, quiet = TRUE,
                    webAddress = 'https://debtmanagement.treasury.govt.nz/investor-resources/data'
                    ){
    
    if (is.null(destDir))
        destDir <- tempdir()
    
    x <- tolower(gsub(' |\\-', '', x))
    
    x <- gsub('tenderissuance', 'tender', x)
    x <- gsub('government', 'g', x)
    x <- gsub('govt', 'g', x)
    x <- gsub('bonds', 'b', x)
    x <- gsub('treasury', 't', x)
    x <- gsub('bills', 'b', x)
    x <- gsub('european', 'e', x)
    x <- gsub('commercial', 'c', x)
    x <- gsub('paper', 'p', x)
    x <- gsub('inflation', 'i', x)
    x <- gsub('indexed', 'i', x)
    x <- gsub('non', 'n', x)
    x <- gsub('market', 'm', x)
    x <- gsub('issuance', 'i', x)
    
    if (grepl('repurchases', x, ignore.case = TRUE))
        x <- 'gbrepurchases'
    
    x <- match.arg(x,
                   choices = c('gbtenderhistory',
                               'gbonissuehistory',
                               'gbonissue',
                               'gbsyndication',
                               'gbrepurchases',
                               'tbtenderhistory',
                               'tbonissue',
                               'tbonissuehistory',
                               'ecponissue',
                               'iibfactors',
                               'nmihistory'
                               )
                   )
    
    Foo <- try({
        indexPage <- read_html(webAddress)
    })
    
    if (inherits(Foo, 'try-error'))
        stop('Could not read ', webAddress)
    
    desc <- switch(x,
                   'gbtenderhistory'  = 'Government bonds - tender issuance history',
                   'gbonissuehistory' = 'Government Bonds - on issue (history at month end)',
                   'gbonissue'        = 'Government Bonds - on issue (at month end)',
                   'gbsyndication'    = 'Government Bonds - Syndication',
                   'gbrepurchases'    = 'Government bonds - tender history (repurchases)',
                   'tbtenderhistory'  = 'Treasury bills - tender history',
                   'tbonissue'        = 'Treasury Bills - on issue (at month end)',
                   'tbonissuehistory' = 'Treasury Bills - on issue (history at month end)',
                   'ecponissue'       = 'European Commercial Paper -  on issue (at month end)',
                   'iibfactors'       = 'Inflation-indexed bonds - factors spreadsheet ',
                   'nmihistory'       = 'Non market issuance - issuance history'
                   )
    
    node <- html_elements(indexPage,
                          xpath = paste0("//*[text()='", desc, "']")
                          )
    
    url <- html_attr(node, 'href')
    
    destfile <- path.expand(file.path(destDir, basename(url)))
    
    if (!file.exists(destfile) ){
        
        Foo <- try(download.file(url = url,
                                 destfile = destfile,
                                 mode = 'wb',
                                 quiet = quiet
                                 )
                   )
        
        if (inherits(Foo, 'try-error') || Foo > 0){
            stop('Could not download ', url)
        }
        
    }
    
    Foo <- try({
        sheets <- excel_sheets(destfile)
    }, silent = TRUE)
    
    if (inherits(Foo, 'try-error') ){
        stop('readxl could not read ', destfile)
    }
    
    FUN <- switch(readxl::format_from_ext(destfile),
                  xls = readxl::read_xls,
                  xlsx = readxl::read_xlsx
                  )

    skip <- getSkipValue(x, destfile, sheets, FUN)
    
    Foo <- try({
        out <- lapply(seq_along(sheets), function(ii){
            
            suppressWarnings(FUN(destfile, sheet = sheets[ii], skip = skip[ii],
                                 .name_repair = 'minimal'))
            
        })
    }, silent = TRUE)
    
    if (inherits(Foo, 'try-error') ){
        stop('readxl could not read ', destfile)
    }
    
    names(out) <- sheets
    
    for (ii in seq_along(out) ){
        
        out[[ii]] <- as.data.frame(out[[ii]], stringsAsFactors = FALSE)
        
        colnames(out[[ii]]) <- gsub('[[:space:]]+', '_', colnames(out[[ii]]))
    }
    
    if (length(sheets) == 1)
        out <- out[[1]]
    
    ##*******
    
    if (x == 'gbtenderhistory'){
        
    } else if (x == 'gbonissuehistory'){
        
        out[['Month_end']] <- out[['Month_end']][, 1:11]
        
        out[['Month_end']] <- out[['Month_end']][!is.na(out[['Month_end']][, 1]) &
                                                 !(out[['Month_end']][, 1] %in% '-'), ]
        
        out[['Month_end']][, 'Month_end'] <- as.Date(as.numeric(out[['Month_end']][, 'Month_end']),
                                                     origin = '1899-12-30')
        
        tmp <- out[['Month_end']][, 'Maturity_Date']
        use <- grepl('^\\d+$', tmp)
        aa <- tmp[use]
        bb <- tmp[!use]
        aa <- as.Date(as.numeric(aa), origin = '1899-12-30')
        qwer <- grepl('\\-\\d{2}$', bb)
        asdf <- strsplit(bb[qwer], '-', fixed = TRUE)
        for (ii in seq_along(asdf))
            asdf[[ii]][3] <- paste0('20', asdf[[ii]][3])
        bb[qwer] <- sapply(asdf, function(zz){
            paste(zz, collapse = '-')
        })
        bb <- gsub(' ', '-', bb)
        bb <- strptime(bb, format = '%d-%b-%Y')
        
        tmp <- rep(Sys.Date(), length(tmp))
        
        tmp[use] <- aa
        tmp[!use] <- as.Date(bb)
        
        out[['Month_end']][, 'Maturity_Date'] <- tmp
        
        tmp <- out[['Month_end']][, 'Date_First_Issued']
        tmp[tmp == '-'] <- NA
        
        out[['Month_end']][, 'Date_First_Issued'] <- as.Date(strptime(paste0('01/', tmp),
                                                                      format = '%d/%m/%Y'))
        
        out[['Summary']][, 'Month_end'] <- as.Date(out[['Summary']][, 'Month_end'])
        
    } else if (x == 'gbonissue'){
        
        new.out <- list()
        
        new.out[['On_Issue']] <- out[1:(which(is.na(out[, 1]))[1] - 1), ]
        
        out <- out[-(1:(which(is.na(out[, 1]))[1] - 1)), ]
        out <- out[-(1:(which(out[, 1] %in% 'INDEX LINKED BONDS'))), ]
        colnames(out) <- unlist(out[1, ])
        out <- out[-1, ]
        
        new.out[['IIB']] <- out[1:(which(is.na(out[, 1]))[1] - 1), ]
        
        out <- out[-(1:(which(is.na(out[, 1]))[1] - 1)), ]
        out <- out[-(1:(which(out[, 1] %in% 'SUMMARY'))), ]
        
        out <- out[, !is.na(out[1, ])]
        
        colnames(out) <- c('Description', 'Date', 'Value')
        
        new.out[['Summary']] <- out
        
        out <- new.out
        
        for (ii in seq_along(out) ){
            colnames(out[[ii]]) <- gsub('[[:space:]]+', '_', colnames(out[[ii]]))
        }
        
        out$On_Issue$Date_first_Issued <- strptime(paste0('01/', out$On_Issue$Date_first_Issued),
                                                   format = '%d/%m/%Y')
        out$On_Issue$Maturity <- as.Date(as.numeric(out$On_Issue$Maturity),
                                         origin = '1899-12-30')
        
        out$IIB$Maturity <- as.Date(as.numeric(out$IIB$Maturity),
                                    origin = '1899-12-30')
        out$IIB$Last_Coupon <- as.Date(as.numeric(out$IIB$Last_Coupon),
                                       origin = '1899-12-30')
        out$IIB$Next_Coupon <- as.Date(as.numeric(out$IIB$Next_Coupon),
                                       origin = '1899-12-30')
        
        out$Summary[, 'Date'] <- as.Date(as.numeric(out$Summary[, 'Date']),
                                         origin = '1899-12-30')
        
    } else if (x == 'gbsyndication'){
        
        for (cc in c('Asia', 'Australia', 'Europe_UK', 'New_Zealand', 'North_America', 'Other', 'Asset_Manager_Central_Bank', 'Bank_Balance_Sheet', 'Bank_Trading_Book', 'Hedge_Fund'))
            out[[1]][[cc]] <- NA
        
        for (ii in grep('^\\d{4}$', sheets) ){

            tmp <- rep(Sys.Date(), 3)
            tmp[1] <- as.Date(as.numeric(out[[ii]][1, 1]), origin = '1899-12-30')
            tmp[2] <- as.Date(as.numeric(out[[ii]][1, 2]), origin = '1899-12-30')
            tmp[3] <- out[[ii]][1, 3]
            
            loc <- sapply(1:nrow(out[[1]]), function(z){

                oo <- as.Date(c(out[[1]][z, 1], out[[1]][z, 2], out[[1]][z, 3]))
                all(oo == tmp)

            })

            if (sum(loc) == 1){

                out[[1]][loc, 'Asia'] <- out[[ii]][grepl('Asia', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'Australia'] <- out[[ii]][grepl('Australia', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'Europe_UK'] <- out[[ii]][grepl('Europe', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'New_Zealand'] <- out[[ii]][grepl('New.*Zealand', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'North_America'] <- out[[ii]][grepl('North.*America', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'Other'] <- out[[ii]][grepl('Other', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'Asset_Manager_Central_Bank'] <- out[[ii]][grepl('Asset.*Manager.*Central.*Bank', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'Bank_Balance_Sheet'] <- out[[ii]][grepl('Bank.*Balance.*Sheet', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'Bank_Trading_Book'] <- out[[ii]][grepl('Bank.*Trading.*Book', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[1]][loc, 'Hedge_Fund'] <- out[[ii]][grepl('Hedge.*Fund', out[[ii]][, 1], ignore.case = TRUE), 2]

            } 
            
        }

        for (ii in grep('^\\d{4}[[:space:]]*Tap$', sheets) ){

            tmp <- rep(Sys.Date(), 3)
            tmp[1] <- as.Date(as.numeric(out[[ii]][1, 1]), origin = '1899-12-30')
            tmp[2] <- as.Date(as.numeric(out[[ii]][1, 2]), origin = '1899-12-30')
            tmp[3] <- out[[ii]][1, 3]
            
            loc <- sapply(1:nrow(out[[2]]), function(z){

                oo <- as.Date(c(out[[2]][z, 1], out[[2]][z, 2], out[[2]][z, 3]))
                all(oo == tmp)

            })

            if (sum(loc) == 1){

                out[[2]][loc, 'Asia'] <- out[[ii]][grepl('Asia', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'Australia'] <- out[[ii]][grepl('Australia', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'Europe_UK'] <- out[[ii]][grepl('Europe', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'New_Zealand'] <- out[[ii]][grepl('New.*Zealand', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'North_America'] <- out[[ii]][grepl('North.*America', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'Other'] <- out[[ii]][grepl('Other', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'Asset_Manager_Central_Bank'] <- out[[ii]][grepl('Asset.*Manager.*Central.*Bank', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'Bank_Balance_Sheet'] <- out[[ii]][grepl('Bank.*Balance.*Sheet', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'Bank_Trading_Book'] <- out[[ii]][grepl('Bank.*Trading.*Book', out[[ii]][, 1], ignore.case = TRUE), 2]
                out[[2]][loc, 'Hedge_Fund'] <- out[[ii]][grepl('Hedge.*Fund', out[[ii]][, 1], ignore.case = TRUE), 2]

            } 
            
        }

        out <- out[1:2]
        
    } else if (x == 'gbrepurchases'){
        
    } else if (x == 'tbtenderhistory'){
        
    } else if (x == 'tbonissue'){

        tmp <- which(!is.na(out$Maturity))[1]

        if (tmp > 1)
            out <- out[-(seq_len(tmp - 1)), ]
        
    } else if (x == 'tbonissuehistory'){
        
        out[['Month_end']] <- out[['Month_end']][!is.na(out[['Month_end']][, 1]) & !(out[['Month_end']][, 1] %in% '-'), ]
        
        tmp <- out[['Month_end']][, 'Month_End']
        use <- grepl('^\\d+$', tmp)
        aa <- tmp[use]
        bb <- tmp[!use]
        aa <- as.Date(as.numeric(aa), origin = '1899-12-30')
        qwer <- grepl('\\-\\d{2}$', bb)
        asdf <- strsplit(bb[qwer], '-', fixed = TRUE)
        for (ii in seq_along(asdf))
            asdf[[ii]][3] <- paste0('20', asdf[[ii]][3])
        bb[qwer] <- sapply(asdf, function(zz){
            paste(zz, collapse = '-')
        })
        bb <- gsub(' ', '-', bb)
        bb <- strptime(bb, format = '%d-%b-%Y')
        
        tmp <- rep(Sys.Date(), length(tmp))
        
        tmp[use] <- aa
        tmp[!use] <- as.Date(bb)
        
        out[['Month_end']][, 'Month_End'] <- tmp
        
        ##
        
        tmp <- out[['Month_end']][, 'Maturity']
        use <- grepl('^\\d+$', tmp)
        aa <- tmp[use]
        bb <- tmp[!use]
        aa <- as.Date(as.numeric(aa), origin = '1899-12-30')
        qwer <- grepl('\\-\\d{2}$', bb)
        asdf <- strsplit(bb[qwer], '-', fixed = TRUE)
        for (ii in seq_along(asdf))
            asdf[[ii]][3] <- paste0('20', asdf[[ii]][3])
        bb[qwer] <- sapply(asdf, function(zz){
            paste(zz, collapse = '-')
        })
        bb <- gsub(' ', '-', bb)
        bb <- strptime(bb, format = '%d-%b-%Y')
        
        tmp <- rep(Sys.Date(), length(tmp))
        
        tmp[use] <- aa
        tmp[!use] <- as.Date(bb)
        
        out[['Month_end']][, 'Maturity'] <- tmp
        
    } else if (x == 'ecponissue'){
        
        out <- out[!is.na(out[, 2]), ]
        out[, 'Month_End'] <- as.Date(as.numeric(out[, 'Month_End']),
                                      origin = '1899-12-30')
        
    } else if (x == 'iibfactors'){
        
        for (ii in seq_along(out) ){

            out[[ii]]$Bond <- gsub('[[:space:]]*IIB[[:space:]]*Factors', '', sheets[ii])
            out[[ii]] <- out[[ii]][, c(4, 1:3)]

        }

        out <- do.call(rbind.data.frame, out)
        sheets <- sheets[1]
        
    } else if (x == 'nmihistory'){
        
        colnames(out) <- c('Settlement_Date', 'Maturity_Date', 'Instrument_Type',
                           'Institution', 'Amount_(Millions)')
        
        out <- out[!is.na(out[, 1]), ]
        
    }

    if (length(sheets) == 1){
        rownames(out) <- NULL
    } else {
        for (ii in seq_along(out))
            rownames(out[[ii]]) <- NULL
    }
    
    out
    
}
