
#' CG_Nmargin
#'
#' @param nsid Netsuit ID
#' @param nspwd Netsuit password
#'
#' @return result
#' @export
#'
#' @examples
CG_Nmargin <- function(nsid,nspwd){
library(RODBC)
library(dplyr)
library(stringr)
rplall <- read.csv("C:/Users/yao.Guan/Desktop/usefulData/rplall.csv")
cuft <- read.csv("C:/Users/yao.Guan/Desktop/usefulData/CUFT88.csv")
cuft <- cuft[,c("Name","Cuft")]
channel <- odbcConnect("NetSuite", uid=nsid, pwd=nspwd)
CG88 <- read.csv("N:/E Commerce/Public Share/Dot Com - Wayfair/Sales predict/Increasing inventory/fileCG88auto.csv")
CG88 <- CG88[,c(1,(ncol(CG88)-90)
                :ncol(CG88))]
dim(CG88)
CG88 <-
  CG88 %>%
  mutate(sum90 = rowSums(.[2:ncol(CG88)])) %>%
  select(WF.Partner.SKU,sum90)

seCG88 <- merge(CG88,rplall,by.x = "WF.Partner.SKU",by.y = "pname")
seCG88 <- aggregate(sum90~ Name, data = seCG88,sum)
seCG88 <- merge(seCG88,cuft,by="Name",all.x = TRUE)
seCG88$storage <- seCG88$sum90*seCG88$Cuft*.22/30

##get price
itemprice<-sqlQuery(channel,
                    "select  i.Name, p.ITEM_UNIT_PRICE
                    from Items i, ITEM_PRICES p
                    where p.NAME = 'Fern FOB Cost + $2/cuft' and p.ITEM_ID = Item_ID
                    ")
  
##get sales
SIteminvoic <- sqlQuery(channel,
                        "select
                        ITEMS.Name as name,
                        sum(TRANSACTION_LINES.ITEM_COUNT) as 'qty',
                        sum(TRANSACTION_LINES.AMOUNT) as 'amount'

                        from
                        TRANSACTION_LINES
                        join (TRANSACTIONS join ENTITY on TRANSACTIONS.ENTITY_ID=ENTITY.ENTITY_ID)
                        on TRANSACTION_LINES.TRANSACTION_ID=TRANSACTIONS.TRANSACTION_ID
                        join LOCATIONS on TRANSACTION_LINES.LOCATION_ID=LOCATIONS.LOCATION_ID
                        join ITEMS on TRANSACTION_LINES.ITEM_ID=ITEMS.ITEM_ID
                        where
                        TRANSACTION_LINES.AMOUNT is not null and
                        TRANSACTIONS.TRANSACTION_TYPE='Sales Order' and
                        TRANSACTIONS.STATUS='Billed' and

                        ENTITY.Name in ('Wayfair.com',
                        'Castlegate - CA Perris',
                        'Castlegate - CA Perris SP',
                        'Castlegate - GA McDonough',
                        'Castlegate - KY Erlanger',
                        'Castlegate - KY Florence',
                        'Castlegate - KY Hebron',
                        'Castlegate - NJ Cranbury',
                        'Castlegate - NJ Cranbury LP',
                        'Castlegate - TX Lancaster SP',
                        'Castlegate - UT',
                        'JossandMain.com') and
                        DATE_CREATED BETWEEN TO_DATE('2018-OCT-01', 'YYYY-MON-DD')
                        AND TO_DATE('2018-DEC-31', 'YYYY-MON-DD')

                        group by ITEMS.Name
                        ")

##aged result
result <- merge(seCG88,itemprice, by= "Name",all.x = TRUE )
result <- merge(result,SIteminvoic, by.x = "Name", by.y = "name", all.x = TRUE)
result$qty <- -result$qty
result$amount <- -result$amount
result$stOverP <- result$storage/(result$amount-result$qty*result$ITEM_UNIT_PRICE)
result <- filter(result,storage!=0)
result <- unique(result)


write.csv(result,paste0("N:/E Commerce/Public Share/Dot Com - Wayfair/CG Slow Item/",Sys.Date(),"StorageMargin.csv"))


### monthly
rplall <- read.csv("C:/Users/yao.Guan/Desktop/usefulData/rplall.csv")
cuft <- read.csv("C:/Users/yao.Guan/Desktop/usefulData/CUFT88.csv")
cuft <- cuft[,c("Name","Cuft")]
channel <- odbcConnect("NetSuite", uid="yao.guan@top-line.com", pwd="NetYG@Chicago")
CG88 <- read.csv("N:/E Commerce/Public Share/Dot Com - Wayfair/Sales predict/Increasing inventory/fileCG88auto.csv")
CG88 <- CG88[,c(1,(ncol(CG88)-30)
                :ncol(CG88))]
dim(CG88)
CG88 <-
  CG88 %>%
  mutate(sum90 = rowSums(.[2:ncol(CG88)])) %>%
  select(WF.Partner.SKU,sum90)

seCG88 <- merge(CG88,rplall,by.x = "WF.Partner.SKU",by.y = "pname")
seCG88 <- aggregate(sum90~ Name, data = seCG88,sum)
seCG88 <- merge(seCG88,cuft,by="Name",all.x = TRUE)
seCG88$storage <- seCG88$sum90*seCG88$Cuft*.22/30

SIteminvoic <- sqlQuery(channel,
                        "select
                        ITEMS.Name as name,
                        sum(TRANSACTION_LINES.ITEM_COUNT) as 'qty',
                        sum(TRANSACTION_LINES.AMOUNT) as 'amount'

                        from
                        TRANSACTION_LINES
                        join (TRANSACTIONS join ENTITY on TRANSACTIONS.ENTITY_ID=ENTITY.ENTITY_ID)
                        on TRANSACTION_LINES.TRANSACTION_ID=TRANSACTIONS.TRANSACTION_ID
                        join LOCATIONS on TRANSACTION_LINES.LOCATION_ID=LOCATIONS.LOCATION_ID
                        join ITEMS on TRANSACTION_LINES.ITEM_ID=ITEMS.ITEM_ID
                        where
                        TRANSACTION_LINES.AMOUNT is not null and
                        TRANSACTIONS.TRANSACTION_TYPE='Sales Order' and
                        TRANSACTIONS.STATUS='Billed' and

                        ENTITY.Name in ('Wayfair.com',
                        'Castlegate - CA Perris',
                        'Castlegate - CA Perris SP',
                        'Castlegate - GA McDonough',
                        'Castlegate - KY Erlanger',
                        'Castlegate - KY Florence',
                        'Castlegate - KY Hebron',
                        'Castlegate - NJ Cranbury',
                        'Castlegate - NJ Cranbury LP',
                        'Castlegate - TX Lancaster SP',
                        'Castlegate - UT',
                        'JossandMain.com') and
                        DATE_CREATED BETWEEN TO_DATE('2018-DEC-01', 'YYYY-MON-DD')
                        AND TO_DATE('2018-DEC-31', 'YYYY-MON-DD')

                        group by ITEMS.Name
                        ")

##aged result
result <- merge(seCG88,itemprice, by= "Name",all.x = TRUE )
result <- merge(result,SIteminvoic, by.x = "Name", by.y = "name", all.x = TRUE)
result$qty <- -result$qty
result$amount <- -result$amount
result$stOverP <- result$storage/(result$amount-result$qty*result$ITEM_UNIT_PRICE)
result <- filter(result,storage!=0)
result <- unique(result)
write.csv(result,paste0("N:/E Commerce/Public Share/Dot Com - Wayfair/CG Slow Item/",Sys.Date(),"(M)StorageMargin.csv"))
}

