#' WFsales
#'
#' @param nsid 
#' @param nspwd 
#' @param early 
#' @param late 
#' @param item 
#'
#' @return a data frame about the sales history
#' @export WFinvoices
#'
#' @examples WFsales(nsid,nspwd,90,0, c("664WH","777[BED]"))
#' 
WFsales <- function(nsid,nspwd,early,late, item){
  library(RODBC)
  library(dplyr)
  library(stringr)
channel <- odbcConnect("NetSuite", uid=nsid, pwd=nspwd)
  cost<-sqlQuery(channel,
                 "select  i.Name, p.ITEM_UNIT_PRICE, i.WAYFAIR_COM_PARTNER_SKU,  i.WAYFAIR_COM_SKU
                 from Items i, ITEM_PRICES p
                 where p.NAME = 'Fern FOB Cost + $2/cuft' and p.ITEM_ID = Item_ID
                 ")
  
 WFprice<-sqlQuery(channel,
                    "select  i.Name, p.ITEM_UNIT_PRICE, i.WAYFAIR_COM_PARTNER_SKU
                    from Items i, ITEM_PRICES p
                    where p.NAME = 'Dot Com (Wayfair.com)' and p.ITEM_ID = Item_ID
                    ")
  
  if(item=="n"){
  WFinvoices<-sqlQuery(channel,paste0(
    "select
    ENTITY.Name as WH,
    TRANSACTIONS.TRANID, 
    TRANSACTIONS.RELATED_TRANID,
    TRANSACTION_LINES.AMOUNT,
    TRANSACTION_LINES.ITEM_COUNT,
    DATE_CREATED,
    ITEMS.Name
    from
    TRANSACTION_LINES
    join (TRANSACTIONS join ENTITY on TRANSACTIONS.ENTITY_ID=ENTITY.ENTITY_ID)
    on TRANSACTION_LINES.TRANSACTION_ID=TRANSACTIONS.TRANSACTION_ID
    join LOCATIONS on TRANSACTION_LINES.LOCATION_ID=LOCATIONS.LOCATION_ID
    join ITEMS on TRANSACTION_LINES.ITEM_ID=ITEMS.ITEM_ID
    where
    TRANSACTION_LINES.AMOUNT is not null and
    TRANSACTIONS.TRANSACTION_TYPE='Invoice' and
    TRANSACTIONS.STATUS in ('Open','Paid In Full') and
    ENTITY.ENTITY_ID in (
    '9352',
    '8790',
    '20567',
    '20587',
    '22701',
    '14877',
    '10228',
    '11125',
    '11124',
    '10290',
    '12124',
    '14525') and
    DATE_CREATED BETWEEN TO_DATE( '" , Sys.Date()-early ,"', 'YYYY-MM-DD')
    AND TO_DATE('" , Sys.Date()-late ,"', 'YYYY-MM-DD')
    order by ENTITY.Name" ))
  }else{
    tm <- list()
    for (i in 1:length(n)){
    tm[[i]]<-sqlQuery(channel,paste0(
      "select
      ENTITY.Name as WH,
      TRANSACTIONS.TRANID, 
      TRANSACTIONS.RELATED_TRANID,
      TRANSACTION_LINES.AMOUNT,
      TRANSACTION_LINES.ITEM_COUNT,
      DATE_CREATED,
      ITEMS.Name
      from
      TRANSACTION_LINES
      join (TRANSACTIONS join ENTITY on TRANSACTIONS.ENTITY_ID=ENTITY.ENTITY_ID)
      on TRANSACTION_LINES.TRANSACTION_ID=TRANSACTIONS.TRANSACTION_ID
      join LOCATIONS on TRANSACTION_LINES.LOCATION_ID=LOCATIONS.LOCATION_ID
      join ITEMS on TRANSACTION_LINES.ITEM_ID=ITEMS.ITEM_ID
      where
      TRANSACTION_LINES.AMOUNT is not null and
      TRANSACTIONS.TRANSACTION_TYPE='Invoice' and
      TRANSACTIONS.STATUS in ('Open','Paid In Full') and
      ENTITY.ENTITY_ID in (
      '9352',
      '8790',
      '20567',
      '20587',
      '22701',
      '14877',
      '10228',
      '11125',
      '11124',
      '10290',
      '12124',
      '14525') and
      ITEMS.Name = ",n[i]," and
      DATE_CREATED BETWEEN TO_DATE( '" , Sys.Date()-early ,"', 'YYYY-MM-DD')
      AND TO_DATE('" , Sys.Date()-late ,"', 'YYYY-MM-DD')
      order by ENTITY.Name" ))
    }
    WFinvoices=tm[[1]]
    for (i in 2:length(n)){
    WFinvoices <-rbind( WFinvoices,tm[[i]])
    }
    }
WFinvoices$Name <- as.character(WFinvoices$Name)
cost$Name <- as.character(cost$Name)

WFinvoices <- merge(WFinvoices,cost, by="Name", all.x = TRUE, all.y = FALSE)
WFinvoices$margin <- ((WFinvoices$AMOUNT/WFinvoices$ITEM_COUNT)-WFinvoices$ITEM_UNIT_PRICE)/(WFinvoices$AMOUNT/WFinvoices$ITEM_COUNT)

return(WFinvoices)

}
