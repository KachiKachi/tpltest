
#' Title
#'
#' @param nsid Netsuit ID
#' @param nspwd Netsuit password
#'
#' @return
#' @export
#'
#' @examples
wow <- function(nsid,nspwd){
rm(list=ls())
library(RODBC)
library(dplyr)
library(stringr)
w <- as.numeric(format(Sys.Date(), "%W"))

todate <- read.csv("C:/Users/yao.Guan/Desktop/usefulData/todate1819.csv")
todate$w18 <- as.character(todate$w18)
todate$w19 <- as.character(todate$w19)

##get week data
channel <- odbcConnect("NetSuite", uid=nsid, pwd=nspwd)

thisweek<-sqlQuery(channel,paste0(
  "select
  ENTITY.Name as name,
  sum(TRANSACTION_LINES.ITEM_COUNT) as 'qty',
  sum(TRANSACTION_LINES.AMOUNT) as 'amount',
  count(DISTINCT TRANSACTION_LINES.TRANSACTION_ID) as 'order'
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
  '5559',
  '9926',
  '17569',
  '9519',
  '20010',
  '5788',
  '20567',
  '16499',
  '18380',
  '18615',
  '5387',
  '21336',
  '5376',
  '20587',
  '5933',
  '5453',
  '22361',
  '5568',
  '20011',
  '22701',
  '7346',
  '15590',
  '8513',
  '8571',
  '9445',
  '6328',
  '7190',
  '3948',
  '3950',
  '3953',
  '3958',
  '3959',
  '3961',
  '3962',
  '3963',
  '3967',
  '3968',
  '14877',
  '8610',
  '15932',
  '18668',
  '10228',
  '11066',
  '11079',
  '9322',
  '11125',
  '11124',
  '4489',
  '7996',
  '12868',
  '12902',
  '8620',
  '8772',
  '8790',
  '9352',
  '10290',
  '12124',
  '14525',
  '16499') and
  DATE_CREATED BETWEEN TO_DATE( '" , todate[w,2] ,"', 'YYYY-MON-DD')
  AND TO_DATE('" , todate[w+1,2] ,"', 'YYYY-MON-DD')
  
  group by ENTITY.Name
  order by ENTITY.Name" ))

write.csv(thisweek,paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/2019/",Sys.Date(),"dotcom19W.csv"),row.names = FALSE)

lastyear<-sqlQuery(channel,paste0(
  "select
  ENTITY.Name as name,
  sum(TRANSACTION_LINES.ITEM_COUNT) as 'qty',
  sum(TRANSACTION_LINES.AMOUNT) as 'amount',
  count(distinct TRANSACTION_LINES.TRANSACTION_ID) as 'order'
  from
  TRANSACTION_LINES
  join (TRANSACTIONS join ENTITY on TRANSACTIONS.ENTITY_ID=ENTITY.ENTITY_ID)
  on TRANSACTION_LINES.TRANSACTION_ID=TRANSACTIONS.TRANSACTION_ID
  join LOCATIONS on TRANSACTION_LINES.LOCATION_ID=LOCATIONS.LOCATION_ID
  join ITEMS on TRANSACTION_LINES.ITEM_ID=ITEMS.ITEM_ID
  where
  TRANSACTION_LINES.AMOUNT is not null and
  TRANSACTIONS.STATUS in ('Open','Paid In Full') and
  ENTITY.ENTITY_ID in (
  '5559',
  '9926',
  '17569',
  '9519',
  '20010',
  '5788',
  '20567',
  '16499',
  '18380',
  '18615',
  '5387',
  '21336',
  '5376',
  '20587',
  '5933',
  '5453',
  '22361',
  '5568',
  '20011',
  '22701',
  '7346',
  '15590',
  '8513',
  '8571',
  '9445',
  '6328',
  '7190',
  '3948',
  '3950',
  '3953',
  '3958',
  '3959',
  '3961',
  '3962',
  '3963',
  '3967',
  '3968',
  '14877',
  '8610',
  '15932',
  '18668',
  '10228',
  '11066',
  '11079',
  '9322',
  '11125',
  '11124',
  '4489',
  '7996',
  '12868',
  '12902',
  '8620',
  '8772',
  '8790',
  '9352',
  '10290',
  '12124',
  '14525','16499') and
  DATE_CREATED BETWEEN TO_DATE( '" , todate[w,1] ,"', 'YYYY-MON-DD')
  AND TO_DATE('" , todate[w+1,1] ,"', 'YYYY-MON-DD')
  group by ENTITY.Name
  order by ENTITY.Name" ))

write.csv(lastyear,paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/2018/",Sys.Date(),"dotcom18W.csv"),row.names = FALSE)

##calculate
list18 <- list.files("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/2018/")
list19 <- list.files("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/2019/")
this <- list19[length(list19)]
last <- list19[length(list19)-1]
this18 <- list18[length(list18)]
this <- read.csv(paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/2019/",this))
last <- read.csv(paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/2019/",last))
this18 <-read.csv(paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/2018/",this18))


names(this)[-1] <- paste0("y19w",w,names(this)[-1])
names(last)[-1] <- paste0("y19w",w-1,names(last)[-1])
names(this18)[-1] <- paste0("y18w",w,names(this18)[-1])
result <- merge(this18,last,by="name",all.x = TRUE, all.y = TRUE)
result <- merge(result,this,by="name",all.x = TRUE, all.y = TRUE)
for (i in 2:ncol(result)){
  result[,i] <- as.numeric(as.character(result[,i]))
  result[,i] <- abs(result[,i])
}

result2 <- result
result2[is.na(result2)] <- 0

total <- as.data.frame(t(colSums(result2[,-1])))
total$name <- "TOTAL"
result <- rbind(result,total)

WF <- c("Castlegate - CA Perris",
        "Castlegate - CA Perris SP",
        "Castlegate - CAN Toronto",
        "Castlegate - GA McDonough",
        "Castlegate - KY Erlanger",
        "Castlegate - KY Florence",
        "Castlegate - KY Hebron",
        "Castlegate - NJ Cranbury",
        "Castlegate - NJ Cranbury LP",
        "Castlegate - TX Lancaster SP",
        "JossandMain.com",
        "Wayfair.com"
)

CG <- c("Castlegate - CA Perris",
        "Castlegate - CA Perris SP",
        "Castlegate - GA McDonough",
        "Castlegate - KY Erlanger",
        "Castlegate - KY Florence",
        "Castlegate - KY Hebron",
        "Castlegate - NJ Cranbury",
        "Castlegate - NJ Cranbury LP",
        "Castlegate - TX Lancaster SP")

os <- c("Overstock.Com",
        "Overstock.Com - 3PL (NEW)"
)



wflist <- filter(result,name%in%WF)
wflist[is.na(wflist)] <- 0
wflist <- as.data.frame(t(colSums(wflist[,-1])))
wflist$name <- "Wayfair"


CGlist <- filter(result,name%in%CG)
CGlist[is.na(CGlist)] <- 0
CGlist <- as.data.frame(t(colSums(CGlist[,-1])))
CGlist$name <- "Castlegate"

oslist <- filter(result,name%in%os)
oslist[is.na(oslist)] <- 0
oslist <- as.data.frame(t(colSums(oslist[,-1])))
oslist$name <- "Overstock"

result <- rbind(result,wflist,CGlist,oslist)

result$YoYqty <- (result[,8]-result[,2])/result[,2]
result$YoYamount <- (result[,9]-result[,3])/result[,3]
result$YoYorder <- (result[,10]-result[,4])/result[,4]

result$WoWqty <- (result[,8]-result[,5])/result[,5]
result$WoWamount <- (result[,9]-result[,6])/result[,6]
result$WoWorder <- (result[,10]-result[,7])/result[,7]









########################################################


tillnow19<-sqlQuery(channel,paste0(
  "select
  ENTITY.Name as name,
  sum(TRANSACTION_LINES.ITEM_COUNT) as 'qty',
  sum(TRANSACTION_LINES.AMOUNT) as 'amount',
  count(DISTINCT TRANSACTION_LINES.TRANSACTION_ID) as 'order'
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
  '5559',
  '9926',
  '17569',
  '9519',
  '20010',
  '5788',
  '20567',
  '16499',
  '18380',
  '18615',
  '5387',
  '21336',
  '5376',
  '20587',
  '5933',
  '5453',
  '22361',
  '5568',
  '20011',
  '22701',
  '7346',
  '15590',
  '8513',
  '8571',
  '9445',
  '6328',
  '7190',
  '3948',
  '3950',
  '3953',
  '3958',
  '3959',
  '3961',
  '3962',
  '3963',
  '3967',
  '3968',
  '14877',
  '8610',
  '15932',
  '18668',
  '10228',
  '11066',
  '11079',
  '9322',
  '11125',
  '11124',
  '4489',
  '7996',
  '12868',
  '12902',
  '8620',
  '8772',
  '8790',
  '9352',
  '10290',
  '12124',
  '14525',
  '16499') and
  DATE_CREATED BETWEEN TO_DATE( '2019-Jan-1', 'YYYY-MON-DD')
  AND TO_DATE('" , Sys.Date()-2 ,"', 'YYYY-MM-DD')
  
  group by ENTITY.Name
  order by ENTITY.Name" ))

write.csv(tillnow19,paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/p2019/",Sys.Date(),w,"dotcom19W.csv"),row.names = FALSE)


tillnow18<-sqlQuery(channel,paste0(
  "select
  ENTITY.Name as name,
  sum(TRANSACTION_LINES.ITEM_COUNT) as 'qty',
  sum(TRANSACTION_LINES.AMOUNT) as 'amount',
  count(DISTINCT TRANSACTION_LINES.TRANSACTION_ID) as 'order'
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
  '5559',
  '9926',
  '17569',
  '9519',
  '20010',
  '5788',
  '20567',
  '16499',
  '18380',
  '18615',
  '5387',
  '21336',
  '5376',
  '20587',
  '5933',
  '5453',
  '22361',
  '5568',
  '20011',
  '22701',
  '7346',
  '15590',
  '8513',
  '8571',
  '9445',
  '6328',
  '7190',
  '3948',
  '3950',
  '3953',
  '3958',
  '3959',
  '3961',
  '3962',
  '3963',
  '3967',
  '3968',
  '14877',
  '8610',
  '15932',
  '18668',
  '10228',
  '11066',
  '11079',
  '9322',
  '11125',
  '11124',
  '4489',
  '7996',
  '12868',
  '12902',
  '8620',
  '8772',
  '8790',
  '9352',
  '10290',
  '12124',
  '14525',
  '16499') and
  DATE_CREATED BETWEEN TO_DATE( '2018-Jan-1', 'YYYY-MON-DD')
  AND TO_DATE('" , Sys.Date()-2 -365 ,"', 'YYYY-MM-DD')
  
  group by ENTITY.Name
  order by ENTITY.Name" ))

write.csv(tillnow18,paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/p2018/",Sys.Date(),w,"dotcom18W.csv"),row.names = FALSE)





###calc2
list18 <- list.files("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/p2018/")
list19 <- list.files("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/p2019/")
this <- list19[length(list19)]
this18 <- list18[length(list18)]
this <- read.csv(paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/p2019/",this))
this18 <-read.csv(paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/p2018/",this18))


names(this)[-1] <- paste0("y19tow",w,names(this)[-1])
names(this18)[-1] <- paste0("y18tow",w,names(this18)[-1])
resultp <- merge(this18,this,by="name",all.x = TRUE, all.y = TRUE)
for (i in 2:ncol(resultp)){
  resultp[,i] <- as.numeric(as.character(resultp[,i]))
  resultp[,i] <- abs(resultp[,i])
}

result2 <- resultp
result2[is.na(result2)] <- 0

total <- as.data.frame(t(colSums(result2[,-1])))
total$name <- "TOTAL"
resultp <- rbind(resultp,total)

WF <- c("Castlegate - CA Perris",
        "Castlegate - CA Perris SP",
        "Castlegate - CAN Toronto",
        "Castlegate - GA McDonough",
        "Castlegate - KY Erlanger",
        "Castlegate - KY Florence",
        "Castlegate - KY Hebron",
        "Castlegate - NJ Cranbury",
        "Castlegate - NJ Cranbury LP",
        "Castlegate - TX Lancaster SP",
        "JossandMain.com",
        "Wayfair.com"
)

CG <- c("Castlegate - CA Perris",
        "Castlegate - CA Perris SP",
        "Castlegate - GA McDonough",
        "Castlegate - KY Erlanger",
        "Castlegate - KY Florence",
        "Castlegate - KY Hebron",
        "Castlegate - NJ Cranbury",
        "Castlegate - NJ Cranbury LP",
        "Castlegate - TX Lancaster SP")

os <- c("Overstock.Com",
        "Overstock.Com - 3PL (NEW)"
)



wflist <- filter(resultp,name%in%WF)
wflist[is.na(wflist)] <- 0
wflist <- as.data.frame(t(colSums(wflist[,-1])))
wflist$name <- "Wayfair"


CGlist <- filter(resultp,name%in%CG)
CGlist[is.na(CGlist)] <- 0
CGlist <- as.data.frame(t(colSums(CGlist[,-1])))
CGlist$name <- "Castlegate"

oslist <- filter(resultp,name%in%os)
oslist[is.na(oslist)] <- 0
oslist <- as.data.frame(t(colSums(oslist[,-1])))
oslist$name <- "Overstock"

resultp <- rbind(resultp,wflist,CGlist,oslist)

resultp$YoYqtyall <- (resultp[,5]-resultp[,2])/resultp[,2]
resultp$YoYamountall <- (resultp[,6]-resultp[,3])/resultp[,3]
resultp$YoYorderall <- (resultp[,7]-resultp[,4])/resultp[,4]

result <- merge(result,resultp,by="name",all.x=TRUE)

###calc2
write.csv(result,paste0("N:/E Commerce/Public Share/Dot Com Weekly Evaluation -Invoice/Evaluation/",Sys.Date(),"W",w,"WOW.csv"),row.names = FALSE)

}

