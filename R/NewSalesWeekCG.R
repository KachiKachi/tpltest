##update TLinv first
#'
#' Title NewSalesWeekCG    
#' 
#' Description This file is based on Jessy's original slow seller report. Will tell you howmany weeks the inventory can feed for each items.
#' And will tell you the result for current inventory, the inventroy after 30 days and the inventory after 60 days. 
#' You can also find CG's feeding week.
#' All the sales week are calculate base on either the recent one mont's sales record or 6 month's sales record.
#' You can also find the sales speed and in-stock-rate in the output report.
#' 
#' @param nsid Netsuit ID
#' @param nspwd Netsuit password
#' 
#'
#' @return a csv file in N:/E Commerce/Public Share/Dot Com - Promotion indicator/
#' @export CGmerge a report shows howmany weeks the inventory will last
#'
#' @examples
SaleWeekCG <- function(nsid,nspwd){


library(xlsx)
library(dplyr)
library(data.table)
library(RODBC)
library(stringr)
library(readxl)
library(lubridate)

channel <- odbcConnect("NetSuite", uid=nsid, pwd=pwd)
RPL <- read.csv("C:/Users/yao.Guan/Desktop/usefulData/RPL.csv")

# get speed of last month --sp ###############################################################################
m <- as.numeric(format(Sys.Date(), "%m"))
m=m-1
#m=m-1
a <- ymd("2019-01-01") %m+% months(0:11)
b <- ymd("2019-01-31") %m+% months(0:11)


new<-sqlQuery(channel,paste0(
  "select
  ITEMS.ITEM_ID as name,
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
  locations.LOC_ID in ('CA-F','IL-S') and
  ENTITY.ENTITY_ID not in (
  '20567',
  '21336',
  '20587',
  '22701',
  '14877',
  '10228',
  '11125',
  '11124',
  '10290',
  '12124',
  '14525'
  ) and
  DATE_LAST_MODIFIED BETWEEN TO_DATE( '" , a[m] ,"', 'YYYY-MM-DD')
  AND TO_DATE('" , b[m]+1 ,"', 'YYYY-MM-DD')
  group by ITEMS.ITEM_ID" ))


kit2 <- sqlQuery(channel, 
                 " select ITEMS.ITEM_ID, Name,WAYFAIR_COM_PARTNER_SKU, Items.WEIGHT as CUFT,MOQ_PRODUCTION_CG,MOQ_BY_LOCATION_CG,
                 Items.FTY_CODE, Items.COUNTRY_OF_MANUFACTURE as Kmanufacturer ,mname,ctns,mMOQ,
                 mNumber,mLMOQ,FactoryCode,mm.MANUFACTURER ,mm.COUNTRY_OF_MANUFACTURE , PORT_OF_ORIGIN.LIST_ITEM_NAME,mm.Port
                 
                 From Items
                 Join ITEM_STATUS on  STATUS_ID = ITEM_STATUS.LIST_ID
                 left Join PORT_OF_ORIGIN on Items.PORT_OF_ORIGIN_ID = PORT_OF_ORIGIN. LIST_ID
                 left Join 
                 (select PARENT_ID, Name as mname, Item_ID as mid,
                 ii.CTNSPC as ctns, MOQ_PRODUCTION_CG as mMOQ,MOQ_BY_LOCATION_CG as mLMOQ,
                 ii.FTY_CODE as FactoryCode,MANUFACTURER,COUNTRY_OF_MANUFACTURE,
                 PORT_OF_ORIGIN_ID as pid,quantity as mNumber,ii.LIST_ITEM_NAME as Port
                 from item_group 
                 left Join (Select *  from Items left join PORT_OF_ORIGIN  
                 on Items.PORT_OF_ORIGIN_ID = PORT_OF_ORIGIN. LIST_ID) ii
                 on item_group.MEMBER_ID = ii.Item_ID
                 ) mm 
                 on Items. Item_ID = mm.PARENT_ID
                 
                 where ITEMS.ISINACTIVE = 'No'
                 ")

kit <- kit2[,c("ITEM_ID","Name","CUFT","mname","ctns","mNumber")]
kit$Name <- as.character(kit$Name)
kit$mname <- as.character(kit$mname)
i <- is.na(kit$mname)
kit$mname[i] <- kit$Name[i]
kit$ctns[i] <- 1
kit$mNumber[i] <- 1
speedm <- merge(kit,new, by.x = "ITEM_ID",by.y = "name", all.x = TRUE)
speedm$qty <- abs(speedm$qty)
speedm$amount <- abs(speedm$amount)
i <- is.na(speedm$qty)
speedm$qty[i] <- 0
speedm$amount[i] <- 0
speedm$mqty <- speedm$mNumber*speedm$qty
speedm <- merge(speedm,RPL,by.x = "mname", by.y = "Name",all.x = TRUE)

speedm$X <- as.character(speedm$X)
i <- !is.na(speedm$X)
speedm$mname[i] <- speedm$X[i]
sapply(speedm, class)

sp <- aggregate(mqty~mname, data = speedm,sum)
sp <- filter(sp,mqty!=0)
m1week <- as.numeric( b[m]-a[m])/7
sp$mqty <- sp$mqty/m1week
colnames(sp) <- c("mname", "spm"  )

# get basehead ###############################################################################################
rplbasehead <- select(speedm,Name,mname,CUFT)


# get inventory ###############################################################################################

bgfolder="N:/E Commerce/Public Share/Dot Com - Wayfair/InvLists"
a1= list.files(paste0(bgfolder))
b1=a1[length(a1)]
bg <- read_xlsx(paste0(bgfolder,"/",b1),sheet=1)
bg <- filter(bg[,c("Name","Type","Total")],Type=="Inventory Item"&Total!=0)
bg <- merge(bg,RPL,by.x = "Name", by.y = "Name",all.x = TRUE,all.y = FALSE)
bg$X <- as.character(bg$X)
i <- !is.na(bg$X)
bg$Name[i] <- bg$X[i]
bg <- aggregate(Total~Name,data = bg,sum)


# get speed of last 6 month ###################################################################################

c <- ymd("2018-08-01") %m+% months(0:24)
d <- ymd("2019-01-31") %m+% months(0:11)

new<-sqlQuery(channel,paste0(
  "select
  ITEMS.ITEM_ID as name,
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
  locations.LOC_ID in ('CA-F','IL-S') and
  ENTITY.ENTITY_ID not in (
  '20567',
  '21336',
  '20587',
  '22701',
  '14877',
  '10228',
  '11125',
  '11124',
  '10290',
  '12124',
  '14525'
  ) and
  DATE_LAST_MODIFIED BETWEEN TO_DATE( '" , c[m] ,"', 'YYYY-MM-DD')
  AND TO_DATE('" , d[m]+1 ,"', 'YYYY-MM-DD')
  
  group by ITEMS.ITEM_ID" ))

kit <- kit2[,c("ITEM_ID","Name","CUFT","mname","ctns","mNumber")]
kit$Name <- as.character(kit$Name)
kit$mname <- as.character(kit$mname)
i <- is.na(kit$mname)
kit$mname[i] <- kit$Name[i]
kit$ctns[i] <- 1
kit$mNumber[i] <- 1
speedm <- merge(kit,new, by.x = "ITEM_ID",by.y = "name", all.x = TRUE)
speedm$qty <- abs(speedm$qty)
speedm$amount <- abs(speedm$amount)
i <- is.na(speedm$qty)
speedm$qty[i] <- 0
speedm$amount[i] <- 0
speedm$mqty <- speedm$mNumber*speedm$qty
speedm <- merge(speedm,RPL,by.x = "mname", by.y = "Name",all.x = TRUE)

speedm$X <- as.character(speedm$X)
i <- !is.na(speedm$X)
speedm$mname[i] <- speedm$X[i]
sapply(speedm, class)

sp6 <- aggregate(mqty~mname, data = speedm,sum)
sp6 <- filter(sp6,mqty!=0)
m6week <- as.numeric( d[m]-c[m])/7
sp6$mqty <- sp6$mqty/m6week
colnames(sp6) <- c("mname", "spm6"  )
# get instock rate ##############################################################################################
m <- as.numeric(format(Sys.Date(), "%m"))
m=m-1
a <- ymd("2019-01-01") %m+% months(0:11)
b <- ymd("2019-01-31") %m+% months(0:11)
c <- ymd("2018-08-01") %m+% months(0:24)
d <- ymd("2019-01-31") %m+% months(0:11)

tlinv <- read.csv("N:/E Commerce/Public Share/Dot Com - Wayfair/Sales predict/Increasing inventory/fileTL88auto.csv")
sapply(tlinv,class)
i=is.na(tlinv)
tlinv[i] <- 0
a[m]
b[m]
c[m]
d[m]

invday <- colnames(tlinv)

mlist <- ymd(a[m]) %m+% days(0:as.numeric( b[m]-a[m]))
m6list <- ymd(c[m]) %m+% days(0:as.numeric( d[m]-c[m]))

mlist <- as.character(mlist)
mlist <- paste0("X",str_sub(mlist,6,7),str_sub(mlist,-2,-1),str_sub(mlist,1,4))
mlist <- intersect(mlist,invday)
minv <- tlinv[,mlist]
i <- minv<5
minv[i] <- 0
i <- minv!=0
minv[i] <- 1

minv <- cbind(tlinv$Item, minv)
names(minv)[1] <- "Item"
minv <- merge(minv,RPL,by.x = "Item",by.y = "Name", all.x = TRUE)
minv$Item <- as.character(minv$Item)
minv$X <- as.character(minv$X)
i <- !is.na(minv$X)
minv$Item[i] <- minv$X[i]
minv <- minv[,c(1:ncol(minv)-1)]
minv <- aggregate(.~Item,data = minv,max)
md <- ncol(minv)-1
minv <- 
  minv %>%
  mutate(smrate = rowSums(.[2:ncol(minv)]))
msrate <- select(minv,Item,smrate)
msrate$smrate <- msrate$smrate/md


m6list <- as.character(m6list)
m6list <- paste0("X",str_sub(m6list,6,7),str_sub(m6list,-2,-1),str_sub(m6list,1,4))
m6list <- intersect(m6list,invday)
m6inv <- tlinv[,m6list]
i <- m6inv<5
m6inv[i] <- 0
i <- m6inv!=0
m6inv[i] <- 1


m6inv <- cbind(tlinv$Item, m6inv)
names(m6inv)[1] <- "Item"
m6inv <- merge(m6inv,RPL,by.x = "Item",by.y = "Name", all.x = TRUE)
m6inv$Item <- as.character(m6inv$Item)
m6inv$X <- as.character(m6inv$X)
i <- !is.na(m6inv$X)
m6inv$Item[i] <- m6inv$X[i]
m6inv <- m6inv[,c(1:(ncol(m6inv)-1))]
m6inv <- aggregate(.~Item,data = m6inv,max)
m6d <- ncol(m6inv)-1
m6inv <- 
  m6inv %>%
  mutate(sm6rate = rowSums(.[2:ncol(m6inv)]))
m6srate <- select(m6inv,Item,sm6rate)
m6srate$sm6rate <- m6srate$sm6rate/m6d


# merge for current iventory ##############################################################################
mergec <- merge(rplbasehead,bg, by.x = "mname",by.y = "Name", all.x = TRUE)
names(mergec)[4] <- "InvToday"
mergec <- merge(mergec,sp,by="mname",all.x = TRUE)
mergec <- merge(mergec,sp6,by="mname",all.x = TRUE)
mergec <- merge(mergec,msrate,by.x ="mname", by.y = "Item", all.x = TRUE, all.y = FALSE)
mergec <- merge(mergec,m6srate,by.x ="mname",by.y = "Item", all.x = TRUE, all.y = FALSE)
mergec$spm <- mergec$spm/mergec$smrate
mergec$spm6 <- mergec$spm6/mergec$sm6rate


head(m6srate)
head(mergec)
dim(mergec)

# get next month income#####################################################################################

income30<-sqlQuery(channel,paste0(
  "select
  ITEMS.NAME as name,
  
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
  TRANSACTIONS.TRANSACTION_TYPE='Purchase Order' and 
  TRANSACTIONS.STATUS='Pending Receipt' and 
  ENTITY.Name <> 'Homelegance Inc. (AGA Domestic Po)' and
  locations.LOC_ID in ('CA-F','IL-S') and
  ENTITY.ENTITY_ID not in (
  '20567',
  '21336',
  '20587',
  '22701',
  '14877',
  '10228',
  '11125',
  '11124',
  '10290',
  '12124',
  '14525'
  ) and
  TRANSACTION_LINES.RECEIVEBYDATE BETWEEN TO_DATE( '" , a[m]-100 ,"', 'YYYY-MM-DD')
  AND TO_DATE('" , Sys.Date() +20 ,"', 'YYYY-MM-DD')
  
  group by ITEMS.NAME" ))

colnames(income30) <- c('name', 'qty30', 'amount30')
merge30 <- merge(mergec,income30, by.x = "mname",by.y = "name", all.x = TRUE)
head(merge30)

income60<-sqlQuery(channel,paste0(
  "select
  ITEMS.NAME as name,
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
  TRANSACTIONS.TRANSACTION_TYPE='Purchase Order' and 
  TRANSACTIONS.STATUS='Pending Receipt' and 
  ENTITY.Name <> 'Homelegance Inc. (AGA Domestic Po)' and
  locations.LOC_ID in ('CA-F','IL-S') and
  ENTITY.ENTITY_ID not in (
  '20567',
  '21336',
  '20587',
  '22701',
  '14877',
  '10228',
  '11125',
  '11124',
  '10290',
  '12124',
  '14525'
  ) and
  TRANSACTION_LINES.RECEIVEBYDATE BETWEEN TO_DATE( '" , Sys.Date() +20 ,"', 'YYYY-MM-DD')
  AND TO_DATE('" , Sys.Date()+50 ,"', 'YYYY-MM-DD')
  
  group by ITEMS.NAME" ))

colnames(income60) <- c('name', 'qty60', 'amount60')
merge60 <- merge(merge30,income60, by.x = "mname",by.y = "name", all.x = TRUE)

######calculate weeks
mergeall <- select(merge60,-amount30,-amount60)
mergeall$qty30[is.na(mergeall$qty30)] <- 0
mergeall$qty60[is.na(mergeall$qty60)] <- 0
mergeall$InvToday[is.na(mergeall$InvToday )] <- 0
mergeall$spm[is.na(mergeall$spm )] <- 0
mergeall$spm6[is.na(mergeall$spm6 )] <- 0
mergeall$smrate[is.na(mergeall$smrate )] <- 0
mergeall$sm6rate[is.na(mergeall$sm6rate )] <- 0
mergeall$spm[is.infinite(mergeall$spm )] <- 0
mergeall$spm6[is.infinite(mergeall$spm6 )] <- 0

mergeall$weekonM <- mergeall$InvToday/mergeall$spm
mergeall$weekon6M <- mergeall$InvToday/mergeall$spm6
mergeall$weekonM[is.na(mergeall$weekonM)] <- 0

Kitm <- aggregate(weekonM~Name,data = mergeall,min)
Kitm6 <- aggregate(weekon6M~Name,data = mergeall,min)
colnames(Kitm) <- c("Name","Kitm")
colnames(Kitm6) <- c("Name","Kitm6")

mergeall <- merge(mergeall,Kitm,by= "Name",all.x = TRUE)
mergeall <- merge(mergeall,Kitm6,by= "Name",all.x = TRUE)
sapply(mergeall, class)

attach(mergeall)
i <- spm<spm6
mergeall$incomeweek30 <- qty30/spm
mergeall$incomeweek30[i] <- qty30[i]/spm6[i]
j <- qty30==0
mergeall$incomeweek30[j] <- 0

mergeall$incomeweek60 <- qty60/spm
mergeall$incomeweek60[i] <- qty60[i]/spm6[i]
j <- qty60==0
mergeall$incomeweek60[j] <- 0

i <- spm<spm6
mergeall$incomeweek30 <- qty30/spm
mergeall$incomeweek30[i] <- qty30[i]/spm6[i]
mergeall$incomeweek60 <- qty60/spm
mergeall$incomeweek60[i] <- qty60[i]/spm6[i]
i <- smrate<sm6rate
mergeall$incomeweek30[i] <- qty30[i]/spm6[i]
mergeall$incomeweek60[i] <- qty60[i]/spm6[i]
i <- smrate>sm6rate
mergeall$incomeweek30[i] <- qty30[i]/spm[i]
mergeall$incomeweek60[i] <- qty60[i]/spm[i]



j <- qty30==0
mergeall$incomeweek30[j] <- 0
j <- qty60==0
mergeall$incomeweek60[j] <- 0




i <- spm<spm6
mergeall$selectM <- weekonM
mergeall$selectM[i] <- weekon6M[i]
i <- smrate<sm6rate
mergeall$selectM[i] <- weekon6M[i]
i <- smrate>sm6rate
mergeall$selectM[i] <- weekonM[i]

i <- mergeall$selectM<2
mergeall$incomebase <- mergeall$selectM-2
mergeall$incomebase[i] <- 0
mergeall$week30 <- mergeall$incomebase+mergeall$incomeweek30

i <- mergeall$week30 < 4 & !is.na(mergeall$week30)
mergeall$week60 <- mergeall$week30+mergeall$incomeweek60 - 4
mergeall$week60[i] <- mergeall$incomeweek60[i]


Kitweek30 <- aggregate(week30~Name, data = mergeall,min)
colnames(Kitweek30) <- c(  'Name'   , 'kitweek30')
Kitweek60 <- aggregate(week60~Name, data = mergeall,min)
colnames(Kitweek60) <- c(  'Name'   , 'kitweek60')

mergeall <- merge(mergeall,Kitweek30,by = "Name", all.x = TRUE)
mergeall <- merge(mergeall,Kitweek60,by = "Name", all.x = TRUE)


selectkit <- aggregate(selectM~Name, data = mergeall,min)
colnames(selectkit) <- c("Name","SelectMKit")
mergeall <- merge(mergeall,selectkit,by = "Name", all.x = TRUE)

#######CG sale number
sapply(mergeall, class)
head(mergeall)
colnames(mergeall)
mergeall


###tab OOS
i <- mergeall$InvToday ==0
mergeall$weekonM[i] <- 'OOS'
mergeall$ weekon6M[i] <- 'OOS'
mergeall$ Kitm[i] <- 'OOS'
mergeall$ Kitm6[i] <- 'OOS'
mergeall$SelectMKit[i] <- "OOS"
i <- mergeall$InvToday + mergeall$incomeweek30 == 0
mergeall$ week30[i] <- 'OOS'
mergeall$ kitweek30[i] <- 'OOS'
i <- mergeall$InvToday==0 & mergeall$incomeweek30 < 2
mergeall$ week30[i] <- 'OOS'
mergeall$ kitweek30[i] <- 'OOS'

i <- mergeall$InvToday + mergeall$incomeweek60+ mergeall$incomeweek30 == 0
mergeall$ week60[i] <- 'OOS'
mergeall$ kitweek60[i] <- 'OOS'
i <- mergeall$InvToday==0 & mergeall$incomeweek30 < 4 & mergeall$incomeweek60<2
mergeall$ week60[i] <- 'OOS'
mergeall$ kitweek60[i] <- 'OOS'

i <- mergeall$kitweek30==0
i <- mergeall$kitweek60==0
mergeall$kitweek30[i] <- "OOS"
mergeall$kitweek60[i] <- "OOS"
i <- mergeall$Kitm==0
i <- mergeall$Kitm6==0
mergeall$Kitm[i] <- "OOS"
mergeall$Kitm6[i] <- "OOS"



###mergeRPLname
mergeall <- merge(mergeall,RPL,by="Name",all.x = TRUE)
mergeall$X <- as.character(mergeall$X)
i=is.na(mergeall$X)
mergeall$X[i] <- mergeall$Name[i]

###merge p name
pname <- sqlQuery(channel, "select Name, WAYFAIR_COM_PARTNER_SKU as PartName From Items")
mergeall <- merge(mergeall,pname,by.x ="X",by.y="Name",all.x = TRUE,all.y = FALSE)


#####finalname
mergeall <- select(mergeall,Name = Name ,RPLName=X, PartName ,Member_name = mname, CUFT, Inventory_Today= InvToday,
                   Speed_Month=spm, Speed_6Month=spm6, Instockrate_LastMonth=smrate,      
                   Instockrate_Last6Month=sm6rate,  
                   Week_baseonLastmonth=weekonM, Week_baseonLast6month=weekon6M, 
                   KitWeek_baseonLastmonth=Kitm, KitWeek_baseonLast6month = Kitm6,  
                   Select_Week=selectM, KitSelect_Week=SelectMKit,
                   Qty_Next30days=qty30,Qty_Next60days=qty60, 
                   Week_Next30days=incomeweek30, Week_Next60days=incomeweek60,
                   incomebase,  Week30thday=week30, Week60thday=week60, Kitweek30thday=kitweek30,
                   Kitweek60thday=kitweek60) 


###################CG inv
CGmerge <- mergeall

WH <- list.files("N:/E Commerce/Public Share/Dot Com - Wayfair/CG daily Inventory")
dat <- WH[length(WH)-2]
file <- read.csv(paste0("N:/E Commerce/Public Share/Dot Com - Wayfair/CG daily Inventory/",dat))
file$CGinbount <- file$Qty.In.Forward.Transit+file$Qty.Inbound
CGinv <- aggregate(cbind(Total.Qty.In.Warehouse,Qty.On.Hold,CGinbount)~Part.Number, data = file, sum)
colnames(CGinv) <- c("PartName","CGinv","CGhold","CGinbount")
CGmerge <- merge(CGmerge,CGinv,by="PartName",all.x = TRUE,all.y = TRUE)


###wfspeed1month
new<-sqlQuery(channel,paste0(
  "select
  ITEMS.Name as Name,
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
  '14525'
  ) and
  DATE_LAST_MODIFIED BETWEEN TO_DATE( '" , a[m] ,"', 'YYYY-MM-DD')
  AND TO_DATE('" , b[m]+1 ,"', 'YYYY-MM-DD')
  
  group by ITEMS.Name" ))


newu=new
newu$Name <- as.character(newu$Name)
newu <- merge(newu,RPL, by="Name",all.x = TRUE)
newu$X <- as.character(newu$X)
i=is.na(newu$X)
newu$X[i] <- newu$Name[i]
newu <- aggregate(qty~X,newu,sum)


newu$CG1mSpeed=abs(newu$qty/(as.numeric(b[m]-a[m])/7))
WF1mspeed <- select(newu,Name=X,CG1mSpeed)

head(WF1mspeed)

###wfspeed6month
new<-sqlQuery(channel,paste0(
  "select
  ITEMS.Name as Name,
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
  '14525'
  ) and
  DATE_LAST_MODIFIED BETWEEN TO_DATE( '" , c[m] ,"', 'YYYY-MM-DD')
  AND TO_DATE('" , d[m]+1 ,"', 'YYYY-MM-DD')
  
  group by ITEMS.Name" ))


newu=new
newu$Name <- as.character(newu$Name)
newu <- merge(newu,RPL, by="Name",all.x = TRUE)
newu$X <- as.character(newu$X)
i=is.na(newu$X)
newu$X[i] <- newu$Name[i]
newu <- aggregate(qty~X,newu,sum)


newu$CG6mSpeed=abs(newu$qty/(as.numeric(d[m]-c[m])/7))
WF6mspeed <- select(newu,Name=X,CG6mSpeed)

head(WF6mspeed)

###CGmerge+speed
CGmerge <- merge(CGmerge,WF1mspeed,by.x = "RPLName",by.y = "Name",all.x = TRUE,all.y = TRUE)
CGmerge <- merge(CGmerge,WF6mspeed,by.x = "RPLName",by.y = "Name",all.x = TRUE,all.y = TRUE)


#####get cg+TPL stockrate ##############################################################################################
m <- as.numeric(format(Sys.Date(), "%m"))
m=m-1
a <- ymd("2019-01-01") %m+% months(0:11)
b <- ymd("2019-01-31") %m+% months(0:11)
c <- ymd("2018-08-01") %m+% months(0:24)
d <- ymd("2019-01-31") %m+% months(0:11)

tlinv <- read.csv("N:/E Commerce/Public Share/Dot Com - Wayfair/Sales predict/Increasing inventory/fileAll88auto.csv")
sapply(tlinv,class)
i=is.na(tlinv)
tlinv[i] <- 0
a[m]
b[m]
c[m]
d[m]

invday <- tolower( colnames(tlinv))
colnames(tlinv) <- invday

mlist <- ymd(a[m]) %m+% days(0:as.numeric( b[m]-a[m]))
m6list <- ymd(c[m]) %m+% days(0:as.numeric( d[m]-c[m]))

mlist <- as.character(mlist)
mlist <- paste0("x",str_sub(mlist,1,4),str_sub(mlist,6,7),str_sub(mlist,-2,-1))
mlist <- intersect(mlist,invday)
minv <- tlinv[,mlist]
i <- minv<5
minv[i] <- 0
i <- minv!=0
minv[i] <- 1

minv <- cbind(tlinv$item, minv)
names(minv)[1] <- "Item"
minv <- merge(minv,RPL,by.x = "Item",by.y = "Name", all.x = TRUE)
minv$Item <- as.character(minv$Item)
minv$X <- as.character(minv$X)
i <- !is.na(minv$X)
minv$Item[i] <- minv$X[i]
minv <- minv[,c(1:ncol(minv)-1)]
minv <- aggregate(.~Item,data = minv,max)
md <- ncol(minv)-1
minv <- 
  minv %>%
  mutate(smrate = rowSums(.[2:ncol(minv)]))
msrate <- select(minv,RPLName=Item,Allmrate=smrate)
msrate$Allmrate <- as.numeric(as.character(msrate$Allmrate))
msrate$Allmrate <- msrate$Allmrate/md


m6list <- as.character(m6list)
m6list <- paste0("x",str_sub(m6list,1,4),str_sub(m6list,6,7),str_sub(m6list,-2,-1))
m6list <- intersect(m6list,invday)
m6inv <- tlinv[,m6list]
i <- m6inv<5
m6inv[i] <- 0
i <- m6inv!=0
m6inv[i] <- 1


m6inv <- cbind(tlinv$item, m6inv)
names(m6inv)[1] <- "Item"
m6inv <- merge(m6inv,RPL,by.x = "Item",by.y = "Name", all.x = TRUE)
m6inv$Item <- as.character(m6inv$Item)
m6inv$X <- as.character(m6inv$X)
i <- !is.na(m6inv$X)
m6inv$Item[i] <- m6inv$X[i]
m6inv <- m6inv[,c(1:(ncol(m6inv)-1))]
m6inv <- aggregate(.~Item,data = m6inv,max)
m6d <- ncol(m6inv)-1
m6inv <- 
  m6inv %>%
  mutate(sm6rate = rowSums(.[2:ncol(m6inv)]))

m6srate <- select(m6inv,RPLName=Item,Allm6rate=sm6rate)
m6srate$Allm6rate <- as.numeric(as.character(m6srate$Allm6rate))
m6srate$Allm6rate <- m6srate$Allm6rate/m6d

CGmerge <- merge(CGmerge,msrate,by="RPLName",all.x = TRUE)
CGmerge <- merge(CGmerge,m6srate,by="RPLName",all.x = TRUE)

###get week number
sapply(CGmerge, class)

i <- is.na(CGmerge$CGinv)
CGmerge$CGinv[i] <- 0

i <- is.na(CGmerge$CGhold)
CGmerge$CGhold[i] <- 0

i <- is.na(CGmerge$CGinbount)
CGmerge$CGinbount[i] <- 0

i <- is.na(CGmerge$CG1mSpeed)
CGmerge$CG1mSpeed[i] <- 0

i <- is.na(CGmerge$CG6mSpeed)
CGmerge$CG6mSpeed[i] <- 0

i <- is.na(CGmerge$Allmrate)
CGmerge$Allmrate[i] <- 0

i <- is.na(CGmerge$Allm6rate)
CGmerge$Allm6rate[i] <- 0


CGmerge$CGohselectWeek <- CGmerge$CGinv/(CGmerge$CG1mSpeed/CGmerge$Allmrate)
CGmerge$CGcomselectWeek <- CGmerge$CGinbount/(CGmerge$CG1mSpeed/CGmerge$Allmrate)

i <- CGmerge$Allmrate<CGmerge$Allm6rate
i[is.na(i)] <- FALSE
CGmerge$CGohselectWeek[i] <- CGmerge$CGinv[i]/(CGmerge$CG6mSpeed[i]/CGmerge$Allm6rate[i])
CGmerge$CGcomselectWeek[i] <- CGmerge$CGinbount[i]/(CGmerge$CG6mSpeed[i]/CGmerge$Allm6rate[i]) 

i <- is.nan(CGmerge$CGohselectWeek)
CGmerge$CGohselectWeek[i] <- 0
i <- is.nan(CGmerge$CGcomselectWeek)
CGmerge$CGcomselectWeek[i] <- 0

CGmerge$allcgweek <- CGmerge$CGohselectWeek+CGmerge$CGcomselectWeek

####write
write.csv(CGmerge,paste0("N:/E Commerce/Public Share/Dot Com - Promotion indicator/", Sys.Date(),"(CG)Estimate sale week.csv"))
paste0("N:/E Commerce/Public Share/Dot Com - Promotion indicator/", Sys.Date(),"(CG)Estimate sale week.csv")
}
