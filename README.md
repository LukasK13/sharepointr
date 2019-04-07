# sharepointr
A R package for reading from and writing to SharePoint lists.

## Installation
`devtools::install_github("lukask13/sharepointr")`

## Connect to a SharePoint server
Connect to a local SharePoint server:

`sp_con = sp_connection("https://your.sharepoint.server/subpage1/", "YourUsername", "YourPassword", Office365 = F)`

Connect to a Office365 SharePoint server:

`sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)`

## Read from Sharepoint
### List all available lists
The results can be returned either as raw list or as character vector.
```
sp_con = sp_connection("https://your.sharepoint.server/subpage1/", "YourUsername", "YourPassword", Office365 = F)
lists = sp_getLists(sp_con)
```

### Get list metainfo
A list can be either adressed by name or by the list ID. The result can be obtained as raw list or as parsed list.
```
sp_con = sp_connection("https://your.sharepoint.server/subpage1/", "YourUsername", "YourPassword", Office365 = F)
lists = sp_getListMetadata(sp_con, "yourList")
```

### Get list columns
A list can be either adressed by name or by the list ID. It is possible to obtain only names of visible columns.
```
sp_con = sp_connection("https://your.sharepoint.server/subpage1/", "YourUsername", "YourPassword", Office365 = F)
columns = sp_getListColumns(sp_con, "yourList")
```

### Read data from list
A list can be either adressed by name or by the list ID. It is possible to allow expanding of deferred tags.
```
sp_con = sp_connection("https://your.sharepoint.server/subpage1/", "YourUsername", "YourPassword", Office365 = F)
data = sp_readListData(sp_con, "yourList")
```

### Pipeline tools for reading
The pipeline toos for reading from a SharePoint list are inspired by the amazing work of Hadley Wickham and his package dplyr.
Therefore, the syntax is similar to the syntax of dplyr.
```
sp_con = sp_connection("https://your.sharepoint.server/subpage1/", "YourUsername", "YourPassword", Office365 = F)
sp_list = sp_list(sp_con, "yourList") %>% sp_filter(yourColumn1 = "yourValue") %>% sp_select(yourColumn2) %>% sp_arrange(yourColumn2) %>% sp_collect()
```

### Read user information
Obtaining and filtering user information is possible by:
```
sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
user = sp_users(sp_con, LoginName = "Username1")
```

## Write to a SharePoint list
New data can be written to a SharePoint list by:
```
sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
sp_writeListData(sp_con, "yourList", list(column1 = "content1", column2 = 2, column3 = 1.23, user = "Username"))
