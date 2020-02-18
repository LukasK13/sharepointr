#' List available SharePoint lists
#'
#' This method lists all available lists for a
#' SharePoint connection.
#'
#' @param con A SharePoint connection returned
#' by sp_connection()
#' @param raw Return response raw formatted. Default is
#' FALSE
#'
#' @return SharePoint response as list (raw = TRUE) or
#' the available lists as character vector (raw = FALSE).
#' @export
sp_getLists <- function(con, raw = F) {
  response = sp_request(con, "lists")
  if (raw) {
    return(response)
  }

  names = sp_changeEscaping(if (con$Office365) response$content$value$Title else response$content$d$results$Title)
  ids = sp_changeEscaping(if (con$Office365) response$content$value$Id else response$content$d$results$Id)
  names(ids) = names
  return(ids)
}

#' Get metadata of a SharePoint list
#'
#' This method lists all available metadata for a
#' SharePoint list.
#'
#' @param con A SharePoint connection
#' @param listName Name of the SharePoint list
#' @param raw Return response raw formatted. Default is
#' FALSE
#'
#' @return SharePoint response as list (raw = TRUE) or
#' the decoded response content as list (raw = FALSE).
#' @export
sp_getListMetadata <- function(con, listName = NULL, listID = NULL, raw = F) {
  if ((is.null(listName) && is.null(listID)) || (!is.null(listName) && !is.null(listID))) stop("Either listName or listID must be provided")
  request = URLencode(paste0("lists/", if (!is.null(listName)) paste0("getbytitle('", listName) else paste0("getbyid('", listID), "')"))
  response = sp_request(con, request)
  return(if (raw) response else (if (con$Office365) response$content else response$content$d))
}

#' List available SharePoint list columns
#'
#' This method lists all available columns of a
#' SharePoint list.
#'
#' @param con A SharePoint connection
#' @param listName Name of the SharePoint list
#' @param listID ID of the SharePoint list
#' @param raw Return response raw formatted. Default is
#' FALSE
#'
#' @return SharePoint response as list (raw = TRUE) or
#' the available columns as character vector (raw = FALSE).
#' @export
sp_getListColumns <- function(con, listName = NULL, listID = NULL, raw = F, hidden = F) {
  if ((is.null(listName) && is.null(listID)) || (!is.null(listName) && !is.null(listID))) stop("Either listName or listID must be provided")
  request = URLencode(paste0("lists/", if (!is.null(listName)) paste0("getbytitle('", listName) else paste0("getbyid('", listID), "')/fields", if (!hidden) "?$filter=Hidden eq false and ReadOnlyField eq false"))
  response = sp_request(con, request)
  return(if (raw) response else sp_changeEscaping(if (con$Office365) response$content$value$Title else response$content$d$results$Title))
}

#' Read data from a SharePoint list
#'
#' This method allows to retrieve all data from a
#' SharePoint list.
#'
#' @param con A SharePoint connection
#' @param listName Name of the SharePoint list
#' @param expand Retrieve data by using deferred tags
#' (takes longer, but lists more results)
#' @export
sp_readListData <- function(con, listName = NULL, listID = NULL, expand = F) {
  if ((is.null(listName) && is.null(listID)) || (!is.null(listName) && !is.null(listID))) stop("Either listName or listID must be provided")
  response = sp_getListColumns(con, listName = listName, listID = listID, raw = T, hidden = F)
  if (response$status_code == 200) {
    columnNamesInternal = if (con$Office365) response$content$value$InternalName[!response$content$value$FromBaseType | response$content$value$InternalName == "Title"] else response$content$d$results$InternalName
    columnNames = if (con$Office365) response$content$value$Title[!response$content$value$FromBaseType | response$content$value$InternalName == "Title"] else response$content$d$results$Title
    response = sp_request(con, URLencode(paste0("lists/", if (!is.null(listName)) paste0("getbytitle('", listName) else paste0("getbyid('", listID), "')/items")))
    if (response$status_code == 200) {
      data = data.frame()
      repeat({
        if (expand && !is.null(unname(unlist(if (con$Office365) response$content$value$FieldValuesAsText else response$content$d$results$FieldValuesAsText)))) {
          items = unname(unlist(if (con$Office365) response$content$value$FieldValuesAsText else response$content$d$results$FieldValuesAsText))
          data_temp = Reduce(rbind, lapply(items, function(item) {
            response = sp_request(con, item)
            if (response$status_code == 200) {
              names(response$content$d) = gsub("_x005f", "", names(response$content$d))
              data = as.data.frame(t(data.frame(unlist(response$content$d[columnNamesInternal]))))
              rownames(data) = NULL
              colnames(data) = columnNames[columnNamesInternal %in% colnames(data)]
              return(data)
            }
          }))
        } else {
          data_temp = as.data.frame(if (con$Office365) response$content$value else response$content$d$results)
          colnames(data_temp) = gsub("^OData_", "", colnames(data_temp))
          data_temp = data_temp[, columnNamesInternal]
          colnames(data_temp) = make.names(columnNames[columnNamesInternal %in% colnames(data_temp)])
        }
        data = if (nrow(data) == 0) data_temp else rbind(
          data.frame(c(data, sapply(colnames(data_temp)[!colnames(data_temp) %in% colnames(data)], function(x) NA))),
          data.frame(c(data_temp, sapply(colnames(data)[!colnames(data) %in% colnames(data_temp)], function(x) NA)))
        )
        if (!is.null(if(con$Office365) response$content$odata.nextLink else response$content$d$`__next`)) {
          response = sp_request(con, if(con$Office365) response$content$odata.nextLink else response$content$d$`__next`)
          if (response$status_code != 200) stop("Invalid response.")
        } else {
          break
        }
      })
      return(data)
    }
  }
}
