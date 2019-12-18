#' Link a SharePoint connection to a list
#'
#' This method links a SharePoint connection to a
#' SharePoint list
#'
#' @param con A SharePoint connection returned
#' by sp_connection()
#' @param listName Name of the SharePoint list
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList")
#' @export
sp_list <- function(con, listName = NULL, listID = NULL) {
  if (!"sp_connection" %in% class(con)) stop("Invalid sharepoint connection.") # Check class of connection object
  if ((is.null(listName) && is.null(listID)) || (!is.null(listName) && !is.null(listID))) stop("Either listName or listID must be provided")
  response = sp_getListColumns(con, listName = listName, listID = listID, raw = T, hidden = T)
  columnNamesInternal = if (con$Office365) response$content$value$InternalName else response$content$d$results$InternalName # Extract internal column names
  columnNames = if (con$Office365) response$content$value$Title else response$content$d$results$Title # Extract external column names
  table = list(con = con, columns = list(columnNamesInternal = columnNamesInternal, columnNames = columnNames), listName = listName, listID = listID, op = list()) # Build list connection object
  class(table) = "sp_list_connection" # Set class
  return(table) # Return list connection object
}

#' Apply filter operations on a SharePoint list
#'
#' This method allows to apply filter operations
#' on a SharePoint list connection. The syntax is
#' inspired by dplyr
#'
#' @param table A SharePoint list connection
#' as returned by sp_list()
#' @param ... Comma separated filter commands
#' white spaces in filters can be escaped using ``.
#' For quoting, !! is used.
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList") %>% sp_filter(Title == "yourTitle")
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList") %>% sp_filter(`Your column` == "yourTitle")
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList") %>% sp_filter(Title == !!yourVariable)
#' @return Modfied SharePoint list connection
#' @export
sp_filter <- function(table, ..., .filter = NULL) {
  if (!"sp_list_connection" %in% class(table)) stop("Invalid SharePoint list.") # Check class of list connection object
  if (is.null(.filter)) {
    filters = lapply(match.call(expand.dots = FALSE)$`...`, function(x) {as.character(bquote(.(x)))}) # Decompose filter command
    if (length(filter) == 0) stop("No filters provided.")
    command = lapply(filters, function(command) { # loop through all filter commands
      command = gsub("`", "", command) # Remove column quoting
      if (length(grep("^!!", command[2])) == 1) { # command is quoted
        command[2] = eval(parse(text = gsub("^!!", "", command[2])), envir = parent.frame(n=3)) # evaluate command
      }
      if (length(grep("^!!", command[3])) == 1) { # command is quoted
        command[3] = eval(parse(text = gsub("^!!", "", command[3])), envir = parent.frame(n=3)) # evaluate command
      }
      if (command[2] %in% table$columns$columnNames) { # command refers to a column name
        command[2] = head(table$columns$columnNamesInternal[table$columns$columnNames %in% command[2]], n = 1) # translate column name to internal column names
      } else { # Command refers to a constant
        command[2] = paste0("'", command[2], "'") # quote constant
      }
      if (command[3] %in% table$columns$columnNames) { # command refers to a column name
        command[3] = head(table$columns$columnNamesInternal[table$columns$columnNames %in% command[3]], n = 1) # translate column name to internal column names
      } else { # Command refers to a constant
        command[3] = paste0("'", command[3], "'") # quote constant
      }
      command = sp_buildFilter(command)
      return(command)
    })
    table$op[[length(table$op) + 1]] = URLencode(paste0("$filter=", paste0("(", unlist(command), ")", collapse = "and"))) # Add final filter command to the list of operations
  } else {
    table$op[[length(table$op) + 1]] = URLencode(paste0("$filter=", .filter)) # Add final filter command to the list of operations
  }
  return(table) # return tabel conneciton object
}

sp_buildFilter <- function(command) {
  if (command[1] == "==") { # Build command for URL
    command = paste(command[2], "eq", command[3])
  } else if (command[1] == "!=") {
    command = paste(command[2], "ne", command[3])
  } else if (command[1] == "<") {
    command = paste(command[2], "lt", command[3])
  } else if (command[1] == "<=") {
    command = paste(command[2], "le", command[3])
  } else if (command[1] == ">") {
    command = paste(command[2], "gt", command[3])
  } else if (command[1] == ">=") {
    command = paste(command[2], "ge", command[3])
  } else if (command[1] == "startswith") {
    command = paste0("startswith(", command[2], ",", command[3], ")")
  } else if (command[1] == "substringof") {
    command = paste0("substringof(", command[2], ",", command[3], ")")
  } else {
    stop("Unknown operator: ", command[1])
  }
  return(command)
}

#' Select columns from a SharePoint list
#'
#' This method allows to select columns from
#' a SharePoint list
#'
#' @param table A SharePoint list connection
#' as returned by sp_list()
#' @param ... Comma separated select commands.
#' White spaces can be escaped using ``
#'
#' @return Modfied SharePoint list connection
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList") %>% sp_select(Title, `Column with whitespaces`)
#' @export
sp_select <- function(table, ..., .select = NULL) {
  if (!"sp_list_connection" %in% class(table)) stop("Invalid SharePoint list.")
  if (is.null(.select)) {
    selects = as.character(match.call(expand.dots = FALSE)$`...`)
  } else {
    selects = .select
  }
  selects = gsub("`", "", selects)
  if (!all(selects %in% table$columns$columnNames)) stop("Unknown variable selected")
  selects = table$columns$columnNamesInternal[table$columns$columnNames %in% selects]
  table$op[[length(table$op) + 1]] = URLencode(paste0("$select=", paste0(selects, collapse = ",")))
  return(table)
}

#' Arrange a SharePoint list
#'
#' This method allows to arrange a SharePoint list
#'
#' @param table A SharePoint list connection
#' as returned by sp_list()
#' @param ... Comma separated arrange commands
#'
#' @return Modfied SharePoint list connection
#' @examples sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList") %>% sp_arrange(Title, desc(column2))
#' @export
sp_arrange <- function(table, ..., .arrange = NULL) {
  if (!"sp_list_connection" %in% class(table)) stop("Invalid SharePoint list.")
  if (is.null(.arrange)) {
    orders = as.character(match.call(expand.dots = FALSE)$`...`)
    descending = grep("^desc(.{0,})$", orders)
    if (length(descending) > 0) {
      orders[descending] = gsub("^desc\\(", "", gsub("\\)$", "", orders[descending]))
    }
    orders = gsub("`", "", orders)
    if (all(orders %in% table$columns$columnNames)) {
      orders = table$columns$columnNamesInternal[table$columns$columnNames %in% orders]
    } else {
      stop("Unknown variable selected")
    }
    order = rep("", length(orders))
    order[descending] = " desc"
    table$op[[length(table$op) + 1]] = URLencode(paste0("$orderby=", paste0(orders, order, collapse = ",")))
  } else {
    table$op[[length(table$op) + 1]] = URLencode(paste0("$orderby=", .arrange))
  }
  return(table)
}

#' Explain a SharePoint list request
#'
#' This method allows to explain a SharePoint
#' list request
#'
#' @param table A SharePoint list connection
#'
#' @return A SharePoint list connection
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList") %>% sp_select(Title, `Column with whitespaces`) %>% sp_explain
#' @export
sp_explain <- function(table) {
  if (!"sp_list_connection" %in% class(table)) stop("Invalid SharePoint list.")
  request = URLencode(paste0("lists/", if (!is.null(table$listName)) paste0("getbytitle('", table$listName) else paste0("getbyid('", table$listID), "')/items"))
  if (length(table$op) > 0) {
    request = paste0(request, "?", paste0(unlist(table$op), collapse = "&"))
  }
  print(request)
  return(table)
}

#' Collect a SharePoint list
#'
#' This method allows to collect the results
#' of a SharePoint list request
#'
#' @param table A SharePoint list connection
#' as returned by sp_list()
#' @param n Number of rows to collect
#' @param skip Number of rows to skip
#' @param expand Follow deferred links
#' @param verbose Print status information
#'
#' @return Request result as data frame
#' @examples sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_list = sp_list(sp_con, "yourList") %>% sp_select(Title, `Column with whitespaces`) %>% sp_collect()
#' @export
sp_collect = function(table, n = Inf, skip = NULL, expand = T, verbose = T) {
  if (!"sp_list_connection" %in% class(table)) stop("Invalid SharePoint list.")
  request = URLencode(paste0("lists/", if (!is.null(table$listName)) paste0("getbytitle('", table$listName) else paste0("getbyid('", table$listID), "')/items"))
  if (!is.infinite(n)) {
    table$op[[length(table$op) + 1]] = paste0("$top=", n)
  }
  if (!is.null(skip)) {
    table$op[[length(table$op) + 1]] = paste0("$skiptoken=", skip)
  }
  if (length(table$op) > 0) {
    request = paste0(request, "?", paste0(unlist(table$op), collapse = "&"))
  }
  if (!verbose) print(request)
  response = sp_request(table$con, request)
  if (response$status_code == 200) {
    data = data.frame()
    repeat({
      if (is.null(if (table$con$Office365) response$content$value$FieldValuesAsText else response$content$d$results$FieldValuesAsText) || !expand) {
        cols = names(if (table$con$Office365) response$content$value else response$content$d$results)[which(unname(unlist(lapply(if (table$con$Office365) response$content$value else response$content$d$results, typeof))) %in% c("character", "numeric", "integer", "double", "logical"))]
        data_temp = as.data.frame(if (table$con$Office365) response$content$value[cols] else response$content$d$results[cols])
        colnames(data_temp) = gsub("OData_", "", colnames(data_temp))
        data_temp = data_temp[,colnames(data_temp) %in% table$columns$columnNamesInternal]
        for (col in colnames(data_temp)) {
          if (col %in% table$columns$columnNamesInternal) {
            colnames(data_temp)[colnames(data_temp) == col] = table$columns$columnNames[table$columns$columnNamesInternal %in% col]
          }
        }
      } else {
        items = unname(unlist(if (table$con$Office365) response$content$value$FieldValuesAsText else response$content$d$results$FieldValuesAsText))
        data_temp = Reduce(rbind, lapply(items, function(item) {
          response = sp_request(table$con, item)
          if (response$status_code == 200) {
            if (table$con$Office365) {
              names(response$content$value) = gsub("_x005f", "", names(response$content$value))
            } else {
              names(response$content$d) = gsub("_x005f", "", names(response$content$d))
            }
            data = as.data.frame(t(data.frame(unlist(if (table$con$Office365) response$content$value[table$columns$columnNamesInternal] else response$content$d[table$columns$columnNamesInternal]))))
            rownames(data) = NULL
            colnames(data) = table$columns$columnNames[table$columns$columnNamesInternal %in% colnames(data)]
            return(data)
          }
        }))
      }
      data = rbind(data, data_temp)
      if (!is.infinite(n) && nrow(data) >= n) break
      if (!is.null(if(table$con$Office365) response$content$odata.nextLink else response$content$d$`__next`)) {
        response = sp_request(table$con, if(table$con$Office365) response$content$odata.nextLink else response$content$d$`__next`)
        if (response$status_code != 200) stop("Invalid response.")
      } else {
        break
      }
    })
    return(data)
  }
}
