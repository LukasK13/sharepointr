# listName = "Leckageprüfung (Beta)"
# data = list(`Stack-ID` = "NM5-222-280", Title = "Test",
#                   `p(t = 0 s) AIR` = 1200, `p(t = 60 s) AIR` = 1100, `p(t = 120 s)` = 1050,
#                   `p(t = 0 s) H2` = 1200, `p(t = 60 s) H2` = 1150, `p(t = 120 s) H2` = 1050,
#                   `p(t = 0 s) KM` = 1200, `p(t = 60 s) KM` = 1200, `p(t = 120 s) KM` = 900,
#                   `p(t = 0 s) AIR_H2` = 1200, `p(t = 60 s) AIR_H2` = 1150, `p(t = 120 s) AIR_H2` = 1050)

#' Write data to a SharePoint list
#'
#' This method allows to write data to a SharePoint
#' list. Therefore, type checks and lookups will be
#' performed (see details).
#'
#' @param con A SharePoint connection returned
#' by sp_connection()
#' @param listName Name of the SharePoint list
#' to write to
#' @param data List or dataframe of the data to
#' write into the SharePoint list
#' @details In order to upload the data correctly,
#' the following data types are casted automatically
#' \enumerate{
#'   \item Text columns
#'   \item Numeric columns
#'   \item User names
#'   \item DateTime columns
#'   \item Lookup columns linking to another table
#' }
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' sp_writeListData(sp_con, "yourList", list(column1 = "content1", column2 = 2, column3 = 1.23, user = "Username"))
#' @export
sp_writeListData <- function(con, listName = NULL, listID = NULL, data) {
  if (!"sp_connection" %in% class(con)) stop("Invalid sharepoint connection.") # Check class of connection object
  if ((is.null(listName) && is.null(listID)) || (!is.null(listName) && !is.null(listID))) stop("Either listName or listID must be provided")
  meta = sp_getListMetadata(con, listName = listName, listID = listID) # Collect metadata of the list
  fields = sp_getListColumns(con, listName = listName, listID = listID, raw = T)$content$d$results # collect field information of the list
  if (!all(names(data) %in% fields$Title)) stop("Unknown columns provided.") # Check if all provided columns can be translated
  names(data) = unlist(lapply(names(data), function(x) { # Loop through all column names of the provided data
    return(fields$InternalName[fields$Title %in% x]) # Return translated internal name
  }))
  if (!all(fields$InternalName[fields$Required] %in% names(data))) stop("Not all required fields provided.") # Check if all required fields are provided
  if (any(fields$InternalName[fields$ReadOnlyField] %in% names(data))) stop("Cannot write read only field.") # Check if no read-only field is provided
  for (x in names(data)) { # Loop through all column names of the provided data
    type = fields$TypeAsString[fields$InternalName %in% x] # Get type of the column
    if (type == "Text") { # Text column
      if (!is.character(data[[x]])) stop("Cannot coerce to type character.") # Check if value is of type character
    } else if (type == "Number") { # Numeric column
      if (any(is.na(as.numeric(data[[x]])))) stop("Cannot coerce to numeric type.") # Check if value is of type numeric
    } else if (type == "User") { # User column
      if (typeof(data[[x]]) != "numeric") { # User ID is not given
        if (typeof(data[[x]]) == "character") { # Value is of type character
          if (length(grep("[[:alpha:]]{3}[[:digit:]]{4}", data[[x]])) == 1) { # Login name is provided
            user = sp_users(con, .filter = paste0("LoginName eq '", data[[x]], "'")) # Filter user ID
          } else { # User name is provided
            user = sp_users(con, .filter = paste0("Title eq '", data[[x]], "'")) # Filter user ID
          }
          if (nrow(user) == 1) { # User ID was found
            data[[x]] = user$Id # Replace value by user ID
            names(data)[names(data) %in% x] = paste0(x, "Id") # Add Id to column name
          } else { # User ID not found
            stop("Cannot find user ", data[[x]])
          }
        } else stop("Cannot coerce to type user.")
      }
    } else if (type == "DateTime") { # Date Time column
      if (!is.character(data[[x]])) stop("Cannot coerce to type character.") # Check if column is of type character
    } else if (type == "Lookup") { # Lookup column
      lookupList = fields$LookupList[which(names(data) %in% x)] # Extract lookup list
      lookupField = fields$LookupField[which(names(data) %in% x)] # Extract lookup column
      ID = sp_list(con, listID = lookupList) %>% sp_filter(.filter = paste0(lookupField, " eq '", data[[x]], "'")) %>%
        sp_select(.select = "ID") %>% sp_collect() # Request lookup ID
      if (nrow(ID) != 1) stop("Cannot resolve lookup field") # Lookup ID wasn't found
      data[[x]] = ID$ID # Replace value by lookup ID
      names(data)[names(data) %in% x] = paste0(x, "Id") # Add Id to column name
    }
  }
  data = as.list(data) # Convert data to list
  data$`__metadata` = list(type = meta$ListItemEntityTypeFullName) # Add metadata tag to list
  request = URLencode(paste0("lists/", if (!is.null(listName)) paste0("getbytitle('", listName) else paste0("getbyid('", listID), "')/items")) # Concatenate request URL
  sp_request(con, request = request, verb = "POST", body = data) # Post data to SharePoint list
}
