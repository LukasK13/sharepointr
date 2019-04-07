#' Retrieve SharePoint users
#'
#' Retrieve all SharePoint users of the given site
#'
#' @param con A SharePoint connection returned by
#' sp_connection()
#' @param ... Filters
#' @param filter A Character string containing
#' valid SharePoint filters (for example:
#' LoginName eq Username1)
#'
#' @return User information as a dataframe
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' user = sp_users(sp_con, LoginName = "Username1")
#' @export
sp_users <- function(con, ..., .filter = NULL) {
  if (!"sp_connection" %in% class(con)) stop("Invalid sharepoint connection.") # Check class of connection object
  filters = lapply(match.call(expand.dots = FALSE)$`...`, function(x) {as.character(bquote(.(x)))}) # Decompose filter command
  if (length(filters) > 0) {
    command = lapply(filters, sp_buildFilter) # Build SharePoint filters
    .filter = paste0(unlist(command), collapse = "&") # Add final filter command to the list of operations
  }
  request = paste0("web/siteusers", if (!is.null(.filter)) paste0("?$filter=", URLencode(.filter))) # Conactenate request URL
  response = sp_request(con, request) # Request list column names
  data = as.data.frame(response$content$d$results[which(unname(unlist(lapply(response$content$d$results, typeof))) != "list")]) # Convert response to dataframe
  return(data) # Return list connection object
}
