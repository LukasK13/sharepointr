#' Establish a connection to SharePoint
#'
#' This method allows to establish and save a connection to
#' SharePoint online or a SharePoint server. The connection
#' properties are therefore saved in a connection object
#' to be later used for the request
#'
#' @param Address Address of the SharePoint instance to be contacted.
#' For example https://example.sharepoint.com
#' @param Username The username to use for the authentication.
#' If not given, a credential file must be provided.
#' For example: john.doe@@example.com
#' @param Password The password to use for the authentication.
#' If not given, a credential file must be provided.
#' @param credentialFile If no username or password is provided,
#' this file will be used for reading the credentials from.
#' The file must be YAML formatted, containing the keys Username
#' and Password.
#' @param Office365 True (default) if a connection to SharePoint online
#' shall be established, False if a connection to SharePoint server
#' shall be established.
#' @param acceptLanguage The locale to be used for retrieving information.
#' Default value is en and can be defined as in
#' https://developer.mozilla.org/en-US/docs/Web/HTTP/Headers/Accept-Language
#' @examples
#' sp_con = sp_connection("https://your.sharepoint.server/subpage1/", "YourUsername", "YourPassword", Office365 = F, acceptLanguage = "en")
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' @export
sp_connection <- function(Address, Username = NULL, Password = NULL, credentialFile = NULL, Office365 = T, acceptLanguage = "en") {
  if (is.null(Username) | is.null(Password)) { # No username or password is given
    if (is.null(credentialFile)) { # No credential file is given
      stop("Not enough arguments.") # stop and show error message
    } else { # credential file is given
      credentials = yaml::read_yaml(credentialFile) # read credential file
      Username = credentials$Username # save username
      Password = credentials$Password # save password
    }
  }
  if (Office365) { # access sharepoint online
    Address_base = regmatches(Address, regexpr("[[:alnum:]]{1,}\\.sharepoint\\.com", Address)) # remove https:// from address
    request = suppressWarnings(readLines(system.file("saml.xml", package = "sharepointr"))) # read XML soap envelope
    request = gsub("\\{Username\\}", Username, request) # paste username into XML form
    request = gsub("\\{Password\\}", Password, request) # paste password into XML form
    request = gsub("\\{Address\\}", Address_base, request) # paste address into XML form
    response = httr::POST(url = "https://login.microsoftonline.com/extSTS.srf", body = request) # request security token from microsoft online
    if (response$status_code != 200) stop("Receiving security token failed.") # Check if request was successful
    content = as_list(read_xml(rawToChar(response$content))) # decode response content
    token = as.character(content$Envelope$Body$RequestSecurityTokenResponse$RequestedSecurityToken$BinarySecurityToken) # extract security token
    response = httr::POST(paste0("https://", Address_base, "/_forms/default.aspx?wa=wsignin1.0"), body = token, httr::add_headers(Host = Address_base)) # post security token to sharepoint online
    if (response$status_code != 200) stop("Receiving access cookies failed.") # Check if request was successful
    cookie = list(rtFa = response$cookies$value[response$cookies$name %in% "rtFa"], FedAuth = response$cookies$value[response$cookies$name %in% "FedAuth"]) # Extract authentication cookies
    con = list(Username = Username, Address = paste0(Address, if (length(grep("/$", Address)) == 1) "_api/" else "/_api/"), Cookie = cookie, Office365 = T, acceptLanguage = acceptLanguage) # create connection object
  } else { # acces sharepoint server
    con = list(Username = Username, Address = paste0(Address, if (length(grep("/$", Address)) == 1) "_api/" else "/_api/"),
               Password = Password, Office365 = F, acceptLanguage = acceptLanguage) # create connection object
  }
  class(con) = "sp_connection" # set class of connection object
  return(con) # return connection object
}

#' Send a request to the SharePoint API
#'
#' Send a request via GET or POST to the
#' SharePoint RESTful API
#'
#' @param con A SharePoint connection returned
#' by sp_connection()
#' @param request A request URL. Can be either a full URL
#' or the last part of the request URL which will be
#' concatenated to the API-URL of the SharePoint connection
#' @param verb Method to use for the request. Either GET
#' (default) or POST.
#' @param json retreive data using JSON formatting
#' @param body Body of the request (only for POST).
#'
#' @return The response as list
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' lists = sp_request(sp_con, "lists", verb = "GET", json = T, body = NULL)
#' @export
#' @import xml2 httr
sp_request <- function(con, request, verb = "GET", json = T, body = NULL) {
  if (!"sp_connection" %in% class(con)) stop("Invalid sharepoint connection.") # Check class of connection object
  request = URLencode(if (length(grep(con$Address, request)) == 1) request else paste0(con$Address, request)) # create valid rquest url and encode it
  if (con$Office365) { # request data from SharePoint online
    if (tolower(verb) == "get") {
      response = httr::GET(request, add_headers(accept = if (json) "application/json;odata=verbose" else "application/atom+xml"), httr::set_cookies(rtFa = con$Cookie$rtFa, FedAuth = con$Cookie$FedAuth)) # send request
    } else if (tolower(verb) == "post") {
      digest = sp_getRequestDigest(con)
      response = httr::POST(request, httr::add_headers(Accept = "application/json;odata=verbose", "X-RequestDigest" = digest,
                                                       "Content-Type" = "application/json;odata=verbose"),
                            httr::set_cookies(rtFa = con$Cookie$rtFa, FedAuth = con$Cookie$FedAuth),
                            body = as.character(toJSON(body), auto_unbox = T))
    } else {
      stop("Unknown verb.")
    }
  } else { # request data from SharePoint server
    if (tolower(verb) == "get") {
      response = httr::GET(request, httr::authenticate(con$Username, con$Password, "ntlm"),
                           httr::add_headers(accept = if (json) "application/json;odata=verbose" else "application/atom+xml", `accept-language` = con$acceptLanguage)) # send request
    } else if (tolower(verb) == "post") {
      if (length(grep("contextinfo$", request)) == 0) {
        digest = sp_getRequestDigest(con)
        response = httr::POST(request, httr::authenticate(con$Username, con$Password, "ntlm"),
                              httr::add_headers(Accept = "application/json;odata=verbose", "X-RequestDigest" = digest, "Content-Type" = "application/json;odata=verbose"),
                              body = as.character(toJSON(body, auto_unbox = T)))
      } else {
        response = httr::POST(request, httr::authenticate(con$Username, con$Password, "ntlm"),
                              httr::add_headers(Accept = "application/json;odata=verbose", "Content-Type" = "application/json;odata=verbose"),
                              body = as.character(toJSON(body, auto_unbox = T)))
      }
    } else {
      stop("Unknown verb.")
    }
  }
  if (!response$status_code %in% c(200, 201, 202, 203, 204, 205, 206, 207, 208, 226)) stop("Status Code ", response$status_code, ": Request failed.\r\n", rawToChar(response$content)) # Check if request was successful
  if (json || tolower(verb) == "post") { # response is JSON formatted
    if (!jsonlite::validate(rawToChar(response$content))) { # response doesn't contains valid JSON
      stop("Request didn't return valid JSON.") # stop with error message
    }
    response$content = fromJSON(rawToChar(response$content)) # convert response content to R list
  } else {
    response$content = xml2::as_list(xml2::read_xml(rawToChar(response$content))) # convert response content to R list
  }
  return(response) # return response
}

#' Change encoding of a SharePoint response
#'
#' Change encoding of a SharePoint response
#'
#' @param encoded An encoded SharePoint response
#'
#' @return SharePoint response with changed encoding
#' @export
sp_changeEscaping <- function(encoded) {
  return(sapply(stringi::stri_replace_all_fixed(encoded, c("<U+", ">"), c("\\u", ""), vectorize_all = FALSE), URLencode, USE.NAMES = F))
}

#' Get a SharePoint request digest
#'
#' Request and receive a SharePoint request
#' digest for further editing
#'
#' @param con A SharePoint connection returned
#' by sp_connection()
#'
#' @return Request digest as string
#' @examples
#' sp_con = sp_connection("https://yourdomain.sharepoint.com", "YourUsername", "YourPassword", Office365 = T)
#' digest = sp_getRequestDigest(sp_con)
#' @export
sp_getRequestDigest <- function(con) {
  response = sp_request(con, "contextinfo", verb = "POST") # Request request digest
  digest = response$content$d$GetContextWebInformation$FormDigestValue # Extract request digest
  return(digest) # Return request digest
}
