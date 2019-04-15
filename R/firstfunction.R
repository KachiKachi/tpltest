
#' change
#'
#' @param d
#' @param s
#' @param j
#' @param t
#'
#' @return
#' @export
#'
#' @examples
change <- function(d, s,j,t) {
  attach(d)
  i=eval(parse(text = j))
  t[i]=s[i]
  }

# #Ctrl+Shift+Alt+R
# devtools::document()
# devtools::use_package("dplyr")
# Author@R: person("Yao", "Guan", email = "guanyaociqi@gmail.com",
#                   role = c("aut", "cre"))
