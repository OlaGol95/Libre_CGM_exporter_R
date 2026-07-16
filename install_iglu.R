options(repos = c(CRAN = "https://cloud.r-project.org"))

required_packages <- c("iglu")
missing_packages <- required_packages[
  !vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)
]

if (length(missing_packages) > 0) {
  install.packages(missing_packages)
}

invisible(lapply(required_packages, library, character.only = TRUE))
message("R package check complete: ", paste(required_packages, collapse = ", "))
