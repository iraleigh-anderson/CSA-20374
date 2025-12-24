
# ---------------- Project .Rprofile ----------------
# It respects paths defined in .Renviron and sets up your library search path.
# 1) Prepend user library (R_LIBS_USER) to library search path if it exists
userlib <- Sys.getenv("R_LIBS_USER")
if (nzchar(userlib) && dir.exists(userlib)) {
  .libPaths(c(userlib, .libPaths()))
}

# 2) Helper: get required environment variables with a clear error if missing
get_env_path <- function(name) {
  val <- Sys.getenv(name)
  if (!nzchar(val)) stop(sprintf("Missing %s in .Renviron", name), call. = FALSE)
  val
}

# 3) Materialize commonly used paths as variables (read-only at startup)
# These lines create R variables in the global environment when the project opens. Instead of calling Sys.getenv() repeatedly.
PEM.path                      <- get_env_path("PEM_PATH")
BEC13.path                    <- get_env_path("BEC13_PATH")
BECMaster.path                <- get_env_path("BECMASTER_PATH")

TEM1.path                     <- get_env_path("TEM1_PATH")
TEM2.path                     <- get_env_path("TEM2_PATH")
TEM3.path                     <- get_env_path("TEM3_PATH")

TEI.Path                      <- get_env_path("TEI_PATH")
TEI.proj.bound.path           <- get_env_path("TEI_PROJ_BOUND_PATH")

CEF.Human.Disturbance.path    <- get_env_path("CEF_HUMAN_DISTURBANCE_PATH")
My.CEF.Human.Disturbance.path <- get_env_path("MY_CEF_HUMAN_DISTURBANCE_PATH")

Working.Folder                <- get_env_path("WORKING_FOLDER")

# 4) Optional: lightweight package loader for interactive sessions
load_pkgs <- function(pkgs) {
  for (p in pkgs) {
    suppressPackageStartupMessages(
      require(p, character.only = TRUE, quietly = TRUE)
    )
  }
}

# Example usage (comment out if you prefer to load in scripts explicitly)
# load_pkgs(c("tidyverse", "sf", "openxlsx", "terra", "DBI", "odbc", "glue", "lwgeom"))

# 5) Friendly message (turn off if you prefer quiet startup)
packageStartupMessage("Project .Rprofile loaded. Libraries: ", paste(.libPaths(), collapse = " | "))
