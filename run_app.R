#!/usr/bin/env Rscript

# Get environment variables or defaults
host_name <- Sys.getenv("HOST_NAME", "127.0.0.1")
host_port <- as.numeric(Sys.getenv("HOST_PORT", "0"))

# Set browser option
options(browser='open')

# If port is 0, find a random available port
if (host_port == 0) {
  # Use httpuv to find a random available port
  if (requireNamespace("httpuv", quietly = TRUE)) {
    host_port <- httpuv::randomPort()
    message(sprintf("Selected random port: %d", host_port))
  } else {
    # Fallback to a common range
    host_port <- sample(3000:8000, 1)
    message(sprintf("Selected random port (fallback): %d", host_port))
  }
}

# Run the app
message(sprintf("Starting Shiny app on http://%s:%d", host_name, host_port))

# runApp will print "Listening on http://..." with the actual port
shiny::runApp('interview_schedule_app.R', 
              port = host_port,
              host = host_name,
              launch.browser = TRUE)