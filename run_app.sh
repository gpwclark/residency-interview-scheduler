#!/usr/bin/env bash
# Run the Shiny app directly without RStudio

HOST_NAME=${HOST_NAME:-"127.0.0.1"}
HOST_PORT=${HOST_PORT:-"0"}

export HOST_NAME
export HOST_PORT

Rscript run_app.R
