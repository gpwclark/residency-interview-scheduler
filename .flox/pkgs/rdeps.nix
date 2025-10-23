{ pkgs ? import <nixpkgs> {} }:

pkgs.rWrapper.override {
  packages = with pkgs.rPackages; [
    shiny
    DT
    readxl
    dplyr
    lubridate
    shinythemes
    shinyWidgets
    tidyr
    ggplot2
  ];
}