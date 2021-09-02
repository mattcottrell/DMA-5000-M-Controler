# DMA-5000-M-Controler
Graphical User Interface Software Coordinating DMA Instrument Control and Data Curation In Google Sheets

Two files are included here:

DMA_5000_M_Controler.xojo_binary_project

This is source code for a Xojo desktop project that relies on the Chilkat API for communicating with Google sheets.  The sofware presents a graphical user interface for collecting daily brewery cellar data using an Anton Paar DMA 500 M density meter and an Alcolyzer module that measures alcohol content.  Data are saved to a cellar log file in Google sheets. A separate sheet summarizes the cellar log in a format that presents the current state of fermentation for all of the vessels in the cellar.

CompletedCellarWork.xojo_binary_project

This is source code is for a Xojo web app that works in the brewery cellar can use to update cellar log data from a mobile phone.
