# DMA-5000-M-Controler
Graphical User Interface Software Coordinating DMA Instrument Control and Data Curation In Google Sheets

Three files are included here:

1- DMA_5000_M_Controler.xojo_binary_project

This is source code for a Xojo desktop project that relies on the Chilkat API for communicating with Google sheets.  The sofware presents a graphical user interface for collecting daily brewery cellar data using an Anton Paar DMA 500 M density meter and an Alcolyzer module for measuring alcohol content.  Data are saved to a cellar log file in Google sheets. A separate Google sheet summarizes the cellar log in a format that presents the current state of fermentation for all of the vessels in the cellar.


2- CompletedCellarWork.xojo_binary_project

This is source code for a Xojo web app that enables workers in the brewery cellar to update cellar log data manually from a mobile phone.


3- Archive Cellar Log - September 12 2021.gs

This Google Script archives a cellar log to google drive and places an empty cellar log in its place for the next batch that will be brewed into the fermentor.
