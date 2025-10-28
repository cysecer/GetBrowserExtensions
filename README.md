# GetBrowserExtensions
Script to get all Browser Extensions of MS Edge and Chrome and append them to a CSV.
Usage: 
- 1: change $CSVPath to the correct .csv destination. Make sure to use network drive and user which will run it has read/write permissions in case script is ran as a startupscript, to ensure csv is saved to one central location.
- 2: ./GetBrowserExtensions.ps1 -CSV "C:\path\to\file\export.csv"
