Overview:

This PowerShell script collects Windows Event Logs for the previous 7 days from specified servers, exports the logs in CSV format, applies Pivot Tables to create an XLSX version of the logs, and generates a summary text file for quick analysis. The script also maintains a log of its operations through a transcript file.

Features:

Transcript Logging: Records the execution process in a transcript file.
Event Log Collection: Retrieves Application, Security, and System logs from specified servers.
Log Export: Exports the event logs into CSV format with UTF-8 encoding.
Pivot Table Creation: Converts CSV files into Excel (XLSX) format with Pivot Tables for better data analysis.
Summary Report: Generates a summary text file containing an overview of event logs per server.
