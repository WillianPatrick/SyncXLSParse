# SyncXLSParse
.NET C# console application for import sheets from XLS and XLSX files from parametrized directory path to Parse Server Platform (Back4App) for transcript to tables, columns headers and cells rows.  

Tests performed with spreadsheets containing over 600,000 rows, average processing time of 80 rows per second multiplied by upload rate. Example: 80 x 10mb (upload) = 800 lines per second.  

Notes: I'm not considering the processing limits of the client or server hardware, which can be higher or lower depending on the case, this information is just the averages I got from my experience in testing in developer mode!  Thanks ;)
