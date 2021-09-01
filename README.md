# SyncXLSParse
.NET C# console application for import sheets from XLS and XLSX files from parametrized directory path to Parse Server Platform (Back4App) for transcript to tables, columns headers and cells rows.  
<br/>
<br/>
Tests performed with spreadsheets containing over 600,000 rows, average processing time of 80 rows per second multiplied by upload rate. Example: 80 x 10mb (upload) = 800 lines per second.  
<br/>
<b>
Notes: I'm not considering the processing limits of the client or server hardware, which can be higher or lower depending on the case, this information is just the averages I got from my experience in testing in developer mode!  Thanks ;)
</b>
<br/>
Usage:
<code>
SyncXLSParse.exe -ApplicationId "Application ID" -Server "https://parseapi.back4app.com/" -Key ".NET KEY" -username "Name User Registred" -password "***" -XlsSyncFolderPath "Full path directory" -RowsBuffer 2000
</code>
<br/>
Parameters:
<br/>
<ul>
  <li>-ApplicationId : string</li>
  <li>-Server : string</li>
  <li>-Key : string</li>
  <li>-username : string</li>
  <li>-password : string</li>
  <li>-XlsSyncFolderPath : string</li>
  <li>-RowsBuffer : integer</li> 
</ul>
<br/>
<b>
Notes: This aplication version only increment data on Parse database, for update or delete necessary manual actions on parse dashboard.
</b>
