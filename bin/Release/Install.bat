"C:\Windows\Microsoft.NET\Framework\v4.0.30319\installutil.exe" /u "ClientsNowService.exe"
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\installutil.exe" "ClientsNowService.exe"
copy /y *.sql "C:\Avenca\ClientsNowService"
copy /y *.htm* "C:\Avenca\ClientsNowService"
net start ClientsNowService
timeout 5