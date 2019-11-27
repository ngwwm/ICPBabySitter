cd E:\Lemon\Workspace\VisualStudio\ICPBabySitter\ICPBabySitter
copy App.config.bak web.config
C:\Windows\Microsoft.NET\Framework\v4.0.30319\aspnet_regiis.exe -pef "connectionStrings" .
C:\Windows\Microsoft.NET\Framework\v4.0.30319\aspnet_regiis.exe -pef "appSettings" .
copy web.config App.config