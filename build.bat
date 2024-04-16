if exist app.ico (
  C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /t:winexe open.cs /win32icon:app.ico
) else (
  C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /t:winexe open.cs
)
if not %errorlevel% == 0 (
  pause
)
