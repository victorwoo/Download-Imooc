@ECHO OFF
powershell -NoProfile -ExecutionPolicy Unrestricted .\Download-Imooc.ps1 171, 75, 197, 203, 9, 207, 186 -Combine -RemoveOriginal
powershell -NoProfile -ExecutionPolicy Unrestricted .\Download-Imooc.ps1 156 -Combine
PAUSE