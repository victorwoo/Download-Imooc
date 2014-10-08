Download-Imooc 1.7.0 （2014-10-08)
==================================

这是一个从 http://www.imooc.com 教学网站批量下载视频的 PowerShell 脚本。默认下载的是最高清晰度的视频。

按课程专辑 URL 下载
---------------
您可以传入课程专辑的 URL 作为下载参数：

    .\Download-Imooc.ps1 http://www.imooc.com/learn/197

按课程专辑 ID 下载
------------------
可以一口气传入多个课程专辑的 ID 作为参数：

    .\Download-Imooc.ps1 75,197

自动续传
--------
如果不传任何参数的话，将在当前文件夹中搜索已下载的课程，并自动续传。

    .\Download-Imooc.ps1

自动合并视频
------------
如果希望自动合并所有视频，请使用 `-Combine` 参数。该参数可以和其它参数同时使用。

	.\Download-Imooc.ps1 -Combine

若希望自动合并之后删除原分段视频，请使用 `-RemoveOriginal` 参数

    .\Download-Imooc.ps1 -Combine -RemoveOriginal

关于
----
代码中用到了参数分组、`-WhatIf` 处理、语义化注释等技术，供参考。