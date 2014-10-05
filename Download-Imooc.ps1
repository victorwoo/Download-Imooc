<#
	.SYNOPSIS
		这是一个从 http://www.imooc.com 教学网站批量下载视频的 PowerShell 脚本。

	.DESCRIPTION
		支持合并 *.flv、输出清单、断点续传、更新本地目录等功能。默认下载的是最高清晰度的视频。

	.PARAMETER  Uri
		教程专辑的 URL，例如 'http://www.imooc.com/learn/197'。

	.PARAMETER  ID
		教程专辑的 ID，支持多个，例如 75,197,156,203。

	.PARAMETER  Combine
		自动合并 *.flv 视频。

	.PARAMETER  RemoveOriginal
		删除分段视频。只有在指定 -Combine 开关的情况下才生效。

	.EXAMPLE
		PS> .\Download-Imooc.ps1 http://www.imooc.com/learn/75
		根据教程专辑的 URL 来下载。

	.EXAMPLE
		PS> .\Download-Imooc.ps1 75
		根据教程专辑的 ID 来下载。

	.EXAMPLE
		PS> .\Download-Imooc.ps1 75,197,156,203
		根据教程专辑的 ID 来下载。支持输入多个 ID。

	.EXAMPLE
		PS> .\Download-Imooc.ps1 http://www.imooc.com/learn/75 -Combine -RemoveOriginal
		根据教程专辑的 URL 来下载。完成之后合并所有的视频。

	.EXAMPLE
		PS> .\Download-Imooc.ps1 75,197,156,203 -Combine -RemoveOriginal
		根据教程专辑的 ID 来下载。支持输入多个 ID。完成之后合并所有的视频。，并且删除原始的分段视频。

	.EXAMPLE
		PS> .\Download-Imooc.ps1
		不带任何参数运行该脚本，将自动检测当前目录下的所有下载文件夹，同时检测对应的专辑网站。对于本地缺失的视频或网站更新的视频，将自动续传。

	.NOTES
		若在使用中遇到问题，请联系 victorwoo@gmail.com

	.INPUTS
		None

	.OUTPUTS
		None

	.LINK
		https://github.com/victorwoo/Download-Imooc

	.LINK
		http://blog.vichamp.com/powershell/2014/09/26/download-videos-from-imooc-com-by-powershell/

#>

# Require PowerShell 3.0 or higher.

[CmdletBinding(DefaultParameterSetName='URI', SupportsShouldProcess=$true, ConfirmImpact='Medium')]
Param
(
    [Parameter(ParameterSetName='URI', Position = 0, Mandatory = $false, HelpMessage = '请输入专辑 URL')]
    [string]
    $Uri,

    [Parameter(ParameterSetName='ID', Position = 0, Mandatory = $true, HelpMessage = '请输入专辑 ID，多个 ID 请用逗号隔开')]
    [int[]]
    $ID,

    [Switch]
    $Combine,

    [Switch]
    $RemoveOriginal
)

$DebugPreference = 'Continue' # Continue, SilentlyContinue
# $WhatIfPreference = $true # $true, $false

# 修正文件名，将文件系统不支持的字符替换成“.”
function Get-NormalizedFileName {
    Param (
        $FileName
    )

    [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {
        $FileName = $FileName.Replace($_, '.')
    }

    return $FileName
}

# 修正目录名，将文件系统不支持的字符替换成“.”
function Get-NormalizedFolderName {
    Param (
        $FolderName
    )

    [System.IO.Path]::GetInvalidPathChars() | ForEach-Object {
        $FolderName = $FolderName.Replace($_, '.')
    }

    return $FolderName
}

# 从专辑页面中分析标题和视频页面的 ID。
function Get-Videos {
    Param (
        $Uri
    )
    
    $Uri = $Uri.Replace('/view/', '/learn/')
    $Uri = $Uri.Replace('/qa/', '/learn/')
    $Uri = $Uri.Replace('/note/', '/learn/')
    $Uri = $Uri.Replace('/wiki/', '/learn/')
    $response = Invoke-WebRequest $Uri
    $title = $response.ParsedHtml.title

    echo $title
    $links = $response.Links
    $links | ForEach-Object {
        if ($_.href -cmatch '(?m)^/video/(\d+)$') {
            $id = $Matches[1]
            $title = $_.InnerText
            if ($title -cmatch '(?m)^\d.*? \((?<DURING>\d{2}:\d{2})\)\s*$') {
	            $during = $matches['DURING']
            } else {
	            return
            }
            Write-Debug $during
            $during = [System.TimeSpan]::Parse("00:$during")
            return [PSCustomObject][Ordered]@{
                ID = $id;
                Title = $title;
                During = $during;
            }
        }
    }
}

# 获取视频下载地址。
function Get-VideoUri {
    Param (
        [Parameter(ValueFromPipeline=$true)]
        $ID
    )

    $template = 'http://www.imooc.com/course/ajaxmediainfo/?mid={0}&mode=flash'
    $uri = $template -f $ID
    Write-Debug $uri
    $result = Invoke-RestMethod $uri
    if ($result.result -ne 0) {
        Write-Warning $result.result
    }

    $uri = $result.data.result.mpath.'0'

    # 取最高清晰度的版本。
    $uri = $uri.Replace('L.flv', 'H.flv').Replace('M.flv', 'H.flv')
    return $uri
}

# 创建“.url”快捷方式。
function New-ShortCut {
    Param (
        $Title,
        $Uri
    )

    $shell = New-Object -ComObject 'wscript.shell'
    $dir = pwd
    $path = Join-Path $dir "$Title\$Title.url"
    $lnk = $shell.CreateShortcut($path)
    $lnk.TargetPath = $Uri
    $lnk.Save()
}

# 判断 PowerShell 运行时版本。禁止在低版本的环境运行。
function Assert-PSVersion {
    if ($PSVersionTable.PSVersion.Major -lt 3) {
        Write-Error '请安装 PowerShell 3.0 以上的版本。'
        exit
    }
}

# 获取当前目录下已存在的课程。
function Get-ExistingCourses {
    Get-ChildItem -Directory | ForEach-Object {
        $folder = $_
        $expectedFilePath = (Join-Path $folder $folder.Name) + '.url'
        if (Test-Path -PathType Leaf $expectedFilePath) {
            $shell = New-Object -ComObject 'wscript.shell'
            $lnk = $shell.CreateShortcut($expectedFilePath)
            $targetPath = $lnk.TargetPath
            if ($targetPath -cmatch '(?m)\A^http://www\.imooc\.com/\w+/\d+$\z') {
                echo $targetPath
            }
        }
    }
}

# 输出索引文件。
function Out-IndexFile {
    Param ($title, $uri, $videos, $folderName)
    $filePath = Join-Path $folderName 'info.txt'
    
    $title | Set-Content $filePath -Encoding UTF8
    $uri | Add-Content $filePath -Encoding UTF8

    $global:offset = [System.TimeSpan]::Zero
    $videos | Select-Object -Property @{
        Name = 'Start';
        Expression = {
            $global:offset
        };
    }, @{
       Name = 'End';
       Expression = {
            $global:offset += $_.During
            $global:offset
       };
    }, During, Title |
        Format-Table -AutoSize |
        Out-String |
        Add-Content $filePath -Encoding UTF8
}

# 用 FlvBind.exe 合并视频文件。
function Combine-Videos {
    Param ($folderName, $actualDownloadAny, $outputPathes)

    $targetFile = "$folderName\$folderName.flv"
    #if ($Combine -and ($actualDownloadAny -or -not (Test-Path $targetFile))) {
    if ($Combine) {
        # -and ($actualDownloadAny -or -not (Test-Path $targetFile))) {
        if ($actualDownloadAny -or -not (Test-Path $targetFile) -or (Test-Path $targetFile) -and $PSCmdlet.ShouldProcess('分段视频', '合并')) {
            Write-Progress -Activity '下载视频' -Status '合并视频'    
            Write-Output ("合并视频（共 {0:N0} 个）" -f $outputPathes.Count)
            $outputPathes.Insert(0, $targetFile)
        
            $eap = $ErrorActionPreference
            $ErrorActionPreference = "SilentlyContinue"
            .\FlvBind.exe $outputPathes.ToArray()
            $ErrorActionPreference = $eap

            <#
            $outputPathes = $outputPathes | ForEach-Object {
                "`"$_`""
            }
            Start-Process `
                -WorkingDirectory (pwd) `
                -FilePath .\FlvBind.exe `
                -ArgumentList $outputPathes `
                -NoNewWindow `
                -Wait `
                -ErrorAction SilentlyContinue `
                -WindowStyle Hidden
            #>
            if ($?) {
                Write-Output '视频合并成功'
                if ($RemoveOriginal -and $PSCmdlet.ShouldProcess('分段视频', '删除')) {
                    $outputPathes.RemoveAt(0)
                    $outputPathes | ForEach-Object {
                        Remove-Item $_
                    }
                    Write-Output '原始视频删除完毕'
                }
            } else {
                Write-Warning '视频合并失败'
            }
        }
    }
}

# 下载课程。
function Download-Course {
    Param (
        [string]$Uri
    )

    Write-Progress -Activity '下载视频' -Status '分析视频 ID'
    $courseTitle, $videos = Get-Videos -Uri $Uri
    Write-Output "课程名称：$title"
    Write-Debug $courseTitle
    $folderName = Get-NormalizedFolderName $courseTitle
    Write-Debug $folderName
    if (-not (Test-Path $folderName)) { $null = mkdir $folderName }
    New-ShortCut -Title $courseTitle -Uri $Uri

    $outputPathes = New-Object System.Collections.ArrayList
    $actualDownloadAny = $false
    #$videos = $videos | Select-Object -First 3
    $counter = 0
    $videos | ForEach-Object {
        if ($_.Title -cnotmatch '(?m)^\d') {
            return
        }
    
        $title = $_.Title
        Write-Progress -Activity '下载视频' -Status '获取视频地址' -PercentComplete ($counter / $videos.Count / 2 * 100)
        $counter ++
        $videoUrl = Get-VideoUri $_.ID
        $extension = ($videoUrl -split '\.')[-1]

        $title = Get-NormalizedFileName $title
        $outputPath = "$folderName\$title.$extension"
        $null = $outputPathes.Add($outputPath)
        Write-Output $title
        Write-Debug $videoUrl
        Write-Debug $outputPath

        if (Test-Path $outputPath) {
            Write-Debug "目标文件 $outputPath 已存在，自动跳过"
        } else {
            Write-Progress -Activity '下载视频' -Status "下载《$title》视频文件" -PercentComplete ($counter / $videos.Count / 2 * 100)
            $counter ++
            if ($PSCmdlet.ShouldProcess("$videoUrl", 'Invoke-WebRequest')) {
                Invoke-WebRequest -Uri $videoUrl -OutFile $outputPath
                $actualDownloadAny = $true
            }
        }
    }

    Out-IndexFile $courseTitle $Uri $videos $folderName
    Combine-Videos $folderName $actualDownloadAny $outputPathes
}

Assert-PSVersion

# 判断参数集
$chosen= $PSCmdlet.ParameterSetName
if ($chosen -eq 'URI') {
    if ($Uri) {
        Download-Course $Uri
    } else {
        Get-ExistingCourses | ForEach-Object {
            Download-Course $_
        }
    }
}
if ($chosen -eq 'ID') {
    $template = 'http://www.imooc.com/learn/{0}'
    $ID | ForEach-Object {
        $Uri = $template -f $_
        Download-Course $Uri
    }
}