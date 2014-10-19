<#
	.SYNOPSIS
		这是一个从慕课网批量下载视频的 PowerShell 脚本。

	.DESCRIPTION
		慕课网（http://www.imooc.com）是一个大型开放式网络课程教学网站。本脚本用于抓取慕课网的视频资源，并对视频资源进行二次加工整理，使您能更方便地离线观看视频教程。

	.PARAMETER  Uri
		教程专辑的 URL，例如 'http://www.imooc.com/learn/197'。

	.PARAMETER  ID
		教程专辑的 ID，支持多个，例如 75, 197, 156, 203。

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
		PS> .\Download-Imooc.ps1 75, 197, 156, 203
		根据教程专辑的 ID 来下载。支持输入多个 ID。

	.EXAMPLE
		PS> .\Download-Imooc.ps1 http://www.imooc.com/learn/75 -Combine -RemoveOriginal
		根据教程专辑的 URL 来下载。完成之后合并所有的视频。

	.EXAMPLE
		PS> .\Download-Imooc.ps1 75, 197, 156, 203 -Combine -RemoveOriginal
		根据教程专辑的 ID 来下载。支持输入多个 ID。完成之后合并所有的视频。，并且删除原始的分段视频。

	.EXAMPLE
		PS> .\Download-Imooc.ps1
		不带任何参数运行该脚本，将自动检测当前目录下的所有下载文件夹，同时检测对应的专辑网站。对于本地缺失的视频或网站更新的视频，将自动续传。

	.NOTES
		只支持 PowerShell 3.0 及更高版本。
		默认下载的是最高清晰度的视频。
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

[CmdletBinding(DefaultParameterSetName = 'URI',
			   SupportsShouldProcess = $true,
			   ConfirmImpact = 'Medium')]
Param
(
	[Parameter(ParameterSetName = 'URI',
			   Position = 0,
			   Mandatory = $false,
			   ValueFromPipeline = $true,
			   HelpMessage = '请输入专辑 URL')]
	[string]
	$Uri,
	
	[Parameter(ParameterSetName = 'ID',
			   Position = 0,
			   Mandatory = $true,
			   ValueFromPipeline = $true,
			   HelpMessage = '请输入专辑 ID，多个 ID 请用逗号隔开')]
	[int[]]
	$ID,
	
	[Switch]
	$Combine,
	
	[Switch]
	$RemoveOriginal
)

$DebugPreference = 'Continue' # Continue, SilentlyContinue
 #$ProgressPreference='SilentlyContinue'
# $WhatIfPreference = $true # $true, $false

# 修正文件名，将文件系统不支持的字符替换成“.”
function Get-NormalizedFileName
{
	Param (
		$FileName
	)
	
	[System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {
		$FileName = $FileName.Replace($_, '.')
	}
	
    $FileName = $FileName.Replace('+', '.')
	return $FileName
}

# 修正目录名，将文件系统不支持的字符替换成“.”
function Get-NormalizedFolderName
{
	Param (
		$FolderName
	)
	
	[System.IO.Path]::GetInvalidPathChars() | ForEach-Object {
		$FolderName = $FolderName.Replace($_, '.')
	}
	
    $FolderName = $FolderName.Replace('+', '.')
	return $FolderName
}

# 从专辑页面中分析标题和视频页面的 ID。
function Get-CourseInfo
{
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
		if ($_.href -cmatch '(?m)^/video/(\d+)$')
		{
			$id = $Matches[1]
			$title = $_.InnerText
			if ($title -cmatch '(?m)^\d.*? \((?<DURING>\d{2}:\d{2})\)\s*$')
			{
				$during = $matches['DURING']
			}
			else
			{
				return
			}
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
function Get-VideoUri
{
	Param (
		[Parameter(ValueFromPipeline = $true)]
		$ID
	)
	
	$template = 'http://www.imooc.com/course/ajaxmediainfo/?mid={0}&mode=flash'
	$uri = $template -f $ID
	Write-Debug $uri
	$result = Invoke-RestMethod $uri
	
	if ($result.result -ne 0)
	{
		Write-Warning $result.result
	}
	
	$uri = $result.data.result.mpath.'0'
	
	# 取最高清晰度的版本。
	$uri = $uri.Replace('L.flv', 'H.flv').Replace('M.flv', 'H.flv')
    $uri = $uri.Replace('L.mp4', 'H.mp4').Replace('M.mp4', 'H.mp4')
	return $uri
}

# 获取源码下载信息。
function Get-SourceInfo
{
	Param (
		[Parameter(ValueFromPipeline = $true)]
		$ID
	)
	
	$template = 'http://www.imooc.com/video/{0}'
	$uri = $template -f $ID
	Write-Debug $uri
	$response = Invoke-WebRequest $uri
	
	$response.Links | Where-Object { $_.class -eq 'downcode' } | ForEach-Object {
        $source = [PSCustomObject][Ordered]@{
            Title = $_.title;
            Href = $_.href;
        }
        echo $source
    }
}

# 创建“.url”快捷方式。
function New-ShortCut
{
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
function Assert-PSVersion
{
	if ($PSVersionTable.PSVersion.Major -lt 3)
	{
		Write-Error '请安装 PowerShell 3.0 以上的版本。'
		exit
	}
}

# 获取当前目录下已存在的课程。
function Get-ExistingCourses
{
	Get-ChildItem -Directory | ForEach-Object {
		$folder = $_
		$expectedFilePath = (Join-Path $folder $folder.Name) + '.url'
		if (Test-Path -PathType Leaf $expectedFilePath)
		{
			$shell = New-Object -ComObject 'wscript.shell'
			$lnk = $shell.CreateShortcut($expectedFilePath)
			$targetPath = $lnk.TargetPath
			if ($targetPath -cmatch '(?m)\A^http://www\.imooc\.com/\w+/\d+$\z')
			{
				echo $targetPath
			}
		}
	}
}

# 输出索引文件。
function Out-IndexFile
{
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

function Combine-MP4
{
    Param ($sources, $dest)
#mp4box -cat file1 -cat file2 [-new] dest
    $params = @()
    $sources | ForEach-Object {
        $params += '-cat'
        $params += $_
    }

    $params += '-new'
    $params += $dest
    
    $eap = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"
    .\util\MP4Box.exe $params
	$ErrorActionPreference = $eap
	
    return $?
}

function Combine-Flv
{
    Param ($sources, $dest)
    $params = $sources

    $params.Insert(0, $dest)
			
	$eap = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"
	.\util\FlvBind.exe $params.ToArray()
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

    return $?
}

# 用 FlvBind.exe 合并视频文件。
function Combine-Videos
{
	Param ($folderName, $actualDownloadAny, $outputPathes)
	
	#if ($Combine -and ($actualDownloadAny -or -not (Test-Path $targetFile))) {
	if ($Combine)
	{
        if ($outputPathes.Count -eq 0) {
            return
        }
        $extension = [System.IO.Path]::GetExtension($outputPathes[0])
        $targetFile = "$folderName\$folderName$extension"

		# -and ($actualDownloadAny -or -not (Test-Path $targetFile))) {
		if ($actualDownloadAny -or -not (Test-Path $targetFile) -or (Test-Path $targetFile) -and $PSCmdlet.ShouldProcess('分段视频', '合并'))
		{
			Write-Progress `
                -Activity '下载' `
                -Status '合并视频' `
                -CurrentOperation ("合并视频（共 {0:N0} 个）" -f $outputPathes.Count) `
                -Id 2 `

			Write-Output ("合并视频（共 {0:N0} 个）" -f $outputPathes.Count)

            if (Test-Path $targetFile) {
                Remove-Item $targetFile
            }

            if ($extension.ToLower() -eq '.flv') {
                $result = Combine-Flv $outputPathes $targetFile
            } elseif ($extension.ToLower() -eq '.mp4') {
                $result = Combine-MP4 $outputPathes $targetFile
            }
			
			if ($result)
			{
				Write-Output '视频合并成功'
				if ($RemoveOriginal -and $PSCmdlet.ShouldProcess('分段视频', '删除'))
				{
					#$outputPathes.RemoveAt(0)
                    Remove-Item "$folderName\分段视频" -Recurse
					<# $outputPathes | ForEach-Object {
						Remove-Item $_
					} #>
					Write-Output '原始视频删除完毕'
				}
			}
			else
			{
				Write-Warning '视频合并失败'
			}
		}
	}
}

# 下载配套源代码。
function Download-Source
{
	Param (
		$FolderName,
		$Title,
		$Href
	)
	
	if (-not $Title -or -not $Href)
	{
		return
	}
	
	$extension = ($Href -split '\.')[-1]

    if (!(Test-Path "$folderName\源代码")) {
        $null = mkdir "$folderName\源代码"
    }

	$outputPath = "$folderName\源代码\$Title.$extension"
	
    if (!(Test-Path $outputPath)) {
	    if ($PSCmdlet.ShouldProcess("$Href", 'Invoke-WebRequest'))
	    {
		    Invoke-WebRequest -Uri $Href -OutFile $outputPath
	    }
    }
}

# 下载课程。
function Download-Course
{
	Param (
		[string]$Uri
	)
	
    Write-Progress `
        -Activity '下载课程' `
        -Status '分析视频 ID' `
        -CurrentOperation $Uri `
        -PercentComplete ($courcesIndex / $cources.Length * 100) `
        -Id 1

	$courseTitle, [array]$videos = Get-CourseInfo -Uri $Uri
 
	Write-Output "《$courseTitle》"
	$folderName = Get-NormalizedFolderName $courseTitle
	if (-not (Test-Path $folderName)) { $null = mkdir $folderName }
	New-ShortCut -Title $courseTitle -Uri $Uri
	
	$outputPathes = New-Object System.Collections.ArrayList
	$actualDownloadAny = $false
	#$videos = $videos | Select-Object -First 3
	
    Write-Progress `
        -Activity '下载课程' `
        -CurrentOperation $courseTitle `
        -PercentComplete ($courcesIndex / $cources.Length * 100) `
        -Id 1
    $videosIndex = 0
	$videos | ForEach-Object {
		if ($_.Title -cnotmatch '(?m)^\d')
		{
            $videosIndex++
			return
		}

		$title = $_.Title
        Write-Progress `
            -Activity '下载视频' `
            -CurrentOperation $title `
            -PercentComplete ($videosIndex / $videos.Count * 100) `
            -Id 2`
            -ParentId 1

		[array]$sources = Get-SourceInfo $_.ID
        $sourceIndex = 0
		$sources | ForEach-Object {
            Write-Progress `
                -Activity '下载源代码' `
                -CurrentOperation $_.Href `
                -PercentComplete ($sourceIndex / $sources.Length * 100) `
                -Id 3 `
                -ParentId 2

			Download-Source $folderName $_.Title $_.Href
            $sourceIndex++
		}
        echo 源代码下载完成
        Write-Progress `
            -Activity '下载源代码' `
            -Completed `
            -Id 3 `
            -ParentId 2
		
		$videoUrl = Get-VideoUri $_.ID
		$extension = ($videoUrl -split '\.')[-1]
		
		$title = Get-NormalizedFileName $title
        if (!(Test-Path "$folderName\分段视频")) {
            $null = mkdir "$folderName\分段视频"
        }

		$outputPath = "$folderName\分段视频\$title.$extension"
		$null = $outputPathes.Add($outputPath)
		Write-Output $title
		Write-Debug $videoUrl
		Write-Debug $outputPath
		
		if (Test-Path $outputPath)
		{
			Write-Debug "目标文件 $outputPath 已存在，自动跳过"
		}
		else
		{
			Write-Progress `
                -Activity '下载视频' `
                -CurrentOperation "$title" `
                -PercentComplete ($videosIndex / $videos.Count * 100) `
                -Id 2 `
                -ParentId 1

			if ($PSCmdlet.ShouldProcess("$videoUrl", 'Invoke-WebRequest'))
			{
				Invoke-WebRequest -Uri $videoUrl -OutFile $outputPath
				$actualDownloadAny = $true
			}
		}
        $videosIndex++
	}
	
	Out-IndexFile $courseTitle $Uri $videos $folderName
    Write-Progress `
        -Activity '下载视频' `
        -Status '合并视频' `
        -CurrentOperation $title `
        -PercentComplete 100 `
        -Id 2 `
        -ParentId 1

	Combine-Videos $folderName $actualDownloadAny $outputPathes
    Write-Progress `
        -Activity '下载视频' `
        -CurrentOperation $title `
        -Completed `
        -Id 2`
        -ParentId 1
}

Assert-PSVersion

# 判断参数集
$chosen = $PSCmdlet.ParameterSetName
if ($chosen -eq 'URI')
{
	if ($Uri)
	{
        [array]$cources = @($Uri)
		Download-Course $Uri
	}
	else
	{
		[array]$cources = Get-ExistingCourses
        
	}
}

if ($chosen -eq 'ID')
{
	$template = 'http://www.imooc.com/learn/{0}'
    [array]$cources = @()
	$ID | ForEach-Object {
		$Uri = $template -f $_
		$cources += $Uri
	}
}

$courcesIndex = 0
$cources | ForEach-Object {
	Download-Course $_
    $courcesIndex++
}
Write-Progress `
    -Activity '下载课程' `
    -Completed `
    -Id 1

echo '全部课程下载完毕'