# Require PowerShell 3.0 or higher.

[CmdletBinding(DefaultParameterSetName='URI', SupportsShouldProcess=$true, ConfirmImpact='Medium')]
Param
(
    [Parameter(ParameterSetName='URI',Position = 0)]
    [string]
    $Uri, # 'http://www.imooc.com/learn/197'

    [Parameter(ParameterSetName='ID', Position = 0)]
    [int[]]
    $ID, # @(75, 197)

    [Switch]
    $Combine, # = $true

    [Switch]
    $RemoveOriginal
)

# $DebugPreference = 'Continue' # Continue, SilentlyContinue
# $WhatIfPreference = $true # $true, $false

# 修正文件名，将文件系统不支持的字符替换成“.”
function Fix-FileName {
    Param (
        $FileName
    )

    [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {
        $FileName = $FileName.Replace($_, '.')
    }

    return $FileName
}

# 修正目录名，将文件系统不支持的字符替换成“.”
function Fix-FolderName {
    Param (
        $FolderName
    )

    [System.IO.Path]::GetInvalidPathChars() | ForEach-Object {
        $FolderName = $FolderName.Replace($_, '.')
    }

    return $FolderName
}

# 从专辑页面中分析标题和视频页面的 ID。
function Get-ID {
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
            return [PSCustomObject][Ordered]@{
                Title = $_.InnerText;
                ID = $Matches[1]
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

function Assert-PSVersion {
    if (($PSVersionTable.PSCompatibleVersions | Where-Object Major -ge 3).Count -eq 0) {
        Write-Error '请安装 PowerShell 3.0 以上的版本。'
        exit
    }
}

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

# 下载课程。
function Download-Course {
    Param (
        [string]$Uri
    )

    Write-Progress -Activity '下载视频' -Status '分析视频 ID'
    $title, $ids = Get-ID -Uri $Uri
    Write-Output "课程名称：$title"
    Write-Debug $title
    $folderName = Fix-FolderName $title
    Write-Debug $folderName
    if (-not (Test-Path $folderName)) { $null = mkdir $folderName }
    New-ShortCut -Title $title -Uri $Uri

    $outputPathes = New-Object System.Collections.ArrayList
    $actualDownloadAny = $false
    #$ids = $ids | Select-Object -First 3
    $ids | ForEach-Object {
        if ($_.Title -cnotmatch '(?m)^\d') {
            return
        }
    
        $title = $_.Title
        Write-Progress -Activity '下载视频' -Status '获取视频地址'
        $videoUrl = Get-VideoUri $_.ID
        $extension = ($videoUrl -split '\.')[-1]

        $title = Fix-FileName $title
        $outputPath = "$folderName\$title.$extension"
        $null = $outputPathes.Add($outputPath)
        Write-Output $title
        Write-Debug $videoUrl
        Write-Debug $outputPath

        if (Test-Path $outputPath) {
            Write-Warning "目标文件 $outputPath 已存在，自动跳过"
        } else {
            Write-Progress -Activity '下载视频' -Status "下载《$title》视频文件"
            if ($PSCmdlet.ShouldProcess("$videoUrl", 'Invoke-WebRequest')) {
                Invoke-WebRequest -Uri $videoUrl -OutFile $outputPath
                $actualDownloadAny = $true
            }
        }
    }

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
                }
            } else {
                Write-Warning '视频合并失败'
            }
        }
    }
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