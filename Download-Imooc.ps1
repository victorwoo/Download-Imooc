Param (
    [string]
    $Uri = 'http://www.imooc.com/learn/197'
)

$DebugPreference = 'Continue'

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

function Get-ID {
    Param (
        $Uri
    )
    
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

function New-ShortCut {
    Param (
        $Title,
        $Uri
    )

    $shell = New-Object -com 'wscript.shell'
    $lnk = $shell.CreateShortcut("$Title\$Title.url")
    $lnk.TargetPath = $Uri
    $lnk.Save()
}

Write-Progress -Activity '下载视频' -Status '分析视频 ID'
$title, $ids = Get-ID -Uri $Uri
Write-Output "课程名称：$title"
Write-Debug $title
$folderName = Fix-FolderName $title
Write-Debug $folderName
if (-not (Test-Path $folderName)) { $null = mkdir $folderName }
New-ShortCut -Title $title -Uri $Uri

$ids | ForEach-Object {
    if ($_.Title -cnotmatch '(?m)^\d') {
        return
    }
    
    $title = $_.Title
    Write-Progress -Activity '下载视频' -Status '获取视频地址'
    $videoUrl = Get-VideoUri $_.ID
    $extension = ($videoUrl -split '\.')[-1]

    Write-Progress -Activity '下载视频' -Status "下载《$title》视频文件"
    $title = Fix-FileName $title
    $outputPath = "$folderName\$title.$extension"
    Write-Output $title
    Write-Debug $videoUrl
    Write-Debug $outputPath
    Invoke-WebRequest -Uri $videoUrl -OutFile "$folderName\$title.$extension"
}