
#region HTML Report Functions
function New-HtmlReport
{
    param(
        [string]      $Title,
        [string]      $Filename = (join-path $env:TEMP "PSReport_$(Get-Date -Format hhmmss_ddMMyyyy).htm"),
        [switch]      $OpenFile,
        [switch]      $OpenFolder,
        [PSObject]    $Content = $null
    )
    [string]                     $reportContents = $null
    [string]                     $html = $null
    [string]                     $htmFilePath = $null
    [System.Diagnostics.Process] $process = $null

    if ($Content)
    {
        if ($Content -is [ScriptBlock])
        {
            & $Content | ForEach-Object {
                $reportContents += [string]$_
            }
            $reportContents.Trim("`n")
        }
        else
        {
            $reportContents = [string]$Content
        }
        $reportContents = $reportContents.Trim()
    }

    function Get-TableCssSettings
    {
        param(
            [string] $Display = 'none',
            [UInt16] $LeftIndent = 16,
            [switch] $Frame
        )
        @"
    display: $Display;
    position: relative;
    color: #000000;
$(if ($Frame) {
	@'
    background-color: #f9f9f9;
    border-left: #b1babf 1px solid;
    border-right: #b1babf 1px solid;
    border-top: #b1babf 1px solid;
    border-bottom: #b1babf 1px solid;
'@
})
    padding-left: ${LeftIndent}px;
    padding-top: 4px;
    padding-bottom: 5px;
    margin-left: 0px;
    margin-right: 0px;
    margin-bottom: 0px;
"@
    }
    function Get-TableTitleCssSettings
    {
        param(
            [string] $BackgroundColor = '#0061bd'
        )
        @"
    display: block;
    position: relative;
    height: 2em;
    color: #ffffff;
    background-color: $BackgroundColor;
    border-left: #b1babf 1px solid;
    border-right: #b1babf 1px solid;
    border-top: #b1babf 1px solid;
    border-bottom: #b1babf 1px solid;
    padding-left: 5px;
    padding-top: 8px;
    margin-left: 0px;
    margin-right: 0px;
    font-family: Tahoma;
    font-size: 8pt;
    font-weight: bold;
"@
    }
    function Get-SpanCssSettings
    {
        @"
    display: block;
    position: absolute;
    color: #ffffff;
    top: 8px;
    font-family: Tahoma;
    font-size: 8pt;
    font-weight: bold;
    text-decoration: underline;
"@
    }

    $html = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<title>$Title</title>
<meta http-equiv=Content-Type content='text/html; charset=windows-1252'></meta>
<meta name="save" content="history"></meta>
<style type="text/css">
body {
    margin-left: 4pt;
    margin-right: 4pt;
    margin-top: 6pt;
    font-family: Tahoma;
    font-size: 8pt;
    font-weight: normal;
}
h1 {
$(Get-TableTitleCssSettings -BackgroundColor '#0061bd')
}
h2 {
$(Get-TableTitleCssSettings -BackgroundColor '#ad1c18')
}
h3 {
$(Get-TableTitleCssSettings -BackgroundColor '#adadad')
}
span.expandableHeaderLink {
$(Get-SpanCssSettings)
}
span.expandableHeaderLinkRightJustified {
$(Get-SpanCssSettings)
    right: 8px;
}
table {
    table-layout: fixed;
    font-size: 100%;
    width: 100%;
    color: #000000;
}
th {
    color: #0061bd;
    padding-top: 2px;
    padding-bottom: 2px;
    vertical-align: top;
    text-align: left;
}
td {
    padding-top: 2px;
    padding-bottom: 2px;
    vertical-align: top;
}
*{margin:0}
div.visibleSection {
$(Get-TableCssSettings -Display 'block' -Frame)
}
div.hiddenSection {
$(Get-TableCssSettings -Display 'none' -Frame)
}
div.visibleSectionNoIndent {
$(Get-TableCssSettings -Display 'block' -LeftIndent 0 -Frame)
}
div.hiddenSectionNoIndent {
$(Get-TableCssSettings -Display 'none' -LeftIndent 0 -Frame)
}
div.visibleSectionNoFrame {
$(Get-TableCssSettings -Display 'block' -LeftIndent 0)
}
div.hiddenSectionNoFrame {
$(Get-TableCssSettings -Display 'none' -LeftIndent 0)
}
div.filler {
    display: block;
    position: relative;
    color: #ffffff;
    background: none transparent scroll repeat 0% 0%;
    border-left: medium none;
    border-right: medium none;
    border-top: medium none;
    border-bottom: medium none;
    padding-top: 4px;
    margin-left: 0px;
    margin-right: 0px;
    margin-bottom: -1px;
    font: 100%/8px Tahoma;
}
div.save {
    behavior: url(#default#savehistory);
}
</style>
<script type="text/javascript">
function toggleVisibility(tableHeader) {
    if (document.getElementById) {
        var triggerLabel = tableHeader.firstChild;
        while ((triggerLabel) && (triggerLabel.innerHTML != 'show') && (triggerLabel.innerHTML != 'hide')) {
            triggerLabel = triggerLabel.nextSibling
        }
        if (triggerLabel) {
            triggerLabel.innerHTML = (triggerLabel.innerHTML == 'hide' ? 'show' : 'hide');
            associatedTable = tableHeader.nextSibling
            while ((associatedTable) && (!(associatedTable.style))) {
                associatedTable = associatedTable.nextSibling
            }
            if (associatedTable) {
                associatedTable.style.display = (triggerLabel.innerHTML == 'hide' ? 'block' : 'none');
            }
        }
    }
}
if (!document.getElementById) {
    document.write('<style type="text/css">\n'+'\tdiv.hiddenSection {\n\t\tdisplay:block;\n\t}\n'+ '</style>');
}
</script>
</head>
<body>
<b><font face="Arial" size="5">$Title</font></b>
<hr size="8" color="#0061bd"></hr>
<font face="Arial" size="1"><b>Generated with 'PowerShell'})</b></font>
<br />
<font face="Arial" size="1">Report created on $(Get-Date)</font>
<div class="filler"></div>
<div class="filler"></div>
<div class="filler"></div>
<div class="save">
$reportContents
</div>
</body>
</html>
"@
    if (-not $Filename)
    {
        $Filename = "$($env:TEMP)\Report_$(Get-Date -Format hhmmss_ddMMyyyy).htm"
    }
    $html | Out-File -Encoding Unicode -FilePath $Filename
    $htmFilePath = (Get-Item -LiteralPath $Filename -ErrorAction SilentlyContinue).PSPath
    if (-not $htmFilePath)
    {
        throw "File '$Filename' was not created"
    }
    if ($OpenFile)
    {
        if (Test-Path -LiteralPath "Registry::HKEY_CLASSES_ROOT\.htm" -ErrorAction SilentlyContinue)
        {
            Invoke-Item -LiteralPath $htmFilePath
        }
        else
        {
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo.Filename = 'notepad.exe'
            $process.StartInfo.Arguments = "`"$($htmFilePath.Replace('Microsoft.PowerShell.Core\FileSystem::',''))`""
            if (-not $process.Start())
            { 
                throw 'Unable to launch notepad.exe'
            }
        }
    }
    if ($OpenFolder)
    {
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.Filename = 'explorer.exe'
        $process.StartInfo.Arguments = "/select,`"$($htmFilePath.Replace('Microsoft.PowerShell.Core\FileSystem::',''))`""
        if (-not $process.Start())
        { 
            throw 'Unable to launch explorer.exe'
        }
    }
    Get-Item -LiteralPath $htmFilePath
}


function Add-HtmlReportSeparator
{
    @"
<hr />
"@
}

function Add-HtmlReportSubtitle
{
    param(
        $Subtitle = $null
    )
    @"
<table>
<th><u>$Subtitle</u></th>
</table>
"@
}


function Add-HtmlReportSection
{
    param(
        [string]   $Title = $null,
        [UInt16]   $Level = 1,
        [switch]   $NoIndent,
        [switch]   $NoFrame,
        [switch]   $Collapsible,
        [switch]   $Expanded,
        [PSObject] $Content = $null
    )
    [UInt16] $headingLevel = $(if (@(1, 2, 3) -notcontains $Level) {3} else {$Level})
    [string] $sectionClass = 'visibleSection'
    [string] $reportContents = $null

    if ($Title)
    {
        if ($Collapsible)
        {
            if (-not $Expanded)
            {
                $sectionClass = 'hiddenSection'
            }
            @"
<h$headingLevel style="cursor: pointer" onclick="toggleVisibility(this)">
<span class="expandableHeaderLink">$Title</span>
<span class="expandableHeaderLinkRightJustified">$(if ($Expanded) {'hide'} else {'show'})</span>
</h$headingLevel>
"@
        }
        else
        {
            @"
<h$headingLevel>
$Title
</h$headingLevel>
"@
        }
    }
    if ($Content)
    {
        if ($Content -is [ScriptBlock])
        {
            & $Content | ForEach-Object {
                $reportContents += [string]$_
            }
            $reportContents = $reportContents.Trim("`n")
        }
        else
        {
            $reportContents = [string]$Content
        }
        $reportContents = $reportContents.Trim()
        if ($NoFrame)
        {
            @"
<div class="${sectionClass}NoFrame">
$reportContents
</div>
"@
        }
        elseif ($NoIndent)
        {
            @"
<div class="${sectionClass}NoIndent">
$reportContents
</div>
"@
        }
        else
        {
            @"
<div class="$sectionClass">
$reportContents
</div>
"@
        }
    }
    @"
<div class="filler"></div>
"@		
}


function ConvertTo-HtmlReportTable
{
    param(
        [PSObject] $InputObject = $null,
        [String[]] $Property = $null,
        [String[]] $GroupBy = $null,
        [string]   $Title = $null,
        [UInt16]   $Level = 1,
        [string]   $Indent = 'AllLevels',
        [switch]   $PrefixGroupNames,
        [switch]   $Collapsible,
        [switch]   $Expanded,
        [PSObject] $AdditionalContent = $null
    )
    begin
    {
        [PSObject] $processObject = $null
        [array]    $objectCollection = @()
        [String[]] $innerHtml = @()
        [string]   $groupNamePrefix = $null
        [string]   $groupTitle = $null
        [string]   $html = $null

        if (@('None', 'OneLevel', 'AllLevels') -notcontains $Indent)
        {
            throw "Cannot bind parameter ""Indent"". Specify one of the following values and try again. The possible values are ""None"", ""OneLevel"", and ""AllLevels""."
            return
        }
    }
    process
    {
        if ($InputObject -and $_)
        {
            throw 'The input object cannot be bound to any parameters for the command either because the command does not take pipeline input or the input and its properties do not match any of the parameters that take pipeline input.'
            return
        }
        if ($processObject = $(if ($InputObject) {$InputObject} else {$_}))
        {
            $objectCollection += $processObject
        }
    }
    end
    {
        if ($GroupBy)
        {
            $innerHtml = $objectCollection | Group-Object -Property $GroupBy[0] | ForEach-Object {
                $groupNamePrefix = $null
                if ($PrefixGroupNames)
                {
                    $groupNamePrefix = "$($GroupBy[0]): "
                }
                $groupTitle = $(if ($_.Name) {"$groupNamePrefix$($_.Name)"} else {"$groupNamePrefix<i>Value not set</i>"})
                $_.Group | ConvertTo-HtmlReportTable -Property $Property -GroupBy $(if ($GroupBy.Count -gt 1) {$GroupBy[1..$($GroupBy.Count - 1)]} else {$null}) -Title $groupTitle -Level ($Level + 1) -Indent $(if ($Indent -eq 'OneLevel') {'None'} else {$Indent}) -PrefixGroupNames:$PrefixGroupNames -Collapsible -Expanded:$Expanded
            }
            if (-not $innerHtml)
            {
                $innerHtml = @()
            }
            $html = [string]::Join("`n", $innerHtml)
        }
        else
        {
            if ($Property)
            {
                $innerHtml = $objectCollection | ConvertTo-Html -Property $Property
            }
            else
            {
                $innerHtml = $objectCollection | ConvertTo-Html
            }
            if (-not $innerHtml)
            {
                $innerHtml = @()
            }
            $html = [string]::Join("`n", $innerHtml) -replace '(?s).*(<table>.*</table>).*', '$1' -replace "<col>`n", "<col></col>`n"
        }
        if ($AdditionalContent)
        {
            if ($AdditionalContent -is [ScriptBlock])
            {
                $html += & $AdditionalContent
            }
            else
            {
                $html += [string]$AdditionalContent
            }
        }
        Add-HtmlReportSection -Title $Title -Level $Level -NoIndent:$($Indent -eq 'None') -Collapsible:$Collapsible -Expanded:$Expanded -Content $html
    }
}


function ConvertTo-HtmlReportList
{
    param(
        [PSObject] $InputObject = $null,
        [String[]] $Property = $null,
        [String[]] $GroupBy = $null,
        [string]   $Title = $null,
        [UInt16]   $Level = 1,
        [string]   $Indent = 'AllLevels',
        [switch]   $PrefixGroupNames,
        [switch]   $Collapsible,
        [switch]   $Expanded,
        [PSObject] $AdditionalContent = $null
    )
    begin
    {
        [PSObject] $processObject = $null
        [array]    $objectCollection = @()
        [String[]] $innerHtml = @()
        [string]   $groupNamePrefix = $null
        [string]   $groupTitle = $null
        [string]   $html = $null
        [UInt32]   $index = 0
        [string]   $itemHtml = $null

        if (@('None', 'OneLevel', 'AllLevels') -notcontains $Indent)
        {
            throw "Cannot bind parameter ""Indent"". Specify one of the following values and try again. The possible values are ""None"", ""OneLevel"", and ""AllLevels""."
            return
        }
    }
    process
    {
        if ($InputObject -and $_)
        {
            throw 'The input object cannot be bound to any parameters for the command either because the command does not take pipeline input or the input and its properties do not match any of the parameters that take pipeline input.'
            return
        }
        if ($processObject = $(if ($InputObject) {$InputObject} else {$_}))
        {
            $objectCollection += $processObject
        }
    }
    end
    {
        if ($GroupBy)
        {
            $innerHtml = $objectCollection | Group-Object -Property $GroupBy[0] | ForEach-Object {
                $groupNamePrefix = $null
                if ($PrefixGroupNames)
                {
                    $groupNamePrefix = "$($GroupBy[0]): "
                }
                $groupTitle = $(if ($_.Name) {"$groupNamePrefix$($_.Name)"} else {"$groupNamePrefix<i>Value not set</i>"})
                $_.Group | ConvertTo-HtmlReportList -Property $Property -GroupBy $(if ($GroupBy.Count -gt 1) {$GroupBy[1..$($GroupBy.Count - 1)]} else {$null}) -Title $groupTitle -Level ($Level + 1) -Indent $(if ($Indent -eq 'OneLevel') {'None'} else {$Indent}) -PrefixGroupNames:$PrefixGroupNames -Collapsible -Expanded:$Expanded
            }
            if (-not $innerHtml)
            {
                $innerHtml = @()
            }
            $html = [string]::Join("`n", $innerHtml)
        }
        else
        {
            $innerHtml = $(for ($index = 0; $index -lt $objectCollection.Count; $index++)
                {
                    $itemHtml = $(foreach ($item in $(if ($Property) {$Property} else {$objectCollection[$index].PSObject.Properties | Where-Object {$_.IsGettable} | ForEach-Object {$_.Name}}))
                        {
                            @"
<tr>
<th width='25%'><b>${item}:</b></th>
<td width='75%'>$([string]($objectCollection[$index].$item))</td>
</tr>
"@
                        })
                    if ($index -eq ($objectCollection.Count - 1))
                    {
                        $itemHtml
                    }
                    else
                    {
                        @"
$itemHtml
</table>
$(Add-HtmlReportSeparator)
<table>
"@
                    }
                })
            if (-not $innerHtml)
            {
                $innerHtml = @()
            }
            $html = @"
<table>
$([string]::Join("`n",$innerHtml))
</table>
"@
        }
        if ($AdditionalContent)
        {
            if ($AdditionalContent -is [ScriptBlock])
            {
                $html += & $AdditionalContent
            }
            else
            {
                $html += [string]$AdditionalContent
            }
        }
        Add-HtmlReportSection -Title $Title -Level $Level -NoFrame:$($Indent -eq 'None') -Collapsible:$Collapsible -Expanded:$Expanded -Content $html
    }
}