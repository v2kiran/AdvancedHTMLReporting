# AdvancedHTMLReporting

## OverView

A long long time ago there was an editor named "PowerGui" which had the concept of powerpacks or plugins that could be used to perform various tasks.
Once such powerpack was the **advancedhtmlreporting powerpack** using which you could generate pretty HTML reports.

Unfortunately after Dell took over Quest "PowerGui" was abandoned and whats more you cannot find any trace of it on the internet anymore.
I had the need to make some html reports today so I dug-up an old copy of the html powerpack and essentially stripped it so that the functions could be called from the powershell console.

## Import or dot source HTML functions

```powershell
#dot source the functions file
. \AdvancedHTMLReporting.ps1
```

## Usage

```powershell
$title = 'PowerShell Report'
$filepath = 'C:\temp\advreport.htm'

#Services
$Services = gsv b*
$svcProperties = 'Name', 'Displayname', 'Status', 'StartType'
$svcgroupByProperties = 'StartType', 'Status'

#Computer
$compdetails = Get-CimInstance win32_computersystem
$compprops = 'Name', 'Domain', 'Model', 'Manufacturer'

# Common indentation style
$indentationStyle = 'AllLevels'

# Generate the report
New-HtmlReport -Title $title -Filename $filepath -OpenFile  -Content {

    $compdetails | ConvertTo-HtmlReportList -Title 'General' -Property $compprops -Indent $indentationStyle -PrefixGroupNames -Collapsible -Expanded
    $Services | ConvertTo-HtmlReportList -Title 'Services List' -Property 'Name','DisplayName' -GroupBy $svcgroupByProperties -Indent $indentationStyle -PrefixGroupNames -Collapsible -Expanded
    $Services | ConvertTo-HtmlReportTable -Title 'Services' -Property $svcProperties -GroupBy $svcgroupByProperties -Indent $indentationStyle -PrefixGroupNames -Collapsible -Expanded

}
```

## Result

![Result](https://github.com/v2kiran/AdvancedHTMLReporting/blob/master/sample.png)
