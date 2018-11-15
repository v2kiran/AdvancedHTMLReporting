

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
    $Services | ConvertTo-HtmlReportList -Title 'Services List' -Property 'Name', 'DisplayName' -GroupBy $svcgroupByProperties -Indent $indentationStyle -PrefixGroupNames -Collapsible -Expanded
    $Services | ConvertTo-HtmlReportTable -Title 'Services' -Property $svcProperties -GroupBy $svcgroupByProperties -Indent $indentationStyle -PrefixGroupNames -Collapsible -Expanded
   
} 


