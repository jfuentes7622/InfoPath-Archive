<#
.SYNOPSIS
    Offline regression tests for reference discovery and the read-only SPS exporter.
.DESCRIPTION
    Uses generated local fixtures and an in-process SharePoint object facade.
    Never queries SharePoint. Test work files are created below the repository
    and removed in finally. Optionally tests an existing extracted sample.
#>
#requires -Version 5.1
[CmdletBinding()]
param([string]$SampleDirectory = '')
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
$work = Join-Path $repository ('.XsnReferenceTests_' + [guid]::NewGuid().ToString('N'))

class XsnReferenceTestState {
    static [int]$OpenedSites = 0
    static [int]$DisposedSites = 0
    static [int]$OpenedWebs = 0
    static [int]$DisposedWebs = 0
    static [int]$PageReads = 0
    static [type]$LookupValueType
    static [type]$UserValueType
}
class XsnReferenceTestFields {
    [object] GetFieldByInternalName([string]$Name) {
        if ($Name -notin @('ID', 'Value', 'Title', 'Name', 'EMail', 'Tags', 'LookupRef', 'UserRef', 'MultiRef')) { throw "Missing field '$Name'." }
        $typeName = switch ($Name) { 'LookupRef' { 'Lookup' }; 'UserRef' { 'User' }; 'MultiRef' { 'UserMulti' }; default { 'Text' } }
        $field = [pscustomobject]@{ InternalName = $Name; TypeAsString = $typeName }
        $field | Add-Member ScriptMethod GetFieldValue {
            param($Value)
            $parts = $Value -split ';#'
            $values = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $parts.Count; $i += 2) {
                $type = [XsnReferenceTestState]::LookupValueType
                if ($this.TypeAsString -like 'User*') { $type = [XsnReferenceTestState]::UserValueType }
                $typed = [Activator]::CreateInstance($type)
                $typed.LookupId = [int]$parts[$i]
                $null = $values.Add($typed)
            }
            if ($this.TypeAsString -like '*Multi') { return ,$values }
            return $values[0]
        }
        return $field
    }
}
class XsnReferenceTestPage : System.Collections.ArrayList {
    [object]$ListItemCollectionPosition
}
class XsnReferenceTestList {
    [object]$Fields = [XsnReferenceTestFields]::new()
    [object[]]$Rows
    [bool]$Fail
    [object] GetItems([object]$Query) {
        [XsnReferenceTestState]::PageReads++
        if ($this.Fail) { throw 'Synthetic list read failure.' }
        $page = [XsnReferenceTestPage]::new()
        $start = 0
        if ($null -ne $Query.ListItemCollectionPosition) { $start = [int]$Query.ListItemCollectionPosition }
        $end = [Math]::Min($start + [int]$Query.RowLimit, $this.Rows.Count)
        for ($i = $start; $i -lt $end; $i++) { $null = $page.Add($this.Rows[$i]) }
        if ($end -lt $this.Rows.Count) { $page.ListItemCollectionPosition = $end }
        return $page
    }
}
class XsnReferenceTestWeb {
    [string]$Url
    [object]$Site
    [hashtable]$Lists
    [object]$SiteUserInfoList
    [void] Dispose() { [XsnReferenceTestState]::DisposedWebs++ }
}
class XsnReferenceTestSite {
    [string]$Url = 'https://example.test/sites/owner'
    [object]$Web
    [object] OpenWeb([string]$Path, [bool]$Exact) {
        if (-not $Exact) { throw 'Exporter must request the exact web.' }
        [XsnReferenceTestState]::OpenedWebs++
        return $this.Web
    }
    [object] OpenWeb() {
        [XsnReferenceTestState]::OpenedWebs++
        return $this.Web
    }
    [void] Dispose() { [XsnReferenceTestState]::DisposedSites++ }
}

function Assert-XsnReference {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

try {
    New-Item -ItemType Directory -Path $work | Out-Null
    # Parse with the current host; running this script using powershell.exe also
    # verifies the actual Windows PowerShell 5.1 parser/runtime.
    foreach ($file in @('GetXsnReferenceDependencies.ps1', 'GetXsnReferenceData_SPS.ps1', 'Helper\XsnReferenceDiscovery.ps1', 'Helper\TestXsnReferenceExport.ps1')) {
        $tokens = $null; $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repository $file), [ref]$tokens, [ref]$errors)
        Assert-XsnReference ($errors.Count -eq 0) "Parser errors in $file : $errors"
    }
    $fixture = Join-Path $work 'Extracted'
    New-Item -ItemType Directory -Path $fixture | Out-Null
    @'
<xsf:xDocumentClass xmlns:xsf="http://schemas.microsoft.com/office/infopath/2003/solutionDefinition" xmlns:my="urn:main" name="urn:test" solutionVersion="1.2.3.4">
 <xsf:views><xsf:view name="Primary"><xsf:mainpane transform="first.xsl"/></xsf:view><xsf:view name="Second"><xsf:mainpane transform="second.xsl"/></xsf:view></xsf:views>
 <xsf:dataObjects>
  <xsf:dataObject name="Countries"><xsf:query><xsf:sharepointListAdapterRW name="Countries" siteURL="/sites/lookups/" sharePointListID="{11111111-1111-1111-1111-111111111111}"><xsf:field internalName="Value"/><xsf:field internalName="Title"/></xsf:sharepointListAdapterRW></xsf:query></xsf:dataObject>
  <xsf:dataObject name="Invisible"><xsf:query><xsf:sharepointListAdapterRW name="Invisible" siteURL="https://example.test/sites/lookups/" sharePointListID="{22222222-2222-2222-2222-222222222222}"><xsf:field internalName="Title"/></xsf:sharepointListAdapterRW></xsf:query></xsf:dataObject>
 </xsf:dataObjects>
</xsf:xDocumentClass>
'@ | Set-Content -LiteralPath (Join-Path $fixture 'manifest.xsf') -Encoding UTF8
    $view = @'
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:my="urn:main" xmlns:xd="http://schemas.microsoft.com/office/infopath/2003" xmlns:xdXDocument="http://schemas.microsoft.com/office/infopath/2003/xslt/xDocument" xmlns:dfs="urn:dfs" xmlns:d="urn:rows">
 <xsl:template match="/"><xsl:for-each select="my:Root"><xsl:for-each select="my:Rows/my:Row">
  <select xd:binding="my:Country" xd:xctname="dropdown">
   <xsl:for-each select="xdXDocument:GetDOM(&quot;Countries&quot;)/dfs:myFields/dfs:dataFields/d:Item[d:Title != '']">
    <option><xsl:attribute name="value"><xsl:value-of select="d:Value"/></xsl:attribute><xsl:value-of select="d:Title"/></option>
   </xsl:for-each>
  </select>
 </xsl:for-each></xsl:for-each></xsl:template>
</xsl:stylesheet>
'@
    $view | Set-Content -LiteralPath (Join-Path $fixture 'first.xsl') -Encoding UTF8
    # The second declared view has an unsupported expression; it must be reported.
    $view.Replace('my:Country', 'my:Unsupported').Replace('select="d:Title"', 'select="concat(d:Title, d:Value)"') |
        Set-Content -LiteralPath (Join-Path $fixture 'second.xsl') -Encoding UTF8
    $discovery = Join-Path $repository 'GetXsnReferenceDependencies.ps1'
    $dependenciesPath = Join-Path $work 'Dependencies.json'
    $null = & $discovery -ExtractedXsnDirectory $fixture -TemplateUrl 'https://example.test/sites/owner/Library/Forms/template.xsn' -OutputPath $dependenciesPath
    $dep = Get-Content -LiteralPath $dependenciesPath -Raw | ConvertFrom-Json
    Assert-XsnReference ($dep.Templates -is [array] -and $dep.Lists -is [array] -and $dep.UserSites -is [array] -and $dep.Warnings -is [array]) 'Collections must always be arrays.'
    Assert-XsnReference ($dep.Lists.Count -eq 2) 'Invisible list connection is exported.'
    $binding = $dep.Templates[0].Bindings[0]
    Assert-XsnReference ($binding.XPath -ceq '/my:Root/my:Rows/my:Row/my:Country') 'MAIN repeating-field absolute XPath.'
    Assert-XsnReference ($binding.ValueField -ceq 'Value' -and $binding.DisplayField -ceq 'Title') 'Actual value/display mapping preserved.'
    Assert-XsnReference ($binding.ValueXPath -ceq 'd:Value' -and $binding.DisplayXPath -ceq 'd:Title') 'Secondary field XPaths.'
    Assert-XsnReference ($binding.CacheRowXPath -ceq '/dfs:myFields/dfs:dataFields/d:Item') 'Cached row path excludes live cascade predicates.'
    Assert-XsnReference ($dep.UserSites[0] -ceq 'https://example.test/sites/owner') 'Owning web, not lookup web, supplies users.'
    Assert-XsnReference ($dep.Warnings.Count -gt 0) 'All views inspected and unsupported expression reported.'
    $main = [xml]'<my:Root xmlns:my="urn:main"><my:Rows><my:Row><my:Country>USA</my:Country></my:Row></my:Rows></my:Root>'
    $manager = [Xml.XmlNamespaceManager]::new($main.NameTable)
    foreach ($property in $dep.Templates[0].Namespaces.PSObject.Properties) { $manager.AddNamespace($property.Name, $property.Value) }
    Assert-XsnReference ($main.SelectSingleNode($binding.XPath, $manager).InnerText -ceq 'USA') 'Discovered XPath selects a MAIN field with namespace-aware XPath.'
    # Portable processed inventory: an obsolete absolute ManifestPath is not used.
    '{"Labels":[]}' | Set-Content -LiteralPath (Join-Path $fixture 'FieldLabels.json') -Encoding UTF8
    [pscustomobject]@{ TemplateUrl = 'https://example.test/sites/owner/Library/Forms/template.xsn'; SolutionVersion = '1.2.3.4'; SolutionName = 'urn:test'; LabelsRelativePath = 'Extracted\FieldLabels.json'; SharePointVersion = '1.0'; Modified = ''; Status = 'Success'; Sha256 = 'ABCDEF'; ManifestPath = 'Z:\obsolete\manifest.xsf' } |
        Export-Csv -LiteralPath (Join-Path $work 'XsnTemplateVersions.csv') -NoTypeInformation -Encoding UTF8
    $null = & $discovery -XsnInventoryCsvPath (Join-Path $work 'XsnTemplateVersions.csv') -OutputPath (Join-Path $work 'Inventory.json')
    $inventoryDep = Get-Content -LiteralPath (Join-Path $work 'Inventory.json') -Raw | ConvertFrom-Json
    Assert-XsnReference ($inventoryDep.Templates[0].Sha256 -ceq 'ABCDEF') 'Inventory metadata retained after relocation.'
    $null = & $discovery -ExtractedXsnDirectory $fixture -TemplateUrl 'https://example.test/custom.xsn' -UserSiteUrl 'https://example.test/override/' -OutputPath (Join-Path $work 'Override.json')
    $override = Get-Content -LiteralPath (Join-Path $work 'Override.json') -Raw | ConvertFrom-Json
    Assert-XsnReference ($override.UserSites[0] -ceq 'https://example.test/override') 'Explicit user web override.'
    $null = & $discovery -ExtractedXsnDirectory $fixture -TemplateUrl 'https://example.test/custom.xsn' -OutputPath (Join-Path $work 'UnknownWeb.json')
    $unknown = Get-Content -LiteralPath (Join-Path $work 'UnknownWeb.json') -Raw | ConvertFrom-Json
    Assert-XsnReference ($unknown.UserSites -is [array] -and $unknown.UserSites.Count -eq 0 -and $unknown.Warnings.Count -gt 0) 'Unknown owning web is recorded, never guessed.'
    $view.Replace('select="d:Title"', 'select="d:Value"') | Set-Content -LiteralPath (Join-Path $fixture 'second.xsl') -Encoding UTF8
    $null = & $discovery -ExtractedXsnDirectory $fixture -TemplateUrl 'https://example.test/sites/owner/Library/Forms/template.xsn' -OutputPath (Join-Path $work 'ConflictingViews.json')
    $conflicting = Get-Content -LiteralPath (Join-Path $work 'ConflictingViews.json') -Raw | ConvertFrom-Json
    Assert-XsnReference ($conflicting.Templates[0].Bindings.Count -eq 0 -and @($conflicting.Warnings | Where-Object { $_ -like '*ambiguous mappings across views*' }).Count -eq 1) 'Conflicting view labels are rejected rather than guessed.'
    # Missing views are partial failures, with known connections still persisted.
    Remove-Item -LiteralPath (Join-Path $fixture 'second.xsl') -Force
    $discoveryThrew = $false
    try { $null = & $discovery -ExtractedXsnDirectory $fixture -TemplateUrl 'https://example.test/sites/owner/Library/Forms/template.xsn' -OutputPath (Join-Path $work 'PartialDiscovery.json') }
    catch { $discoveryThrew = $true }
    $partialDiscovery = Get-Content -LiteralPath (Join-Path $work 'PartialDiscovery.json') -Raw | ConvertFrom-Json
    Assert-XsnReference ($discoveryThrew -and $partialDiscovery.Lists.Count -eq 2) 'Discovery partial failure persists known dependencies then throws.'
    if ($SampleDirectory) {
        $null = & $discovery -ExtractedXsnDirectory $SampleDirectory -TemplateUrl 'https://example.test/sites/notices/Library/Forms/template.xsn' -OutputPath (Join-Path $work 'Sample.json')
        $sample = Get-Content -LiteralPath (Join-Path $work 'Sample.json') -Raw | ConvertFrom-Json
        $country = @($sample.Templates[0].Bindings | Where-Object { $_.ConnectionName -eq 'NoticeCountries' -and $_.ValueField -eq 'Value' -and $_.DisplayField -eq 'Title' })
        Assert-XsnReference ($country.Count -gt 0) 'Real sample Country Value=>Title discovery.'
        Assert-XsnReference ($sample.Lists.Count -ge 17) 'Real sample all 17 list connections.'
        Write-Host "Sample: $($sample.Lists.Count) lists, $($sample.Templates[0].Bindings.Count) bindings, $($sample.Warnings.Count) diagnostic warnings."
    }
    # Everything below uses a facade. A real Get-SPSite command is never invoked.
    # Reflection.Emit creates only in-memory stand-ins for actual typed field values.
    $assemblyName = [Reflection.AssemblyName]::new('XsnReferenceFieldFacade_' + [guid]::NewGuid().ToString('N'))
    if ($PSVersionTable.PSVersion.Major -le 5) { $assembly = [AppDomain]::CurrentDomain.DefineDynamicAssembly($assemblyName, [Reflection.Emit.AssemblyBuilderAccess]::Run) }
    else { $assembly = [Reflection.Emit.AssemblyBuilder]::DefineDynamicAssembly($assemblyName, [Reflection.Emit.AssemblyBuilderAccess]::Run) }
    $module = $assembly.DefineDynamicModule($assemblyName.Name)
    $lookupBuilder = $module.DefineType('Microsoft.SharePoint.SPFieldLookupValue', [Reflection.TypeAttributes]::Public)
    $null = $lookupBuilder.DefineField('LookupId', [int], [Reflection.FieldAttributes]::Public)
    [XsnReferenceTestState]::LookupValueType = $lookupBuilder.CreateType()
    $userBuilder = $module.DefineType('Microsoft.SharePoint.SPFieldUserValue', [Reflection.TypeAttributes]::Public, [XsnReferenceTestState]::LookupValueType)
    [XsnReferenceTestState]::UserValueType = $userBuilder.CreateType()
    $lookup = [XsnReferenceTestList]::new()
    $lookup.Rows = @(
        @{ ID = 1; Value = 'US'; Title = 'United States'; Tags = @('A', 'B'); LookupRef = '7;#Country'; UserRef = '42;#User'; MultiRef = '42;#User;#43;#Other' },
        @{ ID = 2; Value = 'CA'; Title = 'Canada'; Tags = @('C'); LookupRef = ''; UserRef = ''; MultiRef = '' }
    )
    $broken = [XsnReferenceTestList]::new(); $broken.Fail = $true
    $users = [XsnReferenceTestList]::new()
    $users.Rows = @(@{ ID = 17; Name = 'i:0#.w|DOMAIN\retained'; Title = 'Retained User'; EMail = 'retained@example.test' })
    $fakeWeb = [XsnReferenceTestWeb]::new()
    $fakeWeb.Url = 'https://example.test/sites/owner'
    $fakeWeb.Site = [pscustomobject]@{ Url = $fakeWeb.Url }
    $fakeWeb.Lists = @{ ([guid]'11111111-1111-1111-1111-111111111111') = $lookup; ([guid]'22222222-2222-2222-2222-222222222222') = $broken }
    $fakeWeb.SiteUserInfoList = $users
    $subwebList = [XsnReferenceTestList]::new()
    $subwebList.Rows = @(@{ ID = 9; Value = 'GB'; Title = 'Subweb-only Country' })
    $resourcesWeb = [XsnReferenceTestWeb]::new()
    $resourcesWeb.Url = $fakeWeb.Url + '/resources'
    $resourcesWeb.Site = $fakeWeb.Site
    $resourcesWeb.Lists = @{ ([guid]'11111111-1111-1111-1111-111111111111') = $subwebList }
    $redWeb = [XsnReferenceTestWeb]::new()
    $redWeb.Url = $fakeWeb.Url + '/red'
    $redWeb.Site = $fakeWeb.Site
    $subweb = [XsnReferenceTestWeb]::new()
    $subweb.Url = $fakeWeb.Url + '/subweb'
    $subweb.Site = $fakeWeb.Site
    $fakeWebs = @{}
    foreach ($web in @($fakeWeb, $resourcesWeb, $redWeb, $subweb)) { $fakeWebs[$web.Url] = $web }
    function Get-SPWeb {
        [CmdletBinding()]param([string]$Identity)
        $url = $Identity.TrimEnd('/')
        if (-not $fakeWebs.ContainsKey($url)) { throw "Synthetic exact web resolution failure: $Identity" }
        [XsnReferenceTestState]::OpenedWebs++
        return $fakeWebs[$url]
    }
    function Get-SPSite {
        [CmdletBinding()]param([string]$Identity)
        if ($Identity.TrimEnd('/') -ine $fakeWeb.Url) { throw "Get-SPSite requires a site-collection URL, not a subweb: $Identity" }
        [XsnReferenceTestState]::OpenedSites++
        $site = [XsnReferenceTestSite]::new(); $site.Web = $fakeWeb
        return $site
    }
    function New-Object {
        [CmdletBinding()]param([Parameter(Position = 0)][string]$TypeName)
        if ($TypeName -ne 'Microsoft.SharePoint.SPQuery') { throw "Unexpected server-object construction '$TypeName'." }
        return [pscustomobject]@{ RowLimit = 0; ViewAttributes = ''; ViewFieldsOnly = $false; IncludeMandatoryColumns = $true; Query = ''; ViewFields = ''; ListItemCollectionPosition = $null }
    }
    $exportInput = [ordered]@{
        SchemaVersion = 1; Templates = @(); Warnings = @()
        Lists = @(
            @{ SiteUrl = $fakeWeb.Url; ListId = '11111111-1111-1111-1111-111111111111'; Fields = @('Value', 'Title', 'Tags', 'LookupRef', 'UserRef', 'MultiRef') },
            @{ SiteUrl = $fakeWeb.Url + '/'; ListId = '{11111111-1111-1111-1111-111111111111}'; Fields = @('Title') }
        )
        UserSites = @($fakeWeb.Url, ($fakeWeb.Url + '/subweb'))
    }
    $exportInputPath = Join-Path $work 'ExportInput.json'
    $exportInput | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $exportInputPath -Encoding UTF8
    $exporter = Join-Path $repository 'GetXsnReferenceData_SPS.ps1'
    $exportRoot = & $exporter -DependencyManifestPath $exportInputPath -LocalPath (Join-Path $work 'Exports') -PageSize 1
    $manifest = Get-Content -LiteralPath (Join-Path $exportRoot 'Manifest.json') -Raw | ConvertFrom-Json
    foreach ($helper in @('HtmlRendering.ps1', 'XsnTemplates.ps1', 'XsnHtmlRendering.ps1', 'HtmlReferenceData.ps1')) {
        . (Join-Path $PSScriptRoot $helper)
    }
    $imported = Import-HtmlReferenceData $exportRoot
    $listKey = $fakeWeb.Url + '|11111111-1111-1111-1111-111111111111'
    Assert-XsnReference ($imported.Lists.Count -eq 1 -and $imported.Lists.ContainsKey($listKey)) 'Consumer accepts exporter normalized site/GUID key and deduplicated schema.'
    $importedList = $imported.Lists[$listKey]
    Assert-XsnReference ($importedList.Rows -is [array] -and $importedList.Rows.Count -eq 2) 'Consumer accepts actual many-row list CSV.'
    Assert-XsnReference (($importedList.Rows[0].PSObject.Properties.Name -join ',') -ceq 'ID,LookupRef,MultiRef,Tags,Title,UserRef,Value') 'Consumer preserves actual internal CSV column names.'
    $importedIndex = Get-HtmlReferenceIndex $importedList 'Value' 'Title'
    Assert-XsnReference ($importedIndex.Count -eq 2 -and $importedIndex['US'] -ceq 'United States' -and $importedIndex['CA'] -ceq 'Canada') 'Consumer indexes actual exported values and labels.'
    Assert-XsnReference ($imported.Users.Count -eq 1 -and $imported.Users[0].Rows.Count -eq 1) 'Consumer accepts actual one-row user CSV.'
    Assert-XsnReference (($imported.Users[0].Rows[0].PSObject.Properties.Name -join ',') -ceq 'UserId,LoginName,DisplayName,Email') 'Consumer preserves user CSV columns.'
    Assert-XsnReference ($imported.Users[0].Rows[0].LoginName -ceq 'i:0#.w|DOMAIN\retained' -and $imported.Users[0].Entry.SourceSiteUrls.Count -eq 2) 'Consumer preserves exact login and requested source sites.'
    Assert-XsnReference ($manifest.Lists.Count -eq 1 -and $manifest.Lists[0].RowCount -eq 2) 'List deduplication and paging.'
    Assert-XsnReference ($manifest.Users.Count -eq 1 -and $manifest.Users[0].SourceSiteUrls.Count -eq 2) 'User sources deduplicated at site-collection scope.'
    $userRows = @(Import-Csv -LiteralPath (Join-Path $exportRoot $manifest.Users[0].DataPath))
    Assert-XsnReference ($userRows[0].LoginName -ceq 'i:0#.w|DOMAIN\retained' -and $userRows[0].UserId -eq '17') 'Exact claims login and scoped ID retained.'
    $listRows = @(Import-Csv -LiteralPath (Join-Path $exportRoot $manifest.Lists[0].DataPath))
    Assert-XsnReference ($listRows[0].Tags -ceq 'A;#B' -and $listRows[0].ID -eq '1') 'Meaningful multivalues and ID column.'
    Assert-XsnReference ($listRows[0].LookupRef -ceq '7' -and $listRows[0].UserRef -ceq '42' -and $listRows[0].MultiRef -ceq '42;#43') 'Lookup and User values/collections store IDs, not labels or type names.'
    Assert-XsnReference ($manifest.Dependencies.UserSites.Count -eq 2) 'Entire dependency object retained.'
    $subwebInput = [ordered]@{
        SchemaVersion = 1; Templates = @(); Warnings = @()
        Lists = @(@{ SiteUrl = $resourcesWeb.Url; ListId = '11111111-1111-1111-1111-111111111111'; Fields = @('Value', 'Title') })
        UserSites = @($fakeWeb.Url, $resourcesWeb.Url, $redWeb.Url)
    }
    $subwebInputPath = Join-Path $work 'SubwebInput.json'
    $subwebInput | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $subwebInputPath -Encoding UTF8
    $subwebRoot = & $exporter -DependencyManifestPath $subwebInputPath -LocalPath (Join-Path $work 'Subwebs') -PageSize 1
    $subwebManifest = Get-Content -LiteralPath (Join-Path $subwebRoot 'Manifest.json') -Raw | ConvertFrom-Json
    $subwebRows = @(Import-Csv -LiteralPath (Join-Path $subwebRoot $subwebManifest.Lists[0].DataPath))
    Assert-XsnReference ($subwebManifest.Lists[0].Status -ceq 'Success' -and $subwebRows.Count -eq 1 -and $subwebRows[0].Title -ceq 'Subweb-only Country') 'Lookup export opens the exact subweb, not the collection root with the same list GUID.'
    Assert-XsnReference ($subwebManifest.Lists[0].SiteCollectionUrl -ceq $fakeWeb.Url) 'Subweb list records its owning collection.'
    Assert-XsnReference ($subwebManifest.Users.Count -eq 1 -and $subwebManifest.Users[0].Status -ceq 'Success' -and $subwebManifest.Users[0].SourceSiteUrls.Count -eq 3) 'Root/resources/red user sources deduplicate to their owning collection.'
    $subwebUsers = @(Import-Csv -LiteralPath (Join-Path $subwebRoot $subwebManifest.Users[0].DataPath))
    Assert-XsnReference ($subwebUsers.Count -eq 1 -and $subwebUsers[0].DisplayName -ceq 'Retained User') 'Subweb user sources export the root User Information List.'
    $subwebImported = Import-HtmlReferenceData $subwebRoot
    Assert-XsnReference ($subwebImported.Lists.ContainsKey($resourcesWeb.Url + '|11111111-1111-1111-1111-111111111111') -and $subwebImported.Users[0].Entry.SourceSiteUrls -contains $redWeb.Url) 'Consumer retains subweb lookup keys and person-source scope.'
    $subwebInput.Lists[0].SiteUrl = $resourcesWeb.Url + '/missing'
    $subwebInput.UserSites = @($redWeb.Url + '/missing')
    $subwebInput | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $subwebInputPath -Encoding UTF8
    $missingWebThrew = $false
    try { $null = & $exporter -DependencyManifestPath $subwebInputPath -LocalPath (Join-Path $work 'MissingWeb') }
    catch { $missingWebThrew = $true }
    $missingWebRoot = @(Get-ChildItem -LiteralPath (Join-Path $work 'MissingWeb') -Directory)[0].FullName
    $missingWebManifest = Get-Content -LiteralPath (Join-Path $missingWebRoot 'Manifest.json') -Raw | ConvertFrom-Json
    Assert-XsnReference ($missingWebThrew -and $missingWebManifest.Lists[0].Status -ceq 'Failed' -and $missingWebManifest.Users[0].Status -ceq 'Failed') 'Missing subwebs must fail explicitly rather than export from a parent.'
    $emptyList = [XsnReferenceTestList]::new(); $emptyList.Rows = @()
    $fakeWeb.Lists[([guid]'33333333-3333-3333-3333-333333333333')] = $emptyList
    [ordered]@{
        SchemaVersion = 1; Templates = @(); Warnings = @(); UserSites = @()
        Lists = @(@{ SiteUrl = $fakeWeb.Url; ListId = '33333333-3333-3333-3333-333333333333'; Fields = @('Title') })
    } | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $work 'EmptyInput.json') -Encoding UTF8
    $emptyRoot = & $exporter -DependencyManifestPath (Join-Path $work 'EmptyInput.json') -LocalPath (Join-Path $work 'Empty') -PageSize 1
    $emptyManifest = Get-Content -LiteralPath (Join-Path $emptyRoot 'Manifest.json') -Raw | ConvertFrom-Json
    $emptyCsv = @(Get-Content -LiteralPath (Join-Path $emptyRoot $emptyManifest.Lists[0].DataPath))
    Assert-XsnReference ($emptyManifest.Users -is [array] -and $emptyManifest.Users.Count -eq 0 -and $emptyManifest.Lists[0].RowCount -eq 0 -and $emptyCsv.Count -eq 1 -and $emptyCsv[0] -ceq '"ID","Title"') 'Empty snapshots retain CSV headers and empty manifest arrays.'
    $importedEmpty = Import-HtmlReferenceData $emptyRoot
    $emptyKey = $fakeWeb.Url + '|33333333-3333-3333-3333-333333333333'
    Assert-XsnReference ($importedEmpty.Lists.Count -eq 1 -and $importedEmpty.Lists[$emptyKey].Entry.Status -ceq 'Success' -and $importedEmpty.Lists[$emptyKey].Entry.RowCount -eq 0) 'Consumer accepts successful zero-row exporter manifest entry.'
    Assert-XsnReference ($importedEmpty.Lists[$emptyKey].Rows -is [array] -and $importedEmpty.Lists[$emptyKey].Rows.Count -eq 0 -and $importedEmpty.Users.Count -eq 0) 'Consumer accepts header-only CSV and empty users array.'
    $emptyIndex = Get-HtmlReferenceIndex $importedEmpty.Lists[$emptyKey] 'ID' 'Title'
    Assert-XsnReference ($emptyIndex.Count -eq 0) 'Consumer indexes header-only exported fields without inventing a row.'
    $exportInput.Lists += @{ SiteUrl = $fakeWeb.Url; ListId = '22222222-2222-2222-2222-222222222222'; Fields = @('Title') }
    $exportInput.UserSites += 'https://example.test/unreachable'
    $exportInput | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $exportInputPath -Encoding UTF8
    $threw = $false
    try { $null = & $exporter -DependencyManifestPath $exportInputPath -LocalPath (Join-Path $work 'Partial') -PageSize 1 }
    catch { $threw = $true }
    Assert-XsnReference $threw 'Partial failures throw after persistence.'
    $partial = @(Get-ChildItem -LiteralPath (Join-Path $work 'Partial') -Directory)[0].FullName
    $partialManifest = Get-Content -LiteralPath (Join-Path $partial 'Manifest.json') -Raw | ConvertFrom-Json
    Assert-XsnReference (@($partialManifest.Lists | Where-Object Status -eq 'Failed').Count -eq 1) 'List failure recorded.'
    Assert-XsnReference (@($partialManifest.Users | Where-Object Status -eq 'Failed').Count -eq 1) 'User resolution failure recorded.'
    Assert-XsnReference (Test-Path -LiteralPath (Join-Path $partial 'Export.log')) 'Partial failure log persisted.'
    $importedPartial = Import-HtmlReferenceData $partial
    $failedKey = $fakeWeb.Url + '|22222222-2222-2222-2222-222222222222'
    Assert-XsnReference ($importedPartial.Lists.Count -eq 2 -and $importedPartial.Lists[$listKey].Rows.Count -eq 2) 'Consumer retains successful rows in actual partial export.'
    Assert-XsnReference ($importedPartial.Lists[$failedKey].Entry.Status -ceq 'Failed' -and $importedPartial.Lists[$failedKey].Entry.Error -like '*Synthetic list read failure*') 'Consumer accepts actual failed list entry and error.'
    Assert-XsnReference ($importedPartial.Lists[$failedKey].Rows -is [array] -and $importedPartial.Lists[$failedKey].Rows.Count -eq 0 -and $importedPartial.Lists[$failedKey].Entry.RowCount -eq 0) 'Consumer never treats failed list CSV as successful data.'
    $failedUserImports = @($importedPartial.Users | Where-Object { $_.Entry.Status -eq 'Failed' })
    Assert-XsnReference ($importedPartial.Users.Count -eq 2 -and $failedUserImports.Count -eq 1 -and $failedUserImports[0].Rows.Count -eq 0) 'Consumer accepts failed user resolution alongside successful users.'
    Assert-XsnReference ([XsnReferenceTestState]::OpenedSites -eq [XsnReferenceTestState]::DisposedSites) 'All owned SPSites disposed.'
    Assert-XsnReference ([XsnReferenceTestState]::OpenedWebs -eq [XsnReferenceTestState]::DisposedWebs) 'All owned SPWebs disposed.'
    Assert-XsnReference ([XsnReferenceTestState]::PageReads -ge 6) 'Paged facade exercised.'
    Write-Host 'PASS: offline discovery, real sample (if supplied), exact subwebs, collection-scoped users, SPS facade paging/deduplication/partial failures/disposal, and actual export-to-consumer zero/one/many-row interoperability.'
}
finally {
    if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
}
