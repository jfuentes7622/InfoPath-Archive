function Import-HtmlReferenceData {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if (Test-Path -LiteralPath $Path -PathType Container) { $Path = Join-Path $Path 'Manifest.json' }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Reference-data manifest not found: $Path" }
    $manifest = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -ErrorAction Stop
    foreach ($name in @('SchemaVersion', 'ExportedAtUtc', 'Dependencies', 'Lists', 'Users')) {
        if ($null -eq $manifest -or $manifest.PSObject.Properties.Name -notcontains $name) {
            throw "Reference-data manifest lacks '$name': $Path"
        }
    }
    if ($manifest.SchemaVersion -ne 1 -or $manifest.Dependencies.SchemaVersion -ne 1 -or
        $manifest.Dependencies.Templates -isnot [array] -or $manifest.Lists -isnot [array] -or $manifest.Users -isnot [array]) {
        throw "Unsupported or malformed reference-data manifest: $Path"
    }
    $root = Split-Path ([System.IO.Path]::GetFullPath($Path)) -Parent
    $lists = @{}
    $users = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @($manifest.Lists) + @($manifest.Users)) {
        foreach ($name in @('Status', 'Error', 'DataPath', 'RowCount')) {
            if ($entry.PSObject.Properties.Name -notcontains $name) { throw "Reference-data entry lacks '$name'." }
        }
        if ($entry.Status -notin @('Success', 'Failed')) { throw "Invalid reference-data status '$($entry.Status)'." }
        $rows = @()
        if ($entry.Status -eq 'Success') {
            $csv = Resolve-XsnLocalResource -Root $root -RelativePath $entry.DataPath
            $rows = @(Import-Csv -LiteralPath $csv)
            if ($rows.Count -ne $entry.RowCount) { throw "Reference-data row count differs from manifest: $csv" }
        }
        $record = [pscustomobject]@{ Entry = $entry; Rows = $rows; Indexes = @{} }
        if ($entry.PSObject.Properties.Name -contains 'ListId') {
            $key = Get-HtmlReferenceListKey $entry.SiteUrl $entry.ListId
            if ($lists.ContainsKey($key)) { throw "Duplicate reference-data list: $key" }
            $lists[$key] = $record
        }
        else {
            foreach ($name in @('SiteCollectionUrl', 'SourceSiteUrls')) {
                if ($entry.PSObject.Properties.Name -notcontains $name) { throw "User reference-data entry lacks '$name'." }
            }
            if ($entry.SourceSiteUrls -isnot [array]) { throw 'User SourceSiteUrls must be an array.' }
            foreach ($row in $rows) {
                foreach ($name in @('UserId', 'LoginName', 'DisplayName', 'Email')) {
                    if ($row.PSObject.Properties.Name -notcontains $name) { throw "User reference-data CSV lacks '$name'." }
                }
            }
            $users.Add($record)
        }
    }
    $inventory = @(
        foreach ($template in $manifest.Dependencies.Templates) {
            foreach ($name in @('SolutionName', 'SolutionVersion', 'TemplateUrl', 'Sha256', 'Namespaces', 'Bindings', 'UserSiteUrl')) {
                if ($template.PSObject.Properties.Name -notcontains $name) { throw "Reference-data template lacks '$name'." }
            }
            if ($template.Bindings -isnot [array] -or $template.Namespaces -isnot [pscustomobject]) {
                throw 'Reference-data template Bindings and Namespaces must be an array and object.'
            }
            $document = [System.Xml.XmlDocument]::new()
            $manager = New-HtmlReferenceNamespaceManager $document $template
            foreach ($binding in $template.Bindings) {
                foreach ($name in @('XPath', 'ConnectionName', 'SiteUrl', 'ListId', 'ValueField', 'DisplayField')) {
                    if ($binding.PSObject.Properties.Name -notcontains $name -or [string]::IsNullOrWhiteSpace([string]$binding.$name)) {
                        throw "Reference-data binding lacks '$name'."
                    }
                }
                $null = $document.SelectNodes($binding.XPath, $manager)
            }
            [pscustomobject]@{
                Row = [pscustomobject]@{
                    SolutionName = $template.SolutionName; SolutionVersion = $template.SolutionVersion
                    TemplateUrl = $template.TemplateUrl; Modified = ''
                }
                Version = [version]$template.SolutionVersion
                SharePointVersion = [version]'0.0'
                Definition = $template
            }
        }
    )
    return [pscustomobject]@{ Manifest = $manifest; Inventory = $inventory; Lists = $lists; Users = $users }
}

function Get-HtmlReferenceListKey {
    param([string]$SiteUrl, [string]$ListId)

    $uri = [uri]$SiteUrl
    if (-not $uri.IsAbsoluteUri -or $uri.Scheme -notin @('http', 'https')) { throw "Invalid reference site URL: $SiteUrl" }
    return $uri.AbsoluteUri.TrimEnd('/') + '|' + ([guid]$ListId).ToString('D')
}

function New-HtmlReferenceNamespaceManager {
    param([System.Xml.XmlDocument]$Document, $Template)

    $manager = [System.Xml.XmlNamespaceManager]::new($Document.NameTable)
    foreach ($property in $Template.Namespaces.PSObject.Properties) {
        if ($property.Name -notin @('xml', 'xmlns') -and $property.Name) {
            $manager.AddNamespace($property.Name, [string]$property.Value)
        }
    }
    return ,$manager
}

function Get-HtmlReferenceIndex {
    param($Record, [string]$ValueField, [string]$DisplayField, [switch]$IgnoreCase)

    $key = $ValueField + '|' + $DisplayField + '|' + $IgnoreCase.IsPresent
    if (-not $Record.Indexes.ContainsKey($key)) {
        $comparer = if ($IgnoreCase) { [System.StringComparer]::OrdinalIgnoreCase } else { [System.StringComparer]::Ordinal }
        $index = [System.Collections.Generic.Dictionary[string,object]]::new($comparer)
        foreach ($row in $Record.Rows) {
            foreach ($field in @($ValueField, $DisplayField)) {
                if ($row.PSObject.Properties.Name -notcontains $field) { throw "Reference-data CSV lacks requested field '$field'." }
            }
            $value = [string]$row.$ValueField
            $display = [string]$row.$DisplayField
            if ([string]::IsNullOrWhiteSpace($value) -or [string]::IsNullOrWhiteSpace($display)) { continue }
            if (-not $index.ContainsKey($value)) { $index.Add($value, $display) }
            elseif ($index[$value] -cne $display) { $index[$value] = $null }
        }
        $Record.Indexes[$key] = $index
    }
    return ,$Record.Indexes[$key]
}

function Add-HtmlReferenceAudit {
    param($Audit, [string]$SourceFilePath, [string]$XPath, [string]$StoredValue,
        [string]$DisplayValue, [string]$Status, [string]$Source, [string]$Message = '', [string]$ReferenceKey = '')

    $Audit.Add([pscustomobject][ordered]@{
        SourceFilePath = $SourceFilePath; XPath = $XPath; StoredValue = $StoredValue
        DisplayValue = $DisplayValue; Status = $Status; Source = $Source; ReferenceKey = $ReferenceKey; Message = $Message
    })
}

function Set-HtmlReferenceValues {
    param([System.Xml.XmlDocument]$Document, $ReferenceData, $Audit, [string]$SourceFilePath, $XsnMatch)

    if ($null -eq $ReferenceData) { return }
    $inventory = $ReferenceData.Inventory
    if ($null -ne $XsnMatch -and $null -ne $XsnMatch.Template) {
        $selected = $XsnMatch.Template.Row
        $inventory = @($inventory | Where-Object {
            $_.Row.TemplateUrl -ieq $selected.TemplateUrl -and $_.Row.SolutionName -ceq $selected.SolutionName -and
            $_.Version -eq $XsnMatch.Template.Version -and
            ($selected.PSObject.Properties.Name -notcontains 'Sha256' -or -not $selected.Sha256 -or $_.Definition.Sha256 -ieq $selected.Sha256)
        })
    }
    $match = Find-XsnTemplate -Document $Document -Inventory $inventory
    if ($null -eq $match.Template) {
        Add-HtmlReferenceAudit $Audit $SourceFilePath '' '' '' 'NoTemplate' '' $match.Message
        return
    }
    $template = $match.Template.Definition
    $peers = @($inventory | Where-Object {
        $_.Version -eq $match.Template.Version -and $_.Row.TemplateUrl -ieq $template.TemplateUrl -and
        $_.Row.SolutionName -ceq $template.SolutionName
    })
    if (@($peers.Definition.Sha256 | Sort-Object -Unique).Count -gt 1) {
        Add-HtmlReferenceAudit $Audit $SourceFilePath '' '' '' 'Ambiguous' '' 'Conflicting reference definitions for this template version.'
        return
    }
    $manager = New-HtmlReferenceNamespaceManager $Document $template
    $assignments = [System.Collections.Generic.Dictionary[System.Xml.XmlElement,object]]::new()
    foreach ($binding in $template.Bindings) {
        $key = Get-HtmlReferenceListKey $binding.SiteUrl $binding.ListId
        $record = if ($ReferenceData.Lists.ContainsKey($key)) { $ReferenceData.Lists[$key] } else { $null }
        $index = $null
        if ($null -ne $record -and $record.Entry.Status -eq 'Success') {
            $index = Get-HtmlReferenceIndex $record $binding.ValueField $binding.DisplayField
        }
        foreach ($node in $Document.SelectNodes($binding.XPath, $manager)) {
            if ($node -isnot [System.Xml.XmlElement] -or $node.SelectNodes('*').Count -gt 0) {
                Add-HtmlReferenceAudit $Audit $SourceFilePath $binding.XPath '' '' 'Unsupported' $key 'Binding is not a scalar form element.'
                continue
            }
            $value = $node.InnerText
            if ([string]::IsNullOrWhiteSpace($value) -or $value.Contains(';#')) { continue }
            $labels = [System.Collections.Generic.List[string]]::new()
            foreach ($connection in $Document.SelectNodes("//*[local-name()='DataConnection']")) {
                if ($connection.GetAttribute('Name') -cne $binding.ConnectionName) { continue }
                foreach ($row in $connection.SelectNodes('.//*')) {
                    $stored = @($row.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] -and $_.LocalName -ceq $binding.ValueField })
                    $display = @($row.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] -and $_.LocalName -ceq $binding.DisplayField })
                    if ($stored.Count -eq 1 -and $display.Count -eq 1 -and $stored[0].InnerText -ceq $value -and
                        -not [string]::IsNullOrWhiteSpace($display[0].InnerText)) { $labels.Add($display[0].InnerText) }
                }
            }
            $unique = @($labels | Select-Object -Unique)
            $label = $null
            $source = 'SavedCache'
            $status = 'Resolved'
            if ($unique.Count -gt 1) { $status = 'Ambiguous' }
            elseif ($unique.Count -eq 1) { $label = $unique[0] }
            elseif ($null -ne $index -and $index.ContainsKey($value)) {
                $source = 'Snapshot'
                $label = $index[$value]
                if ($null -eq $label) { $status = 'Ambiguous' }
            }
            else {
                $source = 'Snapshot'
                $status = if ($null -eq $record -or $record.Entry.Status -ne 'Success') { 'MissingExport' } else { 'Unresolved' }
            }
            if ($assignments.ContainsKey($node)) {
                if ($assignments[$node].Display -cne $label -or $assignments[$node].Status -ne $status) {
                    $assignments[$node] = [pscustomobject]@{ Display = $null; Status = 'Ambiguous'; Source = $source; XPath = $binding.XPath; ReferenceKey = $key }
                }
            }
            else {
                $assignments.Add($node, [pscustomobject]@{ Display = $label; Status = $status; Source = $source; XPath = $binding.XPath; ReferenceKey = $key })
            }
        }
    }
    foreach ($node in $assignments.Keys) {
        $assignment = $assignments[$node]
        if ($assignment.Status -eq 'Resolved') { $node.SetAttribute('reference-value', 'urn:infopath-archive:display', $assignment.Display) | Out-Null }
        Add-HtmlReferenceAudit $Audit $SourceFilePath $assignment.XPath $node.InnerText $assignment.Display $assignment.Status $assignment.Source -ReferenceKey $assignment.ReferenceKey
    }
    $userRecords = @($ReferenceData.Users | Where-Object {
        @($_.Entry.SourceSiteUrls | Where-Object { $_.TrimEnd('/') -ieq $template.UserSiteUrl.TrimEnd('/') }).Count -gt 0
    })
    foreach ($account in $Document.SelectNodes("//*[local-name()='AccountId' and not(*) and not(ancestor::*[local-name()='DataConnection'])]")) {
        if ([string]::IsNullOrWhiteSpace($account.InnerText)) { continue }
        $savedName = $account.ParentNode.SelectSingleNode("*[local-name()='DisplayName']")
        $display = $null
        $source = 'Snapshot'
        $status = 'Unresolved'
        if ($null -ne $savedName -and -not [string]::IsNullOrWhiteSpace($savedName.InnerText)) {
            $display = $savedName.InnerText; $source = 'SavedDisplayName'; $status = 'Resolved'
        }
        elseif ($userRecords.Count -eq 1 -and $userRecords[0].Entry.Status -eq 'Success') {
            $field = if ($account.InnerText -match '^\d+$') { 'UserId' } else { 'LoginName' }
            $userIndex = Get-HtmlReferenceIndex $userRecords[0] $field 'DisplayName' -IgnoreCase
            $userKey = $account.InnerText
            if ($field -eq 'LoginName' -and -not $userIndex.ContainsKey($userKey)) {
                if ($userKey.StartsWith('i:0#.w|', [System.StringComparison]::OrdinalIgnoreCase)) { $userKey = $userKey.Substring(7) }
                elseif ($userKey -match '^[^|\\]+\\[^|\\]+$') { $userKey = 'i:0#.w|' + $userKey }
            }
            if ($userIndex.ContainsKey($userKey)) {
                $display = $userIndex[$userKey]
                $status = if ($null -eq $display) { 'Ambiguous' } else { 'Resolved' }
            }
        }
        else { $status = if ($userRecords.Count -gt 1) { 'Ambiguous' } else { 'MissingExport' } }
        if ($status -eq 'Resolved') { $account.SetAttribute('reference-value', 'urn:infopath-archive:display', $display) | Out-Null }
        Add-HtmlReferenceAudit $Audit $SourceFilePath (Get-XsnNodePath $account) $account.InnerText $display $status $source -ReferenceKey $template.UserSiteUrl
    }
}
