function Read-XsnXml {
    param([string]$Path)

    $settings = [System.Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $reader = [System.Xml.XmlReader]::Create($Path, $settings)
    try {
        $document = [System.Xml.XmlDocument]::new()
        $document.XmlResolver = $null
        $document.Load($reader)
        return ,$document
    }
    finally {
        $reader.Dispose()
    }
}

function Get-XsnNodePath {
    param([System.Xml.XmlNode]$Node)

    $parts = [System.Collections.Generic.List[string]]::new()
    while ($Node -is [System.Xml.XmlElement]) {
        $parts.Insert(0, ([uri]::EscapeDataString($Node.get_NamespaceURI()) + ':' + $Node.get_LocalName()))
        $Node = $Node.get_ParentNode()
    }
    return '/' + ($parts -join '/')
}

function Convert-XsnBindingPath {
    param([string]$Expression, [System.Xml.XmlElement]$Node, [string]$Context = '')

    $expression = $Expression.Trim() -replace '\[\d+\]', ''
    if ($expression -match '//|[\[\]\(\)\|\$]' -or [string]::IsNullOrWhiteSpace($expression)) {
        return $null
    }
    $parts = [System.Collections.Generic.List[string]]::new()
    if (-not $expression.StartsWith('/') -and $Context) {
        foreach ($part in $Context.TrimStart('/').Split('/')) { $parts.Add($part) }
    }
    foreach ($part in $expression.TrimStart('/').Split('/')) {
        if ($part -eq '.' -or $part -eq '') { continue }
        if ($part -eq '..') {
            if ($parts.Count -eq 0) { return $null }
            $parts.RemoveAt($parts.Count - 1)
            continue
        }
        $qualified = $part.Split(':')
        if ($qualified.Count -gt 2) { return $null }
        $prefix = if ($qualified.Count -eq 2) { $qualified[0] } else { '' }
        $local = $qualified[$qualified.Count - 1]
        try {
            $null = [System.Xml.XmlConvert]::VerifyNCName($local)
            if ($prefix) { $null = [System.Xml.XmlConvert]::VerifyNCName($prefix) }
        }
        catch [System.Xml.XmlException] {
            return $null
        }
        $namespace = if ($prefix) { $Node.GetNamespaceOfPrefix($prefix) } else { '' }
        if ($prefix -and -not $namespace) { return $null }
        $parts.Add(([uri]::EscapeDataString($namespace) + ':' + $local))
    }
    if ($parts.Count -eq 0) { return '' }
    return '/' + ($parts -join '/')
}

function Get-XsnControlPath {
    param([System.Xml.XmlElement]$Control)

    $ancestors = [System.Collections.Generic.List[System.Xml.XmlElement]]::new()
    $parent = $Control.get_ParentNode()
    while ($parent -is [System.Xml.XmlElement]) {
        $ancestors.Insert(0, $parent)
        $parent = $parent.get_ParentNode()
    }
    $context = ''
    foreach ($ancestor in $ancestors) {
        if ($ancestor.get_NamespaceURI() -ne 'http://www.w3.org/1999/XSL/Transform') { continue }
        $expression = switch ($ancestor.get_LocalName()) {
            'template' { $ancestor.GetAttribute('match') }
            'for-each' { $ancestor.GetAttribute('select') }
            default { '' }
        }
        if (-not $expression) { continue }
        $context = Convert-XsnBindingPath -Expression $expression -Node $ancestor -Context $context
        if ($null -eq $context) { return $null }
    }
    return Convert-XsnBindingPath -Expression $Control.GetAttribute('binding', 'http://schemas.microsoft.com/office/infopath/2003') -Node $Control -Context $context
}

function Get-XsnStaticText {
    param([System.Xml.XmlNode]$Node)

    if ($Node -is [System.Xml.XmlElement]) {
        if ($Node.HasAttribute('binding', 'http://schemas.microsoft.com/office/infopath/2003') -or
            $Node.get_LocalName() -in @('script', 'style', 'input', 'select', 'textarea') -or
            ($Node.get_NamespaceURI() -eq 'http://www.w3.org/1999/XSL/Transform' -and $Node.get_LocalName() -ne 'text')) {
            return ''
        }
    }
    if ($Node.NodeType -in @([System.Xml.XmlNodeType]::Text, [System.Xml.XmlNodeType]::CDATA, [System.Xml.XmlNodeType]::SignificantWhitespace)) {
        return $Node.Value
    }
    $text = @($Node.ChildNodes | ForEach-Object { Get-XsnStaticText -Node $_ }) -join ''
    return ($text -replace '\s+', ' ').Trim()
}

function Get-XsnAncestor {
    param([System.Xml.XmlNode]$Node, [string[]]$Names)
    $parent = $Node.get_ParentNode()
    while ($parent -is [System.Xml.XmlElement]) {
        if ($parent.get_LocalName() -in $Names) { return $parent }
        $parent = $parent.get_ParentNode()
    }
    return $null
}

function Get-XsnControlLabel {
    param([System.Xml.XmlElement]$Control, [System.Xml.XmlNamespaceManager]$Manager)

    $aria = $Control.GetAttribute('aria-label')
    if (-not [string]::IsNullOrWhiteSpace($aria)) {
        return [pscustomobject]@{ Label = $aria.Trim(); Source = 'aria-label' }
    }
    $id = $Control.GetAttribute('id')
    if ($id) {
        $labels = @($Control.OwnerDocument.SelectNodes("//*[local-name()='label'][@for]") | Where-Object { $_.GetAttribute('for') -ceq $id })
        if ($labels.Count -eq 1) {
            $text = Get-XsnStaticText -Node $labels[0]
            if ($text) { return [pscustomobject]@{ Label = $text; Source = 'label-for' } }
        }
    }
    $labelledBy = $Control.GetAttribute('aria-labelledby')
    if ($labelledBy) {
        $texts = @(
            foreach ($labelId in ($labelledBy -split '\s+')) {
                $nodes = @($Control.OwnerDocument.SelectNodes('//*[@id]') | Where-Object { $_.GetAttribute('id') -ceq $labelId })
                if ($nodes.Count -ne 1) { return $null }
                Get-XsnStaticText -Node $nodes[0]
            }
        )
        if ($texts.Count -gt 0 -and -not ($texts -contains '')) {
            return [pscustomobject]@{ Label = $texts -join ' '; Source = 'aria-labelledby' }
        }
    }
    $label = Get-XsnAncestor -Node $Control -Names @('label')
    if ($null -ne $label) {
        $text = Get-XsnStaticText -Node $label
        if ($text) { return [pscustomobject]@{ Label = $text; Source = 'enclosing-label' } }
    }
    $previousLabel = $Control.PreviousSibling
    while ($null -ne $previousLabel -and $previousLabel -isnot [System.Xml.XmlElement]) { $previousLabel = $previousLabel.PreviousSibling }
    if ($null -ne $previousLabel -and $previousLabel.get_LocalName() -eq 'label' -and -not $previousLabel.GetAttribute('for')) {
        $text = Get-XsnStaticText -Node $previousLabel
        if ($text) { return [pscustomobject]@{ Label = $text; Source = 'preceding-label' } }
    }
    $cell = Get-XsnAncestor -Node $Control -Names @('td', 'th')
    if ($null -eq $cell) { return $null }
    $controls = @($cell.SelectNodes('.//*[@xd:binding]', $Manager))
    if ($controls.Count -eq 1) {
        $text = Get-XsnStaticText -Node $cell
        if ($text) { return [pscustomobject]@{ Label = $text; Source = 'same-cell' } }
    }
    if ($controls.Count -ne 1) { return $null }
    $previous = $cell.PreviousSibling
    while ($null -ne $previous -and $previous -isnot [System.Xml.XmlElement]) { $previous = $previous.PreviousSibling }
    if ($null -ne $previous -and $previous.get_LocalName() -in @('td', 'th') -and
        @($previous.SelectNodes('.//*[@xd:binding]', $Manager)).Count -eq 0) {
        $text = Get-XsnStaticText -Node $previous
        if ($text) { return [pscustomobject]@{ Label = $text; Source = 'preceding-cell' } }
    }
    $table = Get-XsnAncestor -Node $cell -Names @('table')
    if ($null -eq $table) { return $null }
    $cells = @($cell.ParentNode.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] -and $_.get_LocalName() -in @('td', 'th') })
    if (@($cells | Where-Object { $_.GetAttribute('colspan') -notin @('', '1') -or $_.GetAttribute('rowspan') -notin @('', '1') }).Count -gt 0) { return $null }
    $column = [array]::IndexOf($cells, $cell)
    $headers = @(
        foreach ($row in $table.SelectNodes(".//*[local-name()='tr']")) {
            if ((Get-XsnAncestor -Node $row -Names @('table')) -ne $table) { continue }
            $rowCells = @($row.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] -and $_.get_LocalName() -in @('td', 'th') })
            if ($rowCells.Count -ne $cells.Count -or
                @($rowCells | Where-Object { $_.GetAttribute('colspan') -notin @('', '1') -or $_.GetAttribute('rowspan') -notin @('', '1') }).Count -gt 0 -or
                @($row.SelectNodes('.//*[@xd:binding]', $Manager)).Count -gt 0 -or
                $row.CreateNavigator().ComparePosition($cell.ParentNode.CreateNavigator()) -ne [System.Xml.XmlNodeOrder]::Before) { continue }
            if ($column -ge 0 -and $column -lt $rowCells.Count) {
                $text = Get-XsnStaticText -Node $rowCells[$column]
                if ($text) { $text }
            }
        }
    ) | Select-Object -Unique
    if (@($headers).Count -eq 1) {
        return [pscustomobject]@{ Label = @($headers)[0]; Source = 'table-header' }
    }
    return $null
}

function Get-XsnViewLabelCandidates {
    param([System.Xml.XmlDocument]$Document, [string]$ViewName, [bool]$IsDefault)

    $manager = [System.Xml.XmlNamespaceManager]::new($Document.NameTable)
    $manager.AddNamespace('xd', 'http://schemas.microsoft.com/office/infopath/2003')
    foreach ($control in $Document.SelectNodes('//*[@xd:binding]', $manager)) {
        $path = Get-XsnControlPath -Control $control
        $label = Get-XsnControlLabel -Control $control -Manager $manager
        [pscustomobject]@{
            Path = $path
            Binding = $control.GetAttribute('binding', 'http://schemas.microsoft.com/office/infopath/2003')
            Label = if ($null -ne $label) { $label.Label } else { '' }
            Source = if ($null -ne $label) { $label.Source } else { '' }
            ViewName = $ViewName
            IsDefault = $IsDefault
            Status = if ($null -eq $path -or -not $path) { 'UnsupportedBinding' } elseif ($null -eq $label) { 'NoLabel' } else { 'Candidate' }
        }
    }
}

function Resolve-XsnViewLabels {
    param([AllowEmptyCollection()][object[]]$Candidates)

    foreach ($group in @($Candidates | Where-Object Status -eq 'Candidate' | Group-Object -Property Path -CaseSensitive)) {
        $preferred = @($group.Group | Where-Object IsDefault)
        if ($preferred.Count -eq 0) { $preferred = @($group.Group) }
        $names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        foreach ($candidate in $preferred) { $null = $names.Add($candidate.Label) }
        [pscustomobject]@{
            Path = $group.Name
            Label = if ($names.Count -eq 1) { $preferred[0].Label } else { '' }
            ViewName = ($preferred.ViewName | Select-Object -Unique) -join '; '
            Source = ($preferred.Source | Select-Object -Unique) -join '; '
            Status = if ($names.Count -eq 1) { 'Resolved' } else { 'Ambiguous' }
        }
    }
}

function Import-XsnInventory {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return ,@() }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "XSN inventory not found: $Path" }
    $inventory = @(Import-Csv -LiteralPath $Path)
    if ($inventory.Count -eq 0) { throw "XSN inventory is empty: $Path" }
    foreach ($column in @('TemplateUrl', 'SolutionVersion', 'SolutionName', 'LabelsRelativePath', 'SharePointVersion', 'Modified', 'Status')) {
        if ($inventory[0].PSObject.Properties.Name -notcontains $column) {
            throw "XSN inventory lacks '$column'. Rerun pipeline step 1 to extract XSN labels."
        }
    }
    $root = Split-Path -Path ([System.IO.Path]::GetFullPath($Path)) -Parent
    $records = @(
        foreach ($row in ($inventory | Where-Object Status -eq 'Success')) {
            $version = $null
            $spVersion = $null
            if (-not [version]::TryParse($row.SolutionVersion, [ref]$version) -or -not [version]::TryParse($row.SharePointVersion, [ref]$spVersion)) {
                throw "Invalid version in XSN inventory for '$($row.TemplateUrl)'."
            }
            if ([string]::IsNullOrWhiteSpace($row.LabelsRelativePath) -or [System.IO.Path]::IsPathRooted($row.LabelsRelativePath)) {
                throw "Invalid relative XSN label-map path: '$($row.LabelsRelativePath)'"
            }
            $labelsPath = [System.IO.Path]::GetFullPath((Join-Path -Path $root -ChildPath $row.LabelsRelativePath))
            if (-not $labelsPath.StartsWith($root.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "XSN label-map path escapes its inventory folder: '$labelsPath'"
            }
            if (-not (Test-Path -LiteralPath $labelsPath -PathType Leaf)) { throw "XSN label map is missing: $labelsPath" }
            $configuration = Get-Content -LiteralPath $labelsPath -Raw | ConvertFrom-Json
            if ($null -eq $configuration -or $configuration.PSObject.Properties.Name -notcontains 'Labels' -or $configuration.Labels -isnot [array]) {
                throw "Invalid XSN label map: $labelsPath"
            }
            $labels = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::Ordinal)
            foreach ($label in $configuration.Labels) {
                if ($label.PSObject.Properties.Name -notcontains 'Path' -or $label.PSObject.Properties.Name -notcontains 'Label' -or
                    [string]::IsNullOrWhiteSpace($label.Path) -or [string]::IsNullOrWhiteSpace($label.Label)) {
                    throw "Invalid field label in '$labelsPath'."
                }
                $labels.Add($label.Path, $label.Label)
            }
            [pscustomobject]@{
                Row = $row
                Version = $version
                SharePointVersion = $spVersion
                Labels = $labels
                LabelsPath = $labelsPath
                InventoryRoot = $root
            }
        }
    )
    if ($records.Count -eq 0) { throw "XSN inventory contains no successful templates: $Path" }
    return ,$records
}

function Get-XsnSolutionAttribute {
    param([System.Xml.XmlDocument]$Document, [string]$Name)
    $instruction = $Document.SelectSingleNode("/processing-instruction('mso-infoPathSolution')")
    if ($null -eq $instruction) { return '' }
    $pattern = '(?i)(?:^|\s){0}\s*=\s*(["''])(.*?)\1' -f [regex]::Escape($Name)
    $match = [regex]::Match($instruction.Value, $pattern)
    if (-not $match.Success) { return '' }
    return [System.Net.WebUtility]::HtmlDecode($match.Groups[2].Value).Trim()
}

function Find-XsnTemplate {
    param([System.Xml.XmlDocument]$Document, [AllowEmptyCollection()][object[]]$Inventory)

    $requested = Get-XsnSolutionAttribute -Document $Document -Name 'solutionVersion'
    $name = Get-XsnSolutionAttribute -Document $Document -Name 'name'
    $href = Get-XsnSolutionAttribute -Document $Document -Name 'href'
    $result = [pscustomobject]@{ RequestedVersion = $requested; SolutionName = $name; Href = $href; Status = 'NoMatch'; Template = $null; Message = '' }
    if ($Inventory.Count -eq 0) {
        $result.Status = 'NoInventory'
        $result.Message = 'No XSN inventory supplied; using XML-name labels.'
        return $result
    }
    $version = $null
    if (-not [version]::TryParse($requested, [ref]$version)) {
        $result.Message = "No usable solutionVersion '$requested'; using XML-name labels."
        return $result
    }
    $candidates = @($Inventory)
    if ($name) {
        $candidates = @($candidates | Where-Object { $_.Row.SolutionName -ceq $name })
    }
    elseif (-not $href) {
        $result.Message = 'No form identity or template href; using XML-name labels.'
        return $result
    }
    if ($href -and $candidates.Count -gt 0) {
        $exactUrls = @($candidates | Where-Object {
            $_.Row.TemplateUrl -ieq $href -or ($href.StartsWith('/') -and ([uri]$_.Row.TemplateUrl).AbsolutePath -ieq $href)
        })
        if (-not $name -and $exactUrls.Count -gt 0) { $candidates = $exactUrls }
        elseif (-not $name) {
            $leaf = [System.IO.Path]::GetFileName(($href -split '\?')[0])
            $byName = @($candidates | Where-Object { [uri]::UnescapeDataString([System.IO.Path]::GetFileName(([uri]$_.Row.TemplateUrl).AbsolutePath)) -ieq [uri]::UnescapeDataString($leaf) })
            $families = @($byName | Group-Object { $_.Row.TemplateUrl + '|' + $_.Row.SolutionName })
            if ($families.Count -eq 1) { $candidates = $byName } else { $candidates = @() }
        }
    }
    $eligible = @($candidates | Where-Object { $_.Version -ge $version } | Sort-Object Version)
    if ($eligible.Count -eq 0) {
        $result.Message = "No same-form XSN at solutionVersion $requested or higher; using XML-name labels."
        return $result
    }
    $nearestVersion = $eligible[0].Version
    $nearest = @($eligible | Where-Object { $_.Version -eq $nearestVersion })
    if ($href) {
        $atHref = @($nearest | Where-Object { $_.Row.TemplateUrl -ieq $href -or ($href.StartsWith('/') -and ([uri]$_.Row.TemplateUrl).AbsolutePath -ieq $href) })
        if ($atHref.Count -gt 0) { $nearest = $atHref }
    }
    $identities = @($nearest | Group-Object { $_.Row.SolutionName })
    if ($identities.Count -ne 1) {
        $result.Message = "Ambiguous form identity at solutionVersion $nearestVersion; using XML-name labels."
        return $result
    }
    $selected = @($nearest | Sort-Object @(
        @{ Expression = { $_.SharePointVersion }; Descending = $true },
        @{ Expression = { $_.Row.Modified }; Descending = $true },
        @{ Expression = { $_.Row.TemplateUrl }; Descending = $false }
    ))[0]
    $result.Template = $selected
    $result.Status = if ($nearestVersion -eq $version) { 'ExactVersion' } else { 'NextVersion' }
    $result.Message = "$($result.Status): requested $requested; using $($selected.Row.SolutionVersion) from $($selected.Row.TemplateUrl)."
    return $result
}

function Set-XsnDisplayLabels {
    param([System.Xml.XmlDocument]$Document, $Match)

    $count = 0
    foreach ($node in $Document.SelectNodes('//*')) {
        $node.RemoveAttribute('label', 'urn:infopath-archive:display')
        if ($null -eq $Match.Template) { continue }
        $path = Get-XsnNodePath -Node $node
        if ($Match.Template.Labels.ContainsKey($path)) {
            $null = $node.SetAttribute('label', 'urn:infopath-archive:display', $Match.Template.Labels[$path])
            $count++
        }
    }
    return $count
}
