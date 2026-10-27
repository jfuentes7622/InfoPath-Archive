#requires -Version 5.1
# Local-only discovery helpers. No XSLT, script, or data connection is executed.
function ConvertTo-XsnReferenceSiteUrl {
    param([string]$Url, [string]$TemplateUrl = '')
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -notin @('http', 'https')) {
        $source = $null
        if (-not $Url.StartsWith('/') -or $Url.StartsWith('//') -or
            -not [uri]::TryCreate($TemplateUrl, [UriKind]::Absolute, [ref]$source) -or $source.Scheme -notin @('http', 'https')) {
            throw "Cannot resolve site URL '$Url' against absolute template URL '$TemplateUrl'."
        }
        $uri = [uri]::new($source, $Url)
    }
    if ($uri.Query -or $uri.Fragment -or $uri.UserInfo) { throw "Expected a site/web URL without credentials, query, or fragment: '$Url'." }
    return $uri.AbsoluteUri.TrimEnd('/')
}

function Get-XsnReferenceUserSite {
    param([string]$TemplateUrl, [string]$Override)
    if ($Override) { return ConvertTo-XsnReferenceSiteUrl $Override }
    $url = ConvertTo-XsnReferenceSiteUrl $TemplateUrl
    $uri = [uri]$url
    # Only the standard document-library template location reliably identifies its web.
    if ($uri.AbsolutePath -notmatch '^(.*)/[^/]+/Forms/[^/]+\.xsn$') {
        throw "Cannot reliably derive owning web from '$TemplateUrl'; supply -UserSiteUrl."
    }
    return ($uri.GetLeftPart([UriPartial]::Authority) + $Matches[1]).TrimEnd('/')
}

function ConvertTo-XsnReferenceXPath {
    param([string]$Expression, [System.Xml.XmlElement]$Node, [hashtable]$Namespaces, [string]$Context = '')
    if ($Expression -eq '/') { return '' }
    if ([string]::IsNullOrWhiteSpace($Expression) -or $Expression -match '//|[\[\]\(\)\|\$*]') {
        throw "Unsupported XPath/context '$Expression'."
    }
    $parts = [System.Collections.Generic.List[string]]::new()
    if (-not $Expression.StartsWith('/') -and $Context) {
        foreach ($part in $Context.TrimStart('/').Split('/')) { $parts.Add($part) }
    }
    foreach ($part in $Expression.TrimStart('/').Split('/')) {
        if ($part -in @('', '.')) { continue }
        if ($part -eq '..') {
            if ($parts.Count -eq 0) { throw "XPath '$Expression' escapes its context." }
            $parts.RemoveAt($parts.Count - 1)
            continue
        }
        if ($part -notmatch '^(@?)(?:([A-Za-z_][\w.-]*):)?([A-Za-z_][\w.-]*)$') { throw "Unsupported XPath step '$part'." }
        $attribute = $Matches[1]; $prefix = $Matches[2]; $local = $Matches[3]
        if ($prefix) {
            $uri = $Node.GetNamespaceOfPrefix($prefix)
            if (-not $uri) { throw "Unbound XPath prefix '$prefix'." }
            $candidate = $prefix
            $suffix = 0
            while ($Namespaces.ContainsKey($candidate) -and $Namespaces[$candidate] -cne $uri) { $candidate = 'refns' + $suffix; $suffix++ }
            $Namespaces[$candidate] = $uri
            $parts.Add("$attribute${candidate}:$local")
        }
        else { $parts.Add("$attribute$local") }
    }
    return '/' + ($parts -join '/')
}

function Get-XsnReferenceContext {
    param([System.Xml.XmlElement]$Node, [hashtable]$Namespaces, [string]$MainNamespace, [string]$RootLocalName, [hashtable]$Cache, [int]$Depth = 0)
    if ($Depth -gt 40) { throw 'Cyclic or deeply nested XSL template contexts.' }
    $ancestors = [System.Collections.Generic.List[System.Xml.XmlElement]]::new()
    $parent = $Node.ParentNode
    while ($parent -is [System.Xml.XmlElement]) { $ancestors.Insert(0, $parent); $parent = $parent.ParentNode }
    $context = ''
    $rooted = $Node.GetAttribute('select').StartsWith('/')
    foreach ($ancestor in $ancestors) {
        if ($ancestor.NamespaceURI -ne 'http://www.w3.org/1999/XSL/Transform') { continue }
        $expression = ''
        if ($ancestor.LocalName -eq 'template') {
            $expression = $ancestor.GetAttribute('match')
            if (-not $expression) { throw 'Named XSL templates have no reliable MAIN context.' }
            if (-not $expression.StartsWith('/')) {
                $key = [string]$ancestor.GetHashCode()
                if (-not $Cache.ContainsKey($key)) {
                    $matchPath = ConvertTo-XsnReferenceXPath $expression $ancestor $Namespaces
                    if ($expression -notmatch '^[\w.-]+:[\w.-]+$') { throw "Unsupported template match '$expression'." }
                    $mode = $ancestor.GetAttribute('mode')
                    $entry = '/' + $expression
                    $isRoot = -not $mode -and $expression.Split(':')[1] -ceq $RootLocalName -and
                        $ancestor.GetNamespaceOfPrefix($expression.Split(':')[0]) -ceq $MainNamespace
                    if ($isRoot) { $Cache[$key] = ConvertTo-XsnReferenceXPath $entry $ancestor $Namespaces }
                    else {
                        if (-not $mode) { throw "Relative template '$expression' has no unique mode or declared root." }
                        $calls = @($Node.OwnerDocument.SelectNodes("//*[local-name()='apply-templates' and namespace-uri()='http://www.w3.org/1999/XSL/Transform']") |
                            Where-Object { $_.GetAttribute('mode') -ceq $mode })
                        $paths = @(
                            foreach ($call in $calls) {
                                $callerContext = Get-XsnReferenceContext $call $Namespaces $MainNamespace $RootLocalName $Cache ($Depth + 1)
                                $selected = ConvertTo-XsnReferenceXPath $call.GetAttribute('select') $call $Namespaces $callerContext
                                if ($selected.Split('/')[-1] -cne $matchPath.TrimStart('/')) { throw "Mode '$mode' selects an incompatible template context." }
                                $selected
                            }
                        ) | Sort-Object -Unique
                        if (@($paths).Count -ne 1) { throw "Ambiguous or missing apply-templates context for mode '$mode'." }
                        $Cache[$key] = @($paths)[0]
                    }
                }
                $expression = $Cache[$key]
            }
            $rooted = $true
        }
        if ($ancestor.LocalName -eq 'for-each') { $expression = $ancestor.GetAttribute('select') }
        if ($expression) { $context = ConvertTo-XsnReferenceXPath $expression $ancestor $Namespaces $context }
    }
    if (-not $rooted) { throw 'No rooted MAIN XSL context.' }
    return $context
}

function Get-XsnReferenceMainPath {
    param([System.Xml.XmlElement]$Control, [hashtable]$Namespaces, [string]$MainNamespace, [string]$RootLocalName, [hashtable]$Cache)
    $binding = $Control.GetAttribute('binding', 'http://schemas.microsoft.com/office/infopath/2003')
    $context = ''
    if (-not $binding.StartsWith('/')) { $context = Get-XsnReferenceContext $Control $Namespaces $MainNamespace $RootLocalName $Cache }
    $path = ConvertTo-XsnReferenceXPath $binding $Control $Namespaces $context
    $first = $path.TrimStart('/').Split('/')[0]
    if (-not $MainNamespace -or $first -notmatch '^([^:]+):' -or $Namespaces[$Matches[1]] -cne $MainNamespace) {
        throw "Binding '$path' is not demonstrably rooted in the MAIN document namespace."
    }
    if ($RootLocalName -and $first.Split(':')[-1] -cne $RootLocalName) { throw "Binding '$path' does not start at the declared MAIN root '$RootLocalName'." }
    return $path
}

function Remove-XsnReferencePredicates {
    param([string]$Expression)
    $depth = 0; $quote = ''; $result = ''
    foreach ($c in $Expression.ToCharArray()) {
        if ($quote) { if ([string]$c -eq $quote) { $quote = '' }; if ($depth -eq 0) { $result += $c }; continue }
        if ($c -eq '"' -or $c -eq "'") { $quote = [string]$c; if ($depth -eq 0) { $result += $c }; continue }
        if ($c -eq '[') { $depth++; continue }
        if ($c -eq ']') { $depth--; if ($depth -lt 0) { throw 'Unbalanced row predicate.' }; continue }
        if ($depth -eq 0) { $result += $c }
    }
    if ($depth -ne 0 -or $quote) { throw 'Unbalanced row expression.' }
    return $result.Trim()
}

function Resolve-XsnReferenceRows {
    param([string]$Expression, [System.Xml.XmlElement]$Node, [int]$Depth = 0)
    if ($Depth -gt 6) { throw 'Cyclic or deeply nested row variable.' }
    if ($Expression -match '^\$([\w.-]+)$') {
        $name = $Matches[1]
        $scope = $Node
        $found = $null
        while ($scope -is [System.Xml.XmlElement] -and $null -eq $found) {
            $vars = @($scope.SelectNodes("*[local-name()='variable' and namespace-uri()='http://www.w3.org/1999/XSL/Transform']") | Where-Object { $_.GetAttribute('name') -ceq $name })
            if ($vars.Count -gt 1) { throw "Ambiguous row variable '$name'." }
            if ($vars.Count -eq 1) { $found = $vars[0] }
            $scope = $scope.ParentNode
        }
        if ($null -eq $found) { throw "Unresolved row variable '$name'." }
        $select = $found.GetAttribute('select')
        # InfoPath's generated uniqueness projection preserves the original rows.
        if ($select -match '^([\w.-]+):node-set\(\$([\w.-]+)\)/\*(?:\[.*\])?$') {
            $prefix = $Matches[1]; $variable = $Matches[2]
            if ($found.GetNamespaceOfPrefix($prefix) -cne 'urn:schemas-microsoft-com:xslt') { throw 'Unsupported node-set extension namespace.' }
            return Resolve-XsnReferenceRows ('$' + $variable) $found ($Depth + 1)
        }
        if ($select) { return Resolve-XsnReferenceRows $select $found ($Depth + 1) }
        $copies = @($found.SelectNodes("*[local-name()='copy-of' and namespace-uri()='http://www.w3.org/1999/XSL/Transform']"))
        if ($copies.Count -ne 1) { throw "Unsupported row variable '$name'." }
        return Resolve-XsnReferenceRows $copies[0].GetAttribute('select') $copies[0] ($Depth + 1)
    }
    if ($Expression -notmatch '^([\w.-]+):GetDOM\((["''])(.*?)\2\)(/.*)$') { throw "Unsupported secondary row source '$Expression'." }
    $prefix = $Matches[1]; $connection = $Matches[3]; $rowPath = $Matches[4]
    if ($Node.GetNamespaceOfPrefix($prefix) -cne 'http://schemas.microsoft.com/office/infopath/2003/xslt/xDocument') { throw 'Unsupported GetDOM extension namespace.' }
    return [pscustomobject]@{ ConnectionName = $connection; RowPath = (Remove-XsnReferencePredicates $rowPath); Node = $Node }
}

function Get-XsnReferenceField {
    param([string]$Expression, [System.Xml.XmlElement]$Node, [hashtable]$Namespaces, [string[]]$Fields)
    if ($Expression -notmatch '^@?(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)$') { throw "Unsupported secondary field expression '$Expression'." }
    $field = $Matches[1]
    if ($Fields -cnotcontains $field) { throw "Secondary field '$field' is not a declared internal list field." }
    $path = (ConvertTo-XsnReferenceXPath $Expression $Node $Namespaces).TrimStart('/')
    return [pscustomobject]@{ Field = $field; XPath = $path }
}

function Get-XsnReferenceTemplate {
    param([string]$Directory, [string]$TemplateUrl, [string]$UserSiteUrl, $Metadata, [System.Collections.Generic.List[string]]$Warnings)
    $manifest = Read-XsnXml (Join-Path $Directory 'manifest.xsf')
    $root = $manifest.DocumentElement
    if ($null -eq $root -or $root.LocalName -ne 'xDocumentClass' -or -not $root.GetAttribute('solutionVersion')) {
        throw 'Expected an InfoPath xDocumentClass manifest with solutionVersion.'
    }
    $namespaces = @{}
    foreach ($attribute in $root.Attributes) {
        if ($attribute.Prefix -eq 'xmlns') { $namespaces[$attribute.LocalName] = $attribute.Value }
    }
    $mainNamespace = $root.GetNamespaceOfPrefix('my')
    $rootLocalName = ''
    $rootDeclarations = @(
        foreach ($file in $root.SelectNodes("//*[local-name()='file']")) {
            $nsProperty = $file.SelectSingleNode("*[local-name()='fileProperties']/*[local-name()='property'][@name='namespace']")
            if ($null -ne $nsProperty -and $nsProperty.GetAttribute('value') -ceq $mainNamespace) {
                foreach ($rootProperty in $file.SelectNodes("*[local-name()='fileProperties']/*[local-name()='property'][@name='rootElement']")) { $rootProperty.GetAttribute('value') }
            }
        }
    ) | Sort-Object -Unique
    $rootDeclarations = @($rootDeclarations)
    if ($rootDeclarations.Count -eq 1) { $rootLocalName = $rootDeclarations[0] }
    $template = [pscustomobject][ordered]@{
        SolutionName = $root.GetAttribute('name'); SolutionVersion = $root.GetAttribute('solutionVersion')
        TemplateUrl = $TemplateUrl; Sha256 = ''; Namespaces = $namespaces; Bindings = @(); UserSiteUrl = ''
    }
    if ($null -ne $Metadata) {
        foreach ($property in @('SolutionName', 'SolutionVersion', 'Sha256')) {
            if ($Metadata.PSObject.Properties.Name -contains $property) { $template.$property = [string]$Metadata.$property }
        }
    }
    try { $template.UserSiteUrl = Get-XsnReferenceUserSite $TemplateUrl $UserSiteUrl }
    catch { $Warnings.Add("$TemplateUrl : $($_.Exception.Message)") }
    $connections = @{}; $lists = [System.Collections.Generic.List[object]]::new()
    foreach ($adapter in $root.SelectNodes("//*[local-name()='sharepointListAdapter' or local-name()='sharepointListAdapterRW']")) {
        try {
            $name = $adapter.GetAttribute('name')
            $owner = $adapter.SelectSingleNode("ancestor::*[local-name()='dataObject'][1]")
            if ($null -ne $owner) { $name = $owner.GetAttribute('name') }
            $idText = $adapter.GetAttribute('sharePointListID')
            if (-not $idText) { $idText = $adapter.GetAttribute('listId') }
            $id = ([guid]$idText).ToString('D')
            $site = ConvertTo-XsnReferenceSiteUrl $adapter.GetAttribute('siteURL') $TemplateUrl
            $fields = @($adapter.SelectNodes(".//*[local-name()='field']") | ForEach-Object {
                $internal = $_.GetAttribute('internalName')
                if (-not $internal) { throw "Connection '$name' has a field without internalName." }
                $internal
            } | Sort-Object -Unique)
            $list = [pscustomobject]@{ SiteUrl = $site; ListId = $id; Fields = @($fields + 'ID' | Sort-Object -Unique) }
            $lists.Add($list)
            if (-not $name) { $Warnings.Add("$TemplateUrl : unnamed list connection exported, but cannot bind controls."); continue }
            if ($connections.ContainsKey($name)) {
                $connections[$name] = $null
                $Warnings.Add("$TemplateUrl : duplicate connection name '$name'; all its control bindings rejected.")
            }
            else { $connections[$name] = $list }
        }
        catch { $Warnings.Add("$TemplateUrl : rejected list connection: $($_.Exception.Message)") }
    }
    foreach ($adapter in $root.SelectNodes("//*[local-name()='dataObject']/*[local-name()='query']/*")) {
        if ($adapter.LocalName -notin @('sharepointListAdapter', 'sharepointListAdapterRW')) {
            $Warnings.Add("$TemplateUrl : unsupported connection adapter '$($adapter.LocalName)' (not exported).")
        }
    }
    $bindings = [System.Collections.Generic.List[object]]::new()
    $hasFailures = $false
    $viewIndex = 0
    foreach ($view in $root.SelectNodes("*[local-name()='views']/*[local-name()='view']")) {
        $viewIndex++
        try {
            $pane = $view.SelectSingleNode("*[local-name()='mainpane']")
            if ($null -eq $pane) { throw 'Declared view has no mainpane.' }
            $resource = $pane.GetAttribute('transform').Replace('/', '\')
            if (-not $resource -or [IO.Path]::IsPathRooted($resource) -or '..' -in $resource.Split('\')) { throw "Unsafe view path '$resource'." }
            $viewPath = Join-Path $Directory ('View{0:D4}.xsl' -f $viewIndex)
            if (-not (Test-Path -LiteralPath $viewPath)) { $viewPath = Join-Path $Directory $resource }
            $document = Read-XsnXml $viewPath
            $manager = [System.Xml.XmlNamespaceManager]::new($document.NameTable)
            $manager.AddNamespace('xd', 'http://schemas.microsoft.com/office/infopath/2003')
            $manager.AddNamespace('xsl', 'http://www.w3.org/1999/XSL/Transform')
            $contextCache = @{}
            foreach ($control in $document.SelectNodes("//*[@xd:binding][local-name()='select' or @xd:xctname='dropdown' or @xd:xctname='combobox' or @xd:xctname='listbox']", $manager)) {
                # Literal choices need no external lookup, but every dynamic source is diagnosed.
                if ($control.OuterXml -notmatch 'GetDOM' -and $control.SelectNodes('.//xsl:for-each', $manager).Count -eq 0) { continue }
                try {
                    $mainPath = Get-XsnReferenceMainPath $control $namespaces $mainNamespace $rootLocalName $contextCache
                    $candidates = [System.Collections.Generic.List[object]]::new()
                    foreach ($loop in $control.SelectNodes('.//xsl:for-each', $manager)) {
                        $options = @($loop.SelectNodes("./*[local-name()='option']"))
                        if ($options.Count -eq 0) { continue }
                        $rows = Resolve-XsnReferenceRows $loop.GetAttribute('select') $loop
                        if (-not $connections.ContainsKey($rows.ConnectionName) -or $null -eq $connections[$rows.ConnectionName]) { throw "Unresolved or ambiguous list connection '$($rows.ConnectionName)'." }
                        $list = $connections[$rows.ConnectionName]
                        foreach ($option in $options) {
                            $values = @($option.SelectNodes('./xsl:attribute[@name="value"]/xsl:value-of', $manager))
                            $displays = @($option.SelectNodes('./xsl:value-of', $manager))
                            if ($values.Count -ne 1 -or $displays.Count -ne 1 -or $option.HasAttribute('value')) { throw 'Ambiguous or computed option value/display (no label guessed).' }
                            $valueNodes = @($option.SelectNodes('./xsl:attribute[@name="value"]/*', $manager))
                            $valueText = @($option.SelectNodes('./xsl:attribute[@name="value"]/text()', $manager) | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Value) })
                            $displayNodes = @($option.ChildNodes | Where-Object {
                                ($_ -is [System.Xml.XmlElement] -and -not ($_.NamespaceURI -eq 'http://www.w3.org/1999/XSL/Transform' -and $_.LocalName -in @('attribute', 'if', 'value-of'))) -or
                                ($_ -isnot [System.Xml.XmlElement] -and -not [string]::IsNullOrWhiteSpace($_.InnerText))
                            })
                            $conditionalOutput = @($option.SelectNodes('./xsl:if/*', $manager) | Where-Object {
                                $_.NamespaceURI -ne 'http://www.w3.org/1999/XSL/Transform' -or $_.LocalName -ne 'attribute' -or $_.GetAttribute('name') -notin @('selected', 'disabled')
                            })
                            $conditionalText = @($option.SelectNodes('./xsl:if/text()', $manager) | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Value) })
                            if ($valueNodes.Count -ne 1 -or $valueText.Count -gt 0 -or $displayNodes.Count -gt 0 -or $conditionalOutput.Count -gt 0 -or $conditionalText.Count -gt 0) { throw 'Composite option labels/values are not supported (no label guessed).' }
                            $value = Get-XsnReferenceField $values[0].GetAttribute('select') $values[0] $namespaces $list.Fields
                            $display = Get-XsnReferenceField $displays[0].GetAttribute('select') $displays[0] $namespaces $list.Fields
                            $cachePath = ''
                            try { $cachePath = ConvertTo-XsnReferenceXPath $rows.RowPath $rows.Node $namespaces }
                            catch { $Warnings.Add("$TemplateUrl : cache row XPath unavailable: $($_.Exception.Message)") }
                            $candidates.Add([pscustomobject][ordered]@{
                                XPath = $mainPath; ConnectionName = $rows.ConnectionName; SiteUrl = $list.SiteUrl; ListId = $list.ListId
                                ValueField = $value.Field; DisplayField = $display.Field
                                CacheRowXPath = $cachePath; ValueXPath = $value.XPath; DisplayXPath = $display.XPath
                            })
                        }
                    }
                    if ($candidates.Count -eq 0) { throw 'No supported dynamic option row source.' }
                    $distinct = @($candidates | ForEach-Object { $_ | ConvertTo-Json -Compress } | Sort-Object -Unique)
                    if ($distinct.Count -ne 1) { throw 'Conflicting dynamic option mappings in a single control.' }
                    $bindings.Add($candidates[0])
                }
                catch { $Warnings.Add("$TemplateUrl view '$($view.GetAttribute('name'))' binding '$($control.GetAttribute('binding','http://schemas.microsoft.com/office/infopath/2003'))': rejected: $($_.Exception.Message)") }
            }
        }
        catch { $Warnings.Add("$TemplateUrl view '$($view.GetAttribute('name'))': failed: $($_.Exception.Message)"); $hasFailures = $true }
    }
    # A field whose views disagree must not be resolved using whichever view happened to run first.
    $template.Bindings = @(
        foreach ($group in ($bindings | Group-Object XPath)) {
            $unique = @($group.Group | ForEach-Object { $_ | ConvertTo-Json -Compress } | Sort-Object -Unique)
            if ($unique.Count -eq 1) { $group.Group[0] }
            else { $Warnings.Add("$TemplateUrl : ambiguous mappings across views for '$($group.Name)' (all rejected).") }
        }
    )
    return [pscustomobject]@{ Template = $template; Lists = $lists.ToArray(); HasFailures = $hasFailures }
}
