function Resolve-XsnLocalResource {
    param([string]$Root, [string]$RelativePath)

    $relative = [uri]::UnescapeDataString($RelativePath).Replace('/', '\')
    if ([string]::IsNullOrWhiteSpace($relative) -or [System.IO.Path]::IsPathRooted($relative) -or
        $relative.Contains(':') -or '..' -in $relative.Split('\')) {
        throw "XSN resource must be a local relative path: '$RelativePath'."
    }
    $rootPath = [System.IO.Path]::GetFullPath($Root).TrimEnd('\')
    $path = [System.IO.Path]::GetFullPath((Join-Path $rootPath $relative))
    if (-not $path.StartsWith($rootPath + '\', [System.StringComparison]::OrdinalIgnoreCase) -or
        -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "XSN rendering resource is missing or outside its package: '$RelativePath'."
    }
    $inspection = $path
    while ($inspection -ne $rootPath) {
        if ((Get-Item -LiteralPath $inspection -Force).Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            throw "XSN resources cannot traverse symbolic links: '$RelativePath'."
        }
        $inspection = Split-Path $inspection -Parent
    }
    return $path
}

function Remove-XsnHiddenStyle {
    param([string]$Style)

    $clean = [regex]::Replace($Style, '(?i)behavior\s*:[^;}]*', '')
    return [regex]::Replace($clean, '(?i)(?:^|;)\s*(?:display\s*:\s*none|visibility\s*:\s*(?:hidden|collapse))\s*(?:!important)?\s*(?=;|$)', '')
}

function Convert-XsnArchiveValueStyle {
    param([string]$Style, [switch]$Selection)

    $properties = 'border(?:-[a-z-]+)?|background(?:-[a-z-]+)?|box-shadow|outline(?:-[a-z-]+)?|(?:-webkit-|-moz-)?appearance|padding(?:-[a-z-]+)?|margin(?:-[a-z-]+)?|height|min-height|max-height|overflow(?:-[xy])?|white-space|text-overflow|opacity|line-height'
    if ($Selection) { $properties += '|width|min-width|max-width' }
    return [regex]::Replace($Style, "(?i)(?:^|;)\s*(?:$properties)\s*:[^;]*(?=;|$)", '')
}

function New-XsnArchiveStylesheet {
    param([System.Xml.XmlDocument]$View)

    $xsl = 'http://www.w3.org/1999/XSL/Transform'
    $xd = 'http://schemas.microsoft.com/office/infopath/2003'
    $display = 'urn:infopath-archive:display'
    $manager = [System.Xml.XmlNamespaceManager]::new($View.NameTable)
    $manager.AddNamespace('xsl', $xsl)
    $manager.AddNamespace('xd', $xd)
    $root = $View.DocumentElement
    if ($root.NamespaceURI -ne $xsl -or $root.LocalName -notin @('stylesheet', 'transform')) {
        throw 'XSN view is not an XSLT stylesheet.'
    }
    if ($View.SelectNodes('//xsl:include | //xsl:import', $manager).Count -gt 0) {
        throw 'XSN views with XSLT includes/imports are not yet supported.'
    }
    $null = $root.SetAttribute('xmlns:archive', $display)
    $excluded = $root.GetAttribute('exclude-result-prefixes')
    $root.SetAttribute('exclude-result-prefixes', ($excluded + ' archive').Trim())

    # Replace bound UI before inspecting expressions, eliminating live lookup and calculation code.
    foreach ($control in @($View.SelectNodes('//*[@xd:binding]', $manager))) {
        if ($null -eq $control.ParentNode -or $control.NamespaceURI -eq $xsl) { continue }
        if ($control.LocalName -in @('table', 'tbody', 'tr', 'td', 'div') -and
            $control.SelectNodes('.//*[@xd:binding]', $manager).Count -gt 0) { continue }
        $binding = $control.GetAttribute('binding', $xd)
        if ([string]::IsNullOrWhiteSpace($binding)) { throw 'XSN control has an empty binding.' }
        $replacement = $View.CreateElement('span', $control.NamespaceURI)
        foreach ($attribute in @($control.Attributes)) {
            if ($attribute.NamespaceURI -eq '' -and $attribute.Name -in @('class', 'style', 'title', 'id')) {
                $replacement.SetAttribute($attribute.Name, $attribute.Value)
            }
        }
        $replacement.SetAttribute('class', ($replacement.GetAttribute('class') + ' archive-value').Trim())
        $type = $control.GetAttribute('type').ToLowerInvariant()
        $isSelection = $control.LocalName -eq 'input' -and $type -in @('checkbox', 'radio')
        $replacement.SetAttribute('style', (Convert-XsnArchiveValueStyle -Style $replacement.GetAttribute('style') -Selection:$isSelection))
        if ($isSelection) {
            $replacement.SetAttribute('class', ($replacement.GetAttribute('class') + ' archive-selection'))
            $condition = $control.SelectSingleNode(".//xsl:if[xsl:attribute[translate(@name,'CHECKED','checked')='checked']]", $manager)
            if ($null -ne $condition) {
                $test = $condition.GetAttribute('test')
            }
            else {
                $onValue = $control.GetAttribute('onValue', $xd)
                if (-not $onValue) { $onValue = $control.GetAttribute('value') }
                if (-not $onValue) { $onValue = 'true' }
                if ($onValue.Contains("'")) { throw 'Unsupported quote in XSN selection value.' }
                $test = "$binding = '$onValue'"
            }
            $choose = $View.CreateElement('xsl', 'choose', $xsl)
            $when = $View.CreateElement('xsl', 'when', $xsl)
            $when.SetAttribute('test', $test)
            $when.InnerText = if ($type -eq 'radio') { '(X)' } else { '[X]' }
            $otherwise = $View.CreateElement('xsl', 'otherwise', $xsl)
            $otherwise.InnerText = if ($type -eq 'radio') { '( )' } else { '[ ]' }
            $null = $choose.AppendChild($when)
            $null = $choose.AppendChild($otherwise)
            $null = $replacement.AppendChild($choose)
        }
        else {
            $apply = $View.CreateElement('xsl', 'apply-templates', $xsl)
            $apply.SetAttribute('select', $binding)
            $apply.SetAttribute('mode', 'archive-field')
            $null = $replacement.AppendChild($apply)
        }
        $null = $control.ParentNode.ReplaceChild($replacement, $control)
    }

    foreach ($node in @($View.SelectNodes('//*'))) {
        if ($null -eq $node.ParentNode) { continue }
        if (($node.NamespaceURI -ne $xsl -and $node.LocalName -in @('script', 'object', 'embed', 'iframe', 'button', 'input', 'select', 'textarea')) -or
            ($node.NamespaceURI -eq 'urn:schemas-microsoft-com:xslt' -and $node.LocalName -eq 'script')) {
            $null = $node.ParentNode.RemoveChild($node)
            continue
        }
        foreach ($attribute in @($node.Attributes)) {
            if ($attribute.Name -match '^(?i:on)' -or $attribute.LocalName -eq 'hidden') {
                $null = $node.Attributes.Remove($attribute)
            }
        }
        if ($node.NamespaceURI -ne $xsl -and $node.HasAttribute('style')) {
            $node.SetAttribute('style', (Remove-XsnHiddenStyle $node.GetAttribute('style')))
        }
        if ($node.NamespaceURI -ne $xsl -and $node.GetAttribute('class') -match '\bxd(?:Section|RepeatingSection|Repeating|LayoutRegion)\b') {
            $sectionStyle = [regex]::Replace($node.GetAttribute('style'), '(?i)(^|;)\s*height\s*:', '$1min-height:')
            $sectionStyle = [regex]::Replace($sectionStyle, '(?i)(?:^|;)\s*(?:max-height|overflow(?:-[xy])?)\s*:[^;]*(?=;|$)', '')
            $node.SetAttribute('style', $sectionStyle)
            $node.SetAttribute('class', ($node.GetAttribute('class') + ' archive-section'))
        }
    }
    foreach ($attribute in @($View.SelectNodes("//xsl:attribute[translate(@name,'STYLE','style')='style']", $manager))) {
        $text = @($attribute.ChildNodes | Where-Object NodeType -eq ([System.Xml.XmlNodeType]::Text) | ForEach-Object Value) -join ''
        $attribute.InnerText = Remove-XsnHiddenStyle $text
    }
    foreach ($style in @($View.SelectNodes("//*[local-name()='style' and namespace-uri()!='$xsl']"))) {
        $style.InnerText = [regex]::Replace($style.InnerText, '(?i)(display\s*:\s*none|visibility\s*:\s*(hidden|collapse))\s*(!important)?\s*;?|behavior\s*:[^;}]*', '')
    }
    # Visibility gates around form content must not hide archival sections.
    foreach ($choose in @($View.SelectNodes('//xsl:choose[.//*[namespace-uri()!="http://www.w3.org/1999/XSL/Transform"]]', $manager))) {
        $parent = $choose.ParentNode
        foreach ($branch in @($choose.ChildNodes)) {
            while ($branch.HasChildNodes) { $null = $parent.InsertBefore($branch.FirstChild, $choose) }
        }
        $null = $parent.RemoveChild($choose)
    }
    foreach ($gate in @($View.SelectNodes('//xsl:if[.//*[namespace-uri()!="http://www.w3.org/1999/XSL/Transform"]]', $manager))) {
        $parent = $gate.ParentNode
        if ($null -eq $parent) { continue }
        while ($gate.HasChildNodes) { $null = $parent.InsertBefore($gate.FirstChild, $gate) }
        $null = $parent.RemoveChild($gate)
    }
    foreach ($choose in @($View.SelectNodes('//xsl:choose[not(xsl:when)]', $manager))) {
        $parent = $choose.ParentNode
        while ($choose.HasChildNodes) { $null = $parent.InsertBefore($choose.FirstChild, $choose) }
        $null = $parent.RemoveChild($choose)
    }
    foreach ($value in $View.SelectNodes('//xsl:value-of/@select', $manager)) {
        if ($value.Value -match '(?<![\w.-])(?!(?:text|node)\s*\()[\w.-]+(?::[\w.-]+)?\s*\(|\+|(?<=\d)-(?=\d)|\s(?:-|div|mod|\*)\s|\$') {
            throw "Unbound XSN display expression requires runtime computation: '$($value.Value)'."
        }
    }

    $templates = [System.Xml.XmlDocument]::new()
    $templates.LoadXml(@'
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:archive="urn:infopath-archive:display" version="1.0">
  <xsl:template match="*" mode="archive-field">
    <xsl:choose>
      <xsl:when test="@archive:kind='link'"><a href="{@archive:href}"><xsl:value-of select="@archive:value"/></a></xsl:when>
      <xsl:when test="@archive:kind='selections'"><xsl:for-each select="archive:option"><xsl:choose><xsl:when test="@selected='true'">[X] </xsl:when><xsl:otherwise>[ ] </xsl:otherwise></xsl:choose><xsl:value-of select="."/><xsl:text> </xsl:text></xsl:for-each></xsl:when>
      <xsl:when test="@archive:kind='rich-text' or *[namespace-uri()='http://www.w3.org/1999/xhtml']"><xsl:apply-templates select="node()" mode="archive-rich"/></xsl:when>
      <xsl:otherwise><xsl:value-of select="@archive:value"/></xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="*" mode="archive-rich">
    <xsl:choose>
      <xsl:when test="contains('|p|div|span|br|b|strong|i|em|u|s|ul|ol|li|table|thead|tbody|tr|th|td|blockquote|pre|h1|h2|h3|h4|',concat('|',local-name(),'|'))">
        <xsl:element name="{local-name()}"><xsl:copy-of select="@colspan|@rowspan"/><xsl:apply-templates select="node()" mode="archive-rich"/></xsl:element>
      </xsl:when>
      <xsl:when test="local-name()='a' and (starts-with(@href,'https://') or starts-with(@href,'http://') or starts-with(@href,'mailto:'))"><a href="{@href}"><xsl:apply-templates select="node()" mode="archive-rich"/></a></xsl:when>
      <xsl:when test="not(contains('|script|style|iframe|object|embed|',concat('|',local-name(),'|')))"><xsl:apply-templates select="node()" mode="archive-rich"/></xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="text()" mode="archive-rich"><xsl:value-of select="."/></xsl:template>
</xsl:stylesheet>
'@)
    foreach ($template in @($templates.DocumentElement.ChildNodes)) {
        $null = $root.AppendChild($View.ImportNode($template, $true))
    }
    $body = $View.SelectSingleNode("//*[local-name()='body' and namespace-uri()!='$xsl']")
    if ($null -ne $body) {
        $style = $View.CreateElement('style', $body.NamespaceURI)
        $style.InnerText = @'
.archive-value {
    display: inline-block;
    box-sizing: border-box;
    border: 0 !important;
    background: transparent !important;
    box-shadow: none !important;
    outline: 0 !important;
    appearance: none !important;
    opacity: 1 !important;
    padding: 2px 4px !important;
    margin: 1px 0 !important;
    height: auto !important;
    min-height: 1em;
    max-height: none !important;
    max-width: 100%;
    overflow: visible !important;
    white-space: pre-wrap !important;
    overflow-wrap: anywhere;
    text-overflow: clip !important;
    line-height: 1.35 !important;
    vertical-align: top;
}
.archive-selection {
    width: auto !important;
    white-space: nowrap !important;
    padding: 2px 3px !important;
}
.archive-section {
    height: auto !important;
    max-height: none !important;
    overflow: visible !important;
}
@media print {
    .archive-value { color: black !important; }
}
'@
        $null = $body.PrependChild($style)
    }
    return ,$View
}

function Write-XsnArchiveHtml {
    param(
        [System.Xml.XmlDocument]$Document,
        [psobject]$Configuration,
        [string]$ResourceDirectory,
        [string]$TemplateHash,
        [string]$HtmlPath,
        [string]$OutputRoot,
        [hashtable]$Cache
    )

    if ($TemplateHash -notmatch '^[A-Za-z0-9_-]+$') { throw 'XSN template asset key must be a nonempty safe directory name.' }
    $manifestPath = Resolve-XsnLocalResource $ResourceDirectory 'manifest.xsf'
    $manifest = Read-XsnXml $manifestPath
    $views = $manifest.DocumentElement.SelectSingleNode("*[local-name()='views']")
    if ($null -eq $views) { throw 'XSN manifest has no views.' }
    $defaultName = $views.GetAttribute('default')
    $view = @($views.SelectNodes("*[local-name()='view']") | Where-Object { $_.GetAttribute('name') -ceq $defaultName })
    if ($view.Count -ne 1) { throw "XSN default view '$defaultName' is missing or ambiguous." }
    $pane = $view[0].SelectSingleNode("*[local-name()='mainpane']")
    if ($null -eq $pane) { throw 'XSN view has no mainpane transform.' }
    $viewPath = Resolve-XsnLocalResource $ResourceDirectory $pane.GetAttribute('transform')
    $cacheKey = $viewPath
    if (-not $Cache.ContainsKey($cacheKey)) {
        $stylesheet = New-XsnArchiveStylesheet (Read-XsnXml $viewPath)
        $transform = [System.Xml.Xsl.XslCompiledTransform]::new()
        $reader = [System.Xml.XmlNodeReader]::new($stylesheet)
        try { $transform.Load($reader, [System.Xml.Xsl.XsltSettings]::Default, $null) }
        finally { $reader.Dispose() }
        $Cache[$cacheKey] = $transform
    }

    $raw = $Document.CloneNode($true)
    $formatted = $Document.CloneNode($true)
    $rawElements = @($raw.SelectNodes('//*'))
    $formattedElements = @($formatted.SelectNodes('//*'))
    Convert-HtmlFieldValues -Document $formatted -Configuration $Configuration
    for ($index = 0; $index -lt $rawElements.Count; $index++) {
        $node = $rawElements[$index]
        $formattedNode = $formattedElements[$index]
        $null = $node.SetAttribute('value', 'urn:infopath-archive:display', $formattedNode.InnerText)
        foreach ($name in @('kind', 'href')) {
            if ($formattedNode.HasAttribute($name, 'urn:infopath-archive:display')) {
                $null = $node.SetAttribute($name, 'urn:infopath-archive:display', $formattedNode.GetAttribute($name, 'urn:infopath-archive:display'))
            }
        }
        if ($formattedNode.GetAttribute('kind', 'urn:infopath-archive:display') -eq 'selections') {
            foreach ($option in $formattedNode.SelectNodes("*[namespace-uri()='urn:infopath-archive:display']")) {
                $null = $node.AppendChild($raw.ImportNode($option, $true))
            }
        }
    }
    $arguments = [System.Xml.Xsl.XsltArgumentList]::new()
    $buffer = [System.Text.StringBuilder]::new()
    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.OmitXmlDeclaration = $true
    $settings.ConformanceLevel = [System.Xml.ConformanceLevel]::Auto
    $settings.CheckCharacters = $true
    # XML serialization allows safe DOM postprocessing before final HTML serialization.
    $writer = [System.Xml.XmlWriter]::Create($buffer, $settings)
    try { $Cache[$cacheKey].Transform($raw, $arguments, $writer) }
    finally { $writer.Dispose() }
    $htmlDocument = [System.Xml.XmlDocument]::new()
    $htmlDocument.XmlResolver = $null
    $htmlDocument.LoadXml($buffer.ToString())
    $assetRoot = Join-Path (Join-Path $OutputRoot 'Assets') $TemplateHash
    $htmlDirectory = Split-Path ([System.IO.Path]::GetFullPath($HtmlPath)) -Parent
    $baseUri = [uri]::new($htmlDirectory.TrimEnd('\') + '\')
    $assetUri = [uri]::new([System.IO.Path]::GetFullPath($assetRoot).TrimEnd('\') + '\')
    $relativeAssets = $baseUri.MakeRelativeUri($assetUri).ToString()
    $resourceBase = Split-Path $viewPath -Parent
    $copyResource = {
        param([string]$Reference, [string]$SourceBase = $resourceBase)
        $pathPart = ($Reference -split '[?#]', 2)[0]
        $source = Resolve-XsnLocalResource $SourceBase $pathPart
        $extension = [System.IO.Path]::GetExtension($source).ToLowerInvariant()
        if ($extension -notin @('.gif', '.png', '.jpg', '.jpeg', '.bmp', '.ico', '.css', '.woff', '.woff2', '.ttf')) {
            throw "Unsupported XSN output asset: '$Reference'."
        }
        $relative = [System.IO.Path]::GetFullPath($ResourceDirectory).TrimEnd('\') + '\'
        $resourceName = $source.Substring($relative.Length)
        $destination = Join-Path $assetRoot $resourceName
        New-Item -Path (Split-Path $destination -Parent) -ItemType Directory -Force | Out-Null
        if ($extension -eq '.css' -and -not (Test-Path -LiteralPath $destination)) {
            $css = Get-Content -LiteralPath $source -Raw
            if ($css -match '(?i)@import|expression\s*\(|behavior\s*:') { throw "Unsupported active CSS in '$Reference'." }
            foreach ($match in [regex]::Matches($css, '(?i)url\(\s*[''"]?([^''")]+)[''"]?\s*\)')) {
                $null = & $copyResource $match.Groups[1].Value.Trim() (Split-Path $source -Parent)
            }
        }
        Copy-Item -LiteralPath $source -Destination $destination -Force
        return $relativeAssets + (($resourceName -split '\\' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
    }
    foreach ($element in @($htmlDocument.SelectNodes('//*'))) {
        foreach ($attribute in @($element.Attributes)) {
            if ($attribute.Name -match '^(?i:on)' -or $attribute.NamespaceURI -eq 'http://schemas.microsoft.com/office/infopath/2003') {
                $null = $element.Attributes.Remove($attribute)
            }
        }
        if ($element.LocalName -in @('script', 'object', 'embed', 'iframe', 'meta', 'base')) {
            $null = $element.ParentNode.RemoveChild($element)
            continue
        }
        foreach ($name in @('src', 'background')) {
            if ($element.HasAttribute($name)) { $element.SetAttribute($name, (& $copyResource $element.GetAttribute($name))) }
        }
        if ($element.LocalName -eq 'link' -and $element.HasAttribute('href')) {
            $element.SetAttribute('href', (& $copyResource $element.GetAttribute('href')))
        }
        if ($element.HasAttribute('href') -and $element.GetAttribute('href') -match '^(?i:javascript|vbscript|data):') {
            $element.RemoveAttribute('href')
        }
        if ($element.HasAttribute('style')) {
            $element.SetAttribute('style', (Remove-XsnHiddenStyle $element.GetAttribute('style')))
        }
    }
    foreach ($node in @($htmlDocument.SelectNodes('//@style | //*[local-name()="style"]/text()'))) {
        if ($node.Value -match '(?i)expression\s*\(|@import') {
            throw 'XSN view contains unsupported active CSS or CSS imports.'
        }
        $node.Value = [regex]::Replace($node.Value, '(?i)url\(\s*[''"]?([^''")]+)[''"]?\s*\)', {
            param($match)
            return 'url("' + (& $copyResource $match.Groups[1].Value.Trim()) + '")'
        })
    }
    $serializer = [System.Xml.Xsl.XslCompiledTransform]::new()
    $identity = [System.Xml.XmlDocument]::new()
    $identity.LoadXml('<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"><xsl:output method="html" encoding="utf-8"/><xsl:template match="@*|node()"><xsl:copy><xsl:apply-templates select="@*|node()"/></xsl:copy></xsl:template></xsl:stylesheet>')
    $serializer.Load($identity)
    $output = [System.Xml.XmlWriter]::Create($HtmlPath, $serializer.OutputSettings)
    try { $serializer.Transform($htmlDocument, $null, $output) }
    finally { $output.Dispose() }
    return $defaultName
}
