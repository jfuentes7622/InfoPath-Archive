function Get-HtmlFormattingXPath {
	$names = '|a|abbr|acronym|address|article|aside|b|bdi|bdo|big|blockquote|body|br|caption|center|cite|code|col|colgroup|dd|del|details|dfn|div|dl|dt|em|figcaption|figure|font|footer|h1|h2|h3|h4|h5|h6|header|hr|html|i|img|ins|kbd|li|main|mark|nav|ol|p|pre|q|s|samp|section|small|span|strike|strong|sub|summary|sup|table|tbody|td|tfoot|th|thead|time|tr|tt|u|ul|var|wbr|script|style|iframe|object|embed|'
	return "namespace-uri() = 'http://www.w3.org/1999/xhtml' or (namespace-uri() = '' and contains('$names', concat('|', translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '|')))"
}

function Test-HtmlFormattingNode {
	param(
		[Parameter(Mandatory = $true)]
		[System.Xml.XmlNode]$Node
	)

	return ($null -ne $Node.SelectSingleNode("self::*[$(Get-HtmlFormattingXPath)]"))
}

function Get-HtmlRuleProperty {
	param([psobject]$Rule, [string]$Name, $Default = $null)

	$property = $Rule.PSObject.Properties[$Name]
	if ($null -ne $property) {
		return $property.Value
	}
	return $Default
}

function Import-HtmlFieldRules {
	param([string]$Path)

	$configuration = [pscustomobject]@{
		Culture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
		Namespaces = @{}
		Fields = @()
	}
	if ([string]::IsNullOrWhiteSpace($Path)) {
		return $configuration
	}
	if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
		throw "HTML field-rules file does not exist: $Path"
	}
	$inputRules = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -ErrorAction Stop
	if ($null -eq $inputRules -or $inputRules -is [array] -or $inputRules -isnot [pscustomobject]) {
		throw 'HTML field rules must be a JSON object.'
	}
	$cultureName = [string](Get-HtmlRuleProperty -Rule $inputRules -Name 'Culture' -Default 'en-US')
	if ([string]::IsNullOrWhiteSpace($cultureName)) {
		throw 'HTML field-rules Culture must not be blank.'
	}
	$configuration.Culture = [System.Globalization.CultureInfo]::GetCultureInfo($cultureName)
	$namespaces = Get-HtmlRuleProperty -Rule $inputRules -Name 'Namespaces'
	if ($null -ne $namespaces) {
		if ($namespaces -isnot [pscustomobject]) {
			throw 'HTML field-rules Namespaces must be an object of prefix/URI pairs.'
		}
		foreach ($property in $namespaces.PSObject.Properties) {
			if ([string]::IsNullOrWhiteSpace([string]$property.Value)) {
				throw "HTML field-rules namespace '$($property.Name)' must not be blank."
			}
			$configuration.Namespaces[$property.Name] = [string]$property.Value
		}
	}
	$fields = Get-HtmlRuleProperty -Rule $inputRules -Name 'Fields'
	if ($fields -isnot [array]) {
		throw 'HTML field-rules Fields must be an array.'
	}
	$testDocument = [System.Xml.XmlDocument]::new()
	$manager = [System.Xml.XmlNamespaceManager]::new($testDocument.NameTable)
	foreach ($prefix in $configuration.Namespaces.Keys) {
		$manager.AddNamespace($prefix, $configuration.Namespaces[$prefix])
	}
	$paths = @{}
	$types = @('Text', 'Calculated', 'Boolean', 'Currency', 'Number', 'Percentage', 'Date', 'DateTime', 'Choice', 'MultiChoice', 'Dropdown', 'Radio', 'Lookup', 'Person', 'Signature', 'Hyperlink', 'Attachment', 'MultipleCheckboxes', 'RichText')
	foreach ($rule in $fields) {
		if ($rule -isnot [pscustomobject]) {
			throw 'Each HTML field rule must be an object.'
		}
		foreach ($property in $rule.PSObject.Properties) {
			if ($property.Name -notin @('XPath', 'Type', 'Format', 'Values', 'UrlXPath', 'DescriptionXPath', 'DisplayNameXPath', 'SignerXPath', 'DateXPath', 'Encoding', 'FileName')) {
				throw "Unknown HTML field-rule property '$($property.Name)'."
			}
		}
		$pathValue = [string](Get-HtmlRuleProperty -Rule $rule -Name 'XPath')
		$type = [string](Get-HtmlRuleProperty -Rule $rule -Name 'Type')
		if ([string]::IsNullOrWhiteSpace($pathValue) -or $types -notcontains $type) {
			throw "Each HTML field rule needs an XPath and a supported Type. Invalid rule: '$pathValue' / '$type'."
		}
		if ($paths.ContainsKey($pathValue)) {
			throw "Duplicate HTML field-rule XPath: $pathValue"
		}
		$paths[$pathValue] = $true
		$null = $testDocument.SelectNodes($pathValue, $manager)
		foreach ($name in @('UrlXPath', 'DescriptionXPath', 'DisplayNameXPath', 'SignerXPath', 'DateXPath')) {
			$expression = Get-HtmlRuleProperty -Rule $rule -Name $name
			if ($null -ne $expression) {
				if ([string]::IsNullOrWhiteSpace([string]$expression)) {
					throw "HTML field-rule '$pathValue' has a blank $name."
				}
				$null = $testDocument.SelectNodes([string]$expression, $manager)
			}
		}
		if ($type -eq 'Signature' -and
			($null -eq (Get-HtmlRuleProperty -Rule $rule -Name 'SignerXPath') -or
			$null -eq (Get-HtmlRuleProperty -Rule $rule -Name 'DateXPath'))) {
			throw "Signature rule '$pathValue' requires SignerXPath and DateXPath."
		}
		$values = Get-HtmlRuleProperty -Rule $rule -Name 'Values'
		if ($null -ne $values -and $values -isnot [pscustomobject]) {
			throw "HTML field-rule '$pathValue' Values must be an object of stored/display value pairs."
		}
		$encoding = Get-HtmlRuleProperty -Rule $rule -Name 'Encoding'
		if ($null -ne $encoding -and ($type -ne 'Attachment' -or $encoding -notin @('base64', 'FileReference'))) {
			throw "Encoding is supported only for Attachment fields and must be base64 or FileReference."
		}
		if ($encoding -eq 'base64') {
			$fileName = [string](Get-HtmlRuleProperty -Rule $rule -Name 'FileName')
			if ([string]::IsNullOrWhiteSpace($fileName) -or $fileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $fileName -in @('.', '..')) {
				throw "Base64 Attachment rule '$pathValue' needs a valid FileName without directory components."
			}
		}
		$format = Get-HtmlRuleProperty -Rule $rule -Name 'Format'
		if ($null -ne $format) {
			if ([string]::IsNullOrWhiteSpace([string]$format)) {
				throw "HTML field-rule '$pathValue' Format must not be blank."
			}
			if ($type -in @('Currency', 'Number', 'Percentage')) {
				$null = ([decimal]1).ToString([string]$format, $configuration.Culture)
			}
			elseif ($type -in @('Date', 'DateTime')) {
				$null = [datetimeoffset]::Now.ToString([string]$format, $configuration.Culture)
			}
			else {
				throw "HTML field-rule '$pathValue' does not support Format for Type '$type'."
			}
		}
	}
	$configuration.Fields = $fields
	return $configuration
}

function Set-HtmlDisplayValue {
	param([System.Xml.XmlElement]$Node, [AllowEmptyString()][string]$Value, [string]$Kind = 'text')

	$Node.InnerText = $Value
	$Node.SetAttribute('kind', 'urn:infopath-archive:display', $Kind) | Out-Null
}

function Convert-HtmlDateValue {
	param([string]$Value, [string]$DateFormat, [string]$DateTimeFormat, [string]$Type = '')

	if ($Value.Trim() -notmatch '^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:\d{2})?$') {
		return $null
	}
	$date = [datetimeoffset]::MinValue
	if (-not [datetimeoffset]::TryParse($Value.Trim(), [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$date)) {
		return $null
	}
	$format = if ($Type -eq 'DateTime' -or ($Type -ne 'Date' -and $Value.Contains('T'))) { $DateTimeFormat } else { $DateFormat }
	return $date.ToString($format, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-HtmlMappedValue {
	param([string]$Value, [psobject]$Rule)

	$values = Get-HtmlRuleProperty -Rule $Rule -Name 'Values'
	if ($null -ne $values) {
		$property = $values.PSObject.Properties[$Value]
		if ($null -ne $property) {
			return [string]$property.Value
		}
	}
	return $Value
}

function Get-HtmlChoiceValues {
	param([string]$Value, [switch]$StripLookupIds)

	$parts = @($Value -split ';#' | Where-Object { $_ -ne '' })
	$lookup = $StripLookupIds -and $parts.Count -ge 2 -and $parts.Count % 2 -eq 0
	for ($index = 0; $lookup -and $index -lt $parts.Count; $index += 2) {
		$lookup = $parts[$index] -match '^\d+$'
	}
	if ($lookup) {
		$parts = @(for ($index = 1; $index -lt $parts.Count; $index += 2) { $parts[$index] })
	}
	return ,$parts
}

function Set-HtmlHyperlink {
	param([System.Xml.XmlElement]$Node, [string]$Url, [string]$Label, [switch]$Attachment)

	$Url = $Url.Trim()
	if ([string]::IsNullOrWhiteSpace($Url) -or $Url -match '[\x00-\x1F\x7F]') {
		throw "HTML hyperlink must not be empty or contain control characters."
	}
	$uri = $null
	if (-not [uri]::TryCreate($Url, [System.UriKind]::RelativeOrAbsolute, [ref]$uri) -or
		($uri.IsAbsoluteUri -and $uri.Scheme -notin @('http', 'https', 'mailto', 'tel') -and -not ($Attachment -and $uri.IsFile))) {
		throw "Unsupported or invalid HTML hyperlink: '$Url'."
	}
	if ($Attachment) {
		if ($uri.IsAbsoluteUri) {
			$Url = $uri.AbsoluteUri
		}
		else {
			$Url = (($Url -replace '\\', '/') -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
		}
	}
	Set-HtmlDisplayValue -Node $Node -Value $Label -Kind 'link'
	$Node.SetAttribute('href', 'urn:infopath-archive:display', $Url) | Out-Null
}

function Export-HtmlFieldAttachments {
	param(
		[System.Xml.XmlDocument]$Document,
		[psobject]$Configuration,
		[string]$SourceFileDirectory,
		[string]$SourceRoot,
		[string]$AttachmentOutputDirectory,
		[string]$RelativePrefix
	)

	$manager = [System.Xml.XmlNamespaceManager]::new($Document.NameTable)
	foreach ($prefix in $Configuration.Namespaces.Keys) {
		$manager.AddNamespace($prefix, $Configuration.Namespaces[$prefix])
	}
	$allowedRoot = [System.IO.Path]::GetFullPath($SourceRoot).TrimEnd('\') + '\'
	foreach ($rule in $Configuration.Fields) {
		if ($rule.Type -ne 'Attachment') { continue }
		foreach ($node in $Document.SelectNodes([string]$rule.XPath, $manager)) {
			if ($node -isnot [System.Xml.XmlElement]) { throw "Attachment rule '$($rule.XPath)' must select elements." }
			$value = $node.InnerText.Trim()
			if (-not $value -or $value.StartsWith('ATTACHMENT::')) { continue }
			$encoding = [string](Get-HtmlRuleProperty -Rule $rule -Name 'Encoding' -Default 'FileReference')
			if ($encoding -eq 'base64') {
				$bytes = [Convert]::FromBase64String($value)
				$desiredName = [string]$rule.FileName
			}
			else {
				$uri = $null
				if ([uri]::TryCreate($value, [System.UriKind]::Absolute, [ref]$uri)) {
					if ($uri.Scheme -in @('http', 'https')) { continue }
					if (-not $uri.IsFile) { throw "Unsupported attachment reference: '$value'." }
					$sourcePath = $uri.LocalPath
				}
				else {
					$sourcePath = Join-Path -Path $SourceFileDirectory -ChildPath $value
				}
				$sourcePath = [System.IO.Path]::GetFullPath($sourcePath)
				if (-not $sourcePath.StartsWith($allowedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
					throw "Attachment reference must remain within the InfoPath source directory: '$value'."
				}
				if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
					throw "Referenced attachment does not exist: '$sourcePath'."
				}
				$inspectionPath = $sourcePath
				while ($inspectionPath.TrimEnd('\') -ne $allowedRoot.TrimEnd('\')) {
					$item = Get-Item -LiteralPath $inspectionPath -Force
					if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
						throw "Attachment references cannot traverse symbolic links or junctions: '$value'."
					}
					$inspectionPath = Split-Path -Path $inspectionPath -Parent
				}
				$desiredName = [System.IO.Path]::GetFileName($sourcePath)
			}
			$destination = New-UniqueAttachmentPath -DirectoryPath $AttachmentOutputDirectory -DesiredName $desiredName
			if ($encoding -eq 'base64') {
				[System.IO.File]::WriteAllBytes($destination, $bytes)
			}
			else {
				Copy-Item -LiteralPath $sourcePath -Destination $destination
			}
			$node.InnerText = 'ATTACHMENT::' + $RelativePrefix.TrimEnd('/') + '/' + [System.IO.Path]::GetFileName($destination)
		}
	}
}

function Set-HtmlCheckboxSelections {
	param([System.Xml.XmlElement]$Node, [string]$Value, [psobject]$Rule)

	$items = @($Value -split ';\s*' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
	if ($items.Count -eq 0 -or @($items | Where-Object { $_ -notmatch '^.+?=\s*(?i:true|false)\s*$' }).Count -gt 0) {
		throw "Invalid multiple-checkbox value in '$($Node.get_Name())': '$Value'. Expected Option=true; Option=false."
	}
	Set-HtmlDisplayValue -Node $Node -Value '' -Kind 'selections'
	foreach ($item in $items) {
		$null = $item -match '^(?<label>.+?)=\s*(?<selected>(?i:true|false))\s*$'
		$label = $Matches.label.Trim()
		if ($null -ne $Rule) {
			$label = Get-HtmlMappedValue -Value $label -Rule $Rule
		}
		$option = $Node.OwnerDocument.CreateElement('option', 'urn:infopath-archive:display')
		$option.SetAttribute('selected', $Matches.selected.ToLowerInvariant())
		$option.InnerText = $label
		$null = $Node.AppendChild($option)
	}
}

function Convert-HtmlFieldValues {
	param(
		[System.Xml.XmlDocument]$Document,
		[psobject]$Configuration,
		[string]$DateFormat = 'MM/dd/yyyy',
		[string]$DateTimeFormat = 'MM/dd/yyyy HH:mm'
	)

	$displayNamespace = 'urn:infopath-archive:display'
	$formattingXPath = Get-HtmlFormattingXPath
	$manager = [System.Xml.XmlNamespaceManager]::new($Document.NameTable)
	foreach ($prefix in $Configuration.Namespaces.Keys) {
		$manager.AddNamespace($prefix, $Configuration.Namespaces[$prefix])
	}
	$selectedRules = [System.Collections.Generic.Dictionary[System.Xml.XmlElement,object]]::new()
	foreach ($rule in $Configuration.Fields) {
		foreach ($node in $Document.SelectNodes([string]$rule.XPath, $manager)) {
			if ($node -isnot [System.Xml.XmlElement] -or
				$null -ne $node.SelectSingleNode("ancestor-or-self::*[$formattingXPath]")) {
				throw "HTML field-rule '$($rule.XPath)' must select form elements, not attributes or rich-text formatting."
			}
			if ($selectedRules.ContainsKey($node)) {
				throw "Overlapping HTML field rules select '$($node.get_Name())' more than once."
			}
			$selectedRules.Add($node, $rule)
		}
	}
	foreach ($node in $selectedRules.Keys) {
		$parent = $node.ParentNode
		while ($parent -is [System.Xml.XmlElement]) {
			if ($selectedRules.ContainsKey($parent)) {
				throw "HTML field rules cannot select both a container and its descendant: '$($node.get_Name())'."
			}
			$parent = $parent.ParentNode
		}
	}
	foreach ($node in $selectedRules.Keys) {
		$rule = $selectedRules[$node]
		$type = [string]$rule.Type
		$value = $node.InnerText
		if ($type -eq 'RichText') {
			$node.SetAttribute('kind', $displayNamespace, 'rich-text') | Out-Null
			continue
		}
		if ([string]::IsNullOrWhiteSpace($value) -or $node.GetAttribute('nil', 'http://www.w3.org/2001/XMLSchema-instance') -in @('true', '1')) {
			Set-HtmlDisplayValue -Node $node -Value ''
			continue
		}
		$format = [string](Get-HtmlRuleProperty -Rule $rule -Name 'Format' -Default '')
		switch ($type) {
			{ $_ -in @('Currency', 'Number', 'Percentage') } {
				$number = [decimal]::Zero
				if (-not [decimal]::TryParse($value.Trim(), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
					throw "Invalid $type value in '$($node.get_Name())': '$value'."
				}
				if (-not $format) {
					$format = switch ($type) { 'Currency' { 'C2' } 'Percentage' { '0.##%' } 'Number' { '#,##0.############################' } }
				}
				Set-HtmlDisplayValue -Node $node -Value $number.ToString($format, $Configuration.Culture)
			}
			{ $_ -in @('Date', 'DateTime') } {
				$dateOnly = if ($format -and $type -eq 'Date') { $format } else { $DateFormat }
				$dateTime = if ($format -and $type -eq 'DateTime') { $format } else { $DateTimeFormat }
				$result = Convert-HtmlDateValue -Value $value -DateFormat $dateOnly -DateTimeFormat $dateTime -Type $type
				if ($null -eq $result) { throw "Invalid $type value in '$($node.get_Name())': '$value'." }
				Set-HtmlDisplayValue -Node $node -Value $result
			}
			'Boolean' {
				if ($value.Trim() -notmatch '^(?i:true|false|1|0)$') { throw "Invalid Boolean value in '$($node.get_Name())': '$value'." }
				$boolean = if ($value.Trim() -in @('true', '1')) { 'true' } else { 'false' }
				Set-HtmlDisplayValue -Node $node -Value $boolean -Kind 'boolean'
			}
			{ $_ -in @('Choice', 'MultiChoice', 'Dropdown', 'Radio', 'Lookup') } {
				$parts = Get-HtmlChoiceValues -Value $value -StripLookupIds:($type -eq 'Lookup')
				$labels = @($parts | ForEach-Object { Get-HtmlMappedValue -Value $_ -Rule $rule })
				Set-HtmlDisplayValue -Node $node -Value ($labels -join ', ')
			}
			'Person' {
				$names = [System.Collections.Generic.List[string]]::new()
				$expression = Get-HtmlRuleProperty -Rule $rule -Name 'DisplayNameXPath'
				if ($null -ne $expression) {
					foreach ($display in $node.SelectNodes([string]$expression, $manager)) {
						if (-not [string]::IsNullOrWhiteSpace($display.InnerText)) { $names.Add($display.InnerText) }
					}
				}
				else {
					$people = @($node.SelectNodes(".//*[local-name()='Person']"))
					if ($people.Count -eq 0) { $people = @($node) }
					foreach ($person in $people) {
						$display = $person.SelectSingleNode(".//*[local-name()='DisplayName']")
						$account = $person.SelectSingleNode(".//*[local-name()='AccountId']")
						if ($null -ne $display -and -not [string]::IsNullOrWhiteSpace($display.InnerText)) {
							$names.Add($display.InnerText)
						}
						elseif ($null -ne $account -and -not [string]::IsNullOrWhiteSpace($account.InnerText)) {
							$names.Add((Get-HtmlMappedValue -Value $account.InnerText -Rule $rule))
						}
					}
				}
				if ($names.Count -eq 0) { $names.Add((Get-HtmlMappedValue -Value $value -Rule $rule)) }
				Set-HtmlDisplayValue -Node $node -Value ($names -join '; ')
			}
			'Signature' {
				$signer = $node.SelectSingleNode([string]$rule.SignerXPath, $manager)
				$signedDate = $node.SelectSingleNode([string]$rule.DateXPath, $manager)
				$pieces = [System.Collections.Generic.List[string]]::new()
				if ($null -ne $signer -and -not [string]::IsNullOrWhiteSpace($signer.InnerText)) {
					$pieces.Add((Get-HtmlMappedValue -Value $signer.InnerText -Rule $rule))
				}
				if ($null -ne $signedDate -and -not [string]::IsNullOrWhiteSpace($signedDate.InnerText)) {
					$result = Convert-HtmlDateValue -Value $signedDate.InnerText -DateFormat $DateFormat -DateTimeFormat $DateTimeFormat
					if ($null -eq $result) { throw "Invalid signature date in '$($node.get_Name())'." }
					$pieces.Add($result)
				}
				if ($pieces.Count -eq 0) { throw "Signature rule '$($rule.XPath)' did not resolve a signer or signed date." }
				Set-HtmlDisplayValue -Node $node -Value ($pieces -join ' - ')
			}
			'Hyperlink' {
				$urlExpression = Get-HtmlRuleProperty -Rule $rule -Name 'UrlXPath'
				$descriptionExpression = Get-HtmlRuleProperty -Rule $rule -Name 'DescriptionXPath'
				$url = $value
				$label = $value
				if ($null -ne $urlExpression) {
					$urlNode = $node.SelectSingleNode([string]$urlExpression, $manager)
					if ($null -eq $urlNode) { throw "Hyperlink rule '$($rule.XPath)' did not resolve UrlXPath." }
					$url = $urlNode.InnerText
					$label = $url
				}
				elseif ($value -match '^(?<url>\S+),\s+(?<label>.+)$') {
					$url = $Matches.url
					$label = $Matches.label
				}
				if ($null -ne $descriptionExpression) {
					$description = $node.SelectSingleNode([string]$descriptionExpression, $manager)
					if ($null -ne $description -and -not [string]::IsNullOrWhiteSpace($description.InnerText)) { $label = $description.InnerText }
				}
				Set-HtmlHyperlink -Node $node -Url $url -Label $label
			}
			'Attachment' {
				if ((Get-HtmlRuleProperty -Rule $rule -Name 'Encoding') -eq 'base64' -and -not $value.StartsWith('ATTACHMENT::')) {
					Set-HtmlDisplayValue -Node $node -Value ([string]$rule.FileName)
					break
				}
				$url = $value -replace '^ATTACHMENT::', ''
				$label = ($url -replace '\\', '/').Split('/')[-1]
				Set-HtmlHyperlink -Node $node -Url $url -Label $label -Attachment
			}
			'MultipleCheckboxes' { Set-HtmlCheckboxSelections -Node $node -Value $value -Rule $rule }
			default { Set-HtmlDisplayValue -Node $node -Value $value }
		}
	}
	foreach ($node in $Document.SelectNodes("//*[not(*) and not(ancestor-or-self::*[$formattingXPath])]")) {
		if ($node.HasAttribute('kind', $displayNamespace) -or $node.NamespaceURI -eq $displayNamespace) { continue }
		$value = $node.InnerText
		if ([string]::IsNullOrWhiteSpace($value)) { continue }
		if ($value.StartsWith('ATTACHMENT::')) {
			$url = $value.Substring(12)
			Set-HtmlHyperlink -Node $node -Url $url -Label (($url -replace '\\', '/').Split('/')[-1]) -Attachment
		}
		elseif ($value -match '^(?<url>(?i:https?://|mailto:|tel:)\S+?),\s+(?<label>.+)$') {
			Set-HtmlHyperlink -Node $node -Url $Matches.url -Label $Matches.label
		}
		elseif ($value -match '^(?i:https?://|mailto:|tel:)\S+$') {
			Set-HtmlHyperlink -Node $node -Url $value -Label $value
		}
		elseif ($value.Contains(';#')) {
			$parts = Get-HtmlChoiceValues -Value $value -StripLookupIds
			Set-HtmlDisplayValue -Node $node -Value ($parts -join ', ')
		}
		elseif ($value -match '^.+?=\s*(?i:true|false)(\s*;|$)' -and @($value -split ';\s*' | Where-Object { $_ -and $_ -notmatch '^.+?=\s*(?i:true|false)\s*$' }).Count -eq 0) {
			Set-HtmlCheckboxSelections -Node $node -Value $value
		}
		else {
			$result = Convert-HtmlDateValue -Value $value -DateFormat $DateFormat -DateTimeFormat $DateTimeFormat
			if ($null -ne $result) { Set-HtmlDisplayValue -Node $node -Value $result }
		}
	}
}
