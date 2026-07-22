function Convert-ToHumanReadableLabel {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ElementName
	)

	# Remove namespace prefixes (my:, d:, etc.)
	$label = $ElementName -replace '^[a-z]+:', ''
	
	# Insert space before uppercase that follows lowercase: "RequestorType" → "Requestor Type"
	$label = $label -replace '(?<=[a-z])(?=[A-Z])', ' '
	
	# Insert space before the last uppercase in a sequence before lowercase: "HTTPResponse" → "HTTP Response"
	$label = $label -replace '(?<=[A-Z])(?=[A-Z][a-z])', ' '
	
	# Trim any excess whitespace
	$label = $label.Trim()
	
	# Return the result (already properly capitalized by camelCase)
	return $label
}

Write-Host "Testing label conversion:"
@(
	"RequestorType",
	"my:RequestorType",
	"HTTPResponse",
	"my:FugitivePersonalInformation",
	"RedNoticeApplication",
	"my:RedNoticeApplication"
) | ForEach-Object {
	$result = Convert-ToHumanReadableLabel -ElementName $_
	Write-Host "  '$_' => '$result'"
}
