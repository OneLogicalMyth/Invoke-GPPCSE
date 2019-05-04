$global:BuildReviewRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
"$(Split-Path -Path $MyInvocation.MyCommand.Path)\Functions\*.ps1" | Resolve-Path | % { . $_.ProviderPath }
