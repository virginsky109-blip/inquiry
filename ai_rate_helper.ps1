param(
    [switch]$Configure,
    [switch]$ClearKey,
    [switch]$Status
)

$ErrorActionPreference = 'Stop'
$AppDir = Join-Path $env:LOCALAPPDATA 'ForwardingQuote'
$KeyFile = Join-Path $AppDir 'openai_api_key.dat'
$ConfigFile = Join-Path $AppDir 'ai_config.json'
$Port = 8751

function Ensure-AppDir {
    if (-not (Test-Path $AppDir)) { New-Item -ItemType Directory -Path $AppDir -Force | Out-Null }
}

$DefaultModel = 'gpt-4o-mini'

function Get-Config {
    Ensure-AppDir
    if (Test-Path $ConfigFile) {
        $cfg = Get-Content -Raw $ConfigFile | ConvertFrom-Json
        # Repair configs saved with a non-existent model name.
        if ([string]::IsNullOrWhiteSpace($cfg.model) -or $cfg.model -eq 'gpt-5.6-luna') { $cfg.model = $DefaultModel }
        return $cfg
    }
    return [pscustomobject]@{ model = $DefaultModel; accessToken = '' }
}

function Get-ApiKey {
    if (-not (Test-Path $KeyFile)) { return $null }
    $secure = ConvertTo-SecureString (Get-Content -Raw $KeyFile)
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
}

if ($Configure) {
    Ensure-AppDir
    Write-Host 'OpenAI API key setup (stored encrypted for this Windows user only).' -ForegroundColor Cyan
    $key = Read-Host 'Paste your OpenAI API key' -AsSecureString
    if ($key.Length -eq 0) { throw 'No API key entered.' }
    ConvertFrom-SecureString $key | Set-Content -Path $KeyFile -Encoding ASCII
    $model = Read-Host "Model (press Enter for $DefaultModel)"
    if ([string]::IsNullOrWhiteSpace($model)) { $model = $DefaultModel }
    $accessToken = [Guid]::NewGuid().ToString('N')
    @{ model = $model; accessToken = $accessToken } | ConvertTo-Json | Set-Content -Path $ConfigFile -Encoding UTF8
    Write-Host 'Saved. You can now run AI_Helper_Start.bat.' -ForegroundColor Green
    Write-Host "Connection code (copy this into the Quote System once): $accessToken" -ForegroundColor Yellow
    exit 0
}

if ($ClearKey) {
    Remove-Item $KeyFile -Force -ErrorAction SilentlyContinue
    Write-Host 'Stored API key was removed.' -ForegroundColor Yellow
    exit 0
}

if ($Status) {
    if (Test-Path $KeyFile) { Write-Host 'API key is configured.' }
    else { Write-Host 'API key is not configured.' }
    exit 0
}

function Send-Json($Context, $StatusCode, $Body) {
    $json = $Body | ConvertTo-Json -Depth 12 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $Context.Response.StatusCode = $StatusCode
    $Context.Response.ContentType = 'application/json; charset=utf-8'
    $Context.Response.Headers.Add('Access-Control-Allow-Origin', '*')
    $Context.Response.Headers.Add('Access-Control-Allow-Headers', 'Content-Type, X-FQ-Token')
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Context.Response.Close()
}

function Extract-OutputText($Response) {
    if ($Response.output_text) { return [string]$Response.output_text }
    foreach ($item in $Response.output) {
        foreach ($part in $item.content) {
            if ($part.type -eq 'output_text' -and $part.text) { return [string]$part.text }
        }
    }
    return $null
}

function Ask-RateMapper($Payload) {
    $apiKey = Get-ApiKey
    if (-not $apiKey) { throw 'API key is not configured. Run AI_Setup.bat first.' }
    $config = Get-Config
    $system = @'
You map freight forwarding rate spreadsheet columns. Return JSON only, with no Markdown.
Never invent a price, currency, route, carrier, charge, or container type.
Decide one kind: rate_master, safe_rate, charge_rule, or unknown.
For rate_master return mapping keys chargeId, mode, pol, pod, container, amount, currency, and optionally min for a minimum charge column.
For safe_rate return mapping keys place, port, size, amount.
For charge_rule return mapping keys id, name, section, modes, terms, directions, unit, currency.
Each mapping value must be either a source column name exactly as supplied, a constant such as "const:KRW", or an empty string if not available.
Also return confidence as a number from 0 to 1 and warnings as an array of short Korean strings.
Valid common chargeId values include OCEAN, THC, DOC, SEAL, TRUCK, CFS, LCL_FREIGHT, AIR_FREIGHT, FUEL, SECURITY, AWB, IMPORT_THC, IMPORT_CUSTOMS, DELIVERY, EXPORT_CUSTOMS.
If chargeId is not present as a source column but every row is one known charge, a const: value is allowed.
'@
    $input = @{ fileName = $Payload.fileName; headers = $Payload.headers; sampleRows = @($Payload.sampleRows | Select-Object -First 30) } | ConvertTo-Json -Depth 8 -Compress
    $body = @{ model = $config.model; input = @(@{ role = 'system'; content = @(@{ type = 'input_text'; text = $system }) }, @{ role = 'user'; content = @(@{ type = 'input_text'; text = $input }) }); text = @{ format = @{ type = 'json_object' } } } | ConvertTo-Json -Depth 12
    $headers = @{ Authorization = "Bearer $apiKey" }
    $response = Invoke-RestMethod -Method Post -Uri 'https://api.openai.com/v1/responses' -Headers $headers -ContentType 'application/json' -Body $body -TimeoutSec 90
    $text = Extract-OutputText $response
    if (-not $text) { throw 'The AI returned no text.' }
    $text = $text -replace '^```json\s*','' -replace '^```\s*','' -replace '\s*```$',''
    return ($text | ConvertFrom-Json)
}

$listener = [Net.HttpListener]::new()
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
$listener.Start()
Write-Host "AI helper is running at http://127.0.0.1:$Port/" -ForegroundColor Green
Write-Host 'Keep this window open while using AI rate mapping. Press Ctrl+C to stop.' -ForegroundColor Yellow

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        try {
            if ($ctx.Request.HttpMethod -eq 'OPTIONS') { Send-Json $ctx 200 @{ ok = $true }; continue }
            if ($ctx.Request.Url.AbsolutePath -eq '/health' -and $ctx.Request.HttpMethod -eq 'GET') {
                Send-Json $ctx 200 @{ ok = $true; configured = [bool](Get-ApiKey); model = (Get-Config).model }; continue
            }
            if ($ctx.Request.Url.AbsolutePath -ne '/analyze' -or $ctx.Request.HttpMethod -ne 'POST') { Send-Json $ctx 404 @{ ok = $false; error = 'Not found.' }; continue }
            $expectedToken = [string](Get-Config).accessToken
            if ([string]::IsNullOrWhiteSpace($expectedToken) -or $ctx.Request.Headers['X-FQ-Token'] -ne $expectedToken) { Send-Json $ctx 401 @{ ok = $false; error = 'Invalid local connection code.' }; continue }
            if ($ctx.Request.ContentLength64 -gt 250000) { Send-Json $ctx 413 @{ ok = $false; error = 'Request is too large.' }; continue }
            $reader = [IO.StreamReader]::new($ctx.Request.InputStream, $ctx.Request.ContentEncoding)
            $payload = ($reader.ReadToEnd() | ConvertFrom-Json)
            $result = Ask-RateMapper $payload
            Send-Json $ctx 200 @{ ok = $true; result = $result }
        } catch {
            Send-Json $ctx 400 @{ ok = $false; error = $_.Exception.Message }
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
