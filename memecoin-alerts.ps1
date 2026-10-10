# Memecoin alerts -> Telegram (Windows PowerShell 5.1, no dependency)
#
#   -DryRun   : print the result in the console, send nothing, do not mark tokens as seen
#   -Test     : send a test message and exit
#   -Setup    : find your chat_id (after you sent /start to your bot) and send a test message
#   -Sync     : share the list of already-sent tokens through git; on the PC, scan only
#               when the cloud loop is down (and restart it)
#   -Install  : register the Windows scheduled task (cloud watchdog, every minute)
#   -Uninstall: remove the scheduled task
#
# Source file is kept ASCII-only on purpose (PowerShell 5.1 misreads UTF-8 without BOM).
param(
    [switch]$DryRun,
    [switch]$Setup,
    [switch]$Test,
    [switch]$Sync,
    [switch]$Install,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Root       = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $Root 'config.json'
$SeenPath   = Join-Path $Root 'seen.txt'
$LogPath    = Join-Path $Root 'alerts.log'
$TaskName   = 'MemecoinAlerts'

# ---------------------------------------------------------------- filters ----
$Filters = @{
    Chains         = @('solana', 'ethereum', 'base', 'bsc')
    MinAgeHours    = 0.33     # 20 minutes: early alerts, but past the very first minutes
    MaxAgeHours    = 72       # still "new"
    MinLiquidity   = 30000    # USD
    MinVolumeH1    = 25000    # USD traded in the last hour
    MinTxnsH1      = 150      # buys + sells in the last hour
    MinBuyRatioH1  = 0.45     # buys / (buys + sells)
    MinMarketCap   = 100000
    MaxMarketCap   = 50000000
    MinLiqToMcap   = 0.02     # liquidity / market cap, thin pools are easy to drain
    MinPriceChgH1  = -30      # % ; skip tokens already dumping
    RequireSocial  = $true    # at least one social link AND one website
    MaxRugScore    = 30       # RugCheck normalised score (Solana), lower = safer
    MaxTaxPct      = 10       # buy / sell tax (EVM)
    MaxPerMessage  = 8
}
$GoPlusChain = @{ ethereum = '1'; bsc = '56'; base = '8453'; arbitrum = '42161'; polygon = '137' }

# ---------------------------------------------------------------- helpers ----
function Write-Log([string]$msg) {
    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Write-Host $line
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

function Get-Json([string]$url) {
    for ($i = 0; $i -lt 3; $i++) {
        try {
            return Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 30 -Headers @{ accept = 'application/json' }
        } catch {
            Start-Sleep -Seconds (2 * ($i + 1))
        }
    }
    return $null
}

function Num($v) {
    if ($null -eq $v) { return 0.0 }
    try { return [double]::Parse([string]$v, [Globalization.CultureInfo]::InvariantCulture) } catch { return 0.0 }
}

function Fmt-Usd([double]$n) {
    if ($n -ge 1e6) { return ('${0:0.##}M' -f ($n / 1e6)) }
    if ($n -ge 1e3) { return ('${0:0.#}k' -f ($n / 1e3)) }
    return ('${0:0}' -f $n)
}

function Esc([string]$s) {
    if (-not $s) { return '' }
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function Invoke-Git {
    $ErrorActionPreference = 'Continue'   # git reports progress on stderr
    & git -C $Root @args 2>$null | Out-Null
    return $LASTEXITCODE
}

function Emoji([int]$code) { return [char]::ConvertFromUtf32($code) }

function Load-Config {
    # Cloud runs (GitHub Actions) pass the credentials as environment variables
    if ($env:TELEGRAM_BOT_TOKEN -and $env:TELEGRAM_CHAT_ID) {
        return [pscustomobject]@{ bot_token = $env:TELEGRAM_BOT_TOKEN; chat_id = $env:TELEGRAM_CHAT_ID }
    }
    if (-not (Test-Path $ConfigPath)) { throw "config.json introuvable dans $Root" }
    return Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Send-Telegram($cfg, [string]$html) {
    $body = @{
        chat_id                  = "$($cfg.chat_id)"
        text                     = $html
        parse_mode               = 'HTML'
        disable_web_page_preview = $true
    } | ConvertTo-Json
    $bytes = [Text.Encoding]::UTF8.GetBytes($body)
    $url = 'https://api.telegram.org/bot{0}/sendMessage' -f $cfg.bot_token
    Invoke-RestMethod -Uri $url -Method Post -Body $bytes -ContentType 'application/json; charset=utf-8' -TimeoutSec 30 | Out-Null
}

# ------------------------------------------------------------------ modes ----
if ($Uninstall) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Tache '$TaskName' supprimee."
    exit 0
}

if ($Install) {
    $script  = $MyInvocation.MyCommand.Path
    # conhost --headless keeps a console window from flashing at every run
    $action  = New-ScheduledTaskAction -Execute 'conhost.exe' -Argument "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`" -Sync"
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 1)
    $set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 4)
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $set -Force | Out-Null
    Write-Host "Tache '$TaskName' creee : execution chaque minute (premiere dans 1 minute)."
    exit 0
}

if ($Setup) {
    $cfg = Load-Config
    if (-not $cfg.bot_token -or $cfg.bot_token -like 'COLLE*') { throw 'Colle d''abord le token du bot dans config.json (champ bot_token).' }
    $upd = Get-Json ('https://api.telegram.org/bot{0}/getUpdates' -f $cfg.bot_token)
    $msgs = @($upd.result | Where-Object { $_.message } | ForEach-Object { $_.message })
    if ($msgs.Count -eq 0) { throw 'Aucun message recu par le bot. Ouvre ton bot dans Telegram, envoie /start, puis relance -Setup.' }
    $chat = $msgs[-1].chat
    $cfg.chat_id = "$($chat.id)"
    $cfg | ConvertTo-Json | Set-Content -Path $ConfigPath -Encoding UTF8
    Send-Telegram $cfg ((Emoji 0x2705) + ' Connexion OK. Tu recevras ici les alertes memecoins.')
    Write-Host "chat_id enregistre ($($chat.id)). Message de test envoye."
    exit 0
}

if ($Test) {
    Send-Telegram (Load-Config) ((Emoji 0x2705) + ' Test OK : les alertes memecoins tournent.')
    Write-Host 'Message de test envoye.'
    exit 0
}

# Alerts show Swiss/French local time, also when the script runs on a UTC server
function Get-LocalTime {
    foreach ($id in 'Europe/Zurich', 'W. Europe Standard Time') {
        try { return [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, $id) } catch { }
    }
    return Get-Date
}

# ------------------------------------------------------------------- scan ----
# On the PC, the cloud loop does the scanning; take over only while it is down
if ($Sync -and -not $env:GITHUB_ACTIONS) {
    $gh = Join-Path $env:ProgramFiles 'GitHub CLI\gh.exe'
    $ErrorActionPreference = 'Continue'
    Push-Location $Root
    $runs = & $gh run list --workflow alerts.yml --limit 5 --json status,updatedAt 2>$null | ConvertFrom-Json
    # a run that ended moments ago is handing over to the next one
    $alive = @($runs | Where-Object {
        $_.status -ne 'completed' -or ([DateTime]::UtcNow - [DateTime]::Parse($_.updatedAt).ToUniversalTime()).TotalMinutes -lt 3
    }).Count
    if ($alive -eq 0) { & $gh workflow run alerts.yml 2>$null | Out-Null }
    Pop-Location
    $ErrorActionPreference = 'Stop'
    if ($alive -gt 0) { exit 0 }
    Write-Log 'cloud inactif : relance demandee, scan depuis le PC'
}

if ($Sync -and (Invoke-Git pull -q --rebase) -ne 0) { [void](Invoke-Git rebase --abort) }

# seen.txt: one "chain:address<TAB>unix time" per line
$seen = @{}
$sent = $false
if (Test-Path $SeenPath) {
    foreach ($line in [IO.File]::ReadAllLines($SeenPath)) {
        $parts = $line.Trim().Split("`t")
        if ($parts.Count -eq 2) { $seen[$parts[0]] = [long]$parts[1] }
    }
}
$now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

# 1. Candidates: tokens with a fresh DexScreener profile or active boosts
$byChain = @{}
foreach ($feed in 'token-profiles/latest/v1', 'token-boosts/latest/v1', 'token-boosts/top/v1') {
    $items = Get-Json "https://api.dexscreener.com/$feed"
    foreach ($it in @($items)) {
        if (-not $it.tokenAddress -or $Filters.Chains -notcontains $it.chainId) { continue }
        if ($seen.ContainsKey("$($it.chainId):$($it.tokenAddress)")) { continue }
        if (-not $byChain.ContainsKey($it.chainId)) { $byChain[$it.chainId] = New-Object 'System.Collections.Generic.HashSet[string]' }
        [void]$byChain[$it.chainId].Add($it.tokenAddress)
    }
}

# 2. Market data, keep the deepest pool of each token
$best = @{}
foreach ($chain in $byChain.Keys) {
    $addrs = @($byChain[$chain])
    for ($i = 0; $i -lt $addrs.Count; $i += 30) {
        $chunk = $addrs[$i..([Math]::Min($i + 29, $addrs.Count - 1))] -join ','
        $pairs = Get-Json "https://api.dexscreener.com/tokens/v1/$chain/$chunk"
        foreach ($p in @($pairs)) {
            if (-not $p.baseToken) { continue }
            $key = "$($chain):$($p.baseToken.address)"
            if (-not $best.ContainsKey($key) -or (Num $p.liquidity.usd) -gt (Num $best[$key].liquidity.usd)) { $best[$key] = $p }
        }
    }
}

# 3. Hype + sanity filters
$picked = @()
foreach ($key in $best.Keys) {
    $p = $best[$key]
    if (-not $p.pairCreatedAt) { continue }
    $ageH  = ($now - ([long]$p.pairCreatedAt / 1000)) / 3600
    $liq   = Num $p.liquidity.usd
    $volH1 = Num $p.volume.h1
    $buys  = Num $p.txns.h1.buys
    $sells = Num $p.txns.h1.sells
    $mcap  = Num $p.marketCap
    if ($mcap -le 0) { $mcap = Num $p.fdv }
    $chgH1 = Num $p.priceChange.h1

    if ($ageH -lt $Filters.MinAgeHours -or $ageH -gt $Filters.MaxAgeHours) { continue }
    if ($liq -lt $Filters.MinLiquidity -or $volH1 -lt $Filters.MinVolumeH1) { continue }
    if (($buys + $sells) -lt $Filters.MinTxnsH1) { continue }
    if (($buys / ($buys + $sells)) -lt $Filters.MinBuyRatioH1) { continue }
    if ($mcap -lt $Filters.MinMarketCap -or $mcap -gt $Filters.MaxMarketCap) { continue }
    if (($liq / $mcap) -lt $Filters.MinLiqToMcap) { continue }
    if ($chgH1 -lt $Filters.MinPriceChgH1) { continue }
    if ($Filters.RequireSocial -and (@($p.info.socials | Where-Object { $_.url }).Count -eq 0 -or @($p.info.websites | Where-Object { $_.url }).Count -eq 0)) { continue }

    $picked += [pscustomobject]@{
        Key = $key; Pair = $p; AgeH = $ageH; Liq = $liq; VolH1 = $volH1
        Txns = $buys + $sells; Mcap = $mcap; ChgH1 = $chgH1; Safety = ''
    }
}

# 4. Contract safety (RugCheck on Solana, GoPlus on EVM). No verdict = not alerted.
$final = @()
foreach ($c in ($picked | Sort-Object VolH1 -Descending)) {
    $chain = $c.Pair.chainId
    $addr  = $c.Pair.baseToken.address
    $ok = $false
    if ($chain -eq 'solana') {
        $r = Get-Json "https://api.rugcheck.xyz/v1/tokens/$addr/report/summary"
        if ($r -and $null -ne $r.score_normalised) {
            $danger = @($r.risks | Where-Object { $_.level -eq 'danger' })
            if ($danger.Count -eq 0 -and (Num $r.score_normalised) -le $Filters.MaxRugScore) {
                $ok = $true
                $c.Safety = 'RugCheck {0}/100, LP verrouillee {1:0}%' -f $r.score_normalised, (Num $r.lpLockedPct)
            }
        }
    } elseif ($GoPlusChain.ContainsKey($chain)) {
        $r = Get-Json ("https://api.gopluslabs.io/api/v1/token_security/{0}?contract_addresses={1}" -f $GoPlusChain[$chain], $addr)
        $g = $null
        if ($r -and $r.result) { $g = $r.result.PSObject.Properties | Select-Object -First 1 -ExpandProperty Value }
        if ($g -and $null -ne $g.is_honeypot) {
            $bad = ($g.is_honeypot -eq '1') -or ($g.cannot_sell_all -eq '1') -or ($g.cannot_buy -eq '1') -or
                   ($g.is_open_source -ne '1') -or ($g.is_mintable -eq '1') -or ($g.hidden_owner -eq '1') -or
                   ($g.transfer_pausable -eq '1') -or ($g.can_take_back_ownership -eq '1') -or
                   ((Num $g.buy_tax) * 100 -gt $Filters.MaxTaxPct) -or ((Num $g.sell_tax) * 100 -gt $Filters.MaxTaxPct)
            if (-not $bad) {
                $ok = $true
                $c.Safety = 'GoPlus OK, taxes {0:0.#}% / {1:0.#}%' -f ((Num $g.buy_tax) * 100), ((Num $g.sell_tax) * 100)
            }
        }
    }
    if ($ok) { $final += $c }
    Start-Sleep -Milliseconds 400
}

Write-Log ("candidats={0}  hype={1}  retenus={2}" -f $best.Count, $picked.Count, $final.Count)

# 5. Message
if ($final.Count -gt 0) {
    $top = @($final | Select-Object -First $Filters.MaxPerMessage)
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine(('{0} <b>{1} nouveau(x) memecoin(s)</b> - {2}' -f (Emoji 0x1F680), $top.Count, (Get-LocalTime).ToString('HH:mm')))
    foreach ($c in $top) {
        $p = $c.Pair
        $links = @()
        $links += '<a href="{0}">DexScreener</a>' -f $p.url
        foreach ($s in @($p.info.socials)) { if ($s.url) { $links += '<a href="{0}">{1}</a>' -f (Esc $s.url), (Esc $s.type) } }
        foreach ($w in @($p.info.websites) | Select-Object -First 1) { if ($w.url) { $links += '<a href="{0}">site</a>' -f (Esc $w.url) } }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(('<b>{0}</b> ({1}) - {2}' -f (Esc $p.baseToken.symbol), (Esc $p.baseToken.name), $p.chainId))
        [void]$sb.AppendLine(('Age {0:0}h | MC {1} | Liq {2}' -f $c.AgeH, (Fmt-Usd $c.Mcap), (Fmt-Usd $c.Liq)))
        [void]$sb.AppendLine(('1h : vol {0} | {1:0} txns | {2:+0.#;-0.#}%' -f (Fmt-Usd $c.VolH1), $c.Txns, $c.ChgH1))
        [void]$sb.AppendLine(('{0} {1}' -f (Emoji 0x1F6E1), $c.Safety))
        [void]$sb.AppendLine(('<code>{0}</code>' -f $p.baseToken.address))
        [void]$sb.AppendLine(($links -join ' | '))
    }
    [void]$sb.AppendLine('')
    [void]$sb.Append('<i>Filtre automatique, pas un conseil. Un token "propre" peut quand meme tomber a zero.</i>')
    $text = $sb.ToString()

    if ($DryRun) {
        Write-Host ''
        Write-Host $text
    } else {
        Send-Telegram (Load-Config) $text
        foreach ($c in $top) { $seen[$c.Key] = $now }
        $sent = $true
    }
}

# 6. Persist seen tokens (7 days), only when something was sent
if ($sent) {
    $keys = [string[]]@($seen.Keys)
    [Array]::Sort($keys, [StringComparer]::Ordinal)
    $lines = foreach ($k in $keys) { if (($now - $seen[$k]) -lt 604800) { "$k`t$($seen[$k])" } }
    [IO.File]::WriteAllText($SeenPath, ((@($lines) -join "`n") + "`n"), (New-Object Text.UTF8Encoding $false))
    if ($Sync) {
        [void](Invoke-Git add seen.txt)
        [void](Invoke-Git commit -q -m 'Update seen tokens')
        if ((Invoke-Git push -q) -ne 0) {
            if ((Invoke-Git pull -q --rebase) -ne 0) { [void](Invoke-Git rebase --abort) }
            [void](Invoke-Git push -q)
        }
    }
}
