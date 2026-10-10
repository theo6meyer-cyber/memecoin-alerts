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
$RejectPath = Join-Path $Root 'rejected.txt'
$LogPath    = Join-Path $Root 'alerts.log'
$TaskName   = 'MemecoinAlerts'

# ---------------------------------------------------------------- filters ----
$Filters = @{
    Chains         = @('solana', 'ethereum', 'base', 'bsc')
    MinAgeHours    = 0.33     # 20 minutes: early alerts, but past the very first minutes
    MaxAgeHours    = 6        # only freshly launched tokens
    MinLiquidity   = 30000    # USD
    MinVolumeH1    = 25000    # USD traded in the last hour
    MinTxnsH1      = 150      # buys + sells in the last hour
    MinBuyRatioH1  = 0.45     # buys / (buys + sells)
    MaxBuyRatioH1  = 0.85     # almost nobody selling = holders may be unable to sell
    MinMarketCap   = 100000
    MaxMarketCap   = 50000000
    MinLiqToMcap   = 0.02     # liquidity / market cap, thin pools are easy to drain
    MinPriceChgH1  = -30      # % ; skip tokens already dumping
    RequireSocial  = $true    # its own X account AND its own website (not a tweet, not a trading link)
    MaxRugScore    = 20       # RugCheck normalised score (Solana), lower = safer
    MaxTaxPct      = 5        # buy / sell tax (EVM)
    MinHolders     = 300
    MaxTop1Pct     = 5        # biggest wallet, pools and lockers excluded
    MaxTop10Pct    = 25       # ten biggest wallets together
    MaxCreatorPct  = 5        # share still held by the creator
    MaxInsiderPct  = 5        # share held by wallets flagged as insiders
    MinLpLockedPct = 80       # liquidity that cannot be pulled
    MaxCreatorTokens = 5      # tokens already launched by the same creator
    WashVolToLiq   = 15       # hourly volume / liquidity above this is flagged as possibly artificial
    RecheckMinutes = 30       # a rejected token is not re-examined before this delay
    MaxPerMessage  = 8
}
$GoPlusChain = @{ ethereum = '1'; bsc = '56'; base = '8453'; arbitrum = '42161'; polygon = '137' }
# A link to one of these is not a project website
$Platforms = 'x.com', 'twitter.com', 't.me', 'pump.fun', 'axiom.trade', 'dexscreener.com', 'gmgn.ai', 'bullx.io',
             'tiktok.com', 'instagram.com', 'youtube.com', 'youtu.be', 'reddit.com', 'facebook.com', 'wikipedia.org', 'truthsocial.com'

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

function Pct([double]$n) { return ('{0:0.#}%' -f $n) }

# The project must have its own X account and its own website
function Test-OwnPresence($info) {
    $account = $false
    foreach ($s in @($info.socials)) {
        if ($s.type -eq 'twitter' -and $s.url -match '^https?://(www\.)?(x|twitter)\.com/([A-Za-z0-9_]{1,15})/?(\?.*)?$') {
            if ('i', 'search', 'hashtag', 'home', 'intent', 'share' -notcontains $Matches[3]) { $account = $true }
        }
    }
    $site = $false
    foreach ($w in @($info.websites)) {
        if (-not $w.url) { continue }
        try { $h = ([Uri]$w.url).Host.ToLower() -replace '^www\.', '' } catch { continue }
        if (-not ($Platforms | Where-Object { $h -eq $_ -or $h.EndsWith(".$_") })) { $site = $true }
    }
    return ($account -and $site)
}

# Tokens sharing the same name. Bigger = another one has more liquidity AND more volume (ours is then
# probably the copy; liquidity alone is not enough, copies are often seeded with idle liquidity).
# Rival = another one is also heavily traded, so the name alone does not tell which is which.
function Get-Copies($p) {
    $res = Get-Json ('https://api.dexscreener.com/latest/dex/search?q=' + [Uri]::EscapeDataString($p.baseToken.name))
    $others = @($res.pairs | Where-Object {
        $_.baseToken -and $_.baseToken.name -eq $p.baseToken.name -and $_.baseToken.address -ne $p.baseToken.address
    })
    $liq = Num $p.liquidity.usd
    $vol = Num $p.volume.h24
    $bigger = @($others | Where-Object { (Num $_.liquidity.usd) -gt $liq -and (Num $_.volume.h24) -gt $vol }).Count -gt 0
    $rival  = @($others | Where-Object { (Num $_.volume.h24) -gt ($vol * 0.25) }).Count -gt 0
    return @{ Count = @($others | ForEach-Object { $_.baseToken.address } | Sort-Object -Unique).Count; Bigger = $bigger; Rival = $rival }
}

# Solana: RugCheck full report. Returns the reasons to reject (empty = pass) and a summary line.
function Test-SolanaToken([string]$addr) {
    $r = Get-Json "https://api.rugcheck.xyz/v1/tokens/$addr/report"
    $s = Get-Json "https://api.rugcheck.xyz/v1/tokens/$addr/report/summary"
    if (-not $r -or -not $s -or $null -eq $r.score_normalised) { return @{ Why = @('pas de rapport RugCheck'); Safety = '' } }
    $why = @()
    if ($r.rugged) { $why += 'marque comme rug' }
    if ($r.mintAuthority) { $why += 'le createur peut creer de nouveaux tokens' }
    if ($r.freezeAuthority) { $why += 'le createur peut geler les portefeuilles' }
    if ((Num $r.transferFee.pct) -gt 0) { $why += 'taxe sur les transferts' }
    foreach ($risk in @($r.risks | Where-Object { $_.level -eq 'danger' })) { $why += "risque RugCheck : $($risk.name)" }
    if ((Num $r.score_normalised) -gt $Filters.MaxRugScore) { $why += "score RugCheck $($r.score_normalised)/100" }
    if ((Num $r.totalHolders) -lt $Filters.MinHolders) { $why += "seulement $($r.totalHolders) detenteurs" }
    $lp = Num $s.lpLockedPct
    if ($lp -lt $Filters.MinLpLockedPct) { $why += "liquidite verrouillee a $(Pct $lp) seulement" }

    # Pools and lockers hold tokens too but are not people
    $skip = @{}
    if ($r.knownAccounts) {
        foreach ($k in $r.knownAccounts.PSObject.Properties) { if ('AMM', 'LOCKER' -contains $k.Value.type) { $skip[$k.Name] = 1 } }
    }
    $holders = @($r.topHolders | Where-Object { $_ -and -not $skip.ContainsKey([string]$_.owner) } | Sort-Object { Num $_.pct } -Descending)
    $top1 = 0.0; $top10 = 0.0; $insiders = 0.0; $creator = 0.0
    if ($holders.Count -gt 0) { $top1 = Num $holders[0].pct }
    foreach ($h in ($holders | Select-Object -First 10)) { $top10 += Num $h.pct }
    foreach ($h in $holders) {
        if ($h.insider) { $insiders += Num $h.pct }
        if ($h.owner -eq $r.creator) { $creator += Num $h.pct }
    }
    $supply = Num $r.token.supply
    if ($supply -gt 0) { $creator = [Math]::Max($creator, (Num $r.creatorBalance) / $supply * 100) }
    if ($top1 -gt $Filters.MaxTop1Pct) { $why += "un portefeuille detient $(Pct $top1)" }
    if ($top10 -gt $Filters.MaxTop10Pct) { $why += "les 10 plus gros detiennent $(Pct $top10)" }
    if ($creator -gt $Filters.MaxCreatorPct) { $why += "le createur detient $(Pct $creator)" }
    if ($insiders -gt $Filters.MaxInsiderPct) { $why += "inities : $(Pct $insiders)" }
    $launched = @($r.creatorTokens | Where-Object { $_ }).Count
    if ($launched -gt $Filters.MaxCreatorTokens) { $why += "createur en serie ($launched tokens)" }

    $safety = 'RugCheck {0}/100 | LP {1} | {2} detenteurs | top10 {3} | createur {4}' -f $r.score_normalised, (Pct $lp), $r.totalHolders, (Pct $top10), (Pct $creator)
    return @{ Why = $why; Safety = $safety }
}

# EVM chains: GoPlus token security
function Test-EvmToken([string]$chain, [string]$addr) {
    $r = Get-Json ("https://api.gopluslabs.io/api/v1/token_security/{0}?contract_addresses={1}" -f $GoPlusChain[$chain], $addr)
    $g = $null
    if ($r -and $r.result) { $g = $r.result.PSObject.Properties | Select-Object -First 1 -ExpandProperty Value }
    if (-not $g -or $null -eq $g.is_honeypot) { return @{ Why = @('pas de rapport GoPlus'); Safety = '' } }
    $why = @()
    $flags = [ordered]@{
        is_honeypot = 'honeypot'; cannot_sell_all = 'vente totale impossible'; cannot_buy = 'achat impossible'
        is_mintable = 'creation de tokens possible'; hidden_owner = 'proprietaire cache'; is_proxy = 'contrat modifiable (proxy)'
        transfer_pausable = 'transferts suspendables'; can_take_back_ownership = 'propriete recuperable'
        is_blacklisted = 'liste noire'; slippage_modifiable = 'taxes modifiables'; personal_slippage_modifiable = 'taxe par portefeuille'
        selfdestruct = 'autodestruction'; owner_change_balance = 'le proprietaire peut modifier les soldes'
        trading_cooldown = 'delai impose entre les ventes'; honeypot_with_same_creator = 'createur deja auteur d''un honeypot'
    }
    foreach ($f in $flags.Keys) { if ($g.$f -eq '1') { $why += $flags[$f] } }
    if ($g.is_open_source -ne '1') { $why += 'code non publie' }
    $buyTax = (Num $g.buy_tax) * 100; $sellTax = (Num $g.sell_tax) * 100
    if ($buyTax -gt $Filters.MaxTaxPct -or $sellTax -gt $Filters.MaxTaxPct) { $why += "taxes $(Pct $buyTax) / $(Pct $sellTax)" }
    if ((Num $g.holder_count) -lt $Filters.MinHolders) { $why += "seulement $($g.holder_count) detenteurs" }
    $creator = [Math]::Max((Num $g.creator_percent), (Num $g.owner_percent)) * 100
    if ($creator -gt $Filters.MaxCreatorPct) { $why += "le createur detient $(Pct $creator)" }
    $people = @($g.holders | Where-Object { $_ -and $_.is_contract -ne 1 -and $_.is_locked -ne 1 } | Sort-Object { Num $_.percent } -Descending)
    $top1 = 0.0; $top10 = 0.0
    if ($people.Count -gt 0) { $top1 = (Num $people[0].percent) * 100 }
    foreach ($h in ($people | Select-Object -First 10)) { $top10 += (Num $h.percent) * 100 }
    if ($top1 -gt $Filters.MaxTop1Pct) { $why += "un portefeuille detient $(Pct $top1)" }
    if ($top10 -gt $Filters.MaxTop10Pct) { $why += "les 10 plus gros detiennent $(Pct $top10)" }
    $lpText = 'LP n/d'
    $lpHolders = @($g.lp_holders | Where-Object { $_ })
    if ($lpHolders.Count -gt 0) {
        $locked = 0.0
        foreach ($h in $lpHolders) { if ($h.is_locked -eq 1) { $locked += (Num $h.percent) * 100 } }
        if ($locked -lt $Filters.MinLpLockedPct) { $why += "liquidite verrouillee a $(Pct $locked) seulement" }
        $lpText = "LP $(Pct $locked)"
    }
    $safety = 'GoPlus OK | taxes {0} / {1} | {2} | {3} detenteurs | top10 {4}' -f (Pct $buyTax), (Pct $sellTax), $lpText, $g.holder_count, (Pct $top10)
    return @{ Why = $why; Safety = $safety }
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
    $ErrorActionPreference = 'Continue'
    $alive = 0
    try {
        # Reuse git's stored GitHub credential (the gh CLI is not signed in inside scheduled tasks)
        $env:GIT_TERMINAL_PROMPT = '0'
        $env:GCM_INTERACTIVE = 'never'
        $repo = ((& git -C $Root remote get-url origin) -replace '^https://github\.com/', '') -replace '\.git$', ''
        $cred = 'protocol=https', 'host=github.com', '' | & git -C $Root credential fill 2>$null
        $token = ($cred | Where-Object { $_ -like 'password=*' }) -replace '^password=', ''
        $api = "https://api.github.com/repos/$repo/actions/workflows/alerts.yml"
        $headers = @{ Authorization = "Bearer $token"; Accept = 'application/vnd.github+json'; 'User-Agent' = 'memecoin-alerts' }
        $runs = (Invoke-RestMethod -Uri "$api/runs?per_page=5" -Headers $headers -TimeoutSec 30).workflow_runs
        # a run that ended moments ago is handing over to the next one
        $alive = @($runs | Where-Object {
            $_.status -ne 'completed' -or ([DateTime]::UtcNow - ([DateTime]$_.updated_at).ToUniversalTime()).TotalMinutes -lt 3
        }).Count
        if ($alive -eq 0) {
            Invoke-RestMethod -Uri "$api/dispatches" -Method Post -Headers $headers -Body '{"ref":"main"}' -ContentType 'application/json' -TimeoutSec 30 | Out-Null
        }
    } catch { }
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

# Tokens that failed the safety checks recently (kept on this machine only)
$rejected = @{}
if (Test-Path $RejectPath) {
    foreach ($line in [IO.File]::ReadAllLines($RejectPath)) {
        $parts = $line.Trim().Split("`t")
        if ($parts.Count -eq 2 -and ($now - [long]$parts[1]) -lt ($Filters.RecheckMinutes * 60)) { $rejected[$parts[0]] = [long]$parts[1] }
    }
}

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
    $buyRatio = $buys / ($buys + $sells)
    if ($buyRatio -lt $Filters.MinBuyRatioH1 -or $buyRatio -gt $Filters.MaxBuyRatioH1) { continue }
    if ($mcap -lt $Filters.MinMarketCap -or $mcap -gt $Filters.MaxMarketCap) { continue }
    if (($liq / $mcap) -lt $Filters.MinLiqToMcap) { continue }
    if ($chgH1 -lt $Filters.MinPriceChgH1) { continue }
    if ($Filters.RequireSocial -and -not (Test-OwnPresence $p.info)) { continue }
    if ($rejected.ContainsKey($key)) { continue }

    $picked += [pscustomobject]@{
        Key = $key; Pair = $p; AgeH = $ageH; Liq = $liq; VolH1 = $volH1
        Txns = $buys + $sells; Mcap = $mcap; ChgH1 = $chgH1; Safety = ''; Warn = @()
    }
}

# 4. Contract, holders and copycat checks. Any doubt = not alerted.
$final = @()
foreach ($c in ($picked | Sort-Object VolH1 -Descending)) {
    $chain = $c.Pair.chainId
    $addr  = $c.Pair.baseToken.address
    if ($chain -eq 'solana') { $check = Test-SolanaToken $addr }
    elseif ($GoPlusChain.ContainsKey($chain)) { $check = Test-EvmToken $chain $addr }
    else { $check = @{ Why = @('pas de controle disponible sur cette blockchain'); Safety = '' } }
    $why = @($check.Why)

    $copies = Get-Copies $c.Pair
    if ($copies.Bigger) { $why += 'copie probable : un token du meme nom est plus gros et plus echange' }
    elseif ($copies.Rival) { $c.Warn += "$($copies.Count) autre(s) token(s) portent ce nom, dont un tres echange : verifie l'adresse" }
    elseif ($copies.Count -gt 0) { $c.Warn += "$($copies.Count) copie(s) portent ce nom : verifie l'adresse" }
    $volToLiq = $c.VolH1 / $c.Liq
    if ($volToLiq -gt $Filters.WashVolToLiq) { $c.Warn += ('volume = {0:0}x la liquidite en 1h, peut etre artificiel' -f $volToLiq) }

    if ($why.Count -eq 0) {
        $c.Safety = $check.Safety
        $final += $c
    } else {
        Write-Log ("rejete {0} : {1}" -f $c.Pair.baseToken.symbol, ($why -join ' ; '))
        if (-not $DryRun) { [IO.File]::AppendAllText($RejectPath, "$($c.Key)`t$now`n") }
    }
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
        foreach ($w in $c.Warn) { [void]$sb.AppendLine(('{0} {1}' -f (Emoji 0x26A0), $w)) }
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
