$ErrorActionPreference = "Stop"

$root = Join-Path $PSScriptRoot "..\com.pricop.powerpoint-tools.sdPlugin\imgs\actions"
New-Item -ItemType Directory -Force -Path $root | Out-Null

$bgDefs = @'
<defs>
  <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
    <stop offset="0" stop-color="#182B37"/>
    <stop offset="0.55" stop-color="#0D1820"/>
    <stop offset="1" stop-color="#081116"/>
  </linearGradient>
  <linearGradient id="teal" x1="0" y1="0" x2="1" y2="1">
    <stop offset="0" stop-color="#52F4DF"/>
    <stop offset="1" stop-color="#12CDB5"/>
  </linearGradient>
  <filter id="glow" x="-40%" y="-40%" width="180%" height="180%">
    <feGaussianBlur stdDeviation="2.2" result="blur"/>
    <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
  </filter>
</defs>
<rect x="5" y="5" width="134" height="134" rx="26" fill="url(#bg)" stroke="#21414E" stroke-width="2"/>
<rect x="7" y="7" width="130" height="130" rx="24" fill="none" stroke="#1BD7C0" stroke-opacity=".16"/>
'@

function Svg([string]$body) {
@"
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 144 144">
$bgDefs
$body
</svg>
"@
}

$white = '#F8FAFC'
$muted = '#365666'

$icons = @{
'align-left' = @"
<path d="M37 37v70" stroke="$white" stroke-width="9" stroke-linecap="round"/>
<rect x="50" y="45" width="42" height="18" rx="7" fill="url(#teal)" filter="url(#glow)"/>
<rect x="50" y="78" width="62" height="18" rx="7" fill="$white"/>
"@
'align-center' = @"
<path d="M72 31v82" stroke="$white" stroke-width="7" stroke-linecap="round" opacity=".95"/>
<rect x="45" y="43" width="54" height="19" rx="7" fill="url(#teal)" filter="url(#glow)"/>
<rect x="34" y="80" width="76" height="19" rx="7" fill="$white"/>
"@
'align-right' = @"
<path d="M107 37v70" stroke="$white" stroke-width="9" stroke-linecap="round"/>
<rect x="52" y="45" width="42" height="18" rx="7" fill="url(#teal)" filter="url(#glow)"/>
<rect x="32" y="78" width="62" height="18" rx="7" fill="$white"/>
"@
'align-top' = @"
<path d="M37 37h70" stroke="$white" stroke-width="9" stroke-linecap="round"/>
<rect x="46" y="50" width="20" height="52" rx="7" fill="url(#teal)" filter="url(#glow)"/>
<rect x="82" y="50" width="20" height="35" rx="7" fill="$muted"/>
"@
'align-middle' = @"
<path d="M31 72h82" stroke="$white" stroke-width="7" stroke-linecap="round"/>
<rect x="43" y="42" width="19" height="60" rx="7" fill="$muted"/>
<rect x="63" y="34" width="19" height="76" rx="7" fill="url(#teal)" filter="url(#glow)"/>
<rect x="83" y="46" width="19" height="52" rx="7" fill="$white"/>
"@
'align-bottom' = @"
<path d="M37 107h70" stroke="$white" stroke-width="9" stroke-linecap="round"/>
<rect x="46" y="42" width="20" height="52" rx="7" fill="url(#teal)" filter="url(#glow)"/>
<rect x="82" y="59" width="20" height="35" rx="7" fill="$white"/>
"@
'distribute-horizontal' = @"
<path d="M27 38v68M117 38v68" stroke="$white" stroke-width="8" stroke-linecap="round"/>
<rect x="43" y="53" width="15" height="38" rx="6" fill="#5AA8A0"/>
<rect x="65" y="45" width="15" height="54" rx="6" fill="url(#teal)" filter="url(#glow)"/>
<rect x="87" y="53" width="15" height="38" rx="6" fill="#5AA8A0"/>
"@
'distribute-vertical' = @"
<path d="M38 27h68M38 117h68" stroke="$white" stroke-width="8" stroke-linecap="round"/>
<rect x="53" y="43" width="38" height="15" rx="6" fill="#5AA8A0"/>
<rect x="45" y="65" width="54" height="15" rx="6" fill="url(#teal)" filter="url(#glow)"/>
<rect x="53" y="87" width="38" height="15" rx="6" fill="#5AA8A0"/>
"@
'same-width' = @"
<rect x="44" y="53" width="56" height="47" rx="8" fill="#15313A" stroke="$white" stroke-width="7"/>
<path d="M31 43v58M113 43v58" stroke="url(#teal)" stroke-width="8" stroke-linecap="round" filter="url(#glow)"/>
<path d="M31 43h10M31 101h10M103 43h10M103 101h10" stroke="url(#teal)" stroke-width="8" stroke-linecap="round"/>
"@
'same-height' = @"
<rect x="53" y="44" width="47" height="56" rx="8" fill="#15313A" stroke="$white" stroke-width="7"/>
<path d="M43 31h58M43 113h58" stroke="url(#teal)" stroke-width="8" stroke-linecap="round" filter="url(#glow)"/>
<path d="M43 31v10M101 31v10M43 103v10M101 103v10" stroke="url(#teal)" stroke-width="8" stroke-linecap="round"/>
"@
'same-size' = @"
<rect x="48" y="48" width="50" height="50" rx="9" fill="#17343D" stroke="$white" stroke-width="7"/>
<path d="M36 48V34h14M108 48V34H94M36 96v14h14M108 96v14H94" stroke="url(#teal)" stroke-width="8" stroke-linecap="round" stroke-linejoin="round" filter="url(#glow)"/>
"@
'rectangle' = @"
<rect x="31" y="46" width="82" height="52" rx="7" fill="#17353D" stroke="$white" stroke-width="8"/>
<path d="M35 92L108 50" stroke="url(#teal)" stroke-width="4" opacity=".35"/>
"@
'rounded-rectangle' = @"
<rect x="31" y="42" width="82" height="60" rx="20" fill="#17353D" stroke="$white" stroke-width="8"/>
<path d="M42 91c16 10 47 10 61-2" stroke="url(#teal)" stroke-width="5" opacity=".55" stroke-linecap="round"/>
"@
'shadow-off' = @"
<rect x="47" y="49" width="57" height="52" rx="10" fill="#1C5B59" opacity=".55"/>
<rect x="35" y="37" width="61" height="55" rx="10" fill="#17343D" stroke="$white" stroke-width="8"/>
<path d="M31 113L113 31" stroke="$white" stroke-width="8" stroke-linecap="round"/>
"@
'shadow-on' = @"
<rect x="50" y="54" width="62" height="55" rx="10" fill="url(#teal)" opacity=".72" filter="url(#glow)"/>
<rect x="32" y="36" width="66" height="58" rx="11" fill="#17343D" stroke="$white" stroke-width="8"/>
"@
'border-off' = @"
<rect x="34" y="34" width="76" height="76" rx="12" fill="#17343D" stroke="$white" stroke-width="7" stroke-dasharray="13 10"/>
<path d="M28 116L116 28" stroke="$white" stroke-width="8" stroke-linecap="round"/>
"@
'border-on' = @"
<rect x="34" y="34" width="76" height="76" rx="12" fill="#17343D" stroke="$white" stroke-width="10"/>
<path d="M38 42h68" stroke="url(#teal)" stroke-width="7" stroke-linecap="round" filter="url(#glow)"/>
"@
'match-style' = @"
<path d="M88 34l22 22-17 17-22-22z" fill="url(#teal)" stroke="$white" stroke-width="5" filter="url(#glow)"/>
<path d="M70 52L43 79c-10 10-4 26 7 26 4 0 8-2 12-6l29-29" fill="none" stroke="$white" stroke-width="8" stroke-linecap="round"/>
<path d="M44 101c14-2 20-8 23-21" stroke="url(#teal)" stroke-width="6" stroke-linecap="round"/>
"@
'clean-boxes' = @"
<path d="M72 31l6 15 15 6-15 6-6 15-6-15-15-6 15-6z" fill="url(#teal)" filter="url(#glow)"/>
<path d="M39 70l5 12 12 5-12 5-5 12-5-12-12-5 12-5z" fill="$white"/>
<path d="M99 79l4 9 9 4-9 4-4 9-4-9-9-4 9-4z" fill="$white"/>
"@
'test-connection' = @"
<rect x="30" y="39" width="84" height="58" rx="10" fill="#17343D" stroke="$white" stroke-width="7"/>
<path d="M65 56l24 16-24 16z" fill="url(#teal)" filter="url(#glow)"/>
<path d="M58 108h28M72 97v11" stroke="$white" stroke-width="7" stroke-linecap="round"/>
"@
}

foreach ($name in $icons.Keys) {
    $content = Svg $icons[$name]
    Set-Content -Path (Join-Path $root "$name.svg") -Value $content -Encoding UTF8
}

Write-Host "Generated $($icons.Count) duotone action icons."
