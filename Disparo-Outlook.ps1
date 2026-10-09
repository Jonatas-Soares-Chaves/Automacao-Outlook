Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

[System.Windows.Forms.Application]::EnableVisualStyles()

$Script:WorkbookPath = ''
$Script:AttachmentFolder = ''
$Script:AttachmentMode = 'Resumo Município-UF'
$Script:DistributionFolder = ''
$Script:InformativeFolder = ''
$Script:InformativeFile1 = ''
$Script:InformativeFile2 = ''
$Script:Rows = New-Object System.Collections.Generic.List[object]
$Script:Columns = @()
$Script:Sheets = @()
$Script:SelectedSheet = ''
$Script:Outlook = $null
$Script:LogPath = Join-Path $PSScriptRoot 'disparo-outlook.log'

$DefaultSubject = 'ENEM - 26 - RESUMO DE CONTRATAÇÃO - {{Cidade}}- {{UF}}'
$DefaultCC = @()

$DefaultHtml = @'
<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14pt; line-height: 1.45; color: #1f1f1f;">
<p>Prezado(a) Coordenador(a) Municipal, boa tarde!</p>

<p>Seguem, em anexo, o resumo de contratação e o informativo com as métricas e os valores da equipe de campo referentes ao evento ENEM 2026, que será realizado nos dias 8 e 15/11/2026.</p>

<p>Os valores constantes no resumo de contratação contabilizam os dois dias de aplicação. A alocação no Orion está liberada. O sistema está configurado para considerar o colaborador nos dois dias de aplicação quando ele for alocado.</p>

<ul>
  <li>Link com tutorial para alocar equipe no sistema Orion: <a href="https://cdn.cebraspe.org.br/Sistema/Manual/Orion/Videos/i/">CLIQUE AQUI</a></li>
  <li>Link para solicitar contratação de equipe extra: <a href="https://forms.cloud.microsoft/pages/responsepage.aspx?id=Io89vZLkoUa_zg75ooRBVEiwC8HQn9NGssLr-hdlLO1UNTBDMVFBOFg0SENNV1UwTDNENk05SVk1Ty4u&route=shorturl">CLIQUE AQUI</a></li>
</ul>

<ol>
  <li>A equipe de aplicação deverá chegar ao local às 9h no 1º dia e às 9h30 no 2º dia.</li>
  <li>Verificar se todos estão vestidos com camisa branca, exceto os Tradutores-Intérpretes de Libras, que devem estar com camisa preta, e calça jeans ou preta, portando documento de identificação original com foto, relógio analógico e caneta preta fabricada em material transparente.</li>
  <li>Conferir o documento de identificação de cada colaborador com os dados da Lista de Frequência e do Termo de Sigilo, Compromisso e Confidencialidade, e solicitar a assinatura em cada dia de aplicação, orientando que utilizem a própria caneta.</li>
  <li>Lembrar que no 2º dia deve ser feito rodízio de sala entre os Aplicadores, quando for o caso, com exceção dos Aplicadores Especializados. Os Chefes de Sala devem permanecer nas mesmas salas do 1º dia.</li>
  <li>Os Coordenadores de Aplicação serão locais. São requisitos para atuar como Coordenador de Aplicação: ensino superior completo, experiência anterior de participação em, no mínimo, 4 (quatro) exames do mesmo porte ou em aplicação de provas de concursos e vestibulares, além de participação e aprovação na capacitação presencial.</li>
  <li>Assistente de Coordenador de Local de Aplicação Principal: ensino superior completo, experiência anterior de participação em, no mínimo, 3 (três) exames do mesmo porte ou em aplicação de provas de concursos, vestibulares, exames e avaliações, além de participação e aprovação na capacitação presencial.</li>
</ol>

<p><strong>Observação:</strong> Coordenador(a), priorize a alocação dos Coordenadores de Aplicação e Assistentes em suas respectivas escolas, visto que somente após essa etapa será possível alocá-los no módulo de capacitação. Após a alocação desses Coordenadores e Assistentes, disponibilizaremos o módulo de capacitação para alocar os Coordenadores Estaduais, Coordenadores Municipais, Coordenadores de Aplicação e Assistentes de Aplicação em seus respectivos hotéis/salas de capacitação. Essa atividade deverá ser executada pelos Coordenadores Estaduais, Coordenadores de Polo e Coordenadores Municipais.</p>

<p>Prazo para conclusão da alocação de Coordenadores e Assistentes no Orion: <strong>14/09/2026</strong>. Após essa data, será realizado o pagamento do Auxílio Deslocamento referente à participação na capacitação.</p>

<p><strong>Atenção:</strong> oriente os Coordenadores e Assistentes a verificarem seus dados bancários cadastrados no Orion e, se necessário, realizarem a atualização antes do prazo. O pagamento do Auxílio Deslocamento será efetuado exclusivamente com base nas informações bancárias registradas no sistema.</p>

<ol start="7">
  <li>Todos os colaboradores que forem atuar no evento deverão ter cadastro no sistema do Cebraspe (Orion) e possuir conta bancária própria, uma vez que o Cebraspe não executa pagamento em conta de terceiro, conta conjunta, quando o colaborador é o 2º titular, ou conta salário. Para evitar qualquer imprevisto no pagamento dos colaboradores, cabe aos Coordenadores garantir que a informação esteja cadastrada corretamente no Orion.</li>
  <li>Os colaboradores que forem atuar na função de Aplicador Especializado deverão disponibilizar, em seu cadastro no Orion, o certificado que comprove tais requisitos.</li>
  <li>Caso haja solicitação de contratação extra, o Coordenador deverá preencher a solicitação no formulário disponibilizado pelo Cebraspe no link <a href="https://forms.cloud.microsoft/pages/responsepage.aspx?id=Io89vZLkoUa_zg75ooRBVEiwC8HQn9NGssLr-hdlLO1UNTBDMVFBOFg0SENNV1UwTDNENk05SVk1Ty4u&route=shorturl"><strong>Contratação de equipe extra - Norte - 26</strong></a>. A solicitação será analisada, podendo ou não ser deferida. É de suma importância que as informações preenchidas estejam corretas; caso contrário, a solicitação não será analisada e será descartada.</li>
</ol>

<p>Prazo para solicitação de contratação extra do estado de <strong>TO</strong>: <strong>18/09/2026</strong>.</p>
<p>Prazo para finalizar a alocação da equipe de campo, incluindo Chefe de Sala, Aplicador e demais colaboradores: <strong>20/09/2026</strong>.</p>
</div>
'@
$DefaultHtml = ''

function Get-DefaultSignatureHtml {
    $logoPath = Join-Path $PSScriptRoot 'cebraspe-logo.png'
    $logoHtml = if (Test-Path -LiteralPath $logoPath -PathType Leaf) {
        '<img src="cid:cebraspe-logo" alt="Cebraspe" width="108" style="display: block; border: 0; outline: none; text-decoration: none;">'
    }
    else {
        '<div style="font-size: 20pt; font-weight: 700; color: #0b4f8a; line-height: 1;">Cebraspe</div>'
    }
    return @"
<div style="font-family: Arial, Helvetica, sans-serif; font-size: 11pt; line-height: 1.35; color: #1f1f1f;">
  <p style="margin: 0 0 14px 0;">Atenciosamente,</p>
  <table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse: collapse; margin: 0 0 18px 0;">
    <tr>
      <td style="padding: 0 22px 0 0; vertical-align: middle;">
        $logoHtml
      </td>
      <td style="border-left: 2px solid #0b4f8a; padding: 0 0 0 22px; vertical-align: middle;">
        <div style="font-size: 14pt; font-weight: 700; color: #0b4f8a;">Jonatas Chaves</div>
        <div style="font-size: 11pt; font-weight: 700; color: #333333; margin-top: 6px;">Assistente de Suporte ao Negócio</div>
        <div style="font-size: 10.5pt; color: #333333; margin-top: 6px;">(61) 98382 2424 | <a href="http://www.cebraspe.org.br/" style="color: #0645ad;">www.cebraspe.org.br</a></div>
      </td>
    </tr>
  </table>
  <p style="font-size: 8pt; line-height: 1.45; color: #555555; margin: 0 0 10px 0;">Esta mensagem possui informação de interesse exclusivo do destinatário. A divulgação, sem justa causa, do conteúdo desta mensagem e de seus anexos constitui crime, nos termos do art. 153 do Código Penal Brasileiro. Caso esta mensagem seja recebida por engano, o destinatário deverá comunicar o fato via e-mail, promovendo, imediatamente, a eliminação do seu respectivo conteúdo.</p>
  <p style="font-size: 8pt; line-height: 1.45; color: #555555; margin: 0;">As informações citadas têm caráter de dado pessoal sensível, de acordo com a Lei nº 13.709, de 14 de agosto de 2018, denominada Lei Geral de Proteção de Dados Pessoais (LGPD), e não podem ser reproduzidas, publicadas ou disponibilizadas, sendo restrito o acesso somente às autoridades competentes para uso nos termos da referida lei.</p>
</div>
"@
}

function Add-InlineLogoAttachment {
    param($Mail)
    $logoPath = Join-Path $PSScriptRoot 'cebraspe-logo.png'
    if (-not (Test-Path -LiteralPath $logoPath -PathType Leaf)) { return }
    $attachment = $Mail.Attachments.Add($logoPath, 1, 0, 'cebraspe-logo.png')
    $pa = $attachment.PropertyAccessor
    $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', 'cebraspe-logo')
    $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3713001F', 'cebraspe-logo.png')
    $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B', $true)
}

function Normalize-Text {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $formD = $Text.Trim().ToUpperInvariant().Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($ch in $formD.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return (($sb.ToString().Normalize([Text.NormalizationForm]::FormC) -replace '[^A-Z0-9]+', ' ') -replace '\s+', ' ').Trim()
}

function Replace-Tokens {
    param([string]$Template, $Row)
    $out = $Template
    foreach ($name in $Script:Columns) {
        $value = [string]$Row.Values[$name]
        $out = $out.Replace('{{' + $name + '}}', $value)
    }
    $out = $out.Replace('{{Coordenador}}', [string]$Row.Name)
    $out = $out.Replace('{{Nome}}', [string]$Row.Name)
    $out = $out.Replace('{{Cidade}}', [string]$Row.City)
    $out = $out.Replace('{{Municipio}}', [string]$Row.City)
    $out = $out.Replace('{{Município}}', [string]$Row.City)
    $out = $out.Replace('{{UF}}', [string]$Row.UF)
    return $out
}

function Merge-Addresses {
    param([string]$ColumnAddresses, [string]$UserAddresses)
    $addresses = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($ColumnAddresses, $UserAddresses) + @($DefaultCC)) {
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        foreach ($address in ($value -split '[,;]')) {
            $clean = $address.Trim().Trim('<', '>')
            if ($clean -and -not $addresses.Contains($clean)) { $addresses.Add($clean) }
        }
    }
    return ($addresses -join '; ')
}

function Get-OutlookApplication {
    if ($Script:Outlook) {
        try {
            $null = $Script:Outlook.Session.Accounts.Count
            return $Script:Outlook
        }
        catch {
            $Script:Outlook = $null
        }
    }
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $null = $app.Session.Accounts.Count
        $Script:Outlook = $app
        return $Script:Outlook
    }
    catch {
        try {
            $app = New-Object -ComObject Outlook.Application
            $null = $app.Session.Accounts.Count
            $Script:Outlook = $app
            return $Script:Outlook
        }
        catch {
            Start-Process outlook.exe -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 4
            $app = New-Object -ComObject Outlook.Application
            $null = $app.Session.Accounts.Count
            $Script:Outlook = $app
            return $Script:Outlook
        }
    }
}

function Get-OutlookAccounts {
    $Script:Outlook = Get-OutlookApplication
    $accounts = @()
    foreach ($account in $Script:Outlook.Session.Accounts) {
        $accounts += [string]$account.SmtpAddress
    }
    return $accounts | Where-Object { $_ } | Sort-Object -Unique
}

function Read-WorkbookRows {
    param([string]$Path, [string]$SheetName)
    $excel = $null
    $wb = $null
    $items = New-Object System.Collections.Generic.List[object]
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Path)
        $ws = $null
        foreach ($sheet in $wb.Worksheets) {
            if (-not $SheetName -or $sheet.Name -eq $SheetName) { $ws = $sheet; break }
        }
        if (-not $ws) { throw "Aba '$SheetName' não encontrada." }
        $used = $ws.UsedRange
        $rowCount = $used.Rows.Count
        $colCount = $used.Columns.Count
        $headers = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $header = [string]$used.Cells.Item(1, $c).Text
            if ([string]::IsNullOrWhiteSpace($header)) { $header = "Coluna$c" }
            $headers += $header.Trim()
        }
        for ($r = 2; $r -le $rowCount; $r++) {
            $values = @{}
            $hasData = $false
            for ($c = 1; $c -le $colCount; $c++) {
                $value = [string]$used.Cells.Item($r, $c).Text
                if (-not [string]::IsNullOrWhiteSpace($value)) { $hasData = $true }
                $values[$headers[$c - 1]] = $value.Trim()
            }
            if ($hasData) {
                $items.Add([pscustomobject]@{ Line = $r; Values = $values }) | Out-Null
            }
        }
        return [pscustomobject]@{ Headers = $headers; Rows = $items }
    }
    finally {
        if ($wb) { $wb.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-SheetNames {
    param([string]$Path)
    $excel = $null
    $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Path)
        $names = @()
        foreach ($ws in $wb.Worksheets) { $names += [string]$ws.Name }
        return $names
    }
    finally {
        if ($wb) { $wb.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Read-CsvRows {
    param([string]$Path)
    $csv = Import-Csv -LiteralPath $Path
    $headers = @()
    if ($csv.Count -gt 0) { $headers = $csv[0].PSObject.Properties.Name }
    $items = New-Object System.Collections.Generic.List[object]
    $line = 2
    foreach ($row in $csv) {
        $values = @{}
        foreach ($h in $headers) { $values[$h] = [string]$row.$h }
        $items.Add([pscustomobject]@{ Line = $line; Values = $values }) | Out-Null
        $line++
    }
    return [pscustomobject]@{ Headers = $headers; Rows = $items }
}

function Pick-Column {
    param([string[]]$Headers, [string[]]$Hints)
    foreach ($hint in $Hints) {
        $hintNorm = Normalize-Text $hint
        foreach ($header in $Headers) {
            if ((Normalize-Text $header) -eq $hintNorm) { return $header }
        }
    }
    foreach ($hint in $Hints) {
        $hintNorm = Normalize-Text $hint
        foreach ($header in $Headers) {
            if ((Normalize-Text $header).Contains($hintNorm)) { return $header }
        }
    }
    return ''
}

function Build-PreparedRows {
    param([string]$EmailColumn, [string]$NameColumn, [string]$CityColumn, [string]$UFColumn, [string]$CCColumn)
    $prepared = New-Object System.Collections.Generic.List[object]
    $commonAttachments = Get-InformativeFiles
    foreach ($raw in $Script:RawRows) {
        $email = [string]$raw.Values[$EmailColumn]
        $name = if ($NameColumn) { [string]$raw.Values[$NameColumn] } else { '' }
        $city = if ($CityColumn) { [string]$raw.Values[$CityColumn] } else { '' }
        $uf = if ($UFColumn) { [string]$raw.Values[$UFColumn] } else { '' }
        $cc = if ($CCColumn) { [string]$raw.Values[$CCColumn] } else { '' }
        $attachments = @()
        $summaryAttachments = @()
        $distributionAttachments = @()
        if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
            $summaryAttachments = @(Find-FlatMunicipalFiles -RootFolder $Script:AttachmentFolder -City $city -UF $uf)
            $distributionAttachments = @(Find-DistributionAttachments -City $city -RootFolder $Script:DistributionFolder)
            $municipalAttachments = @($summaryAttachments + $distributionAttachments)
        }
        else {
            $municipalAttachments = @(Find-Attachments -City $city -UF $uf)
        }
        $attachments += @($municipalAttachments)
        $attachments += @($commonAttachments)
        if (-not $uf -and $Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo' -and $Script:AttachmentFolder) {
            $uf = Infer-UF-FromPath $Script:AttachmentFolder
        }
        if (-not $uf -and $Script:AttachmentMode -eq 'Distribuição por subpastas' -and $Script:AttachmentFolder) {
            $rootName = Split-Path -Leaf $Script:AttachmentFolder
            if ($rootName -match '^[A-Za-z]{2}$') { $uf = $rootName.ToUpperInvariant() }
        }
        if (-not $uf -and $attachments.Count -gt 0) {
            $uf = Infer-UF -FileName ([IO.Path]::GetFileNameWithoutExtension($attachments[0]))
        }
        $prepared.Add([pscustomobject]@{
                Line                       = $raw.Line
                To                         = $email
                CC                         = $cc
                Name                       = $name
                City                       = $city
                UF                         = $uf
                Attachments                = $attachments
                MunicipalAttachmentCount   = $municipalAttachments.Count
                SummaryAttachmentCount     = $summaryAttachments.Count
                DistributionAttachmentCount = $distributionAttachments.Count
                InformativeAttachmentCount = $commonAttachments.Count
                Values                     = $raw.Values
            }) | Out-Null
    }
    return $prepared
}

function Get-InformativeFiles {
    $files = @()
    if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
        if (![string]::IsNullOrWhiteSpace($Script:InformativeFolder) -and (Test-Path -LiteralPath $Script:InformativeFolder -PathType Container)) {
            foreach ($file in Get-ChildItem -LiteralPath $Script:InformativeFolder -File -Recurse | Sort-Object FullName) {
                $files += $file.FullName
            }
        }
        return $files
    }
    foreach ($path in @(
            $Script:InformativeFile1,
            $Script:InformativeFile2
        )) {
        if (![string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Leaf)) {
            $files += $path
        }
    }
    return $files
}

function Get-OutlookSignatureHtml {
    $signatureRoot = Join-Path $env:APPDATA 'Microsoft\Signatures'
    if (-not (Test-Path -LiteralPath $signatureRoot)) { return '' }
    $signatureFile = Get-ChildItem -LiteralPath $signatureRoot -Filter '*.htm' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
    if (-not $signatureFile) { return '' }
    try {
        return Get-Content -LiteralPath $signatureFile.FullName -Raw
    }
    catch {
        return ''
    }
}

function Test-SignatureHtml {
    param([string]$Html)
    if ([string]::IsNullOrWhiteSpace($Html)) { return $false }
    $text = [regex]::Replace($Html, '<[^>]+>', '')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    if (-not [string]::IsNullOrWhiteSpace($text)) { return $true }
    return ($Html -match '(?i)<img\b')
}

function Get-RequiredSignatureHtml {
    param($Mail)
    $signature = Get-OutlookSignatureHtml
    if (Test-SignatureHtml $signature) { return $signature }
    $defaultSignature = Get-DefaultSignatureHtml
    if (Test-SignatureHtml $defaultSignature) { return $defaultSignature }
    throw 'Assinatura obrigatória não encontrada.'
}

function Infer-UF {
    param([string]$FileName)
    if ($FileName -match '(?i)(?:^|[-_\s])([A-Z]{2})(?:$|[-_\s])') { return $matches[1].ToUpperInvariant() }
    return ''
}

function Infer-UF-FromPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $parts = @()
    $current = $Path
    while ($current) {
        $leaf = Split-Path -Leaf $current
        if ($leaf) { $parts += $leaf }
        $parent = Split-Path -Parent $current
        if ($parent -eq $current) { break }
        $current = $parent
    }
    foreach ($part in $parts) {
        if ($part -match '(?i)(?:^|[^A-Z])([A-Z]{2})(?:$|[^A-Z])') { return $matches[1].ToUpperInvariant() }
    }
    return ''
}

function Find-Attachments {
    param([string]$City, [string]$UF)
    if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
        return Find-ThreeTypeAttachments -City $City -UF $UF
    }
    if ($Script:AttachmentMode -eq 'Distribuição por subpastas') {
        return Find-DistributionAttachments -City $City
    }
    return Find-FlatMunicipalFiles -RootFolder $Script:AttachmentFolder -City $City -UF $UF
}

function Find-FlatMunicipalFiles {
    param([string]$RootFolder, [string]$City, [string]$UF)
    $foundFiles = @()
    if ([string]::IsNullOrWhiteSpace($RootFolder)) { return $foundFiles }
    if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) { return $foundFiles }
    $cityNorm = Normalize-Text $City
    $ufNorm = Normalize-Text $UF
    foreach ($file in Get-ChildItem -LiteralPath $RootFolder -File -Recurse) {
        $baseNorm = Normalize-Text $file.BaseName
        if (-not $cityNorm) { continue }
        $cityOk = $baseNorm.StartsWith($cityNorm) -or $baseNorm.Contains($cityNorm)
        $ufOk = $true
        if ($ufNorm) { $ufOk = ($baseNorm -match "(^| )$([regex]::Escape($ufNorm))( |$)") -or $baseNorm.EndsWith(' ' + $ufNorm) }
        if ($cityOk -and $ufOk) { $foundFiles += $file.FullName }
    }
    return $foundFiles
}

function Find-DistributionAttachments {
    param([string]$City, [string]$RootFolder = $Script:AttachmentFolder)
    $foundFiles = @()
    if ([string]::IsNullOrWhiteSpace($RootFolder)) { return $foundFiles }
    if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) { return $foundFiles }
    $cityNorm = Normalize-Text $City
    if (-not $cityNorm) { return $foundFiles }

    $folders = @(Get-ChildItem -LiteralPath $RootFolder -Directory)
    $selectedFolder = $folders | Where-Object { (Normalize-Text $_.Name) -eq $cityNorm } | Select-Object -First 1
    if (-not $selectedFolder) {
        $selectedFolder = $folders | Where-Object {
            $folderNorm = Normalize-Text $_.Name
            $folderNorm.StartsWith($cityNorm) -or $cityNorm.StartsWith($folderNorm) -or $folderNorm.Contains($cityNorm)
        } | Sort-Object Name | Select-Object -First 1
    }
    if (-not $selectedFolder) { return $foundFiles }

    foreach ($file in Get-ChildItem -LiteralPath $selectedFolder.FullName -File -Recurse | Sort-Object FullName) {
        $foundFiles += $file.FullName
    }
    return $foundFiles
}

function Find-ThreeTypeAttachments {
    param([string]$City, [string]$UF)
    $files = @()
    $summaryFiles = @(Find-FlatMunicipalFiles -RootFolder $Script:AttachmentFolder -City $City -UF $UF)
    $distributionFiles = @(Find-DistributionAttachments -City $City -RootFolder $Script:DistributionFolder)
    $files += $summaryFiles
    $files += $distributionFiles
    return @($files | Where-Object { $_ } | Sort-Object -Unique)
}

function Get-EditorHtml {
    try {
        if ($webEditor -and $webEditor.Document -and $webEditor.Document.Body) {
            return [string]$webEditor.Document.Body.InnerHtml
        }
    }
    catch { }
    return [string]$txtHtml.Text
}

function Convert-MarkdownTextToHtml {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $encoded = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $encoded = [regex]::Replace($encoded, '\*\*(.+?)\*\*', '<strong>$1</strong>')
    $encoded = [regex]::Replace($encoded, '__(.+?)__', '<u>$1</u>')
    $encoded = [regex]::Replace($encoded, '(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', '<em>$1</em>')
    $encoded = [regex]::Replace($encoded, '(https?://[^\s<]+)', '<a href="$1">$1</a>')
    $paragraphs = @()
    foreach ($block in ($encoded -split "(\r?\n){2,}")) {
        if ([string]::IsNullOrWhiteSpace($block)) { continue }
        $paragraphs += '<p>' + (($block -replace "\r?\n", '<br>')) + '</p>'
    }
    return ($paragraphs -join [Environment]::NewLine)
}

function Normalize-EditorHtml {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    if ($Text -match '<\s*(p|div|br|strong|b|u|em|i|ul|ol|li|a|table|span|font)\b') { return $Text }
    return Convert-MarkdownTextToHtml $Text
}

function Set-EditorHtml {
    param([string]$Html)
    if (-not $webEditor) { return }
    $Html = Normalize-EditorHtml $Html
    $document = @"
<!doctype html>
<html>
<head>
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <style>
    body {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 14pt;
      line-height: 1.45;
      color: #1f1f1f;
      margin: 12px;
      background: #ffffff;
    }
    p { margin: 0 0 12px 0; }
    ul, ol { margin-top: 0; }
  </style>
</head>
<body contenteditable="true">$Html</body>
</html>
"@
    $webEditor.DocumentText = $document
    $txtHtml.Text = $Html
}

function Update-EditorToolbarState {
    try {
        if (-not $webEditor -or -not $webEditor.Document -or -not $webEditor.Document.DomDocument) { return }
        foreach ($command in @('Bold', 'Italic', 'Underline', 'InsertUnorderedList', 'InsertOrderedList', 'JustifyLeft', 'JustifyCenter', 'JustifyRight')) {
            if (-not $Script:EditorButtonsByCommand.ContainsKey($command)) { continue }
            $button = $Script:EditorButtonsByCommand[$command]
            $active = $false
            try { $active = [bool]$webEditor.Document.DomDocument.queryCommandState($command) } catch { $active = $false }
            if ($active) {
                $button.BackColor = [Drawing.Color]::FromArgb(214, 232, 255)
                $button.FlatStyle = 'Popup'
            }
            else {
                $button.BackColor = [Drawing.SystemColors]::Control
                $button.FlatStyle = 'System'
            }
        }
        $txtHtml.Text = Get-EditorHtml
    }
    catch { }
}

function Register-EditorEvents {
    try {
        if (-not $webEditor -or -not $webEditor.Document) { return }
        $handler = { Update-EditorToolbarState }
        $webEditor.Document.AttachEventHandler('onselectionchange', $handler)
        $webEditor.Document.AttachEventHandler('onkeyup', $handler)
        $webEditor.Document.AttachEventHandler('onmouseup', $handler)
    }
    catch { }
}

function Invoke-EditorCommand {
    param([string]$Command, $Value = $null)
    try {
        if ($webEditor -and $webEditor.Document) {
            [void]$webEditor.Document.ExecCommand($Command, $false, $Value)
            $txtHtml.Text = Get-EditorHtml
            Update-EditorToolbarState
            $webEditor.Focus()
        }
    }
    catch {
        Log-Line "Falha no comando do editor '$Command': $($_.Exception.Message)"
    }
}

function Open-HtmlEditorDialog {
    $dialog = New-Object Windows.Forms.Form
    $dialog.Text = 'Editar HTML do e-mail'
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object Drawing.Size(900, 620)
    $dialog.MinimumSize = New-Object Drawing.Size(720, 460)

    $box = New-Object Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ScrollBars = 'Both'
    $box.AcceptsReturn = $true
    $box.AcceptsTab = $true
    $box.WordWrap = $false
    $box.Font = New-Object Drawing.Font('Consolas', 10)
    $box.Text = Get-EditorHtml
    $box.Dock = 'Fill'
    $dialog.Controls.Add($box)

    $panel = New-Object Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = 46
    $dialog.Controls.Add($panel)

    $btnOk = New-Object Windows.Forms.Button
    $btnOk.Text = 'Aplicar'
    $btnOk.Location = New-Object Drawing.Point(680, 9)
    $btnOk.Size = New-Object Drawing.Size(90, 28)
    $btnOk.Anchor = 'Right,Top'
    $panel.Controls.Add($btnOk)

    $btnCancel = New-Object Windows.Forms.Button
    $btnCancel.Text = 'Cancelar'
    $btnCancel.Location = New-Object Drawing.Point(780, 9)
    $btnCancel.Size = New-Object Drawing.Size(90, 28)
    $btnCancel.Anchor = 'Right,Top'
    $panel.Controls.Add($btnCancel)

    $btnOk.Add_Click({
            Set-EditorHtml $box.Text
            $dialog.DialogResult = 'OK'
            $dialog.Close()
        })
    $btnCancel.Add_Click({
            $dialog.DialogResult = 'Cancel'
            $dialog.Close()
        })

    [void]$dialog.ShowDialog($form)
}

function Compose-Mail {
    param($Row, [switch]$SendNow)
    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'
    try {
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            try {
                $Script:Outlook = Get-OutlookApplication
                $mail = $Script:Outlook.CreateItem(0)
                if (-not $mail) { throw 'O Outlook não retornou um item de e-mail válido.' }
                $mail.BodyFormat = 2
                if ($cmbAccount.Text) {
                    foreach ($account in $Script:Outlook.Session.Accounts) {
                        if ([string]$account.SmtpAddress -eq $cmbAccount.Text) {
                            $mail.SendUsingAccount = $account
                            break
                        }
                    }
                }
                $mail.To = $Row.To
                $mail.CC = Merge-Addresses $Row.CC $txtDefaultCC.Text
                $mail.Subject = Replace-Tokens $txtSubject.Text $Row
                foreach ($att in @($Row.Attachments)) {
                    if (-not (Test-Path -LiteralPath $att -PathType Leaf)) {
                        throw "Anexo não encontrado: $att"
                    }
                    [void]$mail.Attachments.Add($att)
                }
                $body = Replace-Tokens (Get-EditorHtml) $Row
                $signature = Get-RequiredSignatureHtml $mail
                $body = $body + '<br>' + $signature
                Add-InlineLogoAttachment $mail
                $mail.HTMLBody = $body
                if ($SendNow) {
                    if (-not $mail.Recipients.ResolveAll()) { throw "Destinatário não resolvido pelo Outlook: $($Row.To)" }
                    $mail.Save()
                    $mail.Send()
                }
                else {
                    [void]$mail.Recipients.ResolveAll()
                    $mail.Save()
                }
                return
            }
            catch {
                $Script:Outlook = $null
                if ($attempt -eq 2) { throw }
            }
        }
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }
}

function Start-OutlookSendReceive {
    if (-not $Script:Outlook) { $Script:Outlook = Get-OutlookApplication }
    try {
        foreach ($syncObject in $Script:Outlook.Session.SyncObjects) {
            try { $syncObject.Start() } catch { }
        }
        $Script:Outlook.Session.SendAndReceive($false)
        Log-Line 'Comando Enviar/Receber acionado no Outlook.'
    }
    catch {
        Log-Line "Não foi possível acionar Enviar/Receber: $($_.Exception.Message)"
    }
}

function Flush-OutlookOutbox {
    param([int]$MaxSeconds = 45)
    if (-not $Script:Outlook) { $Script:Outlook = Get-OutlookApplication }
    $deadline = (Get-Date).AddSeconds($MaxSeconds)
    $lastCount = -1

    while ((Get-Date) -lt $deadline) {
        $outbox = $Script:Outlook.Session.GetDefaultFolder(4)
        $count = [int]$outbox.Items.Count
        if ($count -eq 0) { return 0 }

        if ($count -ne $lastCount) {
            Log-Line "Aguardando envio: $count item(ns) na Caixa de Saída."
            $lastCount = $count
        }

        $items = @()
        for ($i = 1; $i -le $outbox.Items.Count; $i++) {
            $items += $outbox.Items.Item($i)
        }
        foreach ($item in $items) {
            try { $item.Send() } catch { Log-Line "Falha ao reenviar item da Caixa de Saída: $($_.Exception.Message)" }
        }
        Start-OutlookSendReceive
        Start-Sleep -Milliseconds 500
    }

    return [int]$Script:Outlook.Session.GetDefaultFolder(4).Items.Count
}

function Get-OutboxCount {
    if (-not $Script:Outlook) { $Script:Outlook = Get-OutlookApplication }
    try {
        $outbox = $Script:Outlook.Session.GetDefaultFolder(4)
        return [int]$outbox.Items.Count
    }
    catch {
        return -1
    }
}

function Get-SentCount {
    if (-not $Script:Outlook) { $Script:Outlook = Get-OutlookApplication }
    try {
        $sent = $Script:Outlook.Session.GetDefaultFolder(5)
        return [int]$sent.Items.Count
    }
    catch {
        return -1
    }
}

function Wait-SendInterval {
    param([decimal]$Minutes, [int]$CurrentNumber, [int]$Total)
    if ($Minutes -le 0) { return }
    $seconds = [int][Math]::Ceiling([double]$Minutes * 60)
    for ($remaining = $seconds; $remaining -gt 0; $remaining--) {
        if (($remaining -eq $seconds) -or ($remaining -le 10) -or ($remaining % 30 -eq 0)) {
            Log-Line "Aguardando $remaining segundo(s) para o próximo envio ($($CurrentNumber + 1)/$Total)."
        }
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 1
    }
}

function Log-Line {
    param([string]$Message)
    try {
        Add-Content -LiteralPath $Script:LogPath -Value ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' - ' + $Message) -Encoding UTF8
    }
    catch { }
    $txtLog.AppendText((Get-Date -Format 'HH:mm:ss') + ' - ' + $Message + [Environment]::NewLine)
}

function Test-RowReady {
    param($Row)
    if (-not $Row.To) { return $false }
    if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
        return ($Row.SummaryAttachmentCount -ge 1 -and $Row.DistributionAttachmentCount -ge 1 -and $Row.InformativeAttachmentCount -ge 1)
    }
    return $true
}

function Refresh-Preview {
    $grid.Rows.Clear()
    if (-not $cmbEmail.Text -or -not $cmbCity.Text) { return }
    $Script:Rows = Build-PreparedRows $cmbEmail.Text $cmbName.Text $cmbCity.Text $cmbUF.Text $cmbCC.Text
    foreach ($row in $Script:Rows) {
        $status = if (-not $row.To) {
            'Sem e-mail'
        }
        elseif ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
            $missing = @()
            if ($row.SummaryAttachmentCount -lt 1) { $missing += 'resumo' }
            if ($row.DistributionAttachmentCount -lt 1) { $missing += 'distribuição' }
            if ($row.InformativeAttachmentCount -lt 1) { $missing += 'informativo' }
            if ($missing.Count -gt 0) { 'Sem ' + ($missing -join ', ') } else { 'OK' }
        }
        elseif ($row.Attachments.Count -lt 1) {
            'Sem anexo'
        }
        else {
            'OK'
        }
        [void]$grid.Rows.Add($row.Line, $row.City, $row.UF, $row.Name, $row.To, (Merge-Addresses $row.CC $txtDefaultCC.Text), $row.Attachments.Count, $status)
    }
    $lblStatus.Text = "$($Script:Rows.Count) linha(s) preparada(s)."
}

$form = New-Object Windows.Forms.Form
$form.Text = 'Disparo de e-mails pelo Outlook'
$form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::None
$form.Size = New-Object Drawing.Size(1320, 840)
$form.MinimumSize = New-Object Drawing.Size(1180, 780)
$form.StartPosition = 'CenterScreen'

$title = New-Object Windows.Forms.Label
$title.Text = 'Disparo de e-mails pelo Outlook'
$title.Font = New-Object Drawing.Font('Segoe UI', 16, [Drawing.FontStyle]::Bold)
$title.Location = New-Object Drawing.Point(16, 14)
$title.Size = New-Object Drawing.Size(700, 30)
$form.Controls.Add($title)

$subtitle = New-Object Windows.Forms.Label
$subtitle.Text = 'Excel/CSV -> Outlook local, com anexos automáticos por Município-UF'
$subtitle.Location = New-Object Drawing.Point(18, 48)
$subtitle.Size = New-Object Drawing.Size(700, 22)
$form.Controls.Add($subtitle)

$grpOrigin = New-Object Windows.Forms.GroupBox
$grpOrigin.Text = '1. Origem e conta'
$grpOrigin.Location = New-Object Drawing.Point(16, 78)
$grpOrigin.Size = New-Object Drawing.Size(1268, 118)
$form.Controls.Add($grpOrigin)

$lblSheetFile = New-Object Windows.Forms.Label
$lblSheetFile.Text = 'Planilha:'
$lblSheetFile.Location = New-Object Drawing.Point(16, 30)
$lblSheetFile.Size = New-Object Drawing.Size(100, 22)
$grpOrigin.Controls.Add($lblSheetFile)

$txtWorkbook = New-Object Windows.Forms.TextBox
$txtWorkbook.Location = New-Object Drawing.Point(110, 28)
$txtWorkbook.Size = New-Object Drawing.Size(1020, 24)
$txtWorkbook.ReadOnly = $true
$grpOrigin.Controls.Add($txtWorkbook)

$btnWorkbook = New-Object Windows.Forms.Button
$btnWorkbook.Text = 'Selecionar...'
$btnWorkbook.Location = New-Object Drawing.Point(1140, 26)
$btnWorkbook.Size = New-Object Drawing.Size(105, 28)
$grpOrigin.Controls.Add($btnWorkbook)

$lblSheet = New-Object Windows.Forms.Label
$lblSheet.Text = 'Aba:'
$lblSheet.Location = New-Object Drawing.Point(16, 62)
$lblSheet.Size = New-Object Drawing.Size(100, 22)
$grpOrigin.Controls.Add($lblSheet)

$cmbSheet = New-Object Windows.Forms.ComboBox
$cmbSheet.Location = New-Object Drawing.Point(110, 60)
$cmbSheet.Size = New-Object Drawing.Size(210, 24)
$cmbSheet.DropDownStyle = 'DropDownList'
$grpOrigin.Controls.Add($cmbSheet)

$lblAccount = New-Object Windows.Forms.Label
$lblAccount.Text = 'Conta Outlook:'
$lblAccount.Location = New-Object Drawing.Point(340, 62)
$lblAccount.Size = New-Object Drawing.Size(100, 22)
$grpOrigin.Controls.Add($lblAccount)

$cmbAccount = New-Object Windows.Forms.ComboBox
$cmbAccount.Location = New-Object Drawing.Point(440, 60)
$cmbAccount.Size = New-Object Drawing.Size(690, 24)
$grpOrigin.Controls.Add($cmbAccount)

$btnAccounts = New-Object Windows.Forms.Button
$btnAccounts.Text = 'Atualizar contas'
$btnAccounts.Location = New-Object Drawing.Point(1140, 58)
$btnAccounts.Size = New-Object Drawing.Size(105, 28)
$grpOrigin.Controls.Add($btnAccounts)

$grpCols = New-Object Windows.Forms.GroupBox
$grpCols.Text = '2. Colunas da planilha'
$grpCols.Location = New-Object Drawing.Point(16, 204)
$grpCols.Size = New-Object Drawing.Size(1268, 86)
$form.Controls.Add($grpCols)

$labels = @('E-mail *', 'Cópia (CC)', 'Coordenador/Nome', 'Município/Cidade *', 'UF')
$combos = @()
for ($i = 0; $i -lt $labels.Count; $i++) {
    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = $labels[$i]
    $lbl.Location = New-Object Drawing.Point((16 + ($i * 245)), 30)
    $lbl.Size = New-Object Drawing.Size(220, 18)
    $grpCols.Controls.Add($lbl)
    $cmb = New-Object Windows.Forms.ComboBox
    $cmb.Location = New-Object Drawing.Point((16 + ($i * 245)), 50)
    $cmb.Size = New-Object Drawing.Size(230, 24)
    $grpCols.Controls.Add($cmb)
    $combos += $cmb
}
$cmbEmail, $cmbCC, $cmbName, $cmbCity, $cmbUF = $combos

$grpAttach = New-Object Windows.Forms.GroupBox
$grpAttach.Text = '3. Anexos (Opcionais)'
$grpAttach.Location = New-Object Drawing.Point(16, 298)
$grpAttach.Size = New-Object Drawing.Size(1268, 158)
$form.Controls.Add($grpAttach)

$lblAttachmentMode = New-Object Windows.Forms.Label
$lblAttachmentMode.Text = 'Módulo de anexos:'
$lblAttachmentMode.Location = New-Object Drawing.Point(16, 30)
$lblAttachmentMode.Size = New-Object Drawing.Size(130, 22)
$grpAttach.Controls.Add($lblAttachmentMode)

$cmbAttachmentMode = New-Object Windows.Forms.ComboBox
$cmbAttachmentMode.Location = New-Object Drawing.Point(150, 28)
$cmbAttachmentMode.Size = New-Object Drawing.Size(300, 24)
$cmbAttachmentMode.DropDownStyle = 'DropDownList'
[void]$cmbAttachmentMode.Items.Add('Resumo Município-UF')
[void]$cmbAttachmentMode.Items.Add('Distribuição por subpastas')
[void]$cmbAttachmentMode.Items.Add('Resumo + Distribuição + Informativo')
$cmbAttachmentMode.SelectedItem = $Script:AttachmentMode
$grpAttach.Controls.Add($cmbAttachmentMode)

$lblAttachmentFolder = New-Object Windows.Forms.Label
$lblAttachmentFolder.Text = 'Pasta dos arquivos:'
$lblAttachmentFolder.Location = New-Object Drawing.Point(16, 62)
$lblAttachmentFolder.Size = New-Object Drawing.Size(130, 22)
$grpAttach.Controls.Add($lblAttachmentFolder)

$txtAttachmentFolder = New-Object Windows.Forms.TextBox
$txtAttachmentFolder.Location = New-Object Drawing.Point(150, 60)
$txtAttachmentFolder.Size = New-Object Drawing.Size(860, 24)
$txtAttachmentFolder.ReadOnly = $true
$grpAttach.Controls.Add($txtAttachmentFolder)

$btnAttachmentFolder = New-Object Windows.Forms.Button
$btnAttachmentFolder.Text = 'Selecionar...'
$btnAttachmentFolder.Location = New-Object Drawing.Point(1020, 57)
$btnAttachmentFolder.Size = New-Object Drawing.Size(112, 28)
$grpAttach.Controls.Add($btnAttachmentFolder)

$btnClearFolder = New-Object Windows.Forms.Button
$btnClearFolder.Text = 'Limpar'
$btnClearFolder.Location = New-Object Drawing.Point(1140, 57)
$btnClearFolder.Size = New-Object Drawing.Size(112, 28)
$grpAttach.Controls.Add($btnClearFolder)

$lblInformative1 = New-Object Windows.Forms.Label
$lblInformative1.Text = 'Informativo 1:'
$lblInformative1.Location = New-Object Drawing.Point(16, 94)
$lblInformative1.Size = New-Object Drawing.Size(130, 22)
$grpAttach.Controls.Add($lblInformative1)

$txtInformative1 = New-Object Windows.Forms.TextBox
$txtInformative1.Location = New-Object Drawing.Point(150, 92)
$txtInformative1.Size = New-Object Drawing.Size(860, 24)
$txtInformative1.ReadOnly = $true
$grpAttach.Controls.Add($txtInformative1)

$btnInformative1 = New-Object Windows.Forms.Button
$btnInformative1.Text = 'Selecionar...'
$btnInformative1.Location = New-Object Drawing.Point(1020, 89)
$btnInformative1.Size = New-Object Drawing.Size(112, 28)
$grpAttach.Controls.Add($btnInformative1)

$btnClearInformative1 = New-Object Windows.Forms.Button
$btnClearInformative1.Text = 'Limpar'
$btnClearInformative1.Location = New-Object Drawing.Point(1140, 89)
$btnClearInformative1.Size = New-Object Drawing.Size(112, 28)
$grpAttach.Controls.Add($btnClearInformative1)

$lblInformative2 = New-Object Windows.Forms.Label
$lblInformative2.Text = 'Informativo 2:'
$lblInformative2.Location = New-Object Drawing.Point(16, 126)
$lblInformative2.Size = New-Object Drawing.Size(130, 22)
$grpAttach.Controls.Add($lblInformative2)

$txtInformative2 = New-Object Windows.Forms.TextBox
$txtInformative2.Location = New-Object Drawing.Point(150, 124)
$txtInformative2.Size = New-Object Drawing.Size(860, 24)
$txtInformative2.ReadOnly = $true
$grpAttach.Controls.Add($txtInformative2)

$btnInformative2 = New-Object Windows.Forms.Button
$btnInformative2.Text = 'Selecionar...'
$btnInformative2.Location = New-Object Drawing.Point(1020, 121)
$btnInformative2.Size = New-Object Drawing.Size(112, 28)
$grpAttach.Controls.Add($btnInformative2)

$btnClearInformative2 = New-Object Windows.Forms.Button
$btnClearInformative2.Text = 'Limpar'
$btnClearInformative2.Location = New-Object Drawing.Point(1140, 121)
$btnClearInformative2.Size = New-Object Drawing.Size(112, 28)
$grpAttach.Controls.Add($btnClearInformative2)

$defaultInformative1 = 'C:\Users\jonatas.chaves\Downloads\ENEM - 26 -  INFORMATIVO VALORES COORDENADOR DE APLICAÇAO E ASSISTENTE.pdf'
if (Test-Path -LiteralPath $defaultInformative1 -PathType Leaf) {
    $Script:InformativeFile1 = $defaultInformative1
    $txtInformative1.Text = $Script:InformativeFile1
}

function Update-AttachmentModeUi {
    if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
        $grpAttach.Text = '3. Anexos por município'
        $lblAttachmentFolder.Text = 'Resumo:'
        $lblInformative1.Text = 'Distribuição:'
        $lblInformative2.Text = 'Informativo:'
        $txtInformative1.Text = $Script:DistributionFolder
        $txtInformative2.Text = $Script:InformativeFolder
        $btnInformative1.Text = 'Selecionar...'
        $btnInformative2.Text = 'Selecionar...'
        return
    }

    $grpAttach.Text = '3. Anexos (Opcionais)'
    $lblAttachmentFolder.Text = 'Pasta dos arquivos:'
    $lblInformative1.Text = 'Informativo 1:'
    $lblInformative2.Text = 'Informativo 2:'
    $txtInformative1.Text = $Script:InformativeFile1
    $txtInformative2.Text = $Script:InformativeFile2
    $btnInformative1.Text = 'Selecionar...'
    $btnInformative2.Text = 'Selecionar...'
}

$grpPersonal = New-Object Windows.Forms.GroupBox
$grpPersonal.Text = '4. Personalização'
$grpPersonal.Location = New-Object Drawing.Point(16, 464)
$grpPersonal.Size = New-Object Drawing.Size(1268, 218)
$form.Controls.Add($grpPersonal)

$lblSubject = New-Object Windows.Forms.Label
$lblSubject.Text = 'Modelo do assunto:'
$lblSubject.Location = New-Object Drawing.Point(16, 30)
$lblSubject.Size = New-Object Drawing.Size(130, 22)
$grpPersonal.Controls.Add($lblSubject)

$txtSubject = New-Object Windows.Forms.TextBox
$txtSubject.Location = New-Object Drawing.Point(150, 28)
$txtSubject.Size = New-Object Drawing.Size(1100, 24)
$txtSubject.Text = $DefaultSubject
$grpPersonal.Controls.Add($txtSubject)

$lblHtml = New-Object Windows.Forms.Label
$lblHtml.Text = 'Editor do e-mail:'
$lblHtml.Location = New-Object Drawing.Point(16, 60)
$lblHtml.Size = New-Object Drawing.Size(130, 22)
$grpPersonal.Controls.Add($lblHtml)

$editorToolbar = New-Object Windows.Forms.Panel
$editorToolbar.Location = New-Object Drawing.Point(150, 56)
$editorToolbar.Size = New-Object Drawing.Size(1100, 32)
$grpPersonal.Controls.Add($editorToolbar)

$Script:EditorButtonsByCommand = @{}
$editorButtons = @(
    @('B', 'Bold', 'Negrito'),
    @('I', 'Italic', 'Itálico'),
    @('U', 'Underline', 'Sublinhado'),
    @('• Lista', 'InsertUnorderedList', 'Lista com marcadores'),
    @('1. Lista', 'InsertOrderedList', 'Lista numerada'),
    @('←', 'Outdent', 'Diminuir recuo'),
    @('→', 'Indent', 'Aumentar recuo'),
    @('Esq.', 'JustifyLeft', 'Alinhar à esquerda'),
    @('Centro', 'JustifyCenter', 'Centralizar'),
    @('Dir.', 'JustifyRight', 'Alinhar à direita'),
    @('Link', 'CreateLink', 'Inserir link'),
    @('Limpar', 'RemoveFormat', 'Limpar formatação'),
    @('Desfazer', 'Undo', 'Desfazer'),
    @('Refazer', 'Redo', 'Refazer'),
    @('HTML', 'EditHtml', 'Editar código HTML')
)
$buttonX = 0
foreach ($spec in $editorButtons) {
    $btnEditor = New-Object Windows.Forms.Button
    $btnEditor.Text = $spec[0]
    $btnEditor.Tag = $spec[1]
    $btnEditor.Width = if ($spec[0].Length -gt 6) { 74 } elseif ($spec[0].Length -gt 3) { 62 } else { 42 }
    $btnEditor.Height = 26
    $btnEditor.Location = New-Object Drawing.Point($buttonX, 2)
    $btnEditor.FlatStyle = 'System'
    $btnEditor.UseVisualStyleBackColor = $true
    $toolTip = New-Object Windows.Forms.ToolTip
    $toolTip.SetToolTip($btnEditor, $spec[2])
    $commandName = [string]$spec[1]
    if ($commandName -notin @('CreateLink', 'EditHtml', 'RemoveFormat', 'Undo', 'Redo', 'Indent', 'Outdent')) {
        $Script:EditorButtonsByCommand[$commandName] = $btnEditor
    }
    $btnEditor.Add_Click({
            switch ($commandName) {
                'CreateLink' {
                    $url = [Microsoft.VisualBasic.Interaction]::InputBox('Informe o link completo:', 'Inserir link', 'https://')
                    if (-not [string]::IsNullOrWhiteSpace($url) -and $url -ne 'https://') {
                        Invoke-EditorCommand 'CreateLink' $url
                    }
                }
                'EditHtml' {
                    Open-HtmlEditorDialog
                }
                default {
                    Invoke-EditorCommand $commandName
                }
            }
        }.GetNewClosure())
    $editorToolbar.Controls.Add($btnEditor)
    $buttonX += ($btnEditor.Width + 4)
}

$webEditor = New-Object Windows.Forms.WebBrowser
$webEditor.Location = New-Object Drawing.Point(150, 90)
$webEditor.Size = New-Object Drawing.Size(1100, 62)
$webEditor.ScriptErrorsSuppressed = $true
$webEditor.AllowWebBrowserDrop = $false
$webEditor.Add_DocumentCompleted({
        Register-EditorEvents
        Update-EditorToolbarState
    })
$grpPersonal.Controls.Add($webEditor)

$txtHtml = New-Object Windows.Forms.TextBox
$txtHtml.Location = New-Object Drawing.Point(150, 90)
$txtHtml.Size = New-Object Drawing.Size(1, 1)
$txtHtml.Multiline = $true
$txtHtml.ScrollBars = 'Vertical'
$txtHtml.Font = New-Object Drawing.Font('Consolas', 10)
$txtHtml.Text = $DefaultHtml
$txtHtml.Visible = $false
$grpPersonal.Controls.Add($txtHtml)
Set-EditorHtml $DefaultHtml

$lblDefaultCC = New-Object Windows.Forms.Label
$lblDefaultCC.Text = 'CC para todos (;):'
$lblDefaultCC.Location = New-Object Drawing.Point(16, 160)
$lblDefaultCC.Size = New-Object Drawing.Size(130, 22)
$grpPersonal.Controls.Add($lblDefaultCC)

$txtDefaultCC = New-Object Windows.Forms.TextBox
$txtDefaultCC.Location = New-Object Drawing.Point(150, 158)
$txtDefaultCC.Size = New-Object Drawing.Size(1100, 24)
$txtDefaultCC.Text = ''
$grpPersonal.Controls.Add($txtDefaultCC)

$chkSignature = New-Object Windows.Forms.CheckBox
$chkSignature.Text = 'Assinatura da sessão do Outlook obrigatória'
$chkSignature.Location = New-Object Drawing.Point(150, 188)
$chkSignature.Size = New-Object Drawing.Size(360, 22)
$chkSignature.Checked = $true
$chkSignature.Enabled = $false
$grpPersonal.Controls.Add($chkSignature)

$hint = New-Object Windows.Forms.Label
$hint.Text = 'Use {{Nome}}, {{Cidade}}, {{UF}} ou qualquer {{Coluna}} da planilha.'
$hint.TextAlign = 'MiddleRight'
$hint.Location = New-Object Drawing.Point(850, 188)
$hint.Size = New-Object Drawing.Size(400, 22)
$grpPersonal.Controls.Add($hint)

$grpPreview = New-Object Windows.Forms.GroupBox
$grpPreview.Text = '5. Pré-visualização'
$grpPreview.Location = New-Object Drawing.Point(16, 690)
$grpPreview.Size = New-Object Drawing.Size(1268, 94)
$form.Controls.Add($grpPreview)

$grid = New-Object Windows.Forms.DataGridView
$grid.Location = New-Object Drawing.Point(10, 22)
$grid.Size = New-Object Drawing.Size(1246, 62)
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'FullRowSelect'
$grid.Columns.Add('Linha', 'Linha') | Out-Null
$grid.Columns.Add('Cidade', 'Município') | Out-Null
$grid.Columns.Add('UF', 'UF') | Out-Null
$grid.Columns.Add('Nome', 'Coordenador') | Out-Null
$grid.Columns.Add('Para', 'Para') | Out-Null
$grid.Columns.Add('CC', 'CC') | Out-Null
$grid.Columns.Add('Anexos', 'Anexos') | Out-Null
$grid.Columns.Add('Status', 'Status') | Out-Null
$grpPreview.Controls.Add($grid)

$grpLog = New-Object Windows.Forms.GroupBox
$grpLog.Text = 'Log'
$grpLog.Location = New-Object Drawing.Point(16, 790)
$grpLog.Size = New-Object Drawing.Size(900, 46)
$form.Controls.Add($grpLog)

$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Location = New-Object Drawing.Point(10, 20)
$txtLog.Size = New-Object Drawing.Size(878, 20)
$txtLog.ReadOnly = $true
$txtLog.Multiline = $true
$grpLog.Controls.Add($txtLog)

$lblStatus = New-Object Windows.Forms.Label
$lblStatus.Text = 'Selecione uma planilha para começar.'
$lblStatus.Location = New-Object Drawing.Point(930, 790)
$lblStatus.Size = New-Object Drawing.Size(340, 22)
$form.Controls.Add($lblStatus)

$lblInterval = New-Object Windows.Forms.Label
$lblInterval.Text = 'Intervalo entre envios (min):'
$lblInterval.Location = New-Object Drawing.Point(520, 188)
$lblInterval.Size = New-Object Drawing.Size(205, 22)
$grpPersonal.Controls.Add($lblInterval)

$nudIntervalMinutes = New-Object Windows.Forms.NumericUpDown
$nudIntervalMinutes.Location = New-Object Drawing.Point(730, 185)
$nudIntervalMinutes.Size = New-Object Drawing.Size(96, 28)
$nudIntervalMinutes.Minimum = 0
$nudIntervalMinutes.Maximum = 1440
$nudIntervalMinutes.DecimalPlaces = 1
$nudIntervalMinutes.Increment = 0.5
$nudIntervalMinutes.Value = 0
$grpPersonal.Controls.Add($nudIntervalMinutes)

$btnDrafts = New-Object Windows.Forms.Button
$btnDrafts.Text = 'Criar rascunhos'
$btnDrafts.Location = New-Object Drawing.Point(930, 44)
$btnDrafts.Size = New-Object Drawing.Size(150, 32)
$btnDrafts.Enabled = $false
$form.Controls.Add($btnDrafts)

$btnSend = New-Object Windows.Forms.Button
$btnSend.Text = 'Disparar e-mails'
$btnSend.Location = New-Object Drawing.Point(1092, 44)
$btnSend.Size = New-Object Drawing.Size(178, 32)
$btnSend.Enabled = $false
$form.Controls.Add($btnSend)

$refreshColumns = {
    foreach ($cmb in @($cmbEmail, $cmbCC, $cmbName, $cmbCity, $cmbUF)) {
        $cmb.Items.Clear()
        [void]$cmb.Items.Add('')
        foreach ($col in $Script:Columns) { [void]$cmb.Items.Add($col) }
    }
    $cmbEmail.Text = Pick-Column $Script:Columns @('Email', 'E-mail', 'E mail')
    $cmbName.Text = Pick-Column $Script:Columns @('Nome', 'Coordenador', 'Coordenador/Nome')
    $cmbCity.Text = Pick-Column $Script:Columns @('Cidade', 'Município', 'Municipio')
    $cmbUF.Text = Pick-Column $Script:Columns @('UF', 'Estado')
    $cmbCC.Text = Pick-Column $Script:Columns @('CC', 'Cópia', 'Copia')
    Refresh-Preview
}

$loadData = {
    if (-not $Script:WorkbookPath) { return }
    try {
        if ([IO.Path]::GetExtension($Script:WorkbookPath).ToLowerInvariant() -eq '.csv') {
            $data = Read-CsvRows $Script:WorkbookPath
        }
        else {
            $data = Read-WorkbookRows $Script:WorkbookPath $cmbSheet.Text
        }
        $Script:Columns = @($data.Headers)
        $Script:RawRows = $data.Rows
        & $refreshColumns
        $btnDrafts.Enabled = $true
        $btnSend.Enabled = $true
        Log-Line "Planilha carregada: $($Script:RawRows.Count) linha(s)."
    }
    catch {
        [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Erro ao carregar planilha', 'OK', 'Error') | Out-Null
        Log-Line "Erro ao carregar planilha: $($_.Exception.Message)"
    }
}

$btnWorkbook.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Planilhas (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv'
        if ($dialog.ShowDialog() -eq 'OK') {
            $Script:WorkbookPath = $dialog.FileName
            $txtWorkbook.Text = $Script:WorkbookPath
            $cmbSheet.Items.Clear()
            if ([IO.Path]::GetExtension($Script:WorkbookPath).ToLowerInvariant() -eq '.csv') {
                [void]$cmbSheet.Items.Add('CSV')
                $cmbSheet.SelectedIndex = 0
            }
            else {
                $Script:Sheets = Get-SheetNames $Script:WorkbookPath
                foreach ($s in $Script:Sheets) { [void]$cmbSheet.Items.Add($s) }
                if ($cmbSheet.Items.Count -gt 0) { $cmbSheet.SelectedIndex = 0 }
            }
            & $loadData
        }
    })

$cmbSheet.Add_SelectedIndexChanged({ & $loadData })

$btnAttachmentFolder.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        if ($dialog.ShowDialog() -eq 'OK') {
            $Script:AttachmentFolder = $dialog.SelectedPath
            $txtAttachmentFolder.Text = $Script:AttachmentFolder
            Refresh-Preview
            Log-Line "Pasta de anexos selecionada ($($Script:AttachmentMode))."
        }
    })

$cmbAttachmentMode.Add_SelectedIndexChanged({
        $Script:AttachmentMode = [string]$cmbAttachmentMode.SelectedItem
        if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
            $defaultSummaryFolder = 'C:\Users\jonatas.chaves\Downloads\RESUMOS.CE\RESUMOS.CE'
            $defaultDistributionFolder = 'C:\Users\jonatas.chaves\Downloads\CE DISTRIBUIÇÃO\CE'
            $defaultInformativeFolder = 'C:\Users\jonatas.chaves\Downloads\Informativo'
            if ([string]::IsNullOrWhiteSpace($Script:AttachmentFolder) -and (Test-Path -LiteralPath $defaultSummaryFolder -PathType Container)) {
                $Script:AttachmentFolder = $defaultSummaryFolder
                $txtAttachmentFolder.Text = $Script:AttachmentFolder
            }
            if ([string]::IsNullOrWhiteSpace($Script:DistributionFolder) -and (Test-Path -LiteralPath $defaultDistributionFolder -PathType Container)) {
                $Script:DistributionFolder = $defaultDistributionFolder
            }
            if ([string]::IsNullOrWhiteSpace($Script:InformativeFolder) -and (Test-Path -LiteralPath $defaultInformativeFolder -PathType Container)) {
                $Script:InformativeFolder = $defaultInformativeFolder
            }
        }
        if ($Script:AttachmentMode -eq 'Distribuição por subpastas' -and [string]::IsNullOrWhiteSpace($Script:AttachmentFolder)) {
            $defaultDistributionFolder = 'C:\Users\jonatas.chaves\Downloads\RelatoriosDetalhado\AM'
            if (Test-Path -LiteralPath $defaultDistributionFolder -PathType Container) {
                $Script:AttachmentFolder = $defaultDistributionFolder
                $txtAttachmentFolder.Text = $Script:AttachmentFolder
            }
        }
        Update-AttachmentModeUi
        Refresh-Preview
        Log-Line "Módulo de anexos selecionado: $($Script:AttachmentMode)."
    })

$btnClearFolder.Add_Click({
        $Script:AttachmentFolder = ''
        $txtAttachmentFolder.Text = ''
        Refresh-Preview
        Log-Line "Pasta de anexos removida."
    })

$btnInformative1.Add_Click({
        if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
            $dialog = New-Object Windows.Forms.FolderBrowserDialog
            $dialog.Description = 'Selecionar pasta de distribuição'
            if ($Script:DistributionFolder -and (Test-Path -LiteralPath $Script:DistributionFolder -PathType Container)) {
                $dialog.SelectedPath = $Script:DistributionFolder
            }
            if ($dialog.ShowDialog() -eq 'OK') {
                $Script:DistributionFolder = $dialog.SelectedPath
                $txtInformative1.Text = $Script:DistributionFolder
                Refresh-Preview
                Log-Line "Pasta de distribuição selecionada."
            }
            return
        }
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Arquivos PDF (*.pdf)|*.pdf|Todos os arquivos (*.*)|*.*'
        $dialog.Title = 'Selecionar informativo 1'
        if ($Script:InformativeFile1) { $dialog.FileName = $Script:InformativeFile1 }
        if ($dialog.ShowDialog() -eq 'OK') {
            $Script:InformativeFile1 = $dialog.FileName
            $txtInformative1.Text = $Script:InformativeFile1
            Refresh-Preview
            Log-Line "Informativo 1 selecionado."
        }
    })

$btnClearInformative1.Add_Click({
        if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
            $Script:DistributionFolder = ''
            $txtInformative1.Text = ''
            Refresh-Preview
            Log-Line "Pasta de distribuição removida."
            return
        }
        $Script:InformativeFile1 = ''
        $txtInformative1.Text = ''
        Refresh-Preview
        Log-Line "Informativo 1 removido."
    })

$btnInformative2.Add_Click({
        if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
            $dialog = New-Object Windows.Forms.FolderBrowserDialog
            $dialog.Description = 'Selecionar pasta de informativos'
            if ($Script:InformativeFolder -and (Test-Path -LiteralPath $Script:InformativeFolder -PathType Container)) {
                $dialog.SelectedPath = $Script:InformativeFolder
            }
            if ($dialog.ShowDialog() -eq 'OK') {
                $Script:InformativeFolder = $dialog.SelectedPath
                $txtInformative2.Text = $Script:InformativeFolder
                Refresh-Preview
                Log-Line "Pasta de informativos selecionada."
            }
            return
        }
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Arquivos PDF (*.pdf)|*.pdf|Todos os arquivos (*.*)|*.*'
        $dialog.Title = 'Selecionar informativo 2'
        if ($Script:InformativeFile2) { $dialog.FileName = $Script:InformativeFile2 }
        if ($dialog.ShowDialog() -eq 'OK') {
            $Script:InformativeFile2 = $dialog.FileName
            $txtInformative2.Text = $Script:InformativeFile2
            Refresh-Preview
            Log-Line "Informativo 2 selecionado."
        }
    })

$btnClearInformative2.Add_Click({
        if ($Script:AttachmentMode -eq 'Resumo + Distribuição + Informativo') {
            $Script:InformativeFolder = ''
            $txtInformative2.Text = ''
            Refresh-Preview
            Log-Line "Pasta de informativos removida."
            return
        }
        $Script:InformativeFile2 = ''
        $txtInformative2.Text = ''
        Refresh-Preview
        Log-Line "Informativo 2 removido."
    })

foreach ($cmb in @($cmbEmail, $cmbCC, $cmbName, $cmbCity, $cmbUF)) {
    $cmb.Add_SelectedIndexChanged({ Refresh-Preview })
}

$btnAccounts.Add_Click({
        try {
            $Script:Outlook = $null
            $cmbAccount.Items.Clear()
            $accounts = Get-OutlookAccounts
            foreach ($account in $accounts) { [void]$cmbAccount.Items.Add($account) }
            if ($cmbAccount.Items.Count -gt 0) { $cmbAccount.SelectedIndex = 0 }
            Log-Line "Outlook conectado: $($accounts.Count) conta(s) encontrada(s)."
        }
        catch {
            Log-Line "Não foi possível conectar ao Outlook: $($_.Exception.Message)"
        }
    })

$btnDrafts.Add_Click({
        Refresh-Preview
        $valid = @($Script:Rows | Where-Object { Test-RowReady $_ })
        $bad = @($Script:Rows | Where-Object { -not (Test-RowReady $_) })
        if ($valid.Count -eq 0) {
            [Windows.Forms.MessageBox]::Show("Nenhum rascunho pode ser criado. Confira e-mail e anexos obrigatórios na pré-visualização.", 'Rascunhos não criados', 'OK', 'Warning') | Out-Null
            Log-Line 'Criação de rascunhos bloqueada: nenhuma linha pronta.'
            return
        }
        if ($bad.Count -gt 0) {
            $answer = [Windows.Forms.MessageBox]::Show("Há $($bad.Count) linha(s) sem e-mail ou com anexo obrigatório faltando. Criar rascunhos apenas das linhas prontas?", 'Confirmar rascunhos', 'YesNo', 'Warning')
            if ($answer -ne 'Yes') { return }
        }
        $count = 0
        $failures = 0
        foreach ($row in $valid) {
            try {
                Compose-Mail $row
                $count++
                Log-Line "Rascunho criado: linha $($row.Line), $($row.To), $($row.Attachments.Count) anexo(s)."
            }
            catch {
                $failures++
                Log-Line "Falha ao criar rascunho na linha $($row.Line), $($row.To): $($_.Exception.Message)"
            }
        }
        Log-Line "$count rascunho(s) criado(s), $failures falha(s)."
    })

$btnSend.Add_Click({
        Refresh-Preview
        $valid = @($Script:Rows | Where-Object { Test-RowReady $_ })
        $bad = @($Script:Rows | Where-Object { -not (Test-RowReady $_) })
        if ($valid.Count -eq 0) {
            [Windows.Forms.MessageBox]::Show("Nenhum e-mail está pronto para disparo. Confira e-mail e anexos obrigatórios na pré-visualização.", 'Disparo não iniciado', 'OK', 'Warning') | Out-Null
            Log-Line 'Disparo bloqueado: nenhuma linha pronta.'
            return
        }
        $answer = [Windows.Forms.MessageBox]::Show("Enviar $($valid.Count) e-mail(s) agora? Linhas sem e-mail ou com anexo obrigatório faltando serão ignoradas: $($bad.Count).", 'Confirmar envio', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return }
        $outboxBefore = Get-OutboxCount
        $sentBefore = Get-SentCount
        if ($outboxBefore -ge 0) { Log-Line "Caixa de Saída antes do disparo: $outboxBefore item(ns)." }
        if ($sentBefore -ge 0) { Log-Line "Itens Enviados antes do disparo: $sentBefore item(ns)." }
        $count = 0
        $failures = 0
        $intervalMinutes = [decimal]$nudIntervalMinutes.Value
        for ($rowIndex = 0; $rowIndex -lt $valid.Count; $rowIndex++) {
            $row = $valid[$rowIndex]
            try {
                Compose-Mail $row -SendNow
                $count++
                Log-Line "Entregue ao Outlook: linha $($row.Line), $($row.To), $($row.Attachments.Count) anexo(s)."
            }
            catch {
                $failures++
                Log-Line "Falha na linha $($row.Line), $($row.To): $($_.Exception.Message)"
            }
            if ($rowIndex -lt ($valid.Count - 1)) {
                Wait-SendInterval $intervalMinutes ($rowIndex + 1) $valid.Count
            }
        }
        Start-OutlookSendReceive
        $outboxCount = Flush-OutlookOutbox 45
        $sentAfter = Get-SentCount
        $sentDelta = if ($sentBefore -ge 0 -and $sentAfter -ge 0) { $sentAfter - $sentBefore } else { -1 }
        if ($outboxCount -ge 0) {
            if ($sentDelta -ge 0) {
                Log-Line "Resultado: $count e-mail(s) entregue(s) ao Outlook, $failures falha(s), $outboxCount item(ns) na Caixa de Saída, $sentDelta novo(s) item(ns) em Itens Enviados."
                [Windows.Forms.MessageBox]::Show("$count e-mail(s) foram entregues ao Outlook para envio.`nFalhas: $failures`nItens na Caixa de Saída: $outboxCount`nNovos itens em Itens Enviados: $sentDelta`n`nSe algum item permanecer na Caixa de Saída, abra o Outlook e clique em Enviar/Receber.", 'Disparo concluído', 'OK', 'Information') | Out-Null
            }
            else {
                Log-Line "Resultado: $count e-mail(s) entregue(s) ao Outlook, $failures falha(s), $outboxCount item(ns) na Caixa de Saída."
                [Windows.Forms.MessageBox]::Show("$count e-mail(s) foram entregues ao Outlook para envio.`nFalhas: $failures`nItens na Caixa de Saída: $outboxCount`n`nSe algum item permanecer na Caixa de Saída, abra o Outlook e clique em Enviar/Receber.", 'Disparo concluído', 'OK', 'Information') | Out-Null
            }
        }
        else {
            Log-Line "Resultado: $count e-mail(s) entregue(s) ao Outlook, $failures falha(s)."
            [Windows.Forms.MessageBox]::Show("$count e-mail(s) foram entregues ao Outlook para envio.`nFalhas: $failures", 'Disparo concluído', 'OK', 'Information') | Out-Null
        }
    })

try {
    $accounts = Get-OutlookAccounts
    foreach ($account in $accounts) { [void]$cmbAccount.Items.Add($account) }
    if ($cmbAccount.Items.Count -gt 0) { $cmbAccount.SelectedIndex = 0 }
    Log-Line "Outlook conectado: $($accounts.Count) conta(s) encontrada(s)."
}
catch {
    Log-Line "Outlook ainda não conectado. Abra o Outlook e clique em Atualizar contas."
}

[void]$form.ShowDialog()
