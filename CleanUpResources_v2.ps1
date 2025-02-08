#requires -Module Az.Resources
#requires -Module Az.Accounts

[CmdletBinding(SupportsShouldProcess = $true)]
Param(
    [Parameter()]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $true)]
    [string]$To,
    
    [Parameter(Mandatory = $false,
               HelpMessage = "Cantidad de días que un RG está vencido para que sea incluido en el correo de advertencia")]
    [ValidateRange(1, 180)] 
    [int32]$PastMaxExpiryDays = 1,
    
    [Parameter(Mandatory = $false,
               HelpMessage = "Cantidad máxima de días futuros para que un RG sea incluido en el correo de advertencia")]
    [ValidateRange(180,365)] 
    [int32]$FutureMaxExpiryDays = 180
)

# ----------------------------------------------------------------------------------
# Función: Conexión a Azure
# ----------------------------------------------------------------------------------
function Connect-ToAzure {
    Param(
        [string]$ConnectionName,
        [string]$SubscriptionId
    )

    try {
        Write-Verbose "Obteniendo información de la conexión $ConnectionName..."
        $servicePrincipalConnection = Get-AutomationConnection -Name $ConnectionName
        if (-not $servicePrincipalConnection) {
            throw "No se pudo encontrar la conexión '$ConnectionName'."
        }

        Write-Verbose "Conectando a Azure con Az..."
        Connect-AzAccount `
            -ServicePrincipal `
            -Tenant $servicePrincipalConnection.TenantId `
            -ApplicationId $servicePrincipalConnection.ApplicationId `
            -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint | Out-Null

        Write-Verbose "Estableciendo contexto al SubscriptionId $SubscriptionId..."
        Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
    }
    catch {
        Write-Error $_.Exception.Message
        throw $_.Exception
    }
}

# ----------------------------------------------------------------------------------
# Función: Obtener y filtrar Resource Groups con la etiqueta deleteAfter
# ----------------------------------------------------------------------------------
function Get-ResourceGroupsByDeleteAfter {
    Param(
        [int]$PastDays,
        [int]$FutureDays,
        [string]$IgnoreRegex
    )
    
    $allRGs = Get-AzResourceGroup
    Write-Verbose "Total RGs encontrados: $($allRGs.Count)"

    # Filtrar RGs que tengan la etiqueta 'deleteAfter'
    $deleteTagged = $allRGs | Where-Object { $_.Tags.ContainsKey('deleteAfter') }
    Write-Verbose "RGs con etiqueta deleteAfter: $($deleteTagged.Count)"

    # Filtrar RGs que no la tengan (por info, si se necesita)
    $notDeleteTagged = $allRGs | Where-Object { -not $_.Tags.ContainsKey('deleteAfter') }
    Write-Verbose "RGs sin etiqueta deleteAfter: $($notDeleteTagged.Count)"

    # Proyección a un objeto que sea más fácil de manejar
    # Buscamos 'owner' o 'resourceowner' y guardamos esa info en OwnerEmail.
    $deleteTaggedCasted = $deleteTagged | Select-Object `
        @{ Name = "DeleteAfter";   Expression = {[datetime]$_.Tags['deleteAfter']} },
        @{ Name = "OwnerEmail";    Expression = {
            # Primero probamos con 'owner'
            $tags = $_.Tags
            if ($tags['owner']) {
                $tags['owner']
            }
            elseif ($tags['resourceowner']) {
                $tags['resourceowner']
            }
        }},
        @{ Name = "ResourceGroupName"; Expression = { $_.ResourceGroupName } },
        @{ Name = "ResourceCount"; Expression = { 0 } },
        @{ Name = "Resources";     Expression = { @() } }

    # RGs expirados (pasados de la fecha) -> PastDays debe venir como número negativo
    $expired = $deleteTaggedCasted |
        Where-Object { $_.DeleteAfter -lt (Get-Date).AddDays($PastDays) } |
        Where-Object { $_.ResourceGroupName -notmatch $IgnoreRegex } |
        Sort-Object -Property OwnerEmail

    # RGs con fecha muy lejana
    $tooFarOutExpiry = $deleteTaggedCasted |
        Where-Object { $_.DeleteAfter -gt (Get-Date).AddDays($FutureDays) } |
        Sort-Object -Property DeleteAfter

    return [PSCustomObject]@{
        Expired         = $expired
        TooFarOutExpiry = $tooFarOutExpiry
    }
}

# ----------------------------------------------------------------------------------
# Función: Obtener recursos e incrementar el ResourceCount en cada objeto
# ----------------------------------------------------------------------------------
function Populate-ResourceInfo {
    Param(
        [System.Object[]]$ResourceGroups
    )

    foreach ($rg in $ResourceGroups) {
        Write-Verbose "Obteniendo recursos del grupo: $($rg.ResourceGroupName)"
        try {
            $resources = Get-AzResource -ResourceGroupName $rg.ResourceGroupName
            $rg.Resources = $resources
            $rg.ResourceCount = $resources.Count
        }
        catch {
            Write-Warning "No se pudo obtener recursos para el RG '$($rg.ResourceGroupName)': $($_.Exception.Message)"
        }
    }
}

# ----------------------------------------------------------------------------------
# Función: Construye el contenido HTML del cuerpo basado en una tabla
# ----------------------------------------------------------------------------------
function Build-ResourceGroupTableHtml {
    Param(
        [System.Object[]]$ResourceGroups
    )

    if (-not $ResourceGroups -or $ResourceGroups.Count -eq 0) {
        return "<p>No hay Resource Groups para mostrar.</p>"
    }

    $tableRows = $ResourceGroups |
        ForEach-Object {
            "<tr><td>$($_.ResourceGroupName)</td><td>$($_.OwnerEmail)</td><td>$($_.DeleteAfter)</td><td>$($_.ResourceCount)</td></tr>"
        }

    $html = @"
<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">
    <tr style="font-weight:bold;">
        <th>Resource Group</th>
        <th>Owner</th>
        <th>Fecha Expiración</th>
        <th>Recursos</th>
    </tr>
    $($tableRows -join "`r`n")
</table>
"@
    return $html
}

# ----------------------------------------------------------------------------------
# Función: Envía correo con formato HTML y adjuntos inline
# ----------------------------------------------------------------------------------
function Send-CleanupMail {
    Param(
        [Parameter(Mandatory = $true)] [string]$Subject,
        [Parameter(Mandatory = $true)] [string]$TemplateUrl,
        [Parameter(Mandatory = $true)] [string]$TemplateHeaderGraphicUrl,
        [Parameter(Mandatory = $true)] [string]$TableHtml,
        [Parameter(Mandatory = $true)] [PSCredential]$MailCredentials,
        [Parameter(Mandatory = $true)] [string]$MailServer,
        [Parameter(Mandatory = $true)] [string[]]$To,
        [Parameter()] [string]$DeleteDate = (Get-Date -Format "yyyy-MM-dd")
    )

    # Descargar la plantilla HTML
    Write-Verbose "Descargando plantilla HTML de $TemplateUrl"
    $templateResponse = Invoke-WebRequest -Uri $TemplateUrl -ErrorAction Stop
    $template = $templateResponse.Content

    # Descargar la imagen de cabecera
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) "template.png"
    Write-Verbose "Descargando imagen de cabecera a $tempPath"
    Invoke-WebRequest -Uri $TemplateHeaderGraphicUrl -OutFile $tempPath -ErrorAction Stop

    # Reemplazar en la plantilla
    $body = $template -replace "_TABLE_", $TableHtml -replace "_DATE_", $DeleteDate

    # Enviar correo
    Write-Verbose "Enviando correo a: $($To -join ';')"
    .\Send-MailMessageEx.ps1 `
        -Body $body `
        -Subject $Subject `
        -Credential $MailCredentials `
        -SmtpServer $MailServer `
        -Port 587 `
        -BodyAsHtml `
        -UseSSL `
        -InlineAttachments @{ "tagging.png" = $tempPath } `
        -From $MailCredentials.UserName `
        -To $To `
        -Priority "Low"
}

# ----------------------------------------------------------------------------------
# PRINCIPAL
# ----------------------------------------------------------------------------------

Write-Verbose "Iniciando script de limpieza..."

$connectionName = "AzureRunAsConnection"
$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"

# URLs de plantillas HTML para correos
$TemplateUrl_Expired       = Get-AutomationVariable -Name "TemplateUrl-Cleanup"
$TemplateUrl_TooFar        = Get-AutomationVariable -Name "TemplateUrl-CleanupTooFar"
$TemplateHeaderGraphicUrl  = Get-AutomationVariable -Name "TemplateHeaderGraphicUrl-Cleanup"

# Regex para ignorar ciertos RG (ej: "(Default-|AzureFunctions|Api-Default-).*")
$RGNamesIgnoreRegex        = Get-AutomationVariable -Name "RG_NamesIgnore"

# Credenciales para envío de correo
$mailCreds  = Get-AutomationPSCredential -Name 'Office365'
$mailServer = "smtp.office365.com"

# Conectar a Azure
Connect-ToAzure -ConnectionName $connectionName -SubscriptionId $SubscriptionId

# Calcular días para filtros (convierto PastMaxExpiryDays a negativo)
$pastDaysValue = -[int]$PastMaxExpiryDays

Write-Verbose "Recuperando Resource Groups con deleteAfter..."
$results = Get-ResourceGroupsByDeleteAfter -PastDays $pastDaysValue -FutureDays $FutureMaxExpiryDays -IgnoreRegex $RGNamesIgnoreRegex

$expired         = $results.Expired
$tooFarOutExpiry = $results.TooFarOutExpiry

# ----------------------------------------------------------------------------------
# 1) Procesar los expirados
# ----------------------------------------------------------------------------------
Populate-ResourceInfo -ResourceGroups $expired

if ($expired.Count -gt 0) {
    # Construir la tabla HTML
    $rgTableHtml = Build-ResourceGroupTableHtml -ResourceGroups $expired

    # Generar la lista de destinatarios adicionales (propiedad OwnerEmail)
    # Solo agregamos quienes tengan un correo válido (no vacío).
    $toAffected = $expired | ForEach-Object {
        if ($_.OwnerEmail) {
            "<$($_.OwnerEmail)>"
        }
    } | Where-Object { $_ }  # filtra nulos o vacíos

    $toAffectedString = $toAffected -join ";"

    # Asunto
    $subjectExpired = "$($expired.Count) Resource Groups vencidos"

    # Decide si enviamos a todos o solo al destinatario principal
    if ($WhatIf.IsPresent) {
        Write-Warning "WHATIF activo: el correo real no se enviará a los owners originales."
        $toComb = $To
    }
    else {
        if ($toAffectedString) {
            $toComb = "$To;$toAffectedString"
        }
        else {
            # Si no hay OwnerEmail en ningún RG, solo enviamos a $To
            $toComb = $To
        }
    }
    $toArray = $toComb -split ";"

    if ($PSCmdlet.ShouldProcess("Enviar correo de RGs expirados", "Envío de correo a: $toComb")) {
        Send-CleanupMail -Subject $subjectExpired `
                         -TemplateUrl $TemplateUrl_Expired `
                         -TemplateHeaderGraphicUrl $TemplateHeaderGraphicUrl `
                         -TableHtml $rgTableHtml `
                         -MailCredentials $mailCreds `
                         -MailServer $mailServer `
                         -To $toArray
    }
}
else {
    Write-Verbose "No hay RGs expirados; no se enviará correo para RGs expirados."
}

# ----------------------------------------------------------------------------------
# 2) Procesar los que tienen fecha de expiración muy lejana
# ----------------------------------------------------------------------------------
Populate-ResourceInfo -ResourceGroups $tooFarOutExpiry

if ($tooFarOutExpiry.Count -gt 0) {
    $rgTableHtml = Build-ResourceGroupTableHtml -ResourceGroups $tooFarOutExpiry

    $subjectFuture = "$($tooFarOutExpiry.Count) Resource Groups con Expiración > $FutureMaxExpiryDays días"

    # En este caso, tradicionalmente solo se envía a $To (ajusta si quieres añadir owners)
    $toArray = $To -split ";"

    if ($PSCmdlet.ShouldProcess("Enviar correo de RGs con expiración lejana", "Envío de correo a: $To")) {
        Send-CleanupMail -Subject $subjectFuture `
                         -TemplateUrl $TemplateUrl_TooFar `
                         -TemplateHeaderGraphicUrl $TemplateHeaderGraphicUrl `
                         -TableHtml $rgTableHtml `
                         -MailCredentials $mailCreds `
                         -MailServer $mailServer `
                         -To $toArray
    }
}
else {
    Write-Verbose "No hay RGs con fecha de expiración superior a $FutureMaxExpiryDays días; no se enviará correo."
}

Write-Verbose "Script finalizado."
