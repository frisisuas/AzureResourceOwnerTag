<#
.SYNOPSIS
  Etiqueta automáticamente los Resource Groups (RG) con 'owner' a partir de la actividad reciente
  en el Azure Activity Log. Opcionalmente, envía una notificación por correo.

.DESCRIPTION
  - Se buscan todos los RG sin la etiqueta 'owner'.
  - Para cada RG se revisan los registros del Activity Log de los últimos N días (máx. 14).
  - Si se identifica un Caller con formato de email, se asume como 'owner'.
  - Se aplica la etiqueta 'owner' y 'deleteAfter' (con fecha 1 mes en el futuro).
  - Al final, se notifica por correo a una lista compuesta por un destinatario principal ($To)
    y los owners encontrados, a menos que se use -NoEmail o -WhatIf.

.PARAMETER WhatIf
  Muestra las acciones que se realizarían (aplica ShouldProcess) sin ejecutarlas realmente.

.PARAMETER To
  Dirección o direcciones (separadas por ;) a donde se enviará el informe de nuevos RG etiquetados.

.PARAMETER DayCount
  Días hacia atrás a consultar en el Activity Log. Valor máximo: 14.

.PARAMETER NoEmail
  Si se especifica, no se envía correo, pero igualmente se realiza el etiquetado (salvo que también
  se use -WhatIf).

.EXAMPLE
  .\AutoTagResources_v3.ps1 -To "admin@contoso.com" -DayCount 3

.NOTES
  Requiere:
   - Módulo Az instalado en la cuenta de Azure Automation.
   - Variable de Automatización 'AzureRunAsConnection'.
   - Las variables TemplateUrl, TemplateHeaderGraphicUrl, RG_NamesIgnore, SubscriptionId
     configuradas en Azure Automation.
   - Credencial Office365 (Get-AutomationPSCredential -Name 'Office365').

#requires -Module Az.Accounts
#requires -Module Az.Resources
#>

[CmdletBinding(SupportsShouldProcess = $true)]
Param(
    [Parameter()]
    [switch]$WhatIf,

    [Parameter(Mandatory = $true)]
    [string]$To,

    [Parameter()]
    [ValidateRange(1, 14)]
    [int32]$DayCount = 1,

    [Parameter()]
    [switch]$NoEmail
)

# ----------------------------------------------------------------------------------
# 1) Conectar a Azure
# ----------------------------------------------------------------------------------
Write-Verbose "Iniciando script de autoetiquetado..."

try {
    $connectionName = "AzureRunAsConnection"
    Write-Verbose "Obteniendo la conexión '$connectionName'..."
    $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName
    if (-not $servicePrincipalConnection) {
        throw "No se encontró la conexión '$connectionName'."
    }

    Write-Verbose "Conectando a Azure con credenciales de servicio (módulo Az)..."
    Connect-AzAccount `
        -ServicePrincipal `
        -Tenant $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint | Out-Null

    # Cambiar el contexto a la suscripción
    $SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"
    Write-Verbose "Estableciendo contexto a la suscripción: $SubscriptionId"
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
}
catch {
    Write-Error "Error al iniciar sesión en Azure: $($_.Exception.Message)"
    throw
}

if ($WhatIf) {
    Write-Warning "El script se está ejecutando en modo WhatIf: no se aplicarán cambios definitivos."
}
if ($NoEmail) {
    Write-Warning "Parámetro -NoEmail: no se enviará correo al final."
}

# ----------------------------------------------------------------------------------
# 2) Variables de Automatización y configuración
# ----------------------------------------------------------------------------------
$TemplateUrl              = Get-AutomationVariable -Name "TemplateUrl"
$TemplateHeaderGraphicUrl = Get-AutomationVariable -Name "TemplateHeaderGraphicUrl"
$RGNamesIgnoreRegex       = Get-AutomationVariable -Name "RG_NamesIgnore"

$mailCreds  = Get-AutomationPSCredential -Name 'Office365'
$mailServer = "smtp.office365.com"

# Se asegura que DayCount sea al menos 1
if ($DayCount -lt 1) { $DayCount = 1 }

# La etiqueta 'deleteAfter' se fija 1 mes en el futuro
$deleteDate = (Get-Date).AddMonths(1)

# ----------------------------------------------------------------------------------
# 3) Función: Obtener RG sin la etiqueta 'owner'
# ----------------------------------------------------------------------------------
function Get-ResourceGroupsWithoutOwner {
    Param(
        [string]$RGNamesIgnoreRegex
    )
    
    Write-Verbose "Obteniendo todos los Resource Groups..."
    $allRGs = Get-AzResourceGroup

    Write-Verbose "Total RGs encontrados: $($allRGs.Count)"

    # RGs con etiqueta 'owner'
    $rgWithOwner = $allRGs | Where-Object { $_.Tags.ContainsKey('owner') }
    Write-Verbose "RGs con etiqueta 'owner': $($rgWithOwner.Count)"

    # RGs sin etiqueta 'owner'
    $rgWithoutOwner = $allRGs | Where-Object { -not $_.Tags.ContainsKey('owner') }
    Write-Verbose "RGs sin etiqueta 'owner': $($rgWithoutOwner.Count)"

    # Excluir RGs que cumplan el patrón ignorar
    $rgFinal = $rgWithoutOwner | Where-Object { $_.ResourceGroupName -notmatch $RGNamesIgnoreRegex }
    Write-Verbose "Filtrados (regex $RGNamesIgnoreRegex). RGs resultantes: $($rgFinal.Count)"

    return $rgFinal
}

# ----------------------------------------------------------------------------------
# 4) Función: Buscar posible 'owner' en Activity Log
# ----------------------------------------------------------------------------------
function Get-OwnerFromActivityLog {
    Param(
        [string]$ResourceGroupName,
        [int]$DaysLookback
    )
    
    Write-Verbose "Consultando registros de actividad en RG: $ResourceGroupName (últimos $DaysLookback días)"
    try {
        $logs = Get-AzLog `
            -ResourceGroupName $ResourceGroupName `
            -StartTime (Get-Date).AddDays(-$DaysLookback) `
            -EndTime (Get-Date) `
            -Status "Succeeded" 
    }
    catch {
        Write-Warning "Error al obtener logs para $ResourceGroupName: $($_.Exception.Message)"
        return $null
    }

    if (-not $logs) {
        Write-Verbose "Sin actividad en logs para $ResourceGroupName."
        return $null
    }

    # Filtrar logs con Caller tipo "alguien@dominio"
    $callers = $logs |
        Where-Object { $_.Caller -like "*@*" } |
        Where-Object {
            # Excluir operaciones irrelevantes
            $_.OperationName.Value -ne "Microsoft.Storage/storageAccounts/listKeys/action"
        } |
        Where-Object {
            # Excluir donde ya se manipuló 'tags' con 'alias' (o 'owner')
            -not ($_.Properties.Content.requestbody -like "*tags*alias*" -or
                  $_.Properties.Content.responsebody -like "*tags*alias*")
        } |
        Sort-Object -Property Caller -Unique |
        Select-Object -ExpandProperty Caller

    if ($callers -and $callers.Count -gt 0) {
        # Tomar el primer Caller detectado
        $firstCaller = $callers[0]
        Write-Verbose "Caller detectado: $firstCaller"
        return $firstCaller  # Devolvemos el correo completo
    }
    else {
        Write-Verbose "No se halló Caller válido para $ResourceGroupName."
        return $null
    }
}

# ----------------------------------------------------------------------------------
# 5) Función: Etiquetar RG con 'owner' y 'deleteAfter'
# ----------------------------------------------------------------------------------
function Tag-ResourceGroup {
    Param(
        [string]$ResourceGroupName,
        [string]$OwnerEmail,
        [datetime]$DeleteDate
    )

    Write-Verbose "Etiquetando RG: $ResourceGroupName => owner=$OwnerEmail, deleteAfter=$($DeleteDate.ToString('MM/dd/yy'))"
    if ($PSCmdlet.ShouldProcess($ResourceGroupName, "Aplicar etiquetas")) {
        if (-not $WhatIf) {
            Set-AzResourceGroup -Name $ResourceGroupName -Tag @{
                owner = $OwnerEmail
                deleteAfter = $DeleteDate.ToString("MM/dd/yy")
            } | Out-Null
        }
        else {
            Write-Warning "WhatIf activo: No se aplican cambios en $ResourceGroupName."
        }
    }
}

# ----------------------------------------------------------------------------------
# 6) Función: Enviar correo con los RG recién etiquetados
# ----------------------------------------------------------------------------------
function Send-TagEmail {
    Param(
        [System.Collections.ArrayList]$TaggedResults,
        [string]$TemplateUrl,
        [string]$TemplateHeaderGraphicUrl,
        [DateTime]$DeleteDate,
        [PSCredential]$MailCreds,
        [string]$MailServer,
        [string]$To,
        [switch]$WhatIf
    )

    if ($TaggedResults.Count -eq 0) {
        Write-Verbose "No hay RGs recién etiquetados; no se envía correo."
        return
    }

    # Construir tabla HTML
    $tableRows = $TaggedResults | ForEach-Object {
        "<tr><td>$($_.Name)</td><td>$($_.Owner)</td></tr>"
    } 
    $rgTable = $tableRows -join ""

    # Descargar plantilla HTML
    Write-Verbose "Descargando plantilla desde $TemplateUrl"
    $templateContent = (Invoke-WebRequest -Uri $TemplateUrl -ErrorAction Stop).Content
    
    # Descargar imagen a carpeta temporal
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) "template.png"
    Write-Verbose "Descargando imagen de cabecera a $tempPath"
    Invoke-WebRequest -Uri $TemplateHeaderGraphicUrl -OutFile $tempPath -ErrorAction Stop

    # Reemplazar marcadores _TABLE_ y _DATE_
    $formattedDeleteDate = $DeleteDate.ToString("MM/dd/yy")
    $body = $templateContent `
        -replace "_TABLE_", $rgTable `
        -replace "_DATE_", $formattedDeleteDate

    # Construir lista de destinatarios: $To + owners
    $toAffected = $TaggedResults | ForEach-Object {
        "<$($_.Owner)>"
    }
    $toAll = "$To;$(($toAffected -join ';'))"

    if ($WhatIf) {
        Write-Warning "WHATIF habilitado. Solo se enviaría a $To (no a los owners)."
        $toAll = $To
    }

    $toArray = $toAll -split ";"
    $subject = "$($TaggedResults.Count) new resource groups automatically tagged"

    Write-Verbose "Enviando correo a: $toAll"
    .\Send-MailMessageEx.ps1 `
        -Body $body `
        -Subject $subject `
        -Credential $MailCreds `
        -SmtpServer $MailServer `
        -Port 587 `
        -BodyAsHtml `
        -UseSSL `
        -InlineAttachments @{ "tagging.png" = $tempPath } `
        -From $MailCreds.UserName `
        -To $toArray `
        -Priority "Low"
}

# ----------------------------------------------------------------------------------
# PROCESO PRINCIPAL
# ----------------------------------------------------------------------------------

# 1) Obtener RGs sin 'owner'
$rgWithoutOwner = Get-ResourceGroupsWithoutOwner -RGNamesIgnoreRegex $RGNamesIgnoreRegex

# 2) Recorremos y etiquetamos
Write-Verbose "Buscando posibles owners en logs de los últimos $DayCount días..."
$tagResults = New-Object System.Collections.ArrayList

foreach ($rg in $rgWithoutOwner) {
    $index   = [array]::IndexOf($rgWithoutOwner, $rg)
    $percent = 100 * ($index / $rgWithoutOwner.Count)

    Write-Progress -Activity "Analizando Activity Log" -Status "$($rg.ResourceGroupName)" -PercentComplete $percent

    $possibleOwner = Get-OwnerFromActivityLog -ResourceGroupName $rg.ResourceGroupName -DaysLookback $DayCount
    if ($possibleOwner) {
        Tag-ResourceGroup -ResourceGroupName $rg.ResourceGroupName `
                          -OwnerEmail $possibleOwner `
                          -DeleteDate $deleteDate
        # Agregar al listado final
        [void]$tagResults.Add(
            [pscustomobject]@{
                Name  = $rg.ResourceGroupName
                Owner = $possibleOwner
            }
        )
    }
}

Write-Progress -Activity "Analizando Activity Log" -Completed -Status "Finalizado"

# 3) Si no está el switch -NoEmail, enviamos correo
if (-not $NoEmail) {
    Send-TagEmail -TaggedResults $tagResults `
                  -TemplateUrl $TemplateUrl `
                  -TemplateHeaderGraphicUrl $TemplateHeaderGraphicUrl `
                  -DeleteDate $deleteDate `
                  -MailCreds $mailCreds `
                  -MailServer $mailServer `
                  -To $To `
                  -WhatIf:$WhatIf
}
else {
    Write-Verbose "Se omitirá el envío de correo por -NoEmail."
}

Write-Verbose "Script finalizado. Resource Groups etiquetados:"
$tagResults
