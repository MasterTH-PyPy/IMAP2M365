<#
.SYNOPSIS
  Einmalige Exchange-Online-Vorarbeit fuer IMAP2M365 (app-only XOAUTH2).

.DESCRIPTION
  Registriert den Service-Principal der Entra-App in Exchange Online und vergibt
  Postfach-Rechte fuer die Migrations-Zielpostfaecher. Optional wird der Zugriff
  per ApplicationAccessPolicy auf eine Sicherheitsgruppe der Zielpostfaecher
  begrenzt (empfohlen, damit die App NICHT jedes Postfach im Tenant sieht).

  Voraussetzung in Entra (Portal, einmalig, NICHT in diesem Skript):
    - App-Registrierung anlegen, Client-Secret erzeugen.
    - API-Permissions (Application): "Office 365 Exchange Online" -> IMAP.AccessAsApp
      (fuer das Graph-Target zusaetzlich: Microsoft Graph -> Mail.ReadWrite)
    - Admin-Consent erteilen.

.NOTES
  Benoetigt das Modul ExchangeOnlineManagement. Ausfuehren als Exchange-Admin.
#>

param(
    [Parameter(Mandatory)] [string]   $AppId,                 # ClientId der Entra-App
    [Parameter(Mandatory)] [string]   $ObjectId,              # Enterprise-App (Service-Principal) ObjectId aus Entra
    [Parameter(Mandatory)] [string]   $Organization,          # z. B. contoso.onmicrosoft.com
    [Parameter(Mandatory)] [string[]] $TargetMailboxes,       # UPNs der Zielpostfaecher
    [string] $ScopeGroup                                      # optional: Mailenabled Security Group fuer ApplicationAccessPolicy
)

Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -Organization $Organization -ShowBanner:$false

# 1) Service-Principal in Exchange registrieren (idempotent)
$existing = Get-ServicePrincipal -ErrorAction SilentlyContinue |
            Where-Object { $_.AppId -eq $AppId }
if (-not $existing) {
    Write-Host "Registriere ServicePrincipal fuer AppId $AppId ..."
    New-ServicePrincipal -AppId $AppId -ServiceId $ObjectId -DisplayName "IMAP2M365 Migration" | Out-Null
} else {
    Write-Host "ServicePrincipal existiert bereits - ueberspringe."
}

# 2) Postfach-Rechte vergeben (FullAccess reicht fuer IMAP-APPEND)
foreach ($mbx in $TargetMailboxes) {
    Write-Host "Add-MailboxPermission: $mbx"
    Add-MailboxPermission -Identity $mbx -User $ObjectId -AccessRights FullAccess -AutoMapping:$false -ErrorAction Stop | Out-Null
}

# 3) Optional: Zugriff auf genau diese Postfaecher begrenzen
if ($ScopeGroup) {
    Write-Host "Setze ApplicationAccessPolicy (Restrict) auf Gruppe $ScopeGroup ..."
    New-ApplicationAccessPolicy -AppId $AppId -PolicyScopeGroupId $ScopeGroup `
        -AccessRight RestrictAccess -Description "IMAP2M365 nur auf Migrations-Postfaecher" | Out-Null
    Write-Host "Pruefen mit: Test-ApplicationAccessPolicy -AppId $AppId -Identity <upn>"
}

Write-Host "Fertig. App $AppId kann jetzt per app-only XOAUTH2 (IMAP) auf die Zielpostfaecher zugreifen."
Disconnect-ExchangeOnline -Confirm:$false
