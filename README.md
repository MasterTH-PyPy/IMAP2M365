# IMAP2M365

WinForms-Tool (.NET 10) zum Migrieren von Postfächern von einem beliebigen
IMAP-Server nach Microsoft 365 / Exchange Online — mandantenweit als Admin.

Autor: Rudolf Kasper · Konventionen wie OLK2Any / PST2MBOX (graue Standard-UI,
AES-256-Config, Dry-Run, Versionsanzeige pro Commit).

## Warum kein EWS

EWS ist in Exchange Online seit **1. Oktober 2026** in der phasenweisen
Abschaltung, endgültiger Shutdown **1. April 2027**. Microsoft verweist auf
Microsoft Graph. Für ein neues Tool daher nur **IMAP (XOAUTH2)** oder **Graph**.

## Architektur

```
  IMailSource  ──►  MigrationEngine  ──►  IMailTarget
  (IMAP, MailKit)   (Ordner-Mapping,       ├─ ImapAppendTarget  (MailKit, XOAUTH2)  ← Default
                     Dry-Run, Dedup,        └─ GraphTarget       (Graph, Mail.ReadWrite) ← später
                     Fortschritt)
```

- **Ziel umschaltbar** über `IMailTarget`. IMAP-APPEND überträgt rohes MIME 1:1
  inkl. Flags (gelesen/ungelesen) und `INTERNALDATE` (Empfangsdatum bleibt
  korrekt). Graph re-kodiert Nachrichten beim Import und braucht eigenes
  Throttling-Handling — deshalb als zweiter Provider.
- **Auth: app-only Client-Credentials** (MSAL `ConfidentialClientApplication`).
  Ein Token pro Ziel-Backend-Scope, der SASL-XOAUTH2-Username ist die jeweilige
  Postfach-UPN. Kein interaktiver Login, läuft über Batch/CSV.

## Einmalige Exchange-Online-Vorarbeit

Siehe `prep/Register-ExchangeServicePrincipal.ps1`:
1. Entra-App-Registrierung: Application-Permission **Office 365 Exchange Online →
   `IMAP.AccessAsApp`** (für Graph-Target zusätzlich **Graph → `Mail.ReadWrite`**),
   Admin-Consent erteilen.
2. `New-ServicePrincipal` registriert die App in Exchange.
3. `Add-MailboxPermission` je Zielpostfach (oder scoped per
   `New-ApplicationAccessPolicy`, damit der App-Zugriff auf die Migrations-
   Postfächer begrenzt ist).

## Config (AES-256, wie gehabt)

`config.enc` → TenantId, ClientId, ClientSecret, Quell-IMAP-Defaults,
Postfach-Mapping-CSV-Pfad, Ordner-Ausschlüsse, Dry-Run-Flag.

## Roadmap

- [ ] Core-Interfaces + Modelle
- [ ] ExchangeOnlineTokenProvider (app-only, outlook.office365.com/.default)
- [ ] ImapSource (MailKit, Quelle: Basic oder XOAUTH2)
- [ ] ImapAppendTarget (MailKit XOAUTH2, APPEND mit Flags + INTERNALDATE)
- [ ] MigrationEngine (Ordner-Mapping, Dedup über Message-Id, Dry-Run)
- [ ] WinForms-GUI (Postfachliste/CSV, Mapping, Fortschritt, Log)
- [ ] GraphTarget (zweiter Provider)
- [ ] App-Icon (Umschlag + Pfeil nach Wolke)
