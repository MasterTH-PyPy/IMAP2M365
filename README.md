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

## Lizenzierung (node-locked, offline)

- **Ein RSA-3072-Schlüsselpaar.** Privater Key nur im Generator (`tools/LicenseGenerator`),
  öffentlicher Key im Tool (`src/Core/Licensing/EmbeddedPublicKey.cs`). Das Tool kann
  Lizenzen nur prüfen, nicht fälschen.
- **Computer-ID**: Hash aus MachineGuid + BIOS-UUID + Rechnername (`MachineId`), plus
  VM-Erkennung (VMware/VirtualBox/Hyper-V/KVM/QEMU…). Reiter „Lizenz" zeigt ID/Name/VM.
- **Ablauf**: Kunde gibt Firmendaten ein → „Aktivierungsanfrage erzeugen" (mit dem
  öffentlichen Key hybrid verschlüsselt, nur der Lizenzgeber liest sie) → an Lizenzgeber
  senden. Generator entschlüsselt, zeigt Kundendaten + Computer-ID + VM-Flag, man gibt die
  Kontenzahl ein → signierter Lizenzschlüssel → Kunde aktiviert ihn im Reiter „Lizenz".
- **Verbrauch**: **Vollmigration** zieht pro Zielpostfach eine Lizenz (idempotent über SMTP);
  **Testmodus** kopiert nur die ersten 10 Mails/Postfach und verbraucht nichts. Verbrauchsstand
  liegt AES-GCM-verschlüsselt und an die Computer-ID gebunden unter `ProgramData\IMAP2M365`.
- **Grenzen (ehrlich)**: VM-Umgehung und Reset durch Löschen der Verbrauchsdatei sind offline
  ohne Lizenzserver nicht vollständig verhinderbar; die Hürde ist hoch, aber nicht absolut.
- **Produktivschlüssel**: im Generator `[1] Schlüsselpaar erzeugen`, `public.pem` in
  `EmbeddedPublicKey.cs` eintragen, `private.pem` geheim halten.

## Voraussetzungen

- Windows mit PowerShell 7 (für den EXO-Teil) und dem Modul
  `ExchangeOnlineManagement`. Das Tool versucht die Installation automatisch
  über `Install-PSResource`; klappt das im Host nicht, einmalig manuell:
  `Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber`.
- Für Schritt 1 (App anlegen): Anmeldung als Global Admin oder Cloud Application
  Administrator. Der Login läuft interaktiv im Browser (kein Device-Code).

## Einrichtungs-Wizard (GUI-Reiter „Einrichtung")

Die komplette M365-Vorarbeit steckt im Tool – kein manuelles PowerShell mehr nötig:

1. **Entra-App anlegen** (Graph, `GraphProvisioner`): App-Registrierung, Permission
   `IMAP.AccessAsApp` (+ `Mail.ReadWrite` fürs Graph-Target), Client-Secret,
   Admin-Consent via `appRoleAssignment`. Admin meldet sich einmal per Device-Code an.
2. **Mit EXO verbinden** (PowerShell SDK in-process, `ExoRunner`).
3. **Postfächer laden**: `Get-EXOMailbox` liest **User- und freigegebene Postfächer**
   inkl. IMAP-Status (`Get-EXOCASMailbox`) direkt ins Grid.
4. **Einrichtung ausführen**: `New-ServicePrincipal`, pro Postfach `Set-CASMailbox
   -ImapEnabled` (wichtig bei Shared Mailboxes), `Add-MailboxPermission` (FullAccess),
   optional `New-ApplicationAccessPolicy`, Abschluss-Test. Alles mit **Dry-Run**.

**Shared Mailboxes**: app-only XOAUTH2 ist dafür der richtige Weg – kein interaktiver
Login ins (gesperrte) Postfach nötig, Zugriff rein über die Postfachberechtigung.

**Grid**: Ziel-Postfach + Typ (User/Shared) + IMAP-Status werden ausgelesen; Quell-User
(vorbelegt mit Ziel-SMTP, editierbar), Quell-Kennwort (maskiert) und optionaler
Quell-Server-Override pro Zeile werden erfasst. Zentraler Quell-IMAP-Server gilt als
Default für alle Zeilen. Passwörter verschlüsselt in `config.enc` (AES-256), nie im Repo.

## Roadmap

- [x] Core-Interfaces + Modelle
- [x] ExchangeOnlineTokenProvider (app-only, outlook.office365.com/.default)
- [x] ImapSource (MailKit, Quelle: Basic oder XOAUTH2)
- [x] ImapAppendTarget (MailKit XOAUTH2, APPEND mit Flags + INTERNALDATE)
- [x] MigrationEngine (Ordner-Mapping, Dedup über Message-Id, Dry-Run)
- [x] WinForms-GUI (Einrichtungs-Wizard, Postfach-Grid, Mapping, Fortschritt, Log)
- [x] GraphTarget (zweiter Provider, MIME-Import + Throttling), umschaltbar in der GUI
- [x] Graph: große Nachrichten (>4 MB) über Upload-Session für Anhänge
- [x] App-Icon (Umschlag + Pfeil zur Wolke), `assets/app-icon.svg` → `src/app.ico`

## Graph-Backend: große Nachrichten

Graph begrenzt den Request-Body auf 4 MB und bietet **keine** Upload-Session für rohes
MIME. Daher:
- **≤ 3 MB**: roher MIME-Import (`text/plain`, base64) — höchste Treue.
- **> 3 MB**: Nachricht aus den MIME-Teilen gebaut (Betreff/Body/Empfänger),
  Empfangs-/Sendedatum über MAPI-Extended-Properties (`0x0E06`/`0x0039`), kleine
  Anhänge inline, große Anhänge via `createUploadSession` in 1,25-MiB-Blöcken.
  Treue-Hinweis: exakte MIME-Struktur/Custom-Header gehen bei diesem Pfad verloren —
  wer 1:1-Treue für Großmails braucht, nutzt das IMAP-APPEND-Backend (kein 4-MB-Limit).
