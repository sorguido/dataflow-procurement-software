# Plan: Aggiunta navigazione Prev/Next ai file wiki DataFlow

## Obiettivo
Aggiungere un blocco di navigazione in fondo a ogni pagina guida (IT-01…IT-14 e EN-01…EN-14), escludendo le Home. Nessun altro file viene toccato.

## Stato attuale
- 28 file coinvolti (14 IT + 14 EN)
- Nessun file ha già un blocco di navigazione → nessun controllo di duplicazione necessario ma va fatto per coerenza futura

---

## Mapping completo dei blocchi da aggiungere

### ITALIANO (link senza estensione .md)

| File | Blocco da aggiungere |
|------|---------------------|
| IT-01-Primi-passi.md | `---\n[Pagina successiva →](IT-02-Schermata-principale)` |
| IT-02-Schermata-principale.md | `---\n[← Pagina precedente](IT-01-Primi-passi) \| [Pagina successiva →](IT-03-Creare-una-nuova-RdO)` |
| IT-03-Creare-una-nuova-RdO.md | `---\n[← Pagina precedente](IT-02-Schermata-principale) \| [Pagina successiva →](IT-04-Gestire-una-RdO-esistente)` |
| IT-04-Gestire-una-RdO-esistente.md | `---\n[← Pagina precedente](IT-03-Creare-una-nuova-RdO) \| [Pagina successiva →](IT-05-Analisi-SQDC)` |
| IT-05-Analisi-SQDC.md | `---\n[← Pagina precedente](IT-04-Gestire-una-RdO-esistente) \| [Pagina successiva →](IT-06-Value-Stream-Mapping)` |
| IT-06-Value-Stream-Mapping.md | `---\n[← Pagina precedente](IT-05-Analisi-SQDC) \| [Pagina successiva →](IT-07-KPI-Dashboard)` |
| IT-07-KPI-Dashboard.md | `---\n[← Pagina precedente](IT-06-Value-Stream-Mapping) \| [Pagina successiva →](IT-08-Impostazioni-e-manutenzione)` |
| IT-08-Impostazioni-e-manutenzione.md | `---\n[← Pagina precedente](IT-07-KPI-Dashboard) \| [Pagina successiva →](IT-09-Lavoro-multiutente)` |
| IT-09-Lavoro-multiutente.md | `---\n[← Pagina precedente](IT-08-Impostazioni-e-manutenzione) \| [Pagina successiva →](IT-10-Problemi-comuni-e-soluzioni)` |
| IT-10-Problemi-comuni-e-soluzioni.md | `---\n[← Pagina precedente](IT-09-Lavoro-multiutente) \| [Pagina successiva →](IT-11-Best-practices)` |
| IT-11-Best-practices.md | `---\n[← Pagina precedente](IT-10-Problemi-comuni-e-soluzioni) \| [Pagina successiva →](IT-12-Glossario)` |
| IT-12-Glossario.md | `---\n[← Pagina precedente](IT-11-Best-practices) \| [Pagina successiva →](IT-13-Log-e-diagnostica)` |
| IT-13-Log-e-diagnostica.md | `---\n[← Pagina precedente](IT-12-Glossario) \| [Pagina successiva →](IT-14-Supporto)` |
| IT-14-Supporto.md | `---\n[← Pagina precedente](IT-13-Log-e-diagnostica)` |

### INGLESE

| File | Blocco da aggiungere |
|------|---------------------|
| EN-01-Getting-Started.md | `---\n[Next page →](EN-02-Main-Screen)` |
| EN-02-Main-Screen.md | `---\n[← Previous page](EN-01-Getting-Started) \| [Next page →](EN-03-Create-a-New-RFQ)` |
| EN-03-Create-a-New-RFQ.md | `---\n[← Previous page](EN-02-Main-Screen) \| [Next page →](EN-04-Manage-an-Existing-RFQ)` |
| EN-04-Manage-an-Existing-RFQ.md | `---\n[← Previous page](EN-03-Create-a-New-RFQ) \| [Next page →](EN-05-SQDC-Analysis)` |
| EN-05-SQDC-Analysis.md | `---\n[← Previous page](EN-04-Manage-an-Existing-RFQ) \| [Next page →](EN-06-Value-Stream-Mapping)` |
| EN-06-Value-Stream-Mapping.md | `---\n[← Previous page](EN-05-SQDC-Analysis) \| [Next page →](EN-07-KPI-Dashboard)` |
| EN-07-KPI-Dashboard.md | `---\n[← Previous page](EN-06-Value-Stream-Mapping) \| [Next page →](EN-08-Settings-and-Maintenance)` |
| EN-08-Settings-and-Maintenance.md | `---\n[← Previous page](EN-07-KPI-Dashboard) \| [Next page →](EN-09-Multi-User-Work)` |
| EN-09-Multi-User-Work.md | `---\n[← Previous page](EN-08-Settings-and-Maintenance) \| [Next page →](EN-10-Common-Issues-and-Solutions)` |
| EN-10-Common-Issues-and-Solutions.md | `---\n[← Previous page](EN-09-Multi-User-Work) \| [Next page →](EN-11-Best-Practices)` |
| EN-11-Best-Practices.md | `---\n[← Previous page](EN-10-Common-Issues-and-Solutions) \| [Next page →](EN-12-Glossary)` |
| EN-12-Glossary.md | `---\n[← Previous page](EN-11-Best-Practices) \| [Next page →](EN-13-Logs-and-Diagnostics)` |
| EN-13-Logs-and-Diagnostics.md | `---\n[← Previous page](EN-12-Glossary) \| [Next page →](EN-14-Support)` |
| EN-14-Support.md | `---\n[← Previous page](EN-13-Logs-and-Diagnostics)` |

---

## Steps operativi

**Fase 1 — File italiani (IT-01 → IT-14)** *(14 file, eseguibili in parallelo)*

1. Leggere la fine del file
2. Appendere `\n---\n[blocco]` in fondo
3. Verificare assenza duplicati

**Fase 2 — File inglesi (EN-01 → EN-14)** *(14 file, eseguibili in parallelo con Fase 1)*

Stessa logica di Fase 1.

---

## Verifica finale

1. Aprire a campione 2–3 file (es. IT-01, IT-07, EN-14) → controllare che il blocco sia l'ultima riga
2. Verificare che `---` sia su riga propria, preceduto da riga vuota
3. `grep -r "Pagina successiva\|Next page"` → deve restituire esattamente 26 risultati (IT-01…IT-13 + EN-01…EN-13, escluse le ultime pagine)
4. `grep -r "Pagina precedente\|Previous page"` → deve restituire esattamente 26 risultati (IT-02…IT-14 + EN-02…EN-14, escluse le prime pagine)

---

## Decisioni

- Link senza estensione `.md` (standard GitHub Wiki)
- Separatore `|` semplice nel file (non escaped)
- File non toccati: IT-Home.md, EN-Home.md, Home.md
- Nessun refactor, nessun cambio di naming, nessuna logica aggiuntiva
