Append a prev/next navigation block to the end of every IT-* and EN-* guide page (excluding Home pages) in /home/guido/Repository/dataflow-procurement-software.wiki/.

## Files and navigation blocks

### IT files

IT-01-Primi-passi.md
```
---
[Pagina successiva →](IT-02-Schermata-principale)
```

IT-02-Schermata-principale.md
```
---
[← Pagina precedente](IT-01-Primi-passi) | [Pagina successiva →](IT-03-Creare-una-nuova-RdO)
```

IT-03-Creare-una-nuova-RdO.md
```
---
[← Pagina precedente](IT-02-Schermata-principale) | [Pagina successiva →](IT-04-Gestire-una-RdO-esistente)
```

IT-04-Gestire-una-RdO-esistente.md
```
---
[← Pagina precedente](IT-03-Creare-una-nuova-RdO) | [Pagina successiva →](IT-05-Analisi-SQDC)
```

IT-05-Analisi-SQDC.md
```
---
[← Pagina precedente](IT-04-Gestire-una-RdO-esistente) | [Pagina successiva →](IT-06-Value-Stream-Mapping)
```

IT-06-Value-Stream-Mapping.md
```
---
[← Pagina precedente](IT-05-Analisi-SQDC) | [Pagina successiva →](IT-07-KPI-Dashboard)
```

IT-07-KPI-Dashboard.md
```
---
[← Pagina precedente](IT-06-Value-Stream-Mapping) | [Pagina successiva →](IT-08-Impostazioni-e-manutenzione)
```

IT-08-Impostazioni-e-manutenzione.md
```
---
[← Pagina precedente](IT-07-KPI-Dashboard) | [Pagina successiva →](IT-09-Lavoro-multiutente)
```

IT-09-Lavoro-multiutente.md
```
---
[← Pagina precedente](IT-08-Impostazioni-e-manutenzione) | [Pagina successiva →](IT-10-Problemi-comuni-e-soluzioni)
```

IT-10-Problemi-comuni-e-soluzioni.md
```
---
[← Pagina precedente](IT-09-Lavoro-multiutente) | [Pagina successiva →](IT-11-Best-practices)
```

IT-11-Best-practices.md
```
---
[← Pagina precedente](IT-10-Problemi-comuni-e-soluzioni) | [Pagina successiva →](IT-12-Glossario)
```

IT-12-Glossario.md
```
---
[← Pagina precedente](IT-11-Best-practices) | [Pagina successiva →](IT-13-Log-e-diagnostica)
```

IT-13-Log-e-diagnostica.md
```
---
[← Pagina precedente](IT-12-Glossario) | [Pagina successiva →](IT-14-Supporto)
```

IT-14-Supporto.md
```
---
[← Pagina precedente](IT-13-Log-e-diagnostica)
```

---

### EN files

EN-01-Getting-Started.md
```
---
[Next page →](EN-02-Main-Screen)
```

EN-02-Main-Screen.md
```
---
[← Previous page](EN-01-Getting-Started) | [Next page →](EN-03-Create-a-New-RFQ)
```

EN-03-Create-a-New-RFQ.md
```
---
[← Previous page](EN-02-Main-Screen) | [Next page →](EN-04-Manage-an-Existing-RFQ)
```

EN-04-Manage-an-Existing-RFQ.md
```
---
[← Previous page](EN-03-Create-a-New-RFQ) | [Next page →](EN-05-SQDC-Analysis)
```

EN-05-SQDC-Analysis.md
```
---
[← Previous page](EN-04-Manage-an-Existing-RFQ) | [Next page →](EN-06-Value-Stream-Mapping)
```

EN-06-Value-Stream-Mapping.md
```
---
[← Previous page](EN-05-SQDC-Analysis) | [Next page →](EN-07-KPI-Dashboard)
```

EN-07-KPI-Dashboard.md
```
---
[← Previous page](EN-06-Value-Stream-Mapping) | [Next page →](EN-08-Settings-and-Maintenance)
```

EN-08-Settings-and-Maintenance.md
```
---
[← Previous page](EN-07-KPI-Dashboard) | [Next page →](EN-09-Multi-User-Work)
```

EN-09-Multi-User-Work.md
```
---
[← Previous page](EN-08-Settings-and-Maintenance) | [Next page →](EN-10-Common-Issues-and-Solutions)
```

EN-10-Common-Issues-and-Solutions.md
```
---
[← Previous page](EN-09-Multi-User-Work) | [Next page →](EN-11-Best-Practices)
```

EN-11-Best-Practices.md
```
---
[← Previous page](EN-10-Common-Issues-and-Solutions) | [Next page →](EN-12-Glossary)
```

EN-12-Glossary.md
```
---
[← Previous page](EN-11-Best-Practices) | [Next page →](EN-13-Logs-and-Diagnostics)
```

EN-13-Logs-and-Diagnostics.md
```
---
[← Previous page](EN-12-Glossary) | [Next page →](EN-14-Support)
```

EN-14-Support.md
```
---
[← Previous page](EN-13-Logs-and-Diagnostics)
```

---

## Implementation

Use `run_in_terminal` to run the following bash script from the wiki directory:

```bash
cd /home/guido/Repository/dataflow-procurement-software.wiki && \
printf '\n\n---\n[Pagina successiva →](IT-02-Schermata-principale)' >> IT-01-Primi-passi.md && \
printf '\n\n---\n[← Pagina precedente](IT-01-Primi-passi) | [Pagina successiva →](IT-03-Creare-una-nuova-RdO)' >> IT-02-Schermata-principale.md && \
printf '\n\n---\n[← Pagina precedente](IT-02-Schermata-principale) | [Pagina successiva →](IT-04-Gestire-una-RdO-esistente)' >> IT-03-Creare-una-nuova-RdO.md && \
printf '\n\n---\n[← Pagina precedente](IT-03-Creare-una-nuova-RdO) | [Pagina successiva →](IT-05-Analisi-SQDC)' >> IT-04-Gestire-una-RdO-esistente.md && \
printf '\n\n---\n[← Pagina precedente](IT-04-Gestire-una-RdO-esistente) | [Pagina successiva →](IT-06-Value-Stream-Mapping)' >> IT-05-Analisi-SQDC.md && \
printf '\n\n---\n[← Pagina precedente](IT-05-Analisi-SQDC) | [Pagina successiva →](IT-07-KPI-Dashboard)' >> IT-06-Value-Stream-Mapping.md && \
printf '\n\n---\n[← Pagina precedente](IT-06-Value-Stream-Mapping) | [Pagina successiva →](IT-08-Impostazioni-e-manutenzione)' >> IT-07-KPI-Dashboard.md && \
printf '\n\n---\n[← Pagina precedente](IT-07-KPI-Dashboard) | [Pagina successiva →](IT-09-Lavoro-multiutente)' >> IT-08-Impostazioni-e-manutenzione.md && \
printf '\n\n---\n[← Pagina precedente](IT-08-Impostazioni-e-manutenzione) | [Pagina successiva →](IT-10-Problemi-comuni-e-soluzioni)' >> IT-09-Lavoro-multiutente.md && \
printf '\n\n---\n[← Pagina precedente](IT-09-Lavoro-multiutente) | [Pagina successiva →](IT-11-Best-practices)' >> IT-10-Problemi-comuni-e-soluzioni.md && \
printf '\n\n---\n[← Pagina precedente](IT-10-Problemi-comuni-e-soluzioni) | [Pagina successiva →](IT-12-Glossario)' >> IT-11-Best-practices.md && \
printf '\n\n---\n[← Pagina precedente](IT-11-Best-practices) | [Pagina successiva →](IT-13-Log-e-diagnostica)' >> IT-12-Glossario.md && \
printf '\n\n---\n[← Pagina precedente](IT-12-Glossario) | [Pagina successiva →](IT-14-Supporto)' >> IT-13-Log-e-diagnostica.md && \
printf '\n\n---\n[← Pagina precedente](IT-13-Log-e-diagnostica)' >> IT-14-Supporto.md && \
printf '\n\n---\n[Next page →](EN-02-Main-Screen)' >> EN-01-Getting-Started.md && \
printf '\n\n---\n[← Previous page](EN-01-Getting-Started) | [Next page →](EN-03-Create-a-New-RFQ)' >> EN-02-Main-Screen.md && \
printf '\n\n---\n[← Previous page](EN-02-Main-Screen) | [Next page →](EN-04-Manage-an-Existing-RFQ)' >> EN-03-Create-a-New-RFQ.md && \
printf '\n\n---\n[← Previous page](EN-03-Create-a-New-RFQ) | [Next page →](EN-05-SQDC-Analysis)' >> EN-04-Manage-an-Existing-RFQ.md && \
printf '\n\n---\n[← Previous page](EN-04-Manage-an-Existing-RFQ) | [Next page →](EN-06-Value-Stream-Mapping)' >> EN-05-SQDC-Analysis.md && \
printf '\n\n---\n[← Previous page](EN-05-SQDC-Analysis) | [Next page →](EN-07-KPI-Dashboard)' >> EN-06-Value-Stream-Mapping.md && \
printf '\n\n---\n[← Previous page](EN-06-Value-Stream-Mapping) | [Next page →](EN-08-Settings-and-Maintenance)' >> EN-07-KPI-Dashboard.md && \
printf '\n\n---\n[← Previous page](EN-07-KPI-Dashboard) | [Next page →](EN-09-Multi-User-Work)' >> EN-08-Settings-and-Maintenance.md && \
printf '\n\n---\n[← Previous page](EN-08-Settings-and-Maintenance) | [Next page →](EN-10-Common-Issues-and-Solutions)' >> EN-09-Multi-User-Work.md && \
printf '\n\n---\n[← Previous page](EN-09-Multi-User-Work) | [Next page →](EN-11-Best-Practices)' >> EN-10-Common-Issues-and-Solutions.md && \
printf '\n\n---\n[← Previous page](EN-10-Common-Issues-and-Solutions) | [Next page →](EN-12-Glossary)' >> EN-11-Best-Practices.md && \
printf '\n\n---\n[← Previous page](EN-11-Best-Practices) | [Next page →](EN-13-Logs-and-Diagnostics)' >> EN-12-Glossary.md && \
printf '\n\n---\n[← Previous page](EN-12-Glossary) | [Next page →](EN-14-Support)' >> EN-13-Logs-and-Diagnostics.md && \
printf '\n\n---\n[← Previous page](EN-13-Logs-and-Diagnostics)' >> EN-14-Support.md && \
echo "Done."
```

After running, verify with:
```bash
tail -4 IT-01-Primi-passi.md IT-14-Supporto.md EN-01-Getting-Started.md EN-14-Support.md
```
