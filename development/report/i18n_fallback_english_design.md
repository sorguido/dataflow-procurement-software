# I18N Fallback English — Analisi e Design

## 1. Stato attuale fallback
Il comportamento corrente di `tr(text)` è questo:

- `tr` delega a `TranslationService.translate` (`utils/i18n_utils.py:131-133`, `:37-42`).
- il translator attivo è `gettext(...).gettext` della lingua corrente (`utils/i18n_utils.py:90-99`).
- se la chiave non esiste, gettext restituisce l'input (`msgid`) così com'è.

Quindi il fallback reale oggi è: `fallback = input`, non `fallback = inglese`.

Esempio Derisking:
- input a `tr`: `"Qualificato"` (`services/derisking_dashboard_service.py:15`)
- msgid catalogo: `"Qualified"` (`locale/en/LC_MESSAGES/dataflow.po:1870`)
- risultato: `"Qualificato"` (italiano), anche con UI inglese.

---

## 2. Problema strutturale
Il problema non è solo “chiave mancante”, ma “chiave in lingua sbagliata rispetto al catalogo”:

- catalogo con msgid prevalentemente EN (`"New"`, `"Under Evaluation"`, `"Qualified"`, `"Rejected"`) (`locale/en/...:1864-1874`, `locale/it/...:1878-1888`)
- alcuni canonici business sono IT (Derisking status: `Nuovo/In valutazione/Qualificato/Scartato`) (`models/potential_supplier.py:13-24`)

Se si invoca `tr(canonico_IT)`, il sistema non può garantire fallback EN, perché il msgid non è EN.

---

## 3. Opzioni possibili

### Opzione A
Intervenire globalmente su `tr(...)` con fallback “intelligente” post-lookup.

Pro:
- un solo punto teorico di controllo.

Contro:
- alto rischio regressione sistemica (tutta l'app usa `tr`).
- `output == input` non distingue bene tra “chiave mancante” e “traduzione volutamente identica”.
- non risolve da sola input canonici IT: senza normalizzazione dominio, `tr("Qualificato")` resta italiano.
- in conflitto con vincolo di minimo impatto.

---

### Opzione B
Introdurre un layer intermedio centralizzato di normalizzazione dominio -> msgid EN (prima di `tr`), sul pattern già esistente `normalize_rfq_type + translate_rfq_type` (`utils/i18n_utils.py:161-229`).

Pro:
- coerente con architettura refactor (`tr` resta API ufficiale).
- localizzato e reversibile.
- nessun cambio DB/canonico business.
- evita duplicazione se il mapping vive in un solo punto centrale.
- fallback EN garantito per i domini coperti: l'input a `tr` diventa sempre msgid EN.

Contro:
- richiede censire i domini non EN (inizialmente Derisking status).
- la garanzia “sempre EN” vale per i valori passati tramite il layer, non per input arbitrari bypassati.

---

### Opzione C (se utile)
Aggiungere msgid IT duplicati nel catalogo `.po/.mo` per i valori canonici italiani.

Pro:
- riduce i miss senza toccare codice chiamante.

Contro:
- aumenta entropia del catalogo (sinonimi multipli per stesso concetto).
- non risolve la causa architetturale (mancata separazione canonico/UI).
- difficile governare e scalare; rischio incoerenze future.

---

## 4. Scelta consigliata
Opzione consigliata: **Opzione B**.

Motivazione:
- coerenza con refactor: separazione canonico/UI e traduzione runtime via `tr`.
- rischio minimo: nessuna modifica globale a `tr`, nessun impatto DB.
- impatto controllato: intervento nel percorso di adattamento valore->UI, con riuso centrale del mapping.

---

## 5. Punto di intervento
Punto sicuro: layer di adattamento dati->UI, non business logic e non persistenza.

In pratica (design, non implementazione):
- centralizzare la normalizzazione dei valori non EN in `utils/i18n_utils.py` (stesso principio già usato da RFQ type).
- usare tale normalizzazione nei punti che oggi passano direttamente canonico a `tr`, in primis:
  - populate Derisking sheet (`dataflow.py:1563-1567` -> `services/derisking_dashboard_service.py:15`)
- riuso dello stesso mapping anche dove oggi è duplicato (dialog/export), evitando copie locali.

---

## 6. Strategia di fallback inglese
Strategia consigliata:

1. Normalizzare il valore canonico a una chiave msgid EN stabile (pre-`tr`).
2. Invocare `tr(msgid_en)`.
3. Se chiave assente nel catalogo locale, gettext restituisce `msgid_en` (inglese), quindi fallback EN reale.

Perché pre-`tr` e non post-`tr`:
- pre-`tr` garantisce input coerente con catalogo.
- post-`tr` è ambiguo e fragile (`same input` non significa sempre “miss”).

---

## 7. Impatto su altri moduli

- RFQ:
- nullo/positivo; il pattern è già quello (normalizzazione + traduzione), quindi resta coerente.

- Saving:
- nessun impatto diretto se il perimetro resta sui domini con canonico non EN.

- Cost Avoidance:
- nessun impatto diretto, stesso ragionamento di Saving.

- Derisking:
- impatto diretto positivo: status sempre coerente con lingua UI, fallback EN garantito per chiavi dominio note.

---

## 8. Rischi

- UI:
- basso, se il mapping è centralizzato e applicato solo nei punti di display.

- dati:
- basso, perché il canonico DB non cambia.

- export:
- medio-basso: oggi Derisking export usa mapping locale (`services/excel_export_service.py:491-502`); per evitare divergenze deve convergere sul mapping centrale (stesso dominio), senza cambiare formato export.

- search:
- medio UX: la ricerca Derisking opera su `supplier_status` raw (`services/dashboard_search_service.py:27-31`), quindi la visualizzazione EN può non coincidere con i termini ricercabili se non si allinea anche il criterio di ricerca (fuori scope di questo design).

---

## 9. Strategia di test manuale

1. Lingua EN, Derisking grid: verificare che `Status` mostri sempre label EN anche per record con canonico IT.
2. Lingua IT, Derisking grid: verificare che i valori restino corretti in IT.
3. Inserimento/modifica da dialog Derisking: verificare coerenza tra valore selezionato e valore mostrato in griglia dopo refresh.
4. Export Derisking IT/EN: verificare che non emergano divergenze tra visualizzazione e file export.
5. Smoke test RFQ/Saving/CA: verificare assenza regressioni su traduzioni e filtri esistenti.

---

## 10. Coerenza con refactor i18n 2026-04-10

- allineato: SI
- note:
- mantiene `tr(...)` come API ufficiale.
- preserva separazione canonico business vs label UI.
- evita branch lingua manuali.
- evita mapping duplicati locali, spostando il dominio in punto centrale riusabile.
