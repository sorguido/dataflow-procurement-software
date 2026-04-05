#!/usr/bin/env python3
"""
Script per generare un database di test completo per DataFlow — versione ITALIANA.

Genera lo stesso dataset di generate_test_db.py ma con:
  - Database di output: test_dataflow_full_it.db
  - Campi category e notes dei fornitori potenziali in italiano

NON sovrascrive DB reali né il database EN (test_dataflow_full.db).
"""

import generate_test_db as _base

# ---------------------------------------------------------------------------
# OVERRIDE: percorso database italiano
# ---------------------------------------------------------------------------
_base.DB_PATH = 'test_dataflow_full_it.db'

# ---------------------------------------------------------------------------
# OVERRIDE: fornitori potenziali con category e notes in italiano
# ---------------------------------------------------------------------------
_base.POTENTIAL_SUPPLIERS_DATA = [
    (
        'Acciai Speciali Valpadana SRL', 'Acciaio', 'Qualificato',
        'Marco Ferretti', 'm.ferretti@acciaivalpadana.it', '+39 0444 512300',
        'www.acciaivalpadana.it',
        'Secondo fornitore qualificato per lamiere S355. Audit superato 2024-11. Certificazione EN 10025.',
        '2024-11-15', '2024-11-15T10:30:00',
    ),
    (
        'Eurosteel Componenti SpA', 'Acciaio', 'In valutazione',
        'Laura Bianchi', 'l.bianchi@eurosteel.eu', '+39 030 7742100',
        'www.eurosteel.eu',
        'Fornitore alternativo per barre trafilate C45. In fase di qualifica campioni. Offerta ricevuta 2025-02.',
        '2025-02-10', '2025-02-10T09:00:00',
    ),
    (
        'MetalTech Brescia SRL', 'Lavorazioni meccaniche', 'Qualificato',
        'Giovanni Rossi', 'g.rossi@metaltech-bs.it', '+39 030 3451200',
        'www.metaltech-bs.it',
        'Terzista qualificato per tornitura CNC su SS316L. Capacità produttiva validata. Lead time 10gg.',
        '2024-09-20', '2024-09-20T14:00:00',
    ),
    (
        'Nordic Steel Components AB', 'Acciaio', 'Nuovo',
        'Erik Lindstrom', 'e.lindstrom@nordicsteel.se', '+46 31 7001200',
        'www.nordicsteel.se',
        'Fornitore svedese individuato come alternativa a fornitore est-Europa. Prima visita commerciale 2025-03.',
        '2025-03-05', '2025-03-05T11:00:00',
    ),
    (
        'Trattamenti Termici Padova SpA', 'Trattamenti termici', 'Qualificato',
        'Stefano Meneghetti', 's.meneghetti@ttpd.it', '+39 049 8823400',
        'www.ttpd.it',
        'Backup qualificato per bonifica C45 e cementazione. Accordo stand-by attivo. Capacità 2 t/giorno.',
        '2024-06-01', '2024-06-01T08:30:00',
    ),
    (
        'Galvanica Lombarda SRL', 'Trattamenti superficiali', 'In valutazione',
        'Paola Carminati', 'p.carminati@galvanicalombarda.it', '+39 02 9290500',
        'www.galvanicalombarda.it',
        'Secondo fornitore per zincatura elettrolitica e nichelatura. Campioni inviati 2025-01. Attesa rapporto qualità.',
        '2025-01-20', '2025-01-20T10:00:00',
    ),
    (
        'Verniciatura Industriale Veneta SRL', 'Trattamenti superficiali', 'Qualificato',
        'Antonio Zilio', 'a.zilio@viv-srl.it', '+39 0422 631800',
        'www.viv-srl.it',
        'Fornitore backup per verniciatura epossidica RAL 7035. Omologazione completata 2023-10. Accordo quadro annuale.',
        '2023-10-12', '2023-10-12T15:00:00',
    ),
    (
        'Fonderia Bresciana SpA', 'Fusioni', 'In valutazione',
        'Franco Pezzotti', 'f.pezzotti@fonderiabresciana.it', '+39 030 9821000',
        'www.fonderiabresciana.it',
        'Alternativa italiana per fusioni ghisa GJL-250. Prima campionatura in corso. Prevista entro 2025-05.',
        '2025-02-28', '2025-02-28T09:30:00',
    ),
    (
        'CNC Precision Parts GmbH', 'Lavorazioni meccaniche', 'Qualificato',
        'Klaus Weber', 'k.weber@cnc-precision.de', '+49 711 4503200',
        'www.cnc-precision.de',
        'Partner tedesco per fresatura 4140 e rettifica di precisione. ISO 9001:2015. Lead time 15gg. Backup attivo da 2023.',
        '2023-07-18', '2023-07-18T13:00:00',
    ),
    (
        'Lavorazioni Meccaniche Bergamo SRL', 'Lavorazioni meccaniche', 'Scartato',
        'Roberto Cattaneo', 'r.cattaneo@lmbergamo.it', '+39 035 3310800',
        '',
        'Non superato audit qualità 2024-04. Tolleranze fuori specifica su campioni torniti. Rivalutare dopo azioni correttive.',
        '2024-04-10', '2024-04-10T16:00:00',
    ),
    (
        'Stampaggio Metalli Emilia SRL', 'Stampaggio', 'Qualificato',
        'Elena Monti', 'e.monti@sme-srl.it', '+39 059 8812200',
        'www.sme-srl.it',
        'Secondo fornitore per stampaggio alluminio EN-AC-46000. Re-source completato 2024-03. Volume annuo garantito 4.000 pz.',
        '2024-03-22', '2024-03-22T10:30:00',
    ),
    (
        'Meccanica di Piacenza SRL', 'Lavorazioni meccaniche', 'In valutazione',
        'Andrea Ferri', 'a.ferri@meccanicapc.it', '+39 0523 601500',
        'www.meccanicapc.it',
        'Fornitore locale per riduzione lead time su particolari critici. Visita stabilimento 2025-03-18. Capacità CNC adeguata.',
        '2025-03-18', '2025-03-18T11:00:00',
    ),
    (
        'Forgiatura Toscana SpA', 'Forgiatura', 'Nuovo',
        'Massimo Landi', 'm.landi@forgiaturatoscana.it', '+39 055 8923100',
        'www.forgiaturatoscana.it',
        'Identificato come potenziale alternativa per flangiature forgiate PN16. Richiesta documentazione tecnica 2025-04.',
        '2025-04-02', '2025-04-02T09:00:00',
    ),
    (
        'Officine Guidetti SRL', 'Lavorazioni meccaniche', 'Qualificato',
        'Luca Guidetti', 'l.guidetti@officineguidetti.it', '+39 0372 421600',
        'www.officineguidetti.it',
        'Terzista per rettifica di precisione SK3. Backup attivo su componenti critici. Certificato EN ISO 9001.',
        '2023-11-05', '2023-11-05T14:00:00',
    ),
    (
        'Costruzioni Meccaniche Piemonte SRL', 'Strutture saldate', 'In valutazione',
        'Davide Gallo', 'd.gallo@cmp-srl.it', '+39 011 9043200',
        'www.cmp-srl.it',
        'Secondo fornitore per telai saldati e strutture EN1090. Qualifica saldatori in corso. Prevista entro 2025-06.',
        '2025-01-30', '2025-01-30T10:00:00',
    ),
    (
        'Catene Industriali Veneto SRL', 'Trasmissioni meccaniche', 'Qualificato',
        'Mirko Trevisan', 'm.trevisan@cateneveneto.it', '+39 049 9213400',
        'www.cateneveneto.it',
        'Secondo distributore per catene simplex ISO 606 08B. Accordo multi-sourcing attivo 2024. Stock locale 200 m.',
        '2024-05-14', '2024-05-14T09:00:00',
    ),
    (
        'Ingranaggi Precisi Lombardia SRL', 'Trasmissioni meccaniche', 'Nuovo',
        'Chiara Colombo', 'c.colombo@ingranaggeriprecisi.it', '+39 02 4001500',
        '',
        'Contatto iniziale per ingranaggi cilindrici modulo 3. Richiesta preventivo spedita 2025-03-28.',
        '2025-03-28', '2025-03-28T15:30:00',
    ),
    (
        'Guarnizioni & Tenute SRL', 'Gomma e tenute', 'Qualificato',
        'Fabio Negri', 'f.negri@guarnizionientenute.it', '+39 035 4124500',
        'www.guarnizionientenute.it',
        'Secondo fornitore qualificato per guarnizioni NBR e FKM. Campioni approvati 2024-08. Accordo annuale firmato.',
        '2024-08-05', '2024-08-05T11:00:00',
    ),
    (
        'Cuscinetti Nord Italia Srl', 'Cuscinetti e trasmissioni', 'Qualificato',
        'Simona Riva', 's.riva@cni-srl.it', '+39 02 6682100',
        'www.cni-srl.it',
        'Secondo distributore autorizzato SKF per cuscinetti 6205-2RS e a rulli. Listino SKF fisso 2025. Magazzino Milano.',
        '2024-10-01', '2024-10-01T08:30:00',
    ),
    (
        'Profilati Strutturali Sud SpA', 'Acciaio', 'Scartato',
        'Giuseppe Marino', 'g.marino@profilatisudspa.it', '+39 081 7530900',
        'www.profilatisudspa.it',
        'Scartato per lead time eccessivo (60gg vs 20gg richiesti) e assenza certificazioni EN 10210. Da riconsiderare solo se lead time migliorato.',
        '2024-07-22', '2024-07-22T15:00:00',
    ),
]

if __name__ == '__main__':
    _base.create_test_database()
