# REPORT VENDITE AUTOMATIZZATO PER SETTORE E MULTI-STORE
---
## Descrizione
Automazione completa di un report Excel e PDF per il monitoraggio delle vendite nei punti vendita, suddivise per settore merceologico a partire da esportazione del gestionale aziendale.  
Lo script consente di analizzare in modo dettagliato le performance di ciascun reparto all’interno di ogni negozio, offrendo una visione chiara e sintetica di:

 Quantità vendute e giacenze  
 Ricavi e margine lordo (MLVE)  
 Subtotali per settore e totale complessivo  
 Confronto diretto tra punti vendita  

---

![ANTEPRIMA_REPORT](https://github.com/carchedimarco88-jpg/REPORT-VENDITE-PER-SETTORE-E-MULTI-STORE/raw/main/Immagine%20Esempio%20Report.png)

---

##  Funzionalità principali

-  Importazione automatica da Excel  
-  Pivot dinamico: ogni punto vendita diventa un set di colonne  
-  Calcolo MLVE per riga, subtotale e totale (senza sommare percentuali)  
-  Subtotali per Settore1 con evidenziazione visiva  
-  Totale complessivo con MLVE corretti per ogni PV  
-  Esportazione su template Excel con formattazione professionale  
-  Generazione automatica del PDF via Excel COM  

---

## 🛠 Requisiti

- R (≥ 4.0)  
- Pacchetti:
  - `dplyr`
  - `tidyr`
  - `openxlsx`
  - `RDCOMClient` (solo Windows)

---
## Esempio di utilizzo
Immagina di avere 5 punti vendita e 12 settori. Con questo script puoi generare un report che mostra:

| Settore1               | Tot_Venduti | Tot_ML%VE | STORE 1 Venduti | STORE 1 ML%VE | STORE 2 Venduti | STORE 2 ML%VE | STORE 3 Venduti | STORE 3 ML%VE |
|------------------------|-------------|----------|-----------------|--------------|---------------|------------|-----------------|--------------|
| DERMATOLOGIA           | 1.240       |    32 %    |   540             | 28 %        | 420           | 35 %      | 280              | 34 %          |
| Subtotale DERMATOLOGIA | 1.240       |    32 %    |   540             | 28 %       | 420           | 35  %     | 280              | 34 %         |
| INTEGRATORI            | 980         |    41 %    |   300             | 38 %        | 400           | 43 %      | 280              | 42 %          |
| Subtotale INTEGRATORI  | 980         |    41 %    |   300             | 38 %        | 400           | 43 %      | 280              | 42 %          |
| OMEOPATIA              | 620         |    27 %    |   200             | 25 %        | 240           | 29 %      | 180              | 26 %          |
| Subtotale OMEOPATIA    | 620         |    27 %    |   200             | 25 %        | 240           | 29 %      | 180              | 26 %          |
| **Totale complessivo** | **2.840**   | **34 %** | **1.040**       | **30 %**     | **1.060**     | **36 %**   | **740**          | **34 %**       |


Riduce il tempo di preparazione report da ore a secondi

Evita errori manuali nel calcolo del margine

Presenta i dati in modo leggibile e pronto per la stampa

Permette analisi trasversali: per settore, per PV, per MLVE


📌 NOTA BENE

Per utilizzare lo script è necessario sostituire al suo interno il percorso dei file di Input e del Template nei punti indicati dai commenti
