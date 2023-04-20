# MasterBricklink
<b><i>MasterBricklink, AFOL Tools for Bricklink on Google SpreadSheet</b></i>, nasce dalla mente di <b>GianCann</b> (tra il 2018 e il 2019) per semplificare la gestione dell'inventario di mattoncini LEGO. Dopo aver ereditato il progetto, ho migliorato le funzionalità esistenti e ne ho introdotte di nuove, cercando di rendere la cartella di lavoro il più stand-alone possibile e fruibile anche ai meno pratici.<br></br>

## Installazione
La procedura di "installazione", o meglio di "generazione" del foglio di lavoro richiede il completamento di pochi passi:
<ol>Creare una cartella di lavoro in Google Sheets;</ol>
<ol>Importare (per esempio copiando) i vari script utilizzando Editor di Script/App Script;</ol>
<ol>Dopo aver ricaricato il foglio, nel menù Bricklink Tool, in Regenerate, scegliere Regenerate Settings;</ol>
<ol>Nel foglio Settings inserire le chiavi per le API Bricklink e per le API TurboBricksManager;</ol>
<ol>Proseguire con la generazione degli altri foglio procedendo dall'alto verso il basso sempre nel menù Bricklink Tool, Regenerate.</ol>

## To Do List
Come ogni progetto che si rispetti <b><i>sono ancora molte le idee da realizzare</b></i>:
<li>Error Handling e miglioramenti della user experience;</li>
<li>Multi Part-Out di Sets e relativi filtri;</li>
<li>Multipiattaforma: sincronizzazione con Brickowl, Rebrickable;</li>
<li>Introduzione di auto Export verso Bricklink, senza bisogno di XML;</li>
<li>Introduzione di Backup/Ripristino per migrazioni sicure dei dati tra le versioni;</li>
<li>Miglioramento generale degli script per ottimizzare le performance e la comprensibilità;</li>
<li>Scrivere qualche riga per spiegare le varie funzionalità della cartella di lavoro!</li>

## Changelog
### da 1.0.0 ad oggi (GitHub)
v1.2.0: Introdotto aggiornamento automatico dei Database Parts, Minifigures e Sets grazie a TurboBricksManager. (ZioTitanok)<br>
v1.1.3: Introdotto in Lab funzione per il suggerimento e l'aggiornamento dei prezzi. (ZioTitanok) <br>
v1.1.2: Aggiornamento minore di Lab. (ZioTitanok)<br>
v1.1.1: Miglioramento delle performance degli Import. Aggiornamento minore di PartOut. (ZioTitanok)<br>
v1.1.0: Miglioramento delle performance nella generazione degli XML e nel download delle Price Guide. (ZioTitanok)<br>
v1.0.1: ReadMe e Minor Fixes. (ZioTitanok)<br>
v1.0.0: Introdotte funzionalità di creazione/ripristino dei fogli della cartella di lavoro. (ZioTitanok)<br>

### da 0.0.1 a 0.9.0 (pre-GitHub)
v0.9.0: Aggiornamento automatico dei Database Categorie e Colori. (ZioTitanok)<br>
v0.8.0: Introdotti i filtri in Inventory, PartOut e Lab. (ZioTitanok)<br>
v0.7.0: Introduzione dei Settings e di altre funzioni minori per al user experience. (ZioTitanok)<br>
v0.6.0: Non solo Parts: Minifigures, Sets e tutto il resto possono essere gestiti. (ZioTitanok)<br>
v0.5.0: Introdotte le funzionalità di Import tra Inventory, PartOut e Lab. (ZioTitanok)<br>
v0.4.0: Introdotto PartOut per il download dei part-out dei Sets. (ZioTitanok)<br>
v0.3.1: Introdotto XML Export (Upload/Upgrade) per la sincronizzazione manuale dell'Inventario su Bricklink. (ZioTitanok)<br>
v0.3.0: Introdotto XML Export (Wanted) per la creazione manuale di WantedList su Bricklink (GianCann)<br>
v0.2.0: Miglioramento di Lab ed introduzione dei Database di Parts, Minifigures e Sets. (ZioTitanok)<br>
v0.1.0: Introdotto OAuth1 negli script (eliminato il PHP esterno). (GianCann)<br>
v0.0.2: Introdotto Inventory con download dell'Inventory da Bricklink (via PHP esterno). (GianCann)<br>
v0.0.1: Introdotto Lab con download delle Price Guide delle Parts (via PHP esterno). (GianCann)<br>

## Dedica
Alla memoria di GianCann, che sicuramente avrebbe fatto un lavoro migliore del mio nel portare avanti questo progetto.