# MasterBricklink
<b><i>MasterBricklink, AFOL Tools for Bricklink on Google SpreadSheet</b></i>, nasce dalla mente di <b>GianCann</b> (tra il 2018 e il 2019) per semplificare la gestione dell'inventario di mattoncini LEGO. Dopo aver ereditato il progetto, ho migliorato le funzionalità esistenti e ne ho introdotte di nuove, cercando di rendere la cartella di lavoro il più stand-alone possibile e fruibile anche ai meno pratici.<br></br>


## Installazione
Per iniziare, in una cartella di lavoro di Google Sheets, è necessario importare i vari script utilizzando "Strumenti--> Editor di Script".
Dopo aver importato manualmente i Database di Parts, Minifigures e Sets (scaricabili da Bricklink) si procede in modo automatizzato. Generato Settings, correttamente compliato con i dati necessari per le API Bricklink, si possono creare i Database di Categorie e Colori. E' ora possibile generare tutti gli altri fogli della cartella di lavoro.


## To Do List
Come ogni progetto che si rispetti <b><i>sono ancora molte le idee da realizzare</b></i>:
<li>Automatizzare l'aggiornamento dei Database di Parts, Minifigures e Sets (ad oggi, manuali);</li>
<li>Error Handling e miglioramenti della user experience;</li>
<li>Multi Part-Out di Sets e relativi filtri;</li>
<li>Multipiattaforma: sincronizzazione con Brickowl, Rebrickable;</li>
<li>Introduzione di auto Export verso Bricklink, senza bisogno di XML;</li>
<li>Introduzione di Backup/Ripristino per migrazioni sicure dei dati tra le versioni;</li>
<li>Miglioramento generale degli script per ottimizzare le performance e la comprensibilità;</li>
<li>Scrivere qualche riga per spiegare le varie funzionalità della cartella di lavoro!</li>


## Changelog
### da 1.0.0 ad oggi (GitHub)
1.0.1. ReadMe. (ZioTitanok)<br>
1.0.0: Introdotte funzionalità di creazione/ripristino dei fogli della cartella di lavoro. (ZioTitanok)<br>

### da 0.0.1 a 0.9.0 (pre-GitHub)
0.9.0: Aggiornamento automatico dei Database Categorie e Colori. (ZioTitanok)<br>
0.8.0: Introdotti i filtri in Inventory, PartOut e Lab. (ZioTitanok)<br>
0.7.0: Introduzione dei Settings e di altre funzioni minori per al user experience. (ZioTitanok)<br>
0.6.0: Non solo Parts: Minifigures, Sets e tutto il resto possono essere gestiti. (ZioTitanok)<br>
0.5.0: Introdotte le funzionalità di Import tra Inventory, PartOut e Lab. (ZioTitanok)<br>
0.4.0: Introdotto PartOut per il download dei part-out dei Sets. (ZioTitanok)<br>
0.3.1: Introdotto XML Export (Upload, Upgrade) per la sincronizzazione manuale dell'Inventario su Bricklink. (ZioTitanok)<br>
0.3.0: Introdotto XML Export (Wanted) per la creazione manuale di WantedList su Bricklink (GianCann)<br>
0.2.0: Miglioramento di Lab ed introduzione dei Database di Parts, Minifigures e Sets. (ZioTitanok)<br>
0.1.0: Introdotto OAuth1 negli script (eliminato il PHP esterno). (GianCann)<br>
0.0.2: Introdotto Inventory con download dell'Inventary da Bricklink (via PHP esterno). (GianCann)<br>
0.0.1: Introdotto Lab con download delle Price Guide delle Parts (via PHP esterno). (GianCann)<br>

## Dedica
Alla memoria di GianCann, che sicuramente avrebbe fatto un lavoro migliore del mio nel portare avanti questo progetto.
