Dette skriptet behandler flere Excel-filer ved å utføre spesifikke datamanipulasjoner og lagrer resultatet i en ny Excel-fil. Skriptet er designet for å være brukervennlig og bruker GUI-dialoger for filvalg og lagringssted ved hjelp av Tkinter-biblioteket.

Forutsetninger
For å kjøre dette skriptet trenger du Python 3.x installert på systemet ditt sammen med følgende Python-pakker: pandas og tkinter.

Funksjonalitet
Skriptet utfører følgende operasjoner:

Filvalg for Inndatafiler: Skriptet bruker en filvelger-dialog for å la brukeren velge flere Excel-filer (.xlsx).

Filvalg for Utdatafil: Skriptet bruker en filvelger-dialog for å la brukeren velge lagringssted for den nye Excel-filen.

Les og prosesser hver Excel-fil:

Leser arket "Adjustments" fra hver valgt fil.
Konverterer "DATE"-kolonnen til formatet 'dd.mm.yyyy'.
Tar den absolutte verdien av "AMOUNT"-kolonnen.
Lagrer nødvendige kolonner i en liste.
Kombiner og grupper data:

Kombinerer data fra alle filer til en enkelt DataFrame.
Grupperer data etter "DATE" og "MERCHANT ID", og summerer "AMOUNT".
Formater og lagre data:

Omdøper kolonner til ønsket format og rekkefølge.
Lagre den oppsummerte DataFrame til en ny Excel-fil.
Suksessmelding: En melding skrives til konsollen når filen er lagret.

Bruk
Kjør Skriptet: Kjør skriptet i et Python-miljø.
Velg Inndatafiler: En filvelger-dialog vil dukke opp. Velg de Excel-filene (.xlsx) du ønsker å behandle.
Velg Utdatafil: En filvelger-dialog vil dukke opp. Velg hvor du vil lagre den nye Excel-filen.
Fullføring: En melding i konsollen bekrefter at filen er lagret.
