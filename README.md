# X-Helg geocaching adventskalender statistikk
Statistikk for geocachings-adventskalender. Laget for X-Helg på Helgeland, men kan sikkert også brukes av andre. Første iterasjon skrevet i PowerShell.

# Bruk
Skriptet er skrevet i PowerShell, og benytter SQLite for .NET bibliotek for å aksessere en SQLite-database. Databasen er laget i GSAK.

Man oppretter databasen i GSAK, og legger inn alle cacher i kalenderen. Deretter laster man ned alle logger. Med dette er databasen klar for bruk.

Skriptet har en seksjon øverst som gjør en dataprep:

1. Lager poeng-tabell (med drop if exists), legger inn alle som har logget hittil.
2. Lager en FTF-tabell (med drop if exists), finner og legger til alle FTF'er. Dette er basert på at loggen inneholder den "anerkjente" FTF-koden {\*FTF\*}.
3. Gjør en update for hver cache i kalenderen; endrer placeddate til aktuell dag (1-24 des) og setter User4 feltet til et kortnavn [#n] der n er aktuell dag, 1-24.

Deretter henter skriptet ut logger og tildeler poeng etter følgende kriterier:
- Funn i Desember gir 1 poeng.
- Funn på publiseringsdato gir 2 poeng.
- FTF gir 3 poeng.
- "Co-FTF", dvs flere logger sammen, gir 2 poeng til hver.
- Egen publisering gir 3 poeng.

Til slutt lagres data i en CSV-fil. Man kan så benytte Excel til å lage en "pen" rapport, se vedlagt eksempel.

TODO:
- Vil ha en HTML rapport som kan publiseres rett på nettside.
- Vil ha en dynamisk, sorterbar rapport, med link på hver cache og cacher.
- Vil se om Project GC har noe som kan benyttes.
