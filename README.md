# NAV-kommunepuls

GitHub-klar overvåking av NAVs månedlige Excel-fil «Arbeidssøkere og ledige
stillinger. Kommune og kjennetegn» for:

- 5022 Rennebu
- 5029 Skaun
- 5055 Heim
- 5059 Orkland
- 5061 Rindal

## Hva den gjør

- kontrollerer NAVs kildeside hver virkedag kl. 08.30
- behandler bare en ny månedsfil én gang
- publiserer maksimalt ett nytt RSS-innlegg per måned
- viser helt ledige, ledighetsandel og nye stillinger for alle fem kommunene
- sammenligner antall helt ledige med forrige måned når tallene kan oppgis
- lenker til både NAVs statistikkside og den originale Excel-filen
- sender ingen historisk melding på første kjøring

NAV markerer små eller skjermede tall med `*`. Overvåkeren gjengir dette som
«skjermet av NAV» og forsøker ikke å beregne verdien.

## GitHub

Opprett et nytt offentlig repository og last opp innholdet i denne mappen.
Arbeidsflyten skal ligge på `.github/workflows/nav.yml`. Dersom `.github` blir
skjult ved opplasting, opprett filen manuelt i GitHub og kopier innholdet fra
`WORKFLOW.yml`.

Kjør arbeidsflyten manuelt første gang. Den første kjøringen oppretter bare en
grunnlinje. Senere kjøringer lager ett RSS-innlegg når NAV publiserer en ny
månedsfil.

RSS-adressen blir:

`https://raw.githubusercontent.com/BRUKERNAVN/REPOSITORY/main/public/feed.xml`

Erstatt `BRUKERNAVN` og `REPOSITORY` med GitHub-navnene dine.

## Datakilde

[NAV – Arbeidssøkere](https://www.nav.no/no/nav-og-samfunn/statistikk/arbeidssokere-og-stillinger-statistikk/Annen%20Statistikk%20om%20arbeidsmarkedet)

NAV opplyser om brudd i arbeidssøkerstatistikken fra april 2025. Bruk derfor
lange tidsserier og årssammenligninger med varsomhet.
