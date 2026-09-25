# Riello - XLSX z FTP -> CSV EUR/PLN -> FTP

Workflow uruchamia sie codziennie o 13:00 w strefie `Europe/Warsaw` oraz recznie przez `workflow_dispatch`.

## Sekrety GitHub

Utworz w `Settings -> Secrets and variables -> Actions -> New repository secret`:

- `FTP_HOST` - host FTP, bez `ftp://`
- `FTP_PORT` - zwykle `21` (opcjonalnie, ale workflow przekazuje sekret)
- `FTP_USER` - login FTP
- `FTP_PASSWORD` - haslo FTP
- `FTP_USE_TLS` - `true` dla FTPS explicit albo `false` dla zwyklego FTP
- `RIELLO_1PH_FTP_PATH` - np. `/riello/riello_1ph.xlsx`
- `RIELLO_3PH_FTP_PATH` - np. `/riello/riello_3ph.xlsx`
- `RIELLO_OUTPUT_FTP_PATH` - np. `/riello/riello_prices.csv`

## Wynik

CSV UTF-8, separator `;`:

```csv
Code;PriceEURO;PricePLN
AIDG1K21RU;148.83;653.36
```

`PricePLN = PriceEURO * sredni kurs EUR/PLN z tabeli A NBP`, zaokraglenie do 2 miejsc metodą `ROUND_HALF_UP`.

Przy kazdym uruchomieniu wynikowy CSV na FTP jest zastepowany nowa wersja. Zrodłowe XLSX nie sa modyfikowane.
