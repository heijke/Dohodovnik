## Dohodovník

*Powershell skript pro tvorbu výkazů práce pro dohody o provedení práce pro SSPŠ*

Naprosto neužitečné a zbytečné, ale ušetří pár kliknutí :)

## Užití

Jelikož je skript nepodepsaný, před spuštěním je nutné nastavit vhodnou `ExecutionPolicy` pomocí:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Následně v naklonovaném adresáři upravit stávající, či vytvořit nové, konfigurační soubory (csv)- **details.txt** a **entries.txt**.

### details.txt

Obsahuje osobní údaje vyplňované do předlohy dohody ve formátu:

```
Udaj
Název dohody
Jméno, příjmení zaměstnance
Rodné číslo
Č. účtu/kód banky
Místo narození
Adresa zaměstnance
Zdravotní pojišťovna — název a číslo
```

kde hlavička "Udaj" musí být ponechána.

### entries.txt

Obsahuje položky práce vyplňované do předlohy dohody (min. 2, max. 16) ve formátu:

```
Datum,Cinnost,Hodiny,Pozn
8.9.2021,Vzorová činnost 1,4h
11.9.2021,Vzorová činnost 2,5h,poznámka 
12.9.2021,Vzorová činnost 3,2h
```

kde hlavička "Datum,Cinnost,Hodiny,Pozn" musí být ponechána.

### template.docx

Obsahuje předlohu dohody s přidanými Content Controls, které umožňují dokument vyplnit.

### script.ps1

Užití samotného skriptu je prosté, v naklonovaném adresáři s předlohou dohody a konfigurací spusťte `.\script.ps1`: 

```powershell
PS C:\Users\Jenda\Dohodovnik>.\script.ps1
```

