## Dohodovník

*Powershell skript pro tvorbu výkazů práce pro dohody o provedení práce pro SSPŠ*

Naprosto neužitečné a zbytečné, ale ušetří pár kliknutí :)

## Užití

Jelikož je skript nepodepsaný, před spuštěním je nutné nastavit vhodnou `ExecutionPolicy` pomocí:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### template.docx

Obsahuje předlohu dohody s přidanými Content Controls, které umožňují dokument vyplnit.

### script-gui.ps1

Užití samotného skriptu je prosté, v naklonovaném adresáři s předlohou dohody a konfigurací spusťte `.\script-gui.ps1`: 

```powershell
PS C:\Users\Jenda\Dohodovnik>.\script-gui.ps1
```

