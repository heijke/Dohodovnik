Write-Progress -Activity 'Tvorba dohody' -Status 'Otevírání dokumentu...' -PercentComplete 0
#Variables
$Template = "C:\Users\lucie\Music\template.docx"
$Date = (Get-Date).ToString('dd.MM.yyyy')
$Details = "$pwd\details.txt"
$Entries = "$pwd\entries.txt"
$Name = ((Get-Content -Encoding UTF8 $Details | Select-Object -Skip 2 -First 1).Normalize("FormD") -replace '\p{M}', '').ToLower() –replace “ “,”_”

#Create new COM object
$Word = New-Object -ComObject Word.Application
#Open file for editing
try {$Dohoda = $Word.Documents.Open($Template)} catch {
Write-Error "Soubor je již otevřen jiným procesem, zavřete ho a spusťte skript znovu."
}
Write-Progress -Activity 'Tvorba dohody' -Status 'Dokument otevřen!' -PercentComplete 0

function Save-Fin {
  $Dohoda.SaveAs("$save_path.docx")
  Write-Host "Dokument uložen! K nalezení v: $save_path.docx"
  $Dohoda.SaveAs("$save_path.pdf",17)
  Write-Host "Dokument vyexportován do pdf! K nalezení v: $save_path.pdf"
}

function Save-Doc {
  param ([string]$save_path)
  if (!((Test-Path -LiteralPath "$save_path.docx") -or (Test-Path -LiteralPath "$save_path.pdf"))) {
  Save-Fin
  } else {
    Write-Warning "Dokument se stejným názvem již existuje, nelze uložit."
    $Overwrite = Read-Host "Chcete stávající dokumenty přepsat? [Y/N]"
    switch ($Overwrite) {
      Y { 
        Save-Fin; Break
      }
      Default {
        exit 1; Break
      }
    }
  }
}

function Fill-Details {
  param ([string]$details_path)
  try { $Details = Import-Csv -Encoding UTF8 -LiteralPath "$details_path" } catch {
  Write-Error "Soubor s osobmími údaji nenalezen v cestě : $details_path, zkontrolujte, zda existuje"
  }
  try {
  $i = 1
  $Details.ForEach({
    $Dohoda.Content.ContentControls.item($i).Range.Text = "$($_.Udaj)"
    $Dohoda.Content.ContentControls.item($i).Range.Font.Bold = $false
    $i++
  })
  Write-Progress -Activity 'Tvorba dohody' -Status  "Osobní údaje vyplněny!" -PercentComplete 35 
  } catch {
  Write-Error "Nebylo možné vyplnit osobní údaje."
  exit 1
  }
}

function Fill-Table {
  param ([string]$entries_path)
  try { $Entries = Import-Csv -Encoding UTF8 -LiteralPath $entries_path } catch {
  Write-Error "Soubor s položkami práce nenalezen v cestě : $entries_path, zkontrolujte, zda existuje"
  }
  $Tabulka = $Dohoda.Tables.item(1)
  $i = 2
  #Fill the table
  if($Entries.Count -le 1) {
    Write-Error "Alespoň dvě položky práce musejí být přítomny"
    exit 1
  }
  try {
  $Entries.ForEach({
    $Tabulka.Cell($i,1).Range.Text = "$($_.Datum)"
    $Tabulka.Cell($i,2).Range.Text = "$($_.Cinnost)"
    $Tabulka.Cell($i,3).Range.Text = "$($_.Hodiny)"
    $Tabulka.Cell($i,4).Range.Text = "$($_.Pozn)"
    $x++
  })
  Write-Progress -Activity 'Tvorba dohody' -Status "Položky práce vyplněny!" -PercentComplete 60
  } catch {
  Write-Error "Nebylo možné vyplnit položky práce."
  exit 1
  }
}

function Fill-Date {
  $Dohoda.Content.ContentControls.item(8).Range.Text = $Date
  Write-Progress -Activity 'Tvorba dohody' -Status "Dnešní datum vyplněno!" -PercentComplete 70
}

try {
  Fill-Details($Details)
  Fill-Table($Entries)
  Fill-Date
  Save-Doc("$pwd\${Name}_${Date}")
} catch {
  Write-Error "Něco se pokazilo, skript nemůže dále pokračovat :((("
  Write-Host $_
} finally {
  Write-Progress -Activity 'Tvorba dohody' -Status "Uvolňování paměti..." -PercentComplete 75
  $Dohoda.Close($false)
  $Word.Quit()
  Write-Progress -Activity 'Tvorba dohody' -Status "Uvolňování paměti..." -PercentComplete 85
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()
  Remove-Variable Template, Word, Dohoda, Date, save_path, details_path, i, x -ErrorAction SilentlyContinue
  Write-Progress -Activity 'Tvorba dohody' -Status "Paměť uvolněna!" -PercentComplete 100
}