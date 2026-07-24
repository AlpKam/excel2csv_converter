#Requires -Module ImportExcel

# Excel dosyasýný belirtilen aralýktan okuyup, seçilen verileri UTF-8 kodlu CSV olarak dýþa aktarýr.

# ImportExcel modülü yüklü mü kontrol et
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "`n'ImportExcel' modülü yüklü deðil." -ForegroundColor Yellow
    Write-Host "Kurmak için þu komutu çalýþtýrýn:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name ImportExcel -Scope CurrentUser -Force" -ForegroundColor Cyan
    exit
}

# Kullanýcýdan giriþ dosyasý ve çýkýþ dosyasý yollarýný al
$inputPath = Read-Host -Prompt "Excel dosyasýnýn tam yolunu girin (.xlsx)"
$outputPath = Read-Host -Prompt "Oluþturulacak CSV dosyasýnýn tam yolunu girin (.csv ya da klasör)"

# Giriþ ve çýkýþ yollarýndaki çift týrnaklarý temizle
$inputPath = $inputPath.Trim('"')
$outputPath = $outputPath.Trim('"')

# Eðer çýktý yolu bir klasörse, dosya adý olarak giriþ dosyasýnýn adý kullanýlýr
if (Test-Path -Path $outputPath -PathType Container) {
    $inputFileName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
    $outputPath = Join-Path -Path $outputPath -ChildPath "$inputFileName.csv"
}

# Sütun baþlýklarý
$headers = @(
    "Maðaza Kodu", "Maðaza Adý", "Bölge", "Format", "Unix Name",
    "Vlan41 IP", "Toplam R10 Kasa", "Normal Kasa", "Jet Kasa",
    "Cafe Kasa", "DK Kasa", "Store Server", "Rap Station"
)

# Excel verisini 3. satýrdan itibaren, A-M sütunlarý (1-13) arasýnda al (-NoHeader ile baþlýksýz)
$data = Import-Excel -Path $inputPath -StartRow 3 -StartColumn 1 -EndColumn 13 -NoHeader

# Her satýr için baþlýklarla eþleþtirilmiþ özel nesne oluþtur
$selectedData = $data | ForEach-Object {
    $properties = [ordered]@{}
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $properties[$headers[$i]] = $_."P$($i + 1)"
    }
    [pscustomobject]$properties
}

# CSV olarak dýþa aktar (UTF-8)
try {
    $selectedData | Export-Csv -Path $outputPath -Encoding UTF8 -NoTypeInformation -Force
    Write-Host "`nVeriler baþarýyla dýþa aktarýldý:" -ForegroundColor Green
    Write-Host $outputPath -ForegroundColor Cyan
} catch {
    Write-Host "`nCSV dýþa aktarma baþarýsýz oldu! Hata:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
