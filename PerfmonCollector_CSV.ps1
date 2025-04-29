param([int]$duration = 24)

# Süreyi hesapla
$endTime = (Get-Date).AddHours($duration)

# Script dosyasının çalıştığı dizini bul
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# CSV dosyası adını oluştur
$serverName = $env:COMPUTERNAME
$dateStamp = Get-Date -Format "yyyyMMdd_HHmm"
$csvFileName = "$serverName" + "_" + "$dateStamp.csv"
$csvFilePath = Join-Path $scriptDir $csvFileName

# İzlenecek performans sayaçları
$counters = @(
    '\Processor(_Total)\% Privileged Time',
    '\Processor(_Total)\% Processor Time',
    '\Processor(_Total)\% User Time',
    '\SQLServer:Buffer Manager\Buffer cache hit ratio',
    '\SQLServer:Buffer Manager\Page life expectancy',
    '\SQLServer:Databases(_Total)\Active Transactions',
    '\SQLServer:Databases(_Total)\Transactions/sec',
    '\SQLServer:Databases(_Total)\Write Transactions/sec',
    '\SQLServer:General Statistics\Connection Reset/sec',
    '\SQLServer:General Statistics\Logical Connections',
    '\SQLServer:General Statistics\Transactions',
    '\SQLServer:General Statistics\User Connections',
    '\SQLServer:Locks(_Total)\Number of Deadlocks/sec',
    '\SQLServer:Memory Manager\SQL Cache Memory (KB)',
    '\SQLServer:Memory Manager\Target Server Memory (KB)',
    '\SQLServer:Memory Manager\Total Server Memory (KB)'
)

$sampleInterval = 60  # saniye

# CSV başlığı oluşturulmamışsa ekle
try {
    if (!(Test-Path -Path $csvFilePath)) {
        "CounterName,InstanceName,CounterValue,CollectionDateTime" | Out-File -FilePath $csvFilePath -Encoding UTF8
    }
}
catch {
    Write-Output "CSV başlığı yazılamadı: $_"
    exit 1
}

# Süre dolana kadar veriyi topla
while ((Get-Date) -lt $endTime) {
    try {
        $performanceData = Get-Counter -Counter $counters -SampleInterval 1 -MaxSamples 1
        $counterSamples = $performanceData.CounterSamples

        foreach ($sample in $counterSamples) {
            $counterName = $sample.Path
            $instanceName = $sample.InstanceName
            $counterValue = $sample.CookedValue
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

            $csvRow = "$counterName,$instanceName,$counterValue,$timestamp"

            try {
                $csvRow | Out-File -FilePath $csvFilePath -Encoding UTF8 -Append
            }
            catch {
                Write-Output "Veri yazılamadı: $_"
            }
        }
    }
    catch {
        Write-Output "Sayaç hatası: $_"
    }

    Start-Sleep -Seconds $sampleInterval
}

Write-Output "`n✅ Veri toplama tamamlandı: $csvFilePath"
