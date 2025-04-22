param([int]$duration = 24)

# Süreyi hesapla
$endTime = (Get-Date).AddHours($duration)

# CSV dosyası adını oluştur
$serverName = $env:COMPUTERNAME
$dateStamp = Get-Date -Format "yyyyMMdd_HHmm"
$csvFilePath = "$serverName`_$dateStamp.csv"

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
if (!(Test-Path -Path $csvFilePath)) {
    "CounterName,InstanceName,CounterValue,CollectionDateTime" | Out-File -FilePath $csvFilePath -Encoding UTF8
}

# Süre dolana kadar veriyi topla
while ((Get-Date) -lt $endTime) {
    $performanceData = Get-Counter -Counter $counters -SampleInterval 1 -MaxSamples 1
    $counterSamples = $performanceData.CounterSamples

    foreach ($sample in $counterSamples) {
        $counterName = $sample.Path
        $instanceName = $sample.InstanceName
        $counterValue = $sample.CookedValue
        $timestamp = Get-Date

        $csvRow = "$counterName,$instanceName,$counterValue,$timestamp"
        Add-Content -Path $csvFilePath -Value $csvRow
    }

    Start-Sleep -Seconds $sampleInterval
}

Write-Output "Veri toplama tamamlandı: $csvFilePath"
