Clear-Host

# Caminho do arquivo HTML de saída com timestamp
$outputFile = "C:\diagnostico_avancado_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

# CSS para o relatorio
$style = @"
<style>
  body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background-color: #f4f4f4;
    margin: 20px;
    color: #333;
  }
  h1 {
    color: #004085;
    background-color: #cce5ff;
    padding: 15px;
    border-radius: 8px;
    text-align: center;
  }
  h2 {
    color: #155724;
    background-color: #d4edda;
    padding: 10px;
    border-radius: 8px;
    margin-top: 20px;
  }
  table {
    border-collapse: collapse;
    width: 100%;
    background-color: #ffffff;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    margin-bottom: 20px;
    border-radius: 8px;
    overflow: hidden;
  }
  th {
    background-color: #28a745;
    color: #fff;
    padding: 10px;
  }
  td {
    border: 1px solid #ddd;
    padding: 10px;
  }
  tr:nth-child(even) {
    background-color: #f9f9f9;
  }
  tr:hover {
    background-color: #f1f1f1;
  }
  pre {
    background-color: #f8f9fa;
    padding: 15px;
    border-left: 4px solid #007bff;
    overflow-x: auto;
    font-family: Consolas, 'Courier New', monospace;
    border-radius: 8px;
  }
</style>
"@

Write-Host "`n========================================================================" -ForegroundColor DarkCyan
Write-Host "           DIAGNOSTICO AVANCADO DE REDE - POWERSHELL" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor DarkCyan

# Funcao para capturar o output em texto formatado
function Get-CommandOutputAsString {
    param ($ScriptBlock)
    $sbOutput = & $ScriptBlock 2>&1 | Out-String
    return $sbOutput.Trim()
}

# 1) Teste de conexao (ping)
$pingResult = Get-CommandOutputAsString { Test-Connection -ComputerName google.com -Count 4 }

# 2) Resolucao DNS
$dnsResult = Get-CommandOutputAsString { Resolve-DnsName -Name google.com }

# 3) IPs das interfaces de rede
$ipAddresses = Get-NetIPAddress | Select-Object IPAddress, InterfaceAlias, AddressFamily, PrefixLength, ValidLifetime

# 4) Placas de rede e status
$netAdapters = Get-NetAdapter | Select-Object Name, Status, MacAddress, LinkSpeed

# 5) Teste de porta
$portTest = Get-CommandOutputAsString { Test-NetConnection -ComputerName 8.8.8.8 -Port 53 }

# 6) Conexoes TCP ativas
$tcpConnections = Get-NetTCPConnection | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess

# 7) Servidores DNS configurados
$dnsServers = Get-DnsClientServerAddress | ForEach-Object {
    [PSCustomObject]@{
        InterfaceAlias = $_.InterfaceAlias
        ServerAddresses = ($_.ServerAddresses -join ", ")
    }
}

# 8) Regras de Firewall habilitadas
$firewallRules = Get-NetFirewallRule -Enabled True |
    Select-Object DisplayName, Direction, Action, Profile |
    Sort-Object Direction

# 9) Reiniciando adaptador de rede "Ethernet"
Restart-NetAdapter -Name "Ethernet" -Confirm:$false

# 10) Limpar cache DNS
Clear-DnsClientCache

# Construindo o HTML do relatorio
$htmlReport = @"
<html>
<head>
<meta charset='UTF-8'>
<title>Relatorio Diagnostico Avancado de Rede</title>
$style
</head>
<body>
<h1>Relatorio Diagnostico Avancado de Rede</h1>

<h2>1) Teste de Conexao (Ping para google.com)</h2>
<pre>$pingResult</pre>

<h2>2) Consulta DNS para google.com</h2>
<pre>$dnsResult</pre>

<h2>3) Enderecos IP nas Interfaces de Rede</h2>
$($ipAddresses | ConvertTo-Html -Fragment -Property IPAddress, InterfaceAlias, AddressFamily, PrefixLength, ValidLifetime)

<h2>4) Adaptadores de Rede e Status</h2>
$($netAdapters | ConvertTo-Html -Fragment -Property Name, Status, MacAddress, LinkSpeed)

<h2>5) Teste de conexao com IP e Porta (8.8.8.8:53)</h2>
<pre>$portTest</pre>

<h2>6) Conexoes TCP Ativas</h2>
$($tcpConnections | ConvertTo-Html -Fragment -Property LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess)

<h2>7) Servidores DNS Configurados nas Interfaces</h2>
$($dnsServers | ConvertTo-Html -Fragment -Property InterfaceAlias, ServerAddresses)

<h2>8) Regras de Firewall Ativas</h2>
$($firewallRules | ConvertTo-Html -Fragment -Property DisplayName, Direction, Action, Profile)

<h2>9) Adaptador de Rede Reiniciado</h2>
<pre>O adaptador de rede 'Ethernet' foi reiniciado com sucesso.</pre>

<h2>10) Cache DNS Limpo</h2>
<pre>O cache DNS foi limpo com sucesso.</pre>

</body>
</html>
"@

# Salva o relatorio em arquivo com UTF8
$htmlReport | Out-File $outputFile -Encoding utf8

Write-Host "`n========================================================================" -ForegroundColor DarkCyan
Write-Host " Diagnostico concluido." -ForegroundColor Yellow
Write-Host " Relatorio HTML gerado em: $outputFile" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor DarkCyan
