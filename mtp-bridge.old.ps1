# Ponte entre o Node e o celular (MTP) via Shell.Application.
# Lê comandos JSON (um por linha) do stdin e responde um JSON por linha no stdout.
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Text.Encoding]::UTF8
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$shell = New-Object -ComObject Shell.Application
$script:folders = @{}   # subpasta -> objeto Folder do Shell
$script:items = @{}     # "subpasta|nome" -> FolderItem

function Resolve-MtpFolder($deviceName, $subfolder) {
  if ($script:folders.ContainsKey($subfolder)) { return $script:folders[$subfolder] }
  $dev = $shell.NameSpace(17).Items() | Where-Object { $_.Name -eq $deviceName } | Select-Object -First 1
  if (-not $dev) { throw "MTP_INDISPONIVEL: dispositivo '$deviceName' nao encontrado em Este Computador." }
  $storage = @($dev.GetFolder.Items() | Where-Object { $_.IsFolder }) | Select-Object -First 1
  if (-not $storage) { throw "MTP_INDISPONIVEL: armazenamento do celular inacessivel (desbloqueie a tela e use 'Transferencia de arquivos')." }
  $folder = $storage.GetFolder
  foreach ($part in ($subfolder -split '[\\/]')) {
    if ([string]::IsNullOrWhiteSpace($part)) { continue }
    $next = $folder.Items() | Where-Object { $_.Name -eq $part } | Select-Object -First 1
    if (-not $next) { throw "PASTA_NAO_ENCONTRADA: '$part' (em '$subfolder')." }
    $folder = $next.GetFolder
  }
  $script:folders[$subfolder] = $folder
  return $folder
}

function Get-Item($cmd) {
  $key = "$($cmd.subfolder)|$($cmd.name)"
  if ($script:items.ContainsKey($key)) { return $script:items[$key] }
  $folder = Resolve-MtpFolder $cmd.device $cmd.subfolder
  $it = $folder.Items() | Where-Object { $_.Name -eq $cmd.name } | Select-Object -First 1
  return $it
}

function Invoke-Command2($cmd) {
  switch ($cmd.op) {
    'list' {
      $folder = Resolve-MtpFolder $cmd.device $cmd.subfolder
      $out = New-Object System.Collections.ArrayList
      foreach ($it in @($folder.Items())) {
        if ($it.IsFolder) { continue }
        $script:items["$($cmd.subfolder)|$($it.Name)"] = $it
        $size = $it.ExtendedProperty('System.Size')
        $mod = $it.ExtendedProperty('System.DateModified')
        [void]$out.Add(@{
          name = $it.Name
          size = if ($null -ne $size) { [int64]$size } else { -1 }
          modified = if ($mod) { ([datetime]$mod).ToString('o') } else { $null }
        })
      }
      return @{ files = $out.ToArray() }
    }
    'copy' {
      $it = Get-Item $cmd
      if (-not $it) { throw "ARQUIVO_NAO_ENCONTRADO: '$($cmd.name)' nao esta mais no celular." }
      $dest = $shell.NameSpace($cmd.destDir)
      if (-not $dest) { throw "Pasta destino '$($cmd.destDir)' inacessivel." }
      # 4 = sem barra de progresso, 16 = sim para tudo, 512 = nao criar pastas, 1024 = sem UI de erro
      $dest.CopyHere($it, 4 + 16 + 512 + 1024)
      return @{ ok = $true }
    }
    'delete' {
      $it = Get-Item $cmd
      if (-not $it) { return @{ ok = $true; jaNaoExistia = $true } }
      $it.InvokeVerbEx('delete')
      $script:items.Remove("$($cmd.subfolder)|$($cmd.name)")
      # confirma que sumiu
      $folder = Resolve-MtpFolder $cmd.device $cmd.subfolder
      for ($i = 0; $i -lt 20; $i++) {
        $still = $folder.Items() | Where-Object { $_.Name -eq $cmd.name } | Select-Object -First 1
        if (-not $still) { return @{ ok = $true } }
        Start-Sleep -Milliseconds 250
      }
      throw "Arquivo '$($cmd.name)' ainda aparece no celular apos apagar."
    }
    default { throw "Comando desconhecido: $($cmd.op)" }
  }
}

[Console]::Out.WriteLine('{"ready":true}')
[Console]::Out.Flush()

while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { break }
  if ([string]::IsNullOrWhiteSpace($line)) { continue }
  $resp = $null
  try {
    $cmd = $line | ConvertFrom-Json
    $result = Invoke-Command2 $cmd
    $resp = @{ id = $cmd.id; ok = $true; result = $result }
  } catch {
    $resp = @{ id = $cmd.id; ok = $false; error = $_.Exception.Message }
  }
  [Console]::Out.WriteLine(($resp | ConvertTo-Json -Compress -Depth 6))
  [Console]::Out.Flush()
}
