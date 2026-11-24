const { execFileSync } = require("child_process")
const fs = require("fs")
const path = require("path")

const CONFIG = {
  deviceName: "S20 FE de Carlos",
  subfolders: ["DCIM\\Camera"], // Caminho relativo à raiz do armazenamento interno do celular
  destBase: "D:\\Fotos\\S20 FE",
  deleteAfterConfirm: false,
  logFile: path.join(__dirname, "transferencia-log.json")
}

function runPowerShell(psScript) {
  const wrapped = `$ErrorActionPreference = "Stop"\n` + psScript
  const encoded = Buffer.from(wrapped, "utf16le").toString("base64")
  const out = execFileSync(
    "powershell.exe",
    ["-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
    { encoding: "utf8", maxBuffer: 50 * 1024 * 1024 }
  )
  return out.trim()
}

function loadLog(logPath) {
  if (!fs.existsSync(logPath)) {
    return { transferred: {} }
  }
  try {
    const raw = fs.readFileSync(logPath, "utf8")
    const parsed = JSON.parse(raw)
    return { transferred: parsed.transferred || {} }
  } catch (err) {
    console.warn("Aviso: log corrompido, iniciando log limpo:", err.message)
    return { transferred: {} }
  }
}

function saveLog(logPath, data) {
  const tmpFile = `${logPath}.tmp`
  fs.writeFileSync(tmpFile, JSON.stringify(data, null, 2), "utf8")
  fs.renameSync(tmpFile, logPath)
}

function parseDateFromNameOrString(name, fallbackDateStr) {
  const match = name.match(/(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])/)
  if (match) {
    return new Date(`${match[1]}-${match[2]}-${match[3]}T12:00:00`)
  }
  const parsed = new Date(fallbackDateStr)
  return Number.isNaN(parsed.getTime()) ? new Date() : parsed
}

function extractPsErrorMessage(err) {
  const raw = ((err && (err.stderr || err.stdout || err.message)) || "").toString()
  const matches = [...raw.matchAll(/<S S="Error">([\s\S]*?)<\/S>/g)]
  if (matches.length > 0) {
    return matches
      .map((m) =>
        m[1]
          .replace(/_x000D__x000A_/g, " ")
          .replace(/_x000D_|_x000A_/g, " ")
          .trim()
      )
      .filter(Boolean)
      .join(" | ")
  }
  return raw.replace(/#<\s*CLIXML[\s\S]*$/i, "").trim() || "Erro desconhecido no PowerShell"
}

function showMtpUnavailableMessage(deviceName, detail) {
  console.error("")
  console.error(`Não foi possível acessar os arquivos de "${deviceName}" via MTP.`)
  console.error("Isso geralmente acontece quando:")
  console.error("  1. A tela do celular está bloqueada ou o modo USB está em 'Apenas carregar'.")
  console.error("  2. A sessão MTP travou no Windows (o próprio Explorador de Arquivos não abre as pastas).")
  console.error("  3. O cabo ou porta USB falhou durante a comunicação.")
  console.error("Ação sugerida: desbloqueie o celular, alterne o modo USB para 'Transferência de arquivos'")
  console.error("ou reconecte o cabo / reinicie o aparelho e teste abrir no Explorador antes de rodar novamente.")
  if (detail) {
    console.error(`Detalhe técnico: ${detail}`)
  }
  console.error("")
}

function handleMtpError(contextLabel, deviceName, err) {
  const cleanMsg = extractPsErrorMessage(err)
  if (cleanMsg.includes("MTP_STORAGE_INACCESSIBLE")) {
    showMtpUnavailableMessage(deviceName, cleanMsg)
    return true
  }
  console.error(`${contextLabel}:`, cleanMsg)
  return false
}

function buildResolveFolderSnippet(deviceName, subfolder) {
  const safeDevice = deviceName.replace(/'/g, "''")
  const safeSub = subfolder.replace(/'/g, "''")
  return `
$shell = New-Object -ComObject Shell.Application
$root = $shell.NameSpace(17)
if (-not $root) { throw "MTP_STORAGE_INACCESSIBLE: Nao foi possivel acessar 'Este Computador'." }

$dev = $root.Items() | Where-Object { $_.Name -eq '${safeDevice}' } | Select-Object -First 1
if (-not $dev) { throw "MTP_STORAGE_INACCESSIBLE: Dispositivo '${safeDevice}' nao encontrado." }

$devFolder = $dev.GetFolder
if (-not $devFolder) { throw "MTP_STORAGE_INACCESSIBLE: Nao foi possivel abrir o dispositivo '${safeDevice}'." }

try {
  $storageItem = @($devFolder.Items() | Where-Object { $_.IsFolder }) | Select-Object -First 1
} catch {
  throw "MTP_STORAGE_INACCESSIBLE: Falha ao enumerar armazenamento de '${safeDevice}'."
}

if (-not $storageItem) {
  throw "MTP_STORAGE_INACCESSIBLE: Nenhuma pasta de armazenamento acessivel em '${safeDevice}'."
}

$folder = $storageItem.GetFolder
if (-not $folder) {
  throw "MTP_STORAGE_INACCESSIBLE: Nao foi possivel abrir a pasta do armazenamento interno."
}

try {
  $null = @($folder.Items())
} catch {
  throw "MTP_STORAGE_INACCESSIBLE: Falha ao enumerar itens do armazenamento interno."
}

foreach ($part in ('${safeSub}' -split '\\\\')) {
  if ([string]::IsNullOrWhiteSpace($part)) { continue }
  $next = $folder.Items() | Where-Object { $_.Name -eq $part } | Select-Object -First 1
  if (-not $next) {
    $disponiveis = ($folder.Items() | ForEach-Object { $_.Name }) -join ", "
    throw "Subpasta '$part' nao encontrada em '${safeSub}'. Pastas disponiveis aqui: $disponiveis"
  }
  $folder = $next.GetFolder
  if (-not $folder) {
    throw "MTP_STORAGE_INACCESSIBLE: Nao foi possivel abrir a subpasta '$part'."
  }
}
`
}

function listMtpFiles(deviceName, subfolder) {
  const ps = `
${buildResolveFolderSnippet(deviceName, subfolder)}
$items = @($folder.Items() | Where-Object { -not $_.IsFolder } | ForEach-Object {
  [PSCustomObject]@{
    Name = $_.Name
    Size = [int64]$_.Size
    DateRaw = $folder.GetDetailsOf($_, 3)
  }
})
$items | ConvertTo-Json -Compress
`
  const jsonText = runPowerShell(ps)
  if (!jsonText) return []
  const parsed = JSON.parse(jsonText)
  return Array.isArray(parsed) ? parsed : [parsed]
}

function copyMtpFile(deviceName, subfolder, fileName, destDir) {
  const safeFile = fileName.replace(/'/g, "''")
  const safeDest = destDir.replace(/'/g, "''")
  const ps = `
${buildResolveFolderSnippet(deviceName, subfolder)}
$item = $folder.Items() | Where-Object { $_.Name -eq '${safeFile}' } | Select-Object -First 1
if (-not $item) { throw "Arquivo '${safeFile}' nao encontrado no celular." }

$destFolder = $shell.NameSpace('${safeDest}')
if (-not $destFolder) { throw "Pasta destino '${safeDest}' nao encontrada." }

$destFolder.CopyHere($item, 16)
`
  runPowerShell(ps)
}

function deleteMtpFile(deviceName, subfolder, fileName) {
  const safeFile = fileName.replace(/'/g, "''")
  const ps = `
${buildResolveFolderSnippet(deviceName, subfolder)}
$item = $folder.Items() | Where-Object { $_.Name -eq '${safeFile}' } | Select-Object -First 1
if ($item) {
  $item.InvokeVerb("delete")
}
`
  runPowerShell(ps)
}

function main() {
  fs.mkdirSync(CONFIG.destBase, { recursive: true })
  const state = loadLog(CONFIG.logFile)

  for (const subfolder of CONFIG.subfolders) {
    console.log(`Lendo arquivos de ${CONFIG.deviceName} -> ${subfolder}...`)
    let files = []
    try {
      files = listMtpFiles(CONFIG.deviceName, subfolder)
    } catch (err) {
      handleMtpError(`Falha ao listar ${subfolder}`, CONFIG.deviceName, err)
      continue
    }

    const ordered = files
      .map((f) => ({
        ...f,
        dateObj: parseDateFromNameOrString(f.Name, f.DateRaw)
      }))
      .sort((a, b) => a.dateObj - b.dateObj || a.Name.localeCompare(b.Name))

    console.log(`${ordered.length} arquivos encontrados em ${subfolder}.`)

    for (const file of ordered) {
      const year = String(file.dateObj.getFullYear())
      const month = String(file.dateObj.getMonth() + 1).padStart(2, "0")
      const destDir = path.join(CONFIG.destBase, year, month)
      const destFile = path.join(destDir, file.Name)
      const key = `${subfolder}\\${file.Name}`

      fs.mkdirSync(destDir, { recursive: true })

      let existsValid = fs.existsSync(destFile) && fs.statSync(destFile).size > 0

      if (!existsValid) {
        console.log(`Copiando [${year}\\${month}] ${file.Name}...`)
        try {
          copyMtpFile(CONFIG.deviceName, subfolder, file.Name, destDir)
          existsValid = fs.existsSync(destFile) && fs.statSync(destFile).size > 0
        } catch (err) {
          handleMtpError(`Erro ao copiar ${file.Name}`, CONFIG.deviceName, err)
          continue
        }
      } else {
        console.log(`Já íntegro no destino [${year}\\${month}]: ${file.Name}`)
      }

      if (!existsValid) {
        console.warn(`Arquivo não confirmado após cópia: ${file.Name}`)
        continue
      }

      const localSize = fs.statSync(destFile).size
      state.transferred[key] = {
        destFile,
        size: localSize,
        copiedAt: new Date().toISOString(),
        deletedFromDevice: Boolean(state.transferred[key]?.deletedFromDevice)
      }
      saveLog(CONFIG.logFile, state)

      if (CONFIG.deleteAfterConfirm && !state.transferred[key].deletedFromDevice) {
        try {
          deleteMtpFile(CONFIG.deviceName, subfolder, file.Name)
          state.transferred[key].deletedFromDevice = true
          state.transferred[key].deletedAt = new Date().toISOString()
          saveLog(CONFIG.logFile, state)
          console.log(`  -> Removido do celular: ${file.Name}`)
        } catch (err) {
          handleMtpError(`  -> Falha ao remover ${file.Name} do celular`, CONFIG.deviceName, err)
        }
      }
    }
  }

  console.log("Concluído.")
}

main()