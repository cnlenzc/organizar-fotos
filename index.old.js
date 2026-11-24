// Transfere fotos do celular (MTP) para o HD externo, em ordem de data.
// Cada arquivo: copia para uma pasta temporária -> confere o tamanho -> move para o
// destino final -> só então apaga do celular. Se for interrompido, basta rodar de novo:
// o que já foi apagado do celular já está salvo no HD, e o que sobrou recomeça do zero.
//
// Uso:
//   node index.js              transfere e apaga do celular
//   node index.js --simular    só mostra o que seria feito
//   node index.js --sem-apagar copia, mas não apaga do celular
//   node index.js --limite 5  processa só os 5 mais antigos (teste)
//   node index.js --pastas "DCIM\Camera,DCIM\Screenshots"

const { spawn } = require("child_process")
const fs = require("fs")
const path = require("path")
const readline = require("readline")

const args = process.argv.slice(2)
const argValue = (flag) => {
  const i = args.indexOf(flag)
  return i >= 0 ? args[i + 1] : undefined
}

const CONFIG = {
  deviceName: "S20 FE de Carlos",
  subfolders: (argValue("--pastas") || "DCIM\\Camera").split(",").map((s) => s.trim()).filter(Boolean),
  destBase: "D:\\Fotos\\S20 FE",
  simular: args.includes("--simular"),
  apagar: !args.includes("--sem-apagar"),
  limite: Number(argValue("--limite")) || Infinity, // processa no máximo N arquivos (para teste)
  logFile: path.join(__dirname, "transferencia.log"),
}
const STAGING = path.join(CONFIG.destBase, ".transferindo")

// ---------- log ----------
function log(msg) {
  const line = `${new Date().toISOString()} ${msg}`
  console.log(msg)
  fs.appendFileSync(CONFIG.logFile, line + "\n", "utf8")
}

// ---------- ponte PowerShell (processo único, mantém o celular "aberto") ----------
class MtpBridge {
  constructor() {
    this.pending = new Map()
    this.nextId = 1
  }

  start() {
    return new Promise((resolve, reject) => {
      this.proc = spawn(
        "powershell.exe",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-STA", "-File", path.join(__dirname, "mtp-bridge.ps1")],
        { stdio: ["pipe", "pipe", "pipe"], windowsHide: true }
      )
      this.proc.stdout.setEncoding("utf8")
      this.proc.stderr.on("data", (d) => fs.appendFileSync(CONFIG.logFile, `[ps-stderr] ${d}`, "utf8"))
      const rl = readline.createInterface({ input: this.proc.stdout })
      rl.on("line", (line) => {
        line = line.replace(/^\uFEFF/, "").trim()
        if (!line.startsWith("{")) return
        let msg
        try { msg = JSON.parse(line) } catch { return }
        if (msg.ready) return resolve()
        const p = this.pending.get(msg.id)
        if (!p) return
        this.pending.delete(msg.id)
        clearTimeout(p.timer)
        msg.ok ? p.resolve(msg.result) : p.reject(new Error(msg.error))
      })
      this.proc.on("exit", (code) => {
        const err = new Error(`Processo PowerShell terminou (código ${code})`)
        for (const p of this.pending.values()) { clearTimeout(p.timer); p.reject(err) }
        this.pending.clear()
        reject(err)
      })
    })
  }

  call(cmd, timeoutMs = 120000) {
    const id = this.nextId++
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        reject(new Error(`Tempo esgotado em '${cmd.op}' (${cmd.name || cmd.subfolder})`))
      }, timeoutMs)
      this.pending.set(id, { resolve, reject, timer })
      this.proc.stdin.write(JSON.stringify({ id, device: CONFIG.deviceName, ...cmd }) + "\n")
    })
  }

  stop() {
    if (this.proc && !this.proc.killed) {
      this.proc.stdin.end()
      setTimeout(() => this.proc.kill(), 2000).unref()
    }
  }
}

// ---------- datas ----------
// Samsung: 20240714_083309.jpg, Screenshot_20240714-083309_App.jpg, etc.
function fileDate(name, modifiedIso) {
  const m = name.match(/(20\d{2})(\d{2})(\d{2})[_-]?(\d{2})?(\d{2})?(\d{2})?/)
  if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31) {
    return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0))
  }
  const d = modifiedIso ? new Date(modifiedIso) : null
  return d && !Number.isNaN(d.getTime()) ? d : new Date(0)
}

// ---------- utilitários de arquivo ----------
const sizeOf = (p) => { try { return fs.statSync(p).size } catch { return -1 } }

async function waitForSize(file, expected, timeoutMs) {
  const start = Date.now()
  let last = -1, stableSince = Date.now()
  while (Date.now() - start < timeoutMs) {
    const s = sizeOf(file)
    if (s === expected && expected >= 0) return true
    if (expected < 0 && s > 0) {           // tamanho desconhecido: espera estabilizar
      if (s !== last) { last = s; stableSince = Date.now() }
      else if (Date.now() - stableSince > 3000) return true
    }
    await new Promise((r) => setTimeout(r, 200))
  }
  return false
}

// Nome livre no destino, sem sobrescrever nada: foto.jpg -> foto (1).jpg
function freeName(dir, name) {
  const ext = path.extname(name), base = path.basename(name, ext)
  let candidate = path.join(dir, name), n = 1
  while (fs.existsSync(candidate)) candidate = path.join(dir, `${base} (${n++})${ext}`)
  return candidate
}

// ---------- principal ----------
let parar = false
process.on("SIGINT", () => {
  if (parar) process.exit(1)
  parar = true
  console.log("\nInterrompendo após o arquivo atual... (Ctrl+C de novo para forçar)")
})

async function processFile(bridge, subfolder, file) {
  const d = file.date
  const destDir = path.join(CONFIG.destBase, String(d.getFullYear()), String(d.getMonth() + 1).padStart(2, "0"))
  const finalPath = path.join(destDir, file.name)
  const label = `[${path.relative(CONFIG.destBase, destDir)}] ${file.name}`

  // Já existe no destino com o mesmo tamanho? Então já foi copiado antes (ex.: interrompido antes de apagar).
  const existing = sizeOf(finalPath)
  let savedAt = null
  if (existing >= 0 && existing === file.size) {
    savedAt = finalPath
    log(`OK (já estava no HD) ${label}`)
  } else {
    if (CONFIG.simular) { log(`SIMULAR copiar ${label}`); return }

    // copia para a pasta temporária (limpa) e confere tamanho
    fs.rmSync(path.join(STAGING, file.name), { force: true })
    const timeout = 60000 + Math.max(0, file.size) / 1024 // ~1 s por MB extra
    await bridge.call({ op: "copy", subfolder, name: file.name, destDir: STAGING }, timeout)
    const tmpPath = path.join(STAGING, file.name)
    if (!(await waitForSize(tmpPath, file.size, timeout))) {
      throw new Error(`cópia incompleta (esperado ${file.size} bytes, obtido ${sizeOf(tmpPath)})`)
    }
    // confere que o arquivo pode ser lido por inteiro
    if (fs.readFileSync(tmpPath).length !== file.size) throw new Error("falha ao reler o arquivo copiado")

    fs.mkdirSync(destDir, { recursive: true })
    savedAt = existing >= 0 ? freeName(destDir, file.name) : finalPath // nunca sobrescreve
    fs.renameSync(tmpPath, savedAt)
    if (sizeOf(savedAt) !== file.size && file.size >= 0) throw new Error("tamanho diferente após mover para o destino")
    log(`COPIADO ${label}${savedAt !== finalPath ? ` -> ${path.basename(savedAt)}` : ""} (${(file.size / 1048576).toFixed(1)} MB)`)
  }

  if (!CONFIG.apagar || CONFIG.simular) return
  if (file.size < 0) { log(`  não apagado (tamanho desconhecido no celular, confira manualmente): ${file.name}`); return }
  await bridge.call({ op: "delete", subfolder, name: file.name }, 30000)
  log(`  apagado do celular: ${file.name}`)
}

async function main() {
  if (!fs.existsSync(path.parse(CONFIG.destBase).root)) {
    console.error(`HD não encontrado (${path.parse(CONFIG.destBase).root}). Conecte o HD e tente de novo.`)
    process.exit(1)
  }
  fs.mkdirSync(STAGING, { recursive: true })
  for (const f of fs.readdirSync(STAGING)) fs.rmSync(path.join(STAGING, f), { force: true, recursive: true }) // restos de execução interrompida

  log(`=== Início ${CONFIG.simular ? "(SIMULAÇÃO) " : ""}${CONFIG.apagar ? "" : "(sem apagar) "}pastas: ${CONFIG.subfolders.join(", ")}`)
  const bridge = new MtpBridge()
  await bridge.start()

  let total = 0, feitos = 0, erros = 0
  try {
    for (const subfolder of CONFIG.subfolders) {
      if (parar) break
      console.log(`Lendo ${subfolder} no celular (pode levar alguns segundos)...`)
      let files
      try {
        ({ files } = await bridge.call({ op: "list", subfolder }, 600000))
      } catch (err) {
        log(`ERRO ao listar ${subfolder}: ${err.message}`)
        erros++
        continue
      }
      files = (files || [])
        .map((f) => ({ ...f, date: fileDate(f.name, f.modified) }))
        .sort((a, b) => a.date - b.date || a.name.localeCompare(b.name))
      log(`${files.length} arquivos em ${subfolder}`)
      total += files.length

      let errosSeguidos = 0
      for (const [i, file] of files.entries()) {
        if (parar || feitos >= CONFIG.limite) break
        process.stdout.write(`(${i + 1}/${files.length}) `)
        try {
          await processFile(bridge, subfolder, file)
          feitos++
          errosSeguidos = 0
        } catch (err) {
          erros++
          errosSeguidos++
          log(`ERRO ${file.name}: ${err.message} (mantido no celular)`)
          if (errosSeguidos >= 5) {
            log("5 erros seguidos — o celular pode ter desconectado ou bloqueado. Parando. Rode de novo para continuar.")
            parar = true
          }
        }
      }
    }
  } finally {
    bridge.stop()
  }
  log(`=== Fim: ${feitos}/${total} processados, ${erros} erro(s)${parar ? " — interrompido, rode de novo para continuar" : ""}`)
  process.exitCode = erros ? 1 : 0
}

main().catch((err) => {
  log(`ERRO FATAL: ${err.message}`)
  process.exit(1)
})
