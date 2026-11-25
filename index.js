// Transfere as fotos e vídeos do celular (DCIM e WhatsApp) para o HD externo, da mais antiga para a mais nova.
// Cada arquivo é MOVIDO do celular para uma pasta temporária (o Windows só apaga do celular
// depois de terminar a cópia) e de lá para D:\Fotos\<celular>\ano\mês ou D:\Videos\<celular>\ano\mês.
// Pode interromper (Ctrl+C) a qualquer momento. Para continuar, basta rodar de novo.

const { spawn } = require("child_process")
const fs = require("fs")
const path = require("path")
const readline = require("readline")

const CELULAR = "S20 FE de Carlos"
// Pastas do celular a transferir, com todas as subpastas (pastas ocultas, com ".", são ignoradas).
// Figurinhas, GIFs, documentos e status do WhatsApp ficam de fora.
const WHATSAPP = "Android\\media\\com.whatsapp\\WhatsApp\\Media"
const WHATSAPP_BUSINESS = "Android\\media\\com.whatsapp.w4b\\WhatsApp Business\\Media"
const PASTAS_CELULAR = [
  "DCIM",
  `${WHATSAPP}\\WhatsApp Images`,
  `${WHATSAPP}\\WhatsApp Video`,
  `${WHATSAPP}\\WhatsApp Video Notes`,
  `${WHATSAPP_BUSINESS}\\WhatsApp Business Images`,
  `${WHATSAPP_BUSINESS}\\WhatsApp Business Video`,
  `${WHATSAPP_BUSINESS}\\WhatsApp Business Video Notes`,
  "Pictures", // inclui Pictures\Whatsapp, Instagram, Telegram, ...
  "Movies\\WhatsApp",
  "Download",
  "Documents",
  "Telegram\\Telegram Images",
  "Movies\\Telegram",
]
const DESTINO_FOTOS = `D:\\Fotos\\${CELULAR}`
const DESTINO_VIDEOS = `D:\\Videos\\${CELULAR}`
const FOTOS = [".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp", ".gif", ".dng"]
const VIDEOS = [".mp4", ".mov", ".3gp", ".mkv", ".avi", ".m4v", ".webm"]
const SIMULAR = process.argv.includes("--simular")
// --limite N: transfere só as N fotos mais antigas
const LIMITE = Number(process.argv[process.argv.indexOf("--limite") + 1]) || Infinity
const TEMP = path.join(DESTINO_FOTOS, ".copiando") // no mesmo HD, então mover de lá é instantâneo

// O Node não acessa o celular (MTP) diretamente, então um PowerShell fica aberto
// recebendo comandos em JSON pelo stdin e respondendo um JSON por linha.
const PS = `
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$shell = New-Object -ComObject Shell.Application
$dev = $shell.NameSpace(17).Items() | Where-Object Name -eq '${CELULAR}' | Select-Object -First 1
if (-not $dev) { [Console]::Out.WriteLine('{"erro":"Celular nao encontrado"}'); exit }
$raiz = @($dev.GetFolder.Items())[0].GetFolder
$itens = @{}
# Abre uma pasta pelo caminho a partir da raiz do armazenamento (ex.: DCIM\\Camera); $null se não existir
function Abrir($caminho) {
  $pasta = $raiz
  foreach ($p in $caminho -split '\\\\') {
    $item = $pasta.Items() | Where-Object Name -eq $p | Select-Object -First 1
    if (-not $item) { return $null }
    $pasta = $item.GetFolder
  }
  return $pasta
}
# Lista as fotos e vídeos da pasta e de todas as subpastas (ignora o que começa com ".")
function Listar($pasta, $relativo, $extensoes) {
  foreach ($it in @($pasta.Items())) {
    if ($it.Name.StartsWith('.')) { continue }
    $caminho = "$relativo\\$($it.Name)"
    if ($it.IsFolder) { Listar $it.GetFolder $caminho $extensoes; continue }
    if ($extensoes -notcontains [IO.Path]::GetExtension($it.Name).ToLower()) { continue }
    $itens[$caminho] = $it
    $data = $it.ExtendedProperty('System.DateModified')
    @{ caminho = $caminho; nome = $it.Name; tamanho = [int64]$it.ExtendedProperty('System.Size')
       data = if ($data) { ([datetime]$data).ToString('o') } else { $null } }
  }
}
while ($null -ne ($linha = [Console]::In.ReadLine())) {
  try {
    $c = $linha | ConvertFrom-Json
    $r = @{ ok = $true }
    if ($c.op -eq 'listar') {
      $r.arquivos = @(foreach ($p in $c.pastas) { $pasta = Abrir $p; if ($pasta) { Listar $pasta $p $c.extensoes } })
    } elseif ($c.op -eq 'mover') {
      $shell.NameSpace($c.destino).MoveHere($itens[$c.caminho], 4 + 16 + 1024)
    }
  } catch { $r = @{ ok = $false; erro = $_.Exception.Message } }
  [Console]::Out.WriteLine(($r | ConvertTo-Json -Compress -Depth 4))
}
`

function abrirCelular() {
  const ps = spawn("powershell.exe", ["-NoProfile", "-EncodedCommand", Buffer.from(PS, "utf16le").toString("base64")])
  let erros = ""
  ps.stderr.on("data", (d) => (erros += d))
  const respostas = readline.createInterface({ input: ps.stdout })[Symbol.asyncIterator]()
  return {
    async enviar(comando) {
      ps.stdin.write(JSON.stringify(comando) + "\n")
      const { value, done } = await respostas.next()
      if (done) throw Object.assign(new Error(`Perdeu a conexão com o celular ${erros.trim()}`), { fatal: true })
      const r = JSON.parse(value)
      if (!r.ok) throw new Error(r.erro)
      return r
    },
    fechar: () => ps.stdin.end(),
  }
}

// Data da foto: pelo nome (20240714_083309.jpg) ou, se não tiver, pela data de modificação.
function dataDaFoto(arquivo) {
  // aceita 20240714_083309, 20240714-083309, 2024-07-14 e IMG-20240714-WA0001
  const m = arquivo.nome.match(/(20\d\d)[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12]\d|3[01])(?:[_ -](\d\d)[-.]?(\d\d)[-.]?(\d\d))?/)
  if (m) return new Date(m[1], m[2] - 1, m[3], m[4] || 0, m[5] || 0, m[6] || 0)
  return new Date(arquivo.data)
}

const tamanho = (arquivo) => (fs.existsSync(arquivo) ? fs.statSync(arquivo).size : -1)

// Um log por dia: D:\transferencia\transferencia-AAAAMMDD.log
const hoje = new Date()
const AAAAMMDD = `${hoje.getFullYear()}${String(hoje.getMonth() + 1).padStart(2, "0")}${String(hoje.getDate()).padStart(2, "0")}`
const ARQUIVO_LOG = path.join("D:\\transferencia", `transferencia-${AAAAMMDD}.log`)
fs.mkdirSync(path.dirname(ARQUIVO_LOG), { recursive: true })

const log = (msg) => {
  console.log(msg)
  fs.appendFileSync(ARQUIVO_LOG, `${new Date().toLocaleString("pt-BR")} ${msg}\n`)
}

const ehVideo = (nome) => VIDEOS.includes(path.extname(nome).toLowerCase())
const mesDe = (quando) => `${quando.getFullYear()}-${String(quando.getMonth() + 1).padStart(2, "0")}`

// Vídeos vão para D:\Videos\<celular>\ano\mês, o resto para D:\Fotos\<celular>\ano\mês.
function destinoDe(nome, quando) {
  const video = ehVideo(nome)
  const mes = path.join(String(quando.getFullYear()), String(quando.getMonth() + 1).padStart(2, "0"))
  return {
    pasta: path.join(video ? DESTINO_VIDEOS : DESTINO_FOTOS, mes),
    rotulo: path.join(video ? "Videos" : "Fotos", mes, nome),
  }
}

// 4194304 -> "4,0 MB"; 3221225472 -> "3,00 GB"
function legivel(bytes) {
  const num = (n, casas) => n.toLocaleString("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas })
  return bytes >= 1024 ** 3 ? `${num(bytes / 1024 ** 3, 2)} GB` : `${num(bytes / 1024 ** 2, 1)} MB`
}

const soma = (lista) => lista.reduce((total, f) => total + Math.max(0, f.tamanho), 0)
const contagem = (lista, singular, plural) =>
  `${lista.length} ${lista.length === 1 ? singular : plural} (${legivel(soma(lista))})`

// Resumo no início: total, fotos x vídeos, e quanto cada mês ocupa (com o total acumulado).
function resumo(lista) {
  const fotos = lista.filter((f) => !f.video)
  const videos = lista.filter((f) => f.video)
  log(`Total a transferir: ${contagem(lista, "arquivo", "arquivos")}`)
  log(`  Fotos:  ${contagem(fotos, "foto", "fotos")}`)
  log(`  Vídeos: ${contagem(videos, "vídeo", "vídeos")}`)
  log("Por mês:")
  const meses = Map.groupBy(lista, (f) => f.mes)
  let acumulado = 0
  for (const [mes, doMes] of meses) {
    acumulado += soma(doMes)
    const f = doMes.filter((a) => !a.video), v = doMes.filter((a) => a.video)
    const partes = [f.length && contagem(f, "foto", "fotos"), v.length && contagem(v, "vídeo", "vídeos")].filter(Boolean)
    log(`  ${mes}: ${partes.join(" + ")} = ${legivel(soma(doMes))} | acumulado ${legivel(acumulado)}`)
  }
}

// Compara dois arquivos byte a byte, em pedaços de 1 MB (vídeos grandes não lotam a memória).
function iguais(a, b) {
  if (tamanho(a) !== tamanho(b)) return false
  const fa = fs.openSync(a, "r"), fb = fs.openSync(b, "r")
  const ba = Buffer.alloc(1 << 20), bb = Buffer.alloc(1 << 20)
  try {
    for (let lidos; (lidos = fs.readSync(fa, ba)) > 0; ) {
      fs.readSync(fb, bb, 0, lidos)
      if (!ba.subarray(0, lidos).equals(bb.subarray(0, lidos))) return false
    }
    return true
  } finally {
    fs.closeSync(fa)
    fs.closeSync(fb)
  }
}

// Leva o arquivo da pasta temporária para a pasta final. Nunca sobrescreve:
// se já existe um idêntico (byte a byte) descarta o repetido; se é diferente, salva como "nome (1).jpg".
function guardar(nome, quando) {
  const temp = path.join(TEMP, nome)
  const { pasta } = destinoDe(nome, quando)
  const ext = path.extname(nome)
  let final = path.join(pasta, nome)
  for (let n = 1; fs.existsSync(final); n++) {
    if (iguais(final, temp)) { fs.rmSync(temp); return "já estava no HD" }
    final = path.join(pasta, `${path.basename(nome, ext)} (${n})${ext}`)
  }
  fs.mkdirSync(pasta, { recursive: true })
  fs.renameSync(temp, final)
  return final === path.join(pasta, nome) ? "OK" : `OK, salva como ${path.basename(final)}`
}

async function main() {
  fs.mkdirSync(TEMP, { recursive: true })

  const celular = abrirCelular()
  log("Lendo a lista de fotos do celular (pode levar alguns minutos)...")
  const { arquivos } = await celular.enviar({ op: "listar", pastas: PASTAS_CELULAR, extensoes: [...FOTOS, ...VIDEOS] })
  const fotos = arquivos
    .map((a) => {
      const quando = dataDaFoto(a)
      return { ...a, quando, mes: mesDe(quando), video: ehVideo(a.nome) }
    })
    .sort((a, b) => a.quando - b.quando || a.nome.localeCompare(b.nome))
  log(`${contagem(fotos, "arquivo", "arquivos")} no celular${SIMULAR ? " (simulação)" : ""}`)
  const porPasta = [...Map.groupBy(fotos, (f) => path.dirname(f.caminho))].sort(([a], [b]) => a.localeCompare(b))
  for (const [pasta, daPasta] of porPasta) {
    log(`  ${pasta}: ${contagem(daPasta, "arquivo", "arquivos")}`)
  }

  // Restos de uma execução interrompida. Uma cópia pela metade é menor que o original, que
  // ainda está no celular: descarta. Se não há no celular um arquivo com o mesmo nome e
  // tamanho diferente, a cópia está completa (o original já saiu do celular): guarda.
  for (const nome of SIMULAR ? [] : fs.readdirSync(TEMP)) {
    const temp = path.join(TEMP, nome)
    const mesmoNome = fotos.filter((f) => f.nome === nome)
    if (mesmoNome.length && !mesmoNome.some((f) => f.tamanho === tamanho(temp))) {
      fs.rmSync(temp, { force: true })
      continue
    }
    const quando = dataDaFoto({ nome, data: fs.statSync(temp).mtime })
    log(`Recuperado da execução anterior: ${destinoDe(nome, quando).rotulo} ${guardar(nome, quando)}`)
  }

  const lista = fotos.slice(0, LIMITE)
  if (lista.length < fotos.length) log(`Limite: só as ${lista.length} mais antigas`)
  resumo(lista)

  for (const [i, foto] of lista.entries()) {
    const origem = path.join(CELULAR, foto.caminho)
    const prefixo = `(${i + 1}/${lista.length}) ${origem} (${legivel(foto.tamanho)}) -> ${destinoDe(foto.nome, foto.quando).rotulo}`
    if (SIMULAR) { log(prefixo); continue }

    try {
      // O Windows só apaga do celular depois de terminar a cópia.
      await celular.enviar({ op: "mover", caminho: foto.caminho, destino: TEMP })
      const recebido = tamanho(path.join(TEMP, foto.nome))
      if (recebido < 0) throw new Error("não chegou no HD")
      const aviso = recebido !== foto.tamanho ? ` ATENÇÃO: tamanho ${recebido} em vez de ${foto.tamanho}, confira` : ""
      log(`${prefixo} ${guardar(foto.nome, foto.quando)}${aviso}`)
    } catch (erro) {
      if (erro.fatal) throw erro
      log(`${prefixo} ERRO: ${erro.message} (continua no celular)`)
    }
  }

  celular.fechar()
  log("Concluído.")
}

main().catch((erro) => {
  log(`ERRO: ${erro.message}`)
  process.exit(1)
})
