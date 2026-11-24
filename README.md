# organizar-fotos

Transfere as fotos e vídeos do celular (Samsung S20 FE) para o HD externo, **da mais antiga para a mais nova**, movendo cada arquivo (o Windows só apaga do celular depois de copiar).

São lidas todas as pastas dentro de `DCIM` (Camera, Screenshots, Screen recordings, Restored, ...), incluindo subpastas. Pastas ocultas (que começam com `.`) são ignoradas.

## Requisitos

- Windows (usa o Explorador de Arquivos via PowerShell para acessar o celular)
- Node.js 21 ou mais novo
- Celular conectado por USB, **desbloqueado** e no modo **Transferência de arquivos**
- HD externo conectado

## Como usar

```
node index.js --simular   # só mostra a lista, na ordem em que seria copiada
node index.js             # copia e apaga do celular
node index.js --limite 20 # só as 20 fotos mais antigas (bom para testar)
```

`--limite` também funciona junto com `--simular`.

Também dá para usar `npm run simular` e `npm start`.

## Onde as fotos e vídeos ficam

Fotos e vídeos vão para pastas separadas, organizadas por ano e mês:

```
D:\Fotos\S20 FE de Carlos\
  2021\
    10\20211011_170615.jpg
    11\20211102_174927.jpg
D:\Videos\S20 FE de Carlos\
  2023\
    01\20230102_103435.mp4
```

São tratados como vídeo os arquivos `.mp4`, `.mov`, `.3gp`, `.mkv`, `.avi`, `.m4v` e `.webm` (lista `VIDEOS` no código). Todo o resto vai para Fotos.

A data sai do nome do arquivo: `20240714_083309.jpg`, `Screenshot_20240714_083309_App.jpg`, `IMG-20240714-WA0001.jpg`, `2024-07-14.jpg`. Se o nome não tiver data, usa a data de modificação do arquivo.

Arquivos de pastas diferentes do celular vão para a mesma pasta `ano\mês`. Se dois tiverem o mesmo nome e conteúdo diferente, o segundo é salvo como `nome (1).jpg`.

## É seguro?

Para cada foto, o script:

1. **Move** a foto do celular para uma pasta temporária no HD (`D:\Fotos\S20 FE de Carlos\.copiando`). É o mesmo "mover" do Explorador de Arquivos: o Windows só apaga a foto do celular depois de terminar a cópia.
2. Confere o tamanho. Se for diferente do que o celular informava, avisa no log com **ATENÇÃO**.
3. Leva a foto para a pasta final (`ano\mês`), em Fotos ou Videos. A pasta `.copiando` fica no mesmo HD, então esse passo é instantâneo, até para vídeos grandes.

Se a foto não chegar ao HD, ela **continua no celular** e o erro aparece no log.

O script nunca sobrescreve nada no HD:
- se já existe uma foto com o mesmo nome e tamanho, a repetida é descartada;
- se existe um arquivo diferente com o mesmo nome, a nova é salva como `nome (1).jpg`.

> Fotos removidas do celular **não vão para a lixeira** do celular.

## Se interromper

Pode fechar ou apertar **Ctrl+C** a qualquer momento. Para continuar, é só rodar de novo: o que ainda está no celular é o que falta.

Se sobrou alguma foto em `.copiando` da execução anterior, o script resolve sozinho:
- se a foto ainda está no celular, a cópia que sobrou estava pela metade e é descartada;
- se não está mais no celular, a foto é levada para a pasta final.

## Log

Tudo fica registrado em `D:\transferencia\transferencia-AAAAMMDD.log`, com um arquivo por dia (ex.: `transferencia-20261124.log`). Se rodar mais de uma vez no mesmo dia, as linhas são acrescentadas ao mesmo arquivo.

No início vem quantos arquivos há em cada pasta do celular e um resumo do que vai ser transferido (respeitando o `--limite`): o total, separado em fotos e vídeos, e quanto cada mês ocupa, com o total acumulado:

```
3254 arquivos no celular
  DCIM\Camera: 2907 arquivos (31,93 GB)
  DCIM\Screenshots: 322 arquivos (166,2 MB)
  ...
Total a transferir: 2907 arquivos (31,93 GB)
  Fotos:  2742 fotos (6,09 GB)
  Vídeos: 165 vídeos (25,83 GB)
Por mês:
  2023-04: 248 fotos (398,3 MB) + 1 vídeo (9,4 MB) = 407,7 MB | acumulado 407,7 MB
  2023-05: 6 fotos (9,2 MB) = 9,2 MB | acumulado 416,9 MB
  ...
```

Depois, uma linha por arquivo, com a origem no celular, o tamanho e o destino:

```
(1/2907) S20 FE de Carlos\DCIM\Camera\20230412_165049.jpg (4,2 MB) -> Fotos\2023\04\20230412_165049.jpg OK
```

## Configuração

No começo do [index.js](index.js):

| Constante       | Padrão              | O que é                                  |
| --------------- | ------------------- | ---------------------------------------- |
| `CELULAR`       | `S20 FE de Carlos`  | Nome do celular em "Este Computador"      |
| `PASTA_CELULAR` | `DCIM`              | Pasta do celular a transferir (com todas as subpastas) |
| `DESTINO_FOTOS` | `D:\Fotos\<CELULAR>` | Pasta das fotos no HD                  |
| `DESTINO_VIDEOS`| `D:\Videos\<CELULAR>`| Pasta dos vídeos no HD                 |
| `VIDEOS`        | `.mp4`, `.mov`, ...  | Extensões tratadas como vídeo          |

Para transferir só uma pasta (ex.: `DCIM\Screenshots`) ou outra pasta (ex.: `Pictures`), troque `PASTA_CELULAR` e rode de novo.

## Problemas comuns

- **"Celular nao encontrado"**: confira se o nome aparece igual em "Este Computador" e se o celular está desbloqueado.
- **Lista vazia ou erro ao ler**: no celular, mude o USB para "Transferência de arquivos" (às vezes fica em "Apenas carregar").
- **"salva como nome (1).jpg"**: já havia outro arquivo com o mesmo nome e tamanho diferente. As duas versões ficam guardadas.

## Como funciona por dentro

O Node não consegue acessar o celular diretamente, porque ele aparece no Windows como dispositivo MTP e não como unidade com letra. Por isso o `index.js` abre um único PowerShell que usa o `Shell.Application` (o mesmo mecanismo do Explorador) e conversa com ele por JSON, uma linha por comando.

O script usa só a operação de mover (`MoveHere`). Ela não pede confirmação; o comando de apagar direto no celular abre uma janela de confirmação do Windows.

A versão anterior do script está guardada em `index.old.js` e `mtp-bridge.old.ps1`.
