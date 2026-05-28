# TRUXT Macro Research Portal — contexto para Claude

## Visão geral
Site estático hospedado no GitHub Pages.  
Repo: `S:\Macro\Site` (branch `main` → GitHub Pages automático)  
Acesso restrito por senha (auth.js).

---

## Arquitetura de arquivos

```
S:\Macro\Site\
├── index.html              ← SPA vanilla JS/CSS, sem frameworks
├── auth.js                 ← lógica de senha
├── config.json             ← relatórios, países, apresentações
├── assets\truxt_logo.jpg
├── data\
│   ├── market.json         ← snapshots recentes (max 90, rolling) — gerado pelo VBA
│   └── market_history.json ← histórico diário EOD — gerado pelo VBA ExportHistorico
├── reports\
│   └── brasil\
│       ├── atividade\  (PIB, IBC-Br, PIM, PMC, PMS, PNAD, Caged_novo)
│       ├── inflacao\   (IPCA_Novo_Fluxo, IPCA15_Novo_Fluxo)
│       └── fiscal\     (Balanco-de-Pagamentos)
├── presentations\          ← PDFs listados em config.json
├── scripts\
│   └── update_market.ps1   ← roda no Main PC via Task Scheduler
└── Market\
    ├── ExportMercado.bas   ← módulo VBA (importar no xlsm)
    └── market_data.xlsm    ← Excel sempre aberto na BBG machine
```

---

## Fluxo de dados de mercado

```
BBG Machine (Excel aberto)
  └── market_data.xlsm
        └── VBA: ExportarMercado   → data\market.json     (botão "▶ Exportar Agora")
        └── VBA: ExportHistorico   → data\market_history.json (botão "▶ Exportar Histórico")

Main PC — Task Scheduler (09:10 / 13:05 / 18:05 dias úteis)
  └── update_market.ps1
        └── git add / commit / push → GitHub Pages
        (publica market.json se modificado em ≤10 min;
         market_history.json se modificado em ≤60 min)
```

**Publicação manual**: o usuário pede ao Claude para fazer git add + commit + push dos arquivos de dados.

---

## Excel: market_data.xlsm

### Sheets
| Sheet | Conteúdo | Formato BDH |
|-------|----------|-------------|
| `DI_Futuro` | DI futuros BM&F | Row 4=tickers, Row 6=fields, Row 7+=date+values |
| `FX` | Pares de câmbio | idem |
| `Treasuries` | Yields EUA | idem |
| `EXPORT` | FX + Treasuries consolidados via INDEX/MATCH | lido pelo VBA |

### VBA (ExportMercado.bas)
- `ExportarMercado`: 
  - Lê FX e Treasuries da aba EXPORT
  - Lê **todos** os vértices DI diretamente do `DI_Futuro` (row 4, tickers OD*)
  - Acrescenta snapshot ao market.json (rolling 90, newest-first)
- `ExportHistorico`:
  - **Incremental**: detecta a última data salva no JSON via InStrRev, só exporta linhas novas
  - Na primeira vez: exporta tudo do histórico BDH
  - Lê DI dinamicamente (todos os OD* do row 4)
  - Grava market_history.json (oldest-first, um snapshot/dia)
- `ParseDILabel(tickerText)`: converte ticker BBG → label legível
  - Ex: `"ODM26 COMB Comdty"` → `"Jun/26"`
  - Códigos de mês: FGHJKMNQUVXZ = Jan…Dez
- Constantes: `JSON_PATH`, `HISTORY_PATH`, `MAX_SNAPSHOTS=90`, `EXPORT_SHEET="EXPORT"`
- Timestamps em pt-BR hardcoded (Excel BBG está em inglês)
- Pula vértices DI com valor = 0 (contratos vencidos/sem dado no BBG)

### Botões na sheet EXPORT
- **▶ Exportar Agora** → `ExportarMercado` (atualiza tabela snapshot do site)
- **▶ Exportar Histórico** → `ExportHistorico` (atualiza gráfico histórico de FX/DI)

### Setup inicial VBA
1. `Alt+F11` → selecionar módulo `ExportMercado` → Delete → `File → Import File → ExportMercado.bas`
2. Colar `Workbook_Open()` no módulo `EstaPastaDeTrabalho` (código no final do .bas)
3. Na primeira vez: rodar `ExportHistorico` manualmente (exporta tudo); nas seguintes é incremental

---

## Site: index.html

SPA com 4 abas: **Países · BCs · Mercado · PPTs**

### Aba Mercado

**Sub-abas**: DI Futuro | FX | Treasuries  
**Seletores globais**: `<input type="date">` snap-a e snap-b (comparação), botão × limpa B  
**Dados**: carrega `market.json` + `market_history.json`, deduplica por data (mais recente wins)  
`findSnap(dateStr)`: aceita qualquer data, cai no dia útil anterior mais próximo  

#### DI Futuro
- Filtros: `Todos` / `≤ 2 anos` / `Só Jan` (pills, variável `mktDIFilter`)
- Gráfico SVG (`drawCurve`) com **todos** os vértices que chegam no JSON (dinâmico)
- Tabela de variação em bps abaixo do gráfico (snapA vs snapB, `renderDITable`)
- `DI_PT_MONTHS`: mapa mês PT → índice para o filtro de data

#### FX
- Toggle **Tabela | Histórico** (`fxHistView = "table"|"hist"`)
- **Tabela**: snapshot atual com var% vs snapB (`renderFX`)
- **Histórico**: série temporal SVG (`drawFXSeries`)
  - Seletor de pares: pills multi-select (7 pares, cores fixas em `FX_PAIRS`)
  - Atalhos de período: 1M / 3M / 6M / 1A / Tudo + pickers De/Até
  - Modos: `Valor` (absoluto) / `Var. %` (% desde início do período)
  - Cursor vertical + tooltip com valores de todos os pares selecionados
  - `fxSelPairs` (Set), `fxHistMode`, `fxActiveRange` — estado global
  - `initFXHistory()` chamado uma vez após loadMarket

#### Treasuries
- Gráfico SVG (`drawCurve`) com yields 2y→30y

### FX_PAIRS (módulo-level)
```javascript
const FX_PAIRS = [
  { key:"usdbrl", label:"USD/BRL", desc:"Dólar / Real",  dec:4, color:"var(--accent)"  },
  { key:"eurbrl", label:"EUR/BRL", desc:"Euro / Real",   dec:4, color:"var(--amber)"   },
  { key:"eurusd", label:"EUR/USD", desc:"Euro / Dólar",  dec:4, color:"var(--green)"   },
  { key:"gbpusd", label:"GBP/USD", desc:"Libra / Dólar", dec:4, color:"var(--purple)"  },
  { key:"usdjpy", label:"USD/JPY", desc:"Dólar / Iene",  dec:2, color:"var(--red)"     },
  { key:"usdcny", label:"USD/CNY", desc:"Dólar / Yuan",  dec:4, color:"#e06c9b"        },
  { key:"dxy",    label:"DXY",     desc:"Índice Dólar",  dec:2, color:"var(--text2)"   },
];
```

### config.json
```json
{
  "countries": ["brasil","eua","zona_euro","china"],
  "reports": [{ "id","title","country","section","file","updated" }],
  "centralBanks": [],
  "speeches": [],
  "presentations": [{ "id","title","file","date","description" }]
}
```
Seções válidas: `atividade`, `inflacao`, `fiscal`

---

## Regras de trabalho

- **Relatórios**: sempre adicionar ao config.json + confirmar que o .html existe em reports/
- **Apresentações (PPTs)**: o usuário sempre pede explicitamente — nunca auto-escanear
- **Git**: commitar apenas arquivos relevantes
  - `data/market.json` e `data/market_history.json`: commitar quando o usuário pedir publicação
  - Nunca commitar `Market/market_data.xlsm`
- **Publicação de dados**: o usuário pede "atualiza o site" → `git add data/market.json data/market_history.json` + commit + push

---

## Pendências / ideias futuras

1. **Histórico de Treasuries**: gráfico de série temporal igual ao FX Histórico (usa os mesmos dados de market_history.json — campo `treasuries`).

2. **Histórico de DI**: mostrar como a curva DI evoluiu ao longo do tempo — seria um chart com datas no eixo X e possibilidade de selecionar vértices.

3. **Novos países/relatórios**: adicionar EUA, Zona do Euro, China (apenas estrutura no config.json + htmls em reports/).

4. **Outros indicadores Brasil**: o usuário pode querer adicionar mais relatórios à medida que forem produzidos.
