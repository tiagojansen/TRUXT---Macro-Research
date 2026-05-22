# Macro Research Portal

Portal interno de análise macroeconômica. SPA mobile-first, sem dependências externas.

## Estrutura

```
Site/
├── index.html                        ← portal principal
├── auth.js                           ← proteção por senha
├── config.json                       ← índice de relatórios
├── .github/workflows/pages.yml       ← deploy automático (GitHub Pages)
└── reports/
    └── brasil/
        └── atividade/
            └── pmc.html              ← PMC — Pesquisa Mensal de Comércio
```

## Uso local

Abra o `index.html` direto no navegador **ou** sirva com qualquer servidor estático:

```bash
# Python (qualquer pasta)
python -m http.server 8080

# Node (npx, sem instalação)
npx serve .
```

Acesse `http://localhost:8080` e use a senha `macro2026`.

## Trocar a senha

Edite a primeira linha de `auth.js`:

```js
const PASSWORD = "macro2026";  // ← troque aqui
```

A senha fica no client-side (adequado para uso interno em rede privada). Para proteção real em ambiente público, implemente autenticação server-side.

## Adicionar relatórios

1. Copie o HTML renderizado para `reports/<país>/<seção>/<id>.html`
2. Adicione uma entrada em `config.json`:

```json
{
  "id": "pib",
  "title": "PIB — Contas Nacionais",
  "country": "brasil",
  "section": "atividade",
  "file": "reports/brasil/atividade/pib.html",
  "updated": "2026-05-22"
}
```

**Países disponíveis:** `brasil`, `eua`, `zona_euro`, `china`, `reino_unido`, `japao`  
**Seções disponíveis:** `atividade`, `inflacao`, `fiscal`

## Deploy no GitHub Pages

### 1. Criar repositório no GitHub

Acesse: https://github.com/new  
- Deixe **privado** (recomendado para uso interno)
- Não inicialize com README (já temos um)

### 2. Conectar e publicar

```bash
cd S:\Macro\Site
git remote add origin https://github.com/SEU_USUARIO/macro-portal.git
git push -u origin main
```

### 3. Habilitar GitHub Pages

No repositório → **Settings** → **Pages**:
- Source: `GitHub Actions`
- O workflow em `.github/workflows/pages.yml` fará o deploy automaticamente a cada `git push`.

### 4. Acessar pelo celular

Após o deploy, o site ficará disponível em:
```
https://SEU_USUARIO.github.io/macro-portal/
```

> **Dica:** adicione aos favoritos na tela inicial do celular (iOS: compartilhar → "Adicionar à Tela de Início"; Android: menu → "Adicionar à tela inicial").

## Atualizar relatórios

```bash
# 1. Renderize o novo relatório (ex: com Quarto)
quarto render PMC.qmd --to html

# 2. Copie para a pasta correta
copy PMC.html S:\Macro\Site\reports\brasil\atividade\pmc.html

# 3. Atualize a data em config.json, então:
cd S:\Macro\Site
git add reports/brasil/atividade/pmc.html config.json
git commit -m "update: PMC maio 2026"
git push
```

O GitHub Actions publica automaticamente em ~1 minuto.
