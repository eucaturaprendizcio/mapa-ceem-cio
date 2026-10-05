# Painel CEEM — Mapa por Estado e Cidade

Painel estático que mostra no mapa do Brasil (e nos mapas de Rondônia e Paraná) as coordenações CEEM: gerente, coordenador operacional e coordenador comercial de cada estado ou cidade.

Os dados vêm da planilha `data/ceem.xlsx`, lida direto no navegador. Não tem backend, banco de dados nem login.

## Estrutura

```
index.html            página com os mapas SVG
css/style.css         estilo
js/script.js          lê a planilha, monta mapas, listas, tooltips e modais
js/cidades.js         lista de cidades do Brasil (busca)
js/rondonia-dados.js  posição das cidades e contorno de Rondônia
js/parana-dados.js    posição das cidades e contorno do Paraná
data/ceem.xlsx        planilha com os dados (fonte única)
```

## Como atualizar os dados

1. Edite o arquivo `data/ceem.xlsx`.
2. Mantenha o nome das abas: `CEEM`, `Cidades RO` e `Cidades PR`.
3. Não mude o cabeçalho nem a ordem das colunas. Os dados começam na 4ª linha.
4. Na aba `CEEM`, não altere o texto da coluna Estado (UF).
5. Nas abas de cidades, a coluna CEEM aceita:
   - Rondônia: `Vilhena`, `Ji-Paraná`, `Porto Velho` ou `Cacoal`
   - Paraná: `Cascavel` ou `Curitiba`
   - `--` para cidade sem CEEM definido
6. Salve, faça commit e push. O deploy atualiza sozinho.

Colunas (a partir da B): CEEM, Cor (#RRGGBB), Gerente (nome, telefone, e-mail), Coord. Operacional (nome, telefone, e-mail), Coord. Comercial (nome, telefone, e-mail).

## Rodar localmente

O navegador bloqueia a leitura da planilha quando o `index.html` é aberto direto do disco (duplo clique). Rode por um servidor local:

```bash
python -m http.server 8000
```

Depois acesse http://localhost:8000

## Deploy na Vercel

1. Suba o projeto para um repositório no GitHub (pode ser privado).
2. Na Vercel, entre com o GitHub e instale o app da Vercel, dando acesso ao repositório.
3. Em **Add New → Project**, importe o repositório.
4. Deixe o framework como **Other**, sem build command e sem output directory.
5. Clique em **Deploy**.

A cada `git push`, a Vercel faz um novo deploy. Funciona também em GitHub Pages, Netlify e Cloudflare Pages.

## Observações

- A biblioteca [SheetJS](https://sheetjs.com/) é carregada por CDN (cdnjs) no `index.html` para ler o `.xlsx`.
- O site não tem controle de acesso. Quem tiver o link consegue ver o painel e baixar `data/ceem.xlsx`, que contém nomes, telefones e e-mails. Se precisar restringir, use a proteção por senha da Vercel (plano Pro) ou outra camada, como o Cloudflare Access.
- Se aparecer o erro "Erro ao carregar a planilha", confira se o arquivo está em `data/ceem.xlsx` e se a página está sendo servida por HTTP, não aberta do disco.
