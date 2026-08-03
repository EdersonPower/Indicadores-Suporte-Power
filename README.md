# Power Analytics — Sprint 3

Portal de Gestão do Suporte da Power System.

## Publicação no GitHub Pages

Substitua o conteúdo do repositório pelos arquivos desta pasta. Mantenha o `index.html` na raiz. O GitHub Pages já configurado atualizará automaticamente.

## Adicionar um trimestre

1. Coloque o arquivo em `data/ANO/`.
2. Adicione o registro em `data/manifest.json`.
3. Faça commit.

Exemplo:

```json
{"path":"data/2026/3trim2026.xlsx","label":"3º Trimestre 2026","year":2026,"quarter":3}
```

## Fotos dos colaboradores

Salve JPGs em `assets/img/team/` usando o primeiro nome em minúsculas:

- `dayane.jpg`
- `luiz.jpg`
- `lukas.jpg`
- `lucas.jpg`
- `gabriel.jpg`
- `pamela.jpg`

Na ausência da foto, o sistema exibe automaticamente um avatar com a inicial.

## Teste local

```bash
python -m http.server 8000
```

Abra `http://localhost:8000`.
