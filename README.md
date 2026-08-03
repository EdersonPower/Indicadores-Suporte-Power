# Power Analytics — Dashboard NPS

Dashboard estático para GitHub Pages que lê planilhas Excel diretamente no navegador.

## Testar localmente

Por segurança, navegadores podem bloquear a leitura automática de arquivos quando o `index.html` é aberto com duplo clique. Use um servidor local:

```bash
python -m http.server 8000
```

Depois abra `http://localhost:8000`.

Também é possível abrir o HTML e usar **Carregar Excel** manualmente.

## Publicar no GitHub Pages

1. Crie um repositório no GitHub.
2. Envie todos os arquivos e pastas deste projeto.
3. Acesse **Settings > Pages**.
4. Selecione **Deploy from a branch**, branch `main`, pasta `/root`.
5. Salve e abra o endereço fornecido pelo GitHub.

## Adicionar um novo trimestre

1. Copie o Excel para `data/2026/`, por exemplo `3trim2026.xlsx`.
2. Edite `data/manifest.json` e adicione:

```json
{"path":"data/2026/3trim2026.xlsx","label":"3º Trimestre 2026"}
```

O Excel deve conter abas mensais no padrão `NPS Jul`, `NPS Ago`, `NPS Set` e preferencialmente uma aba `NPS Trim` ou `NPS Trimestral`.

## Atualização sem editar o projeto

Use o botão **Carregar Excel**. A leitura ocorre somente no navegador; o arquivo não é enviado para servidor algum.

## Bibliotecas

- Tailwind CSS (CDN)
- Chart.js
- SheetJS
