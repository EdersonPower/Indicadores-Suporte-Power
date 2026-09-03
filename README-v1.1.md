# Power Analytics v1.1 — Perfil & Conquistas

Este pacote é uma atualização incremental. Ele NÃO contém suas planilhas, fotos nem `config/team.json`, para não sobrescrever os dados atuais do GitHub.

## Arquivos para substituir/adicionar
- `index.html` — substituir
- `assets/js/bundle.js` — substituir
- `assets/css/main.css` — substituir
- `config/people-profile.json` — adicionar

Depois do commit, faça Ctrl+F5 no navegador.

## O que entrou na v1.1
- Nova página pública **Perfil & Conquistas**.
- Data de aniversário (dia/mês).
- Data de admissão e tempo de empresa calculado automaticamente pela data atual.
- Próximos aniversários e aniversários de empresa (janela de 45 dias).
- Mural automático de conquistas da Caixinha a partir das planilhas históricas.
- Conquistas automáticas a cada ano completo de empresa.
- Somente colaboradores ativos entram no mural.
- Em Gestão da Equipe agora é possível informar admissão, aniversário e data de desligamento; alterações locais seguem o mesmo padrão de localStorage já usado pelo projeto.

## Datas pré-carregadas de `aniversarios.xlsx`
- Dayane — admissão 03/08/2022 — aniversário 12/07
- Lukas — admissão 01/04/2024 — aniversário 13/04
- Lucas — admissão 07/04/2025 — aniversário 20/08
- Ryan — admissão 22/06/2026 — aniversário 05/08
- Thamies — admissão 31/08/2026 — aniversário 16/05

Murilo será reconhecido pelo cadastro/dados atuais do seu GitHub, mas as datas dele aparecerão como não cadastradas até serem informadas na Gestão da Equipe ou adicionadas a `config/people-profile.json`.
