# Power Analytics v1.2 — Cadastro unificado

## Alteração principal
O `config/team.json` passa a ser a única fonte cadastral dos colaboradores.

Campos de perfil incorporados ao `team.json`:
- `admissionDate` — data de admissão (`AAAA-MM-DD`)
- `birthDayMonth` — aniversário (`DD/MM`)
- `terminationDate` — data de desligamento ou `null`
- `achievements` — conquistas manuais, quando houver

O arquivo `config/people-profile.json` não é mais utilizado pelo JavaScript e pode ser removido do repositório.

## Novo colaborador
1. Adicione a foto em `assets/img/team/`.
2. Cadastre a pessoa uma única vez em `config/team.json`.
3. Use `status: "active"` para ativo e `status: "inactive"` para desligado.
4. Os indicadores passam a aparecer quando existirem dados reais nas planilhas.

## Observações
- As metas atuais de taxa de avaliação permanecem em 30%.
- `Livre`, `Murilo` e demais nomes administrativos já filtrados continuam ignorados.
- O histórico continua associado pelo nome do colaborador, evitando transferência indevida quando um ramal é reutilizado.
