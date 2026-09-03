# Power Analytics v1.1.1 — Correção de colaboradores

Atualização incremental sobre a v1.1.

## Regras corrigidas

- `Livre` nunca é tratado como colaborador.
- `Murilo` é ignorado integralmente em perfis, rankings, indicadores, Caixinha, bonificações, avaliações e conquistas.
- `Thamires` é a colaboradora válida, com admissão em 31/08/2026.
- O alias antigo `Thamies` é normalizado automaticamente para `Thamires` para evitar duplicidades.
- Thamires aparece em Perfil & Conquistas mesmo antes de possuir atendimentos.
- Indicadores individuais só aparecem quando existirem registros reais nas planilhas.
- Valores administrativos `Sem colaborador`, `Sem atendente`, `Vago` e `Disponível` também são ignorados.

## Arquivos para atualizar no GitHub

Substitua/adicone somente:

- `assets/js/bundle.js`
- `config/people-profile.json`
- `README-v1.1.1.md` (documentação; não interfere no dashboard)

Não apague Excel, fotos, `manifest.json`, `team.json` ou demais arquivos do projeto.
