# Power Analytics — Sprint 4.0

## Novidades
- Cadastro automático de colaboradores encontrados nas planilhas.
- Arquivo `config/team.json` com ramal, nome, cargo, foto, status e metas.
- Central do Gestor com progresso das metas, alertas e destaques.
- Gestão da Equipe pela interface.
- Status ativo/desligado.
- Metas individuais de atendimentos, avaliações, taxa, média final e Caixinha.
- Alterações salvas no navegador por `localStorage`.
- Colaboradores desligados permanecem no histórico, mas saem dos rankings atuais.

## Publicação
Substitua todos os arquivos do repositório pelo conteúdo deste ZIP e faça um commit.
Depois pressione `Ctrl + F5`.


## Correção 4.1

- Fotos de Gabriel e Luiz incluídas.
- Associação das fotos mantida pelo nome do colaborador.
- Configurações antigas do navegador são migradas sem sobrescrever os caminhos oficiais das fotos.
- Não é necessário limpar o Local Storage para visualizar as imagens.


## Correção 4.2 — Importação sem duplicidade

- Cada período utiliza a chave única `ano + trimestre`.
- Ao importar novamente um trimestre existente, os dados anteriores são substituídos.
- Novos trimestres continuam sendo adicionados normalmente.
- Duplicidades existentes são removidas durante o carregamento inicial.
- A lista permanece ordenada por ano e trimestre.
- O sistema informa se o período foi atualizado ou adicionado.
- O campo de seleção de arquivo é limpo após a importação, permitindo importar novamente o mesmo arquivo.


## Sprint 4.4 — Colaborador do período

- Foto circular com proporção fixa 1:1.
- `object-fit: cover` para evitar distorção ou formato oval.
- Layout horizontal com foto e descrição à esquerda.
- Indicadores compactos alinhados à direita.
- Adaptação responsiva para notebook, tablet e celular.


## Correção 4.4.1

- Incluída a função `executivePeriodLabel`.
- Corrigido o erro que impedia a montagem do dashboard.
- Mantido o novo layout circular do colaborador em destaque.
