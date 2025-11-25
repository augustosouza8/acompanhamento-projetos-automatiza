# Acompanhamento de Projetos – PoC em Flask

Aplicação web usada para organizar projetos em quatro níveis hierárquicos (Projeto → Macroetapa → Etapa → Tarefa), calcular automaticamente datas agregadas, registrar metadados por nível e anotar atualizações semanais. A UI foi mantida propositalmente simples (HTML + CSS inline) para facilitar inspeções por humanos e por LLMs.

---

## 1. Hierarquia e regras de negócio

### 1.1. Níveis hierárquicos

1. **Projeto** – contém macroetapas. Possui metadados gerais (escopo, status, links, responsáveis, contatos) e datas calculadas a partir das macroetapas filhas.
2. **Macroetapa** – pode operar em dois modos exclusivos:
   - **`structure_type = "stages"`** → exige etapas; tarefas ficam dentro das etapas.
   - **`structure_type = "tasks"`** → tarefas são associadas diretamente à macroetapa; não há etapas.
   O usuário escolhe o modo ao clicar em “Criar etapas e depois tarefas” ou “Adicionar tarefas diretamente” e o sistema não permite alternar para um modo incompatível com as entidades já existentes (ex.: impedir trocar para “stages” se houver tarefas diretas).
3. **Etapa** – pertence a uma macroetapa. Ao ser criada, o usuário responde se “Trata-se de um robô ou um sistema?”. As opções são `"robô"`, `"sistema"` ou `"não se aplica"`. Para robô/sistema, os campos de escopo e ferramentas tornam-se obrigatórios; caso contrário, permanecem ocultos.
4. **Tarefa** – nível mais granular. Sempre possui `macrostage_id` e, opcionalmente, `stage_id`. Todas as datas são inseridas neste nível.

### 1.2. Regras de data

1. **Tarefa** – recebe `start_date` e `end_date` diretamente do usuário (opcionais).
2. **Etapa** – `start_date` = menor data das tarefas filhas (desconsiderando `NULL`); `end_date` = maior data das tarefas filhas. Se não houver tarefas datadas, ambos permanecem `NULL`.
3. **Macroetapa** – coleta datas das etapas e das tarefas diretas (quando `structure_type = "tasks"`). Novamente usa o menor início / maior fim.
4. **Projeto** – calcula com base nas macroetapas (menor início / maior fim).

Essas regras estão concentradas em `recalculate_stage`, `recalculate_macrostage`, `recalculate_project` e são invocadas sempre após `create/update/delete` de tarefas, etapas ou macroetapas. O fluxo nunca grava datas manualmente em níveis superiores.

### 1.3. Metadados e campos auxiliares

- **Projeto:** escopo, status, GitHub, coordenador, pessoas de apoio, órgão demandante, setor interno, gestor responsável + contato, gestor técnico + contato.
- **Macroetapa:** `structure_type`, `position` (para ordenação) e datas calculadas.
- **Etapa:** `stage_type`, `scope`, `tools` (lista separada por vírgula) e `other_tools` (texto livre). Esses campos são exibidos logo abaixo do nome da etapa quando `stage_type` ∈ {robô, sistema}.
- **Tarefa:** `position` e datas. Ao clicar em “Atualizações semanais” surge uma tabela com entradas `WeeklyUpdate` (texto + data), também editáveis/excluíveis inline.

---

## 2. Rotas principais

| Método/rota                               | Descrição                                                                 | Observações claves                                                                          |
|-------------------------------------------|----------------------------------------------------------------------------|---------------------------------------------------------------------------------------------|
| `GET /projects`                           | Lista projetos e link para o detalhe                                     | `templates/projects.html`                                                                   |
| `POST /projects/create`                   | Cria projeto (nome obrigatório)                                           | Metadados adicionais são editados na tela detalhada                                         |
| `GET /projects/<project_id>`              | Exibe árvore completa do projeto                                          | `templates/project_detail.html`                                                             |
| `POST /macrostages/create`                | Adiciona macroetapa ao projeto                                            | Após criar, o usuário escolhe o tipo de estrutura                                           |
| `POST /macrostages/<id>/structure`        | Define `structure_type` (etapas ou tarefas diretas)                       | Usa âncoras para manter o scroll                                                            |
| `POST /stages/create`                     | Cria etapa (apenas quando `structure_type="stages"`)                      | Processa `stage_type`, escopo e ferramentas                                                 |
| `POST /tasks/create`                      | Cria tarefa ligada à etapa ou direto à macroetapa                         | Se `stage_id` é informado, usa âncora `#stage-…`; caso contrário, `#macrostage-…`           |
| `POST /macrostages/<id>/tasks/reorder`    | Reordena tarefas diretas da macroetapa                                    | Tabelas HTML com `draggable` alimentam esta rota                                            |
| `POST /stages/<id>/reorder`               | Reordena etapas dentro da macroetapa                                      | Idem                                                                                        |
| `POST /tasks/<stage_id>/reorder`          | Reordena tarefas dentro da etapa                                          | Idem                                                                                        |
| `POST /tasks/<task_id>/weekly_updates/*`  | CRUD de atualizações semanais                                             | Sempre recalcula datas e redireciona para `#task-…`                                         |
| `POST /macrostages/<id>/reorder` etc.     | Outras rotas de reorder e delete (ver `app.py`)                           | Todas usam âncoras                                                                          |


---

## 3. Estrutura de arquivos

```
acompanhamento-projetos-automatiza/
├── app.py                 # Rotas Flask, orquestração e regras de negócio
├── models.py              # Definição dos modelos SQLAlchemy
├── migrations/            # Scripts utilitários para bancos já existentes
├── templates/             # HTML + CSS + JS simples (base, lista de projetos, tela detalhada)
├── pyproject.toml         # Configuração do projeto (usado por uv/pip)
├── requirements.txt       # Dependências para deploy no Render
└── instance/schedule.db   # Banco SQLite (versionado no Git para deploy inicial)
```

> **Migrations:** arquivos `001`–`007` são idempotentes e servem apenas para atualizar bancos legados. Ambientes novos podem ignorá-los porque `app.py` chama `db.create_all()` na inicialização.

---

## 4. Fluxo da interface

1. Acesse `/projects` e cadastre um projeto.
2. Dentro do projeto, utilize o botão “Adicionar macroetapa”.
3. Ao ver a macroetapa recém-criada, escolha uma das opções:
   - **Criar etapas e depois tarefas** → o formulário de etapa aparece logo abaixo, com a pergunta sobre robô/sistema. Após salvar, as etapas ficam listadas numa `<ul>` com drag-and-drop.
   - **Adicionar tarefas diretamente** → habilita uma tabela dedicada a tarefas diretas. Cada linha inclui botões para editar (com inputs inline) e “Atualizações semanais”.
4. Digite datas apenas nas tarefas. As seções superiores atualizam automaticamente.
5. Use os ícones de lápis e as tabelas inline para editar qualquer nível; todos os formulários redirecionam com âncora (`#macrostage-XX`, `#stage-YY`, `#task-ZZ`).

---

## 5. Regras de cálculo (detalhadas)

| Contexto                 | Algoritmo                                                                                                                              | Função responsável                    |
|--------------------------|----------------------------------------------------------------------------------------------------------------------------------------|---------------------------------------|
| Tarefa                   | Datas informadas manualmente. Se o usuário deixar em branco, o campo correspondente permanece `NULL`.                                 | `create_task`, `update_task`          |
| Etapa (`Stage`)          | `start_date = min(task.start_date)` / `end_date = max(task.end_date)` considerando apenas tarefas com valor válido.                   | `recalculate_stage`                   |
| Macroetapa (`MacroStage`)| Mesma lógica, porém considerando **todas** as etapas + todas as tarefas diretas (quando `structure_type="tasks"`).                   | `recalculate_macrostage`              |
| Projeto (`Project`)      | `start_date = min(macrostage.start_date)` / `end_date = max(macrostage.end_date)` ignorando `NULL`.                                   | `recalculate_project`                 |
| Atualizações semanais    | Não influenciam datas. Servem apenas para histórico textual.                                                                           | `create_weekly_update` e derivados    |
| Reordenação              | Cada nível possui `position`. Rotas `/reorder` recebem um `order` (lista de IDs) e atualizam as posições sequencialmente.             | `reorder_macrostages`, `reorder_*`    |

Os recalculados são disparados após qualquer operação que modifique tarefas/etapas/macroetapas. O commit sempre ocorre após `recalculate_*` para garantir consistência.

---

## 6. Execução do projeto

### Via `uv`

```bash
uv venv --python 3.13.7
source .venv/bin/activate         # Windows: .venv\Scripts\activate
uv pip install -r <(uv pip compile pyproject.toml)
uv run app.py
```

### Via `pip`

```bash
python -m venv .venv
source .venv/bin/activate         # Windows: .venv\Scripts\activate
pip install flask flask_sqlalchemy
python app.py
```

Servidor disponível em `http://127.0.0.1:5000`.

> **Para bancos existentes:** execute os arquivos `migrations/<NNN>_*.py` na ordem correta (por ex.: `python migrations/005_allow_tasks_without_stage.py`). Para novos ambientes, não é necessário executar migrations manualmente.

---

## 6.1. Deploy no Render (PoC)

Esta seção descreve como fazer o deploy da aplicação no Render usando SQLite para testes/PoC.

### Pré-requisitos

1. **Garantir que o `schedule.db` está no repositório:**
   - O arquivo `instance/schedule.db` deve estar versionado no Git
   - Se ainda não estiver, após ajustar o `.gitignore`, execute:
     ```bash
     git add -f instance/schedule.db
     git commit -m "Adiciona schedule.db inicial para deploy"
     git push
     ```

### Passos para deploy

1. **Acesse o [Render Dashboard](https://dashboard.render.com/)**
   - Faça login ou crie uma conta

2. **Crie um novo Web Service:**
   - Clique em "New +" → "Web Service"
   - Conecte seu repositório GitHub (autorize o Render se necessário)
   - Selecione o repositório `acompanhamento-projetos-automatiza`

3. **Configure o serviço:**
   - **Name:** escolha um nome (ex: `acompanhamento-projetos`)
   - **Environment:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python app.py`
   - **Plan:** Free (suficiente para PoC)

4. **Deploy:**
   - Clique em "Create Web Service"
   - O Render iniciará o build automaticamente
   - Aguarde o deploy completar (pode levar alguns minutos)

5. **Acesse a aplicação:**
   - Após o deploy, o Render fornecerá uma URL (ex: `https://acompanhamento-projetos.onrender.com`)
   - Acesse a URL e você será redirecionado para a página de projetos/dashboard
   - O banco `instance/schedule.db` do repositório será usado como base inicial

### Importante: Sistema de arquivos efêmero

⚠️ **Atenção:** O sistema de arquivos do Render é **efêmero**. Isso significa que:

- Alterações feitas no banco SQLite durante os testes **podem ser perdidas** em novos deploys
- Cada vez que o serviço reiniciar ou você fizer um novo deploy, o banco volta ao estado do `schedule.db` versionado no Git
- Isso é **aceitável para esta fase de PoC/testes**, mas para produção você deve considerar migrar para PostgreSQL (disponível no Render)

### Atualizando o banco inicial

Se você quiser atualizar o snapshot do banco que será usado no deploy:

1. Faça as alterações localmente
2. Copie o arquivo atualizado:
   ```bash
   cp instance/schedule.db instance/schedule.db
   ```
3. Adicione ao Git e faça commit:
   ```bash
   git add -f instance/schedule.db
   git commit -m "Atualiza snapshot do banco para deploy"
   git push
   ```
4. O Render fará um novo deploy automaticamente (ou você pode acionar manualmente)

---

## 7. Considerações para manutenção e LLMs

- Toda a lógica está em `app.py`; não há Blueprints. LLMs podem seguir o arquivo sequencialmente para entender cada rota.
- IDs HTML (`macrostage-<id>`, `stage-<id>`, `task-<id>`) são críticos para os redirecionamentos com âncora. Ao adicionar novos blocos, mantenha esse padrão.
- Regras de data estão centralizadas em `recalculate_*`. Alterações de negócio devem ser feitas ali para evitar divergências.
- Operações de banco compartilham a helper `redirect_with_anchor`; use-a sempre que houver um `return redirect`. Isso evita que o usuário seja levado ao topo da página.
- O JS de drag-and-drop fica em `templates/project_detail.html` (bloco `extra_scripts`). Ele está desacoplado o suficiente para permitir substituição por bibliotecas mais sofisticadas se necessário.

---

## 8. Roadmap sugerido

1. Autenticação e perfis por equipe.
2. Exportação de dados em CSV/Excel e APIs REST.
3. Interface responsiva e bibliotecas de drag-and-drop com suporte móvel.
4. Testes automatizados para as funções de recalculo.
5. Internacionalização (UI + datas) caso outros idiomas sejam necessários.

Contribuições são bem-vindas; abra uma issue ou PR para discutir melhorias.
