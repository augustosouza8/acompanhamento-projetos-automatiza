"""
Aplicativo Flask simples (PoC) para gerenciar Projetos, Macroetapas, Etapas e Tarefas,
calculando datas automaticamente de baixo para cima.

Execução:
    - Instale dependências: pip install flask flask_sqlalchemy
    - Rode: python app.py
    - Acesse: http://127.0.0.1:5000
"""

from datetime import datetime
from flask import Flask, render_template, redirect, url_for, request, abort
from models import db, Project, MacroStage, Stage, Task, WeeklyUpdate

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///schedule.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# Inicializa o objeto db com o app
db.init_app(app)


PROJECT_STATUS_CHOICES = [
    "a iniciar",
    "em andamento",
    "concluído",
    "suspenso",
    "descartado",
]

MACRO_STRUCTURE_CHOICES = {"stages", "tasks"}
STAGE_TYPE_CHOICES = ["robô", "sistema", "não se aplica"]


def redirect_with_anchor(endpoint: str, anchor: str, **values):
    url = url_for(endpoint, **values)
    if anchor:
        url = f"{url}#{anchor}"
    return redirect(url)


def task_parent_context(task: Task):
    if task.stage:
        return task.stage.macrostage.project.id, f"stage-{task.stage.id}"
    return task.macrostage.project.id, f"macrostage-{task.macrostage.id}"


def parse_date_field(raw_value: str):
    """
    Converte uma string no formato YYYY-MM-DD em objeto date.
    Retorna None caso não informado ou inválido.
    """
    if not raw_value:
        return None

    try:
        return datetime.strptime(raw_value, "%Y-%m-%d").date()
    except ValueError:
        return None

# --------------------------
# Funções auxiliares de datas
# --------------------------

def recalculate_stage(stage: Stage) -> None:
    """
    Recalcula as datas de início e fim de uma Etapa (Stage) com base nas suas Tarefas.

    Regra:
        - start_date = menor data de início entre as tarefas
        - end_date   = maior data de fim entre as tarefas
        - se não houver tarefas com datas válidas, ambos ficam como None
    """
    if stage is None:
        return

    tasks = stage.tasks

    start_dates = [t.start_date for t in tasks if t.start_date is not None]
    end_dates = [t.end_date for t in tasks if t.end_date is not None]

    stage.start_date = min(start_dates) if start_dates else None
    stage.end_date = max(end_dates) if end_dates else None


def recalculate_macrostage(macrostage: MacroStage) -> None:
    """
    Recalcula as datas de início e fim de uma Macroetapa com base nas Etapas.

    Regra:
        - start_date = menor start_date entre as etapas
        - end_date   = maior end_date entre as etapas
        - se não houver etapas com datas válidas, ambos ficam como None
    """
    if macrostage is None:
        return

    stages = macrostage.stages
    tasks = macrostage.tasks

    start_dates = [s.start_date for s in stages if s.start_date is not None]
    end_dates = [s.end_date for s in stages if s.end_date is not None]

    task_starts = [t.start_date for t in tasks if t.start_date is not None]
    task_ends = [t.end_date for t in tasks if t.end_date is not None]

    start_dates.extend(task_starts)
    end_dates.extend(task_ends)

    macrostage.start_date = min(start_dates) if start_dates else None
    macrostage.end_date = max(end_dates) if end_dates else None


def recalculate_project(project: Project) -> None:
    """
    Recalcula as datas de início e fim de um Projeto com base nas Macroetapas.

    Regra:
        - start_date = menor start_date entre as macroetapas
        - end_date   = maior end_date entre as macroetapas
        - se não houver macroetapas com datas válidas, ambos ficam como None
    """
    if project is None:
        return

    macrostages = project.macrostages

    start_dates = [m.start_date for m in macrostages if m.start_date is not None]
    end_dates = [m.end_date for m in macrostages if m.end_date is not None]

    project.start_date = min(start_dates) if start_dates else None
    project.end_date = max(end_dates) if end_dates else None


def recalculate_all_from_stage(stage: Stage) -> None:
    """
    A partir de uma Etapa (Stage), recalcula em cadeia:
        - Stage (etapa)
        - MacroStage (macroetapa)
        - Project (projeto)
    e realiza o commit das alterações.
    """
    if stage is None:
        return

    macrostage = stage.macrostage
    project = macrostage.project if macrostage else None

    # Recalcula em cascata
    recalculate_stage(stage)
    recalculate_macrostage(macrostage)
    recalculate_project(project)

    db.session.commit()


def recalculate_all_from_macrostage(macrostage: MacroStage) -> None:
    if macrostage is None:
        return

    project = macrostage.project

    recalculate_macrostage(macrostage)
    recalculate_project(project)

    db.session.commit()


# --------------
# Rotas principais
# --------------

@app.route("/")
def index():
    """
    Redireciona para a listagem de projetos.
    """
    return redirect(url_for("list_projects"))


@app.route("/projects", methods=["GET"])
def list_projects():
    """
    Lista todos os projetos cadastrados, com suas datas calculadas.
    """
    projects = Project.query.order_by(Project.id).all()
    return render_template("projects.html", projects=projects)


@app.route("/projects/create", methods=["POST"])
def create_project():
    """
    Cria um novo Projeto com base no nome enviado via formulário.
    As datas de início e fim serão calculadas depois, a partir das macroetapas.
    """
    name = request.form.get("name", "").strip()
    if not name:
        # Em uma versão mais elaborada, poderíamos exibir mensagem de erro.
        return redirect(url_for("list_projects"))

    project = Project(name=name)
    db.session.add(project)
    db.session.commit()
    return redirect(url_for("list_projects"))


@app.route("/projects/<int:project_id>", methods=["GET"])
def project_detail(project_id: int):
    """
    Exibe a página de detalhe de um Projeto, com suas macroetapas, etapas e tarefas.
    Permite adicionar Macroetapas, Etapas e Tarefas via formulários simples.
    """
    project = Project.query.get_or_404(project_id)
    # As macroetapas e etapas serão acessadas via relacionamentos no template.
    return render_template(
        "project_detail.html",
        project=project,
        status_choices=PROJECT_STATUS_CHOICES,
        stage_type_choices=STAGE_TYPE_CHOICES,
    )


@app.route("/macrostages/create", methods=["POST"])
def create_macrostage():
    """
    Cria uma nova Macroetapa ligada a um Projeto.
    """
    project_id = request.form.get("project_id")
    name = request.form.get("name", "").strip()

    if not project_id or not name:
        return redirect(url_for("list_projects"))

    project = Project.query.get(project_id)
    if project is None:
        abort(404)

    existing_positions = [m.position or 0 for m in project.macrostages]
    next_position = (max(existing_positions) + 1) if existing_positions else 1

    macrostage = MacroStage(name=name, project=project, position=next_position)
    db.session.add(macrostage)
    db.session.commit()

    # Recalcula datas do projeto (caso a macroetapa venha a ter datas no futuro)
    recalculate_project(project)
    db.session.commit()

    return redirect_with_anchor("project_detail", f"macrostage-{macrostage.id}", project_id=project.id)


@app.route("/stages/create", methods=["POST"])
def create_stage():
    """
    Cria uma nova Etapa ligada a uma Macroetapa.
    """
    macrostage_id = request.form.get("macrostage_id")
    name = request.form.get("name", "").strip()
    stage_type = (request.form.get("stage_type", "") or "não se aplica").strip()
    scope = request.form.get("scope", "").strip()
    tools_selected = request.form.getlist("tools")
    other_tools = request.form.get("other_tools", "").strip()

    if stage_type not in STAGE_TYPE_CHOICES:
        stage_type = "não se aplica"

    if stage_type not in ("robô", "sistema"):
        scope = None
        tools_selected = []
        other_tools = None
    else:
        scope = scope or None
        other_tools = other_tools or None

    tools_str = ",".join(tools_selected) if tools_selected else None

    if not macrostage_id or not name:
        return redirect(url_for("list_projects"))

    macrostage = MacroStage.query.get(macrostage_id)
    if macrostage is None:
        abort(404)

    if macrostage.structure_type == "tasks":
        return redirect(url_for("project_detail", project_id=macrostage.project.id))

    if macrostage.structure_type is None:
        macrostage.structure_type = "stages"

    existing_positions = [s.position or 0 for s in macrostage.stages]
    next_position = (max(existing_positions) + 1) if existing_positions else 1

    stage = Stage(
        name=name,
        macrostage=macrostage,
        position=next_position,
        stage_type=stage_type,
        scope=scope,
        tools=tools_str,
        other_tools=other_tools,
    )
    db.session.add(stage)
    db.session.commit()

    # Recalcula datas da macroetapa/projeto (ainda não há tarefas, então provavelmente continuará None)
    recalculate_macrostage(macrostage)
    recalculate_project(macrostage.project)
    db.session.commit()

    return redirect_with_anchor("project_detail", f"stage-{stage.id}", project_id=macrostage.project.id)


@app.route("/tasks/create", methods=["POST"])
def create_task():
    """
    Cria uma nova Tarefa ligada a uma Etapa.

    Recebe do formulário:
        - stage_id
        - name
        - start_date (YYYY-MM-DD)
        - end_date   (YYYY-MM-DD)
    """
    stage_id = request.form.get("stage_id")
    macrostage_id = request.form.get("macrostage_id")
    name = request.form.get("name", "").strip()
    start_date_str = request.form.get("start_date", "").strip()
    end_date_str = request.form.get("end_date", "").strip()

    if not name:
        return redirect(url_for("list_projects"))

    stage = Stage.query.get(stage_id) if stage_id else None
    if stage_id and stage is None:
        abort(404)

    if stage:
        macrostage = stage.macrostage
    else:
        if not macrostage_id:
            return redirect(url_for("list_projects"))
        macrostage = MacroStage.query.get(macrostage_id)
        if macrostage is None:
            abort(404)

    # Conversão das strings de data (HTML <input type="date">) para objetos date
    start_date = parse_date_field(start_date_str)
    end_date = parse_date_field(end_date_str)

    if stage:
        if stage.macrostage.structure_type == "tasks":
            return redirect(url_for("project_detail", project_id=stage.macrostage.project.id))
        existing_positions = [t.position or 0 for t in stage.tasks]
    else:
        if macrostage.structure_type == "stages":
            return redirect(url_for("project_detail", project_id=macrostage.project.id))
        if macrostage.structure_type is None:
            macrostage.structure_type = "tasks"
        existing_positions = [t.position or 0 for t in macrostage.tasks if t.stage is None]

    next_position = (max(existing_positions) + 1) if existing_positions else 1

    task = Task(
        name=name,
        stage=stage,
        macrostage=macrostage,
        start_date=start_date,
        end_date=end_date,
        position=next_position,
    )
    db.session.add(task)
    db.session.commit()

    # Recalcula datas em cascata (etapa, macroetapa, projeto)
    if stage:
        recalculate_all_from_stage(stage)
        anchor = f"stage-{stage.id}"
    else:
        recalculate_all_from_macrostage(macrostage)
        anchor = f"macrostage-{macrostage.id}"

    return redirect_with_anchor("project_detail", anchor, project_id=macrostage.project.id)


@app.route("/projects/<int:project_id>/update", methods=["POST"])
def update_project(project_id: int):
    """
    Atualiza as informações gerais de um projeto.
    """
    project = Project.query.get_or_404(project_id)
    name = request.form.get("name", "").strip()

    if name:
        project.name = name

    project.scope = request.form.get("scope", "").strip() or None
    status = request.form.get("status", "").strip()
    project.status = status if status in PROJECT_STATUS_CHOICES else None
    project.github_link = request.form.get("github_link", "").strip() or None
    project.coordinator = request.form.get("coordinator", "").strip() or None
    project.automation_support = request.form.get("automation_support", "").strip() or None
    project.requesting_agency = request.form.get("requesting_agency", "").strip() or None
    project.internal_department = request.form.get("internal_department", "").strip() or None
    project.sponsoring_manager = request.form.get("sponsoring_manager", "").strip() or None
    project.sponsoring_manager_contact = request.form.get("sponsoring_manager_contact", "").strip() or None
    project.technical_manager = request.form.get("technical_manager", "").strip() or None
    project.technical_manager_contact = request.form.get("technical_manager_contact", "").strip() or None

    db.session.commit()

    return redirect(url_for("project_detail", project_id=project.id))


@app.route("/projects/<int:project_id>/delete", methods=["POST"])
def delete_project(project_id: int):
    """
    Exclui um projeto e todos os seus itens relacionados.
    """
    project = Project.query.get_or_404(project_id)
    db.session.delete(project)
    db.session.commit()
    return redirect(url_for("list_projects"))


@app.route("/macrostages/<int:macrostage_id>/update", methods=["POST"])
def update_macrostage(macrostage_id: int):
    """
    Atualiza o nome de uma macroetapa.
    """
    macrostage = MacroStage.query.get_or_404(macrostage_id)
    name = request.form.get("name", "").strip()

    if name:
        macrostage.name = name
        db.session.commit()

    return redirect_with_anchor("project_detail", f"macrostage-{macrostage.id}", project_id=macrostage.project.id)


@app.route("/macrostages/<int:macrostage_id>/delete", methods=["POST"])
def delete_macrostage(macrostage_id: int):
    """
    Remove uma macroetapa.
    """
    macrostage = MacroStage.query.get_or_404(macrostage_id)
    project_id = macrostage.project.id
    db.session.delete(macrostage)
    db.session.commit()

    project = Project.query.get(project_id)
    if project:
        recalculate_project(project)
        db.session.commit()

    return redirect_with_anchor("project_detail", f"macrostage-{macrostage_id}", project_id=project_id)


@app.route("/stages/<int:stage_id>/update", methods=["POST"])
def update_stage(stage_id: int):
    """
    Atualiza as informações de uma etapa.
    """
    stage = Stage.query.get_or_404(stage_id)
    name = request.form.get("name", "").strip()
    if name:
        stage.name = name

    stage_type = (request.form.get("stage_type", "") or "não se aplica").strip()
    if stage_type not in STAGE_TYPE_CHOICES:
        stage_type = "não se aplica"

    stage.stage_type = stage_type

    if stage_type in ("robô", "sistema"):
        stage.scope = request.form.get("scope", "").strip() or None
        tools_selected = request.form.getlist("tools")
        stage.tools = ",".join(tools_selected) if tools_selected else None
        stage.other_tools = request.form.get("other_tools", "").strip() or None
    else:
        stage.scope = None
        stage.tools = None
        stage.other_tools = None

    db.session.commit()

    return redirect_with_anchor(
        "project_detail",
        f"stage-{stage.id}",
        project_id=stage.macrostage.project.id,
    )


@app.route("/stages/<int:stage_id>/delete", methods=["POST"])
def delete_stage(stage_id: int):
    """
    Remove uma etapa e suas tarefas.
    """
    stage = Stage.query.get_or_404(stage_id)
    macrostage_id = stage.macrostage.id
    project_id = stage.macrostage.project.id

    for task in list(stage.tasks):
        db.session.delete(task)

    db.session.delete(stage)
    db.session.commit()

    macrostage = MacroStage.query.get(macrostage_id)
    project = Project.query.get(project_id)

    if macrostage:
        recalculate_macrostage(macrostage)
    if project:
        recalculate_project(project)
    db.session.commit()

    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/projects/<int:project_id>/macrostages/reorder", methods=["POST"])
def reorder_macrostages(project_id: int):
    """
    Atualiza a ordem das macroetapas de um projeto específico.
    """
    project = Project.query.get_or_404(project_id)
    payload = request.get_json(silent=True) or {}
    order = payload.get("order")

    if not isinstance(order, list):
        return {"status": "error", "message": "Formato inválido"}, 400

    macrostage_map = {m.id: m for m in project.macrostages}

    position = 1
    for macrostage_id in order:
        try:
            macrostage_id = int(macrostage_id)
        except (TypeError, ValueError):
            continue

        macrostage = macrostage_map.pop(macrostage_id, None)
        if macrostage is None:
            continue

        macrostage.position = position
        position += 1

    # Aplica ordem para quaisquer macroetapas não incluídas no payload
    for macrostage in macrostage_map.values():
        macrostage.position = position
        position += 1

    db.session.commit()
    return {"status": "ok"}


@app.route("/stages/<int:macrostage_id>/reorder", methods=["POST"])
def reorder_stages(macrostage_id: int):
    """
    Atualiza a ordem das etapas dentro de uma macroetapa específica.
    """
    macrostage = MacroStage.query.get_or_404(macrostage_id)
    payload = request.get_json(silent=True) or {}
    order = payload.get("order")

    if not isinstance(order, list):
        return {"status": "error", "message": "Formato inválido"}, 400

    stage_map = {s.id: s for s in macrostage.stages}

    position = 1
    for stage_id in order:
        try:
            stage_id = int(stage_id)
        except (TypeError, ValueError):
            continue

        stage = stage_map.pop(stage_id, None)
        if stage is None:
            continue

        stage.position = position
        position += 1

    for stage in stage_map.values():
        stage.position = position
        position += 1

    db.session.commit()
    return {"status": "ok"}


@app.route("/tasks/<int:stage_id>/reorder", methods=["POST"])
def reorder_tasks(stage_id: int):
    """
    Atualiza a ordem das tarefas dentro de uma etapa específica.
    """
    stage = Stage.query.get_or_404(stage_id)
    payload = request.get_json(silent=True) or {}
    order = payload.get("order")

    if not isinstance(order, list):
        return {"status": "error", "message": "Formato inválido"}, 400

    task_map = {t.id: t for t in stage.tasks}

    position = 1
    for task_id in order:
        try:
            task_id = int(task_id)
        except (TypeError, ValueError):
            continue

        task = task_map.pop(task_id, None)
        if task is None:
            continue

        task.position = position
        position += 1

    for task in task_map.values():
        task.position = position
        position += 1

    db.session.commit()
    return {"status": "ok"}


@app.route("/macrostages/<int:macrostage_id>/tasks/reorder", methods=["POST"])
def reorder_macrostage_tasks(macrostage_id: int):
    """
    Atualiza a ordem das tarefas diretamente ligadas a uma macroetapa.
    """
    macrostage = MacroStage.query.get_or_404(macrostage_id)
    payload = request.get_json(silent=True) or {}
    order = payload.get("order")

    if not isinstance(order, list):
        return {"status": "error", "message": "Formato inválido"}, 400

    direct_tasks = {t.id: t for t in macrostage.tasks if t.stage_id is None}

    position = 1
    for task_id in order:
        try:
            task_id = int(task_id)
        except (TypeError, ValueError):
            continue

        task = direct_tasks.pop(task_id, None)
        if task is None:
            continue

        task.position = position
        position += 1

    for task in direct_tasks.values():
        task.position = position
        position += 1

    db.session.commit()
    return {"status": "ok"}


@app.route("/macrostages/<int:macrostage_id>/structure", methods=["POST"])
def set_macrostage_structure(macrostage_id: int):
    """Define se a macroetapa usará etapas ou tarefas diretas."""
    macrostage = MacroStage.query.get_or_404(macrostage_id)
    requested = (request.form.get("structure_type") or "").strip()

    anchor = f"macrostage-{macrostage.id}"

    if requested not in MACRO_STRUCTURE_CHOICES:
        return redirect_with_anchor("project_detail", anchor, project_id=macrostage.project.id)

    direct_tasks = [t for t in macrostage.tasks if t.stage_id is None]

    if requested == "stages" and direct_tasks:
        return redirect_with_anchor("project_detail", anchor, project_id=macrostage.project.id)

    if requested == "tasks" and macrostage.stages:
        return redirect_with_anchor("project_detail", anchor, project_id=macrostage.project.id)

    macrostage.structure_type = requested
    db.session.commit()
    return redirect_with_anchor("project_detail", anchor, project_id=macrostage.project.id)


@app.route("/tasks/<int:task_id>/update", methods=["POST"])
def update_task(task_id: int):
    """
    Atualiza uma tarefa, incluindo datas de início e fim.
    """
    task = Task.query.get_or_404(task_id)
    stage = task.stage
    macrostage = task.macrostage

    name = request.form.get("name", "").strip()
    start_date_str = request.form.get("start_date", "").strip()
    end_date_str = request.form.get("end_date", "").strip()

    if name:
        task.name = name

    task.start_date = parse_date_field(start_date_str)
    task.end_date = parse_date_field(end_date_str)

    if stage:
        recalculate_all_from_stage(stage)
        project_id = stage.macrostage.project.id
        anchor = f"stage-{stage.id}"
    else:
        recalculate_all_from_macrostage(macrostage)
        project_id = macrostage.project.id
        anchor = f"macrostage-{macrostage.id}"

    return redirect_with_anchor("project_detail", anchor, project_id=project_id)


@app.route("/tasks/<int:task_id>/delete", methods=["POST"])
def delete_task(task_id: int):
    """
    Remove uma tarefa e recalcula as datas associadas.
    """
    task = Task.query.get_or_404(task_id)
    stage_id = task.stage_id
    macrostage_id = task.macrostage_id
    project_id = task.macrostage.project.id

    db.session.delete(task)
    db.session.commit()

    stage = Stage.query.get(stage_id)
    if stage:
        recalculate_all_from_stage(stage)
        anchor = f"stage-{stage.id}"
    else:
        macrostage = MacroStage.query.get(macrostage_id)
        if macrostage:
            recalculate_all_from_macrostage(macrostage)
            anchor = f"macrostage-{macrostage.id}"
        else:
            anchor = ""

    return redirect_with_anchor("project_detail", anchor, project_id=project_id)


@app.route("/tasks/<int:task_id>/weekly_updates/create", methods=["POST"])
def create_weekly_update(task_id: int):
    """
    Cria uma nova atualização semanal para a tarefa informada.
    """
    task = Task.query.get_or_404(task_id)
    content = request.form.get("content", "").strip()
    update_date_str = request.form.get("update_date", "").strip()

    if not content:
        project_id, anchor = task_parent_context(task)
        return redirect_with_anchor("project_detail", anchor, project_id=project_id)

    update = WeeklyUpdate(
        task=task,
        content=content,
        update_date=parse_date_field(update_date_str),
    )
    db.session.add(update)
    db.session.commit()

    project_id, anchor = task_parent_context(task)
    return redirect_with_anchor("project_detail", anchor, project_id=project_id)


@app.route("/weekly_updates/<int:update_id>/update", methods=["POST"])
def update_weekly_update(update_id: int):
    """
    Atualiza o conteúdo ou a data de uma atualização semanal.
    """
    weekly_update = WeeklyUpdate.query.get_or_404(update_id)

    content = request.form.get("content", "").strip()
    update_date_str = request.form.get("update_date", "").strip()

    if content:
        weekly_update.content = content
    weekly_update.update_date = parse_date_field(update_date_str)

    db.session.commit()

    project_id, anchor = task_parent_context(weekly_update.task)
    return redirect_with_anchor("project_detail", anchor, project_id=project_id)


@app.route("/weekly_updates/<int:update_id>/delete", methods=["POST"])
def delete_weekly_update(update_id: int):
    """
    Remove uma atualização semanal específica.
    """
    weekly_update = WeeklyUpdate.query.get_or_404(update_id)
    project_id, anchor = task_parent_context(weekly_update.task)

    db.session.delete(weekly_update)
    db.session.commit()

    return redirect_with_anchor("project_detail", anchor, project_id=project_id)


if __name__ == "__main__":
    # Cria as tabelas antes de iniciar o servidor (substitui o before_first_request)
    with app.app_context():
        db.create_all()

    # Executa o app Flask em modo debug para facilitar testes
    app.run(debug=True)
