"""
Definição dos modelos de dados usando SQLAlchemy (via Flask-SQLAlchemy).

Hierarquia:
    Projeto (Project)
        -> Macroetapas (MacroStage)
            -> Etapas (Stage)
                -> Tarefas (Task)
"""

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Project(db.Model):
    """
    Representa um Projeto.

    Campos:
        - id          : identificador único
        - name        : nome do projeto
        - start_date  : data de início (calculada com base nas macroetapas)
        - end_date    : data de fim (calculada com base nas macroetapas)
    """

    __tablename__ = "projects"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)

    start_date = db.Column(db.Date, nullable=True)
    end_date = db.Column(db.Date, nullable=True)

    # Metadados adicionais
    scope = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(50), nullable=True)
    github_link = db.Column(db.String(255), nullable=True)
    coordinator = db.Column(db.String(255), nullable=True)
    automation_support = db.Column(db.String(255), nullable=True)
    requesting_agency = db.Column(db.String(255), nullable=True)
    internal_department = db.Column(db.String(255), nullable=True)
    sponsoring_manager = db.Column(db.String(255), nullable=True)
    sponsoring_manager_contact = db.Column(db.String(255), nullable=True)
    technical_manager = db.Column(db.String(255), nullable=True)
    technical_manager_contact = db.Column(db.String(255), nullable=True)

    # Um projeto possui várias macroetapas
    macrostages = db.relationship(
        "MacroStage",
        back_populates="project",
        cascade="all, delete-orphan",
        order_by="MacroStage.position, MacroStage.id",
    )

    def __repr__(self) -> str:
        return f"<Project id={self.id} name={self.name!r}>"


class MacroStage(db.Model):
    """
    Representa uma Macroetapa dentro de um Projeto.

    Campos:
        - id          : identificador único
        - project_id  : referência ao projeto pai
        - name        : nome da macroetapa
        - start_date  : data de início (calculada com base nas etapas)
        - end_date    : data de fim (calculada com base nas etapas)
    """

    __tablename__ = "macrostages"

    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)
    name = db.Column(db.String(255), nullable=False)
    position = db.Column(db.Integer, nullable=False, default=0)
    structure_type = db.Column(db.String(20), nullable=True)

    start_date = db.Column(db.Date, nullable=True)
    end_date = db.Column(db.Date, nullable=True)

    # Relação com Project
    project = db.relationship("Project", back_populates="macrostages")

    # Uma macroetapa possui várias etapas
    stages = db.relationship(
        "Stage",
        back_populates="macrostage",
        cascade="all, delete-orphan",
        order_by="Stage.position, Stage.id",
    )

    # Macroetapa também pode possuir tarefas diretamente
    tasks = db.relationship(
        "Task",
        back_populates="macrostage",
        cascade="all, delete-orphan",
        order_by="Task.position, Task.id",
    )

    def __repr__(self) -> str:
        return f"<MacroStage id={self.id} name={self.name!r} project_id={self.project_id}>"


class Stage(db.Model):
    """
    Representa uma Etapa dentro de uma Macroetapa.

    Campos:
        - id             : identificador único
        - macrostage_id  : referência à macroetapa pai
        - name           : nome da etapa
        - start_date     : data de início (calculada com base nas tarefas)
        - end_date       : data de fim (calculada com base nas tarefas)
    """

    __tablename__ = "stages"

    id = db.Column(db.Integer, primary_key=True)
    macrostage_id = db.Column(db.Integer, db.ForeignKey("macrostages.id"), nullable=False)
    name = db.Column(db.String(255), nullable=False)
    position = db.Column(db.Integer, nullable=False, default=0)
    stage_type = db.Column(db.String(20), nullable=True)
    scope = db.Column(db.Text, nullable=True)
    tools = db.Column(db.String(255), nullable=True)
    other_tools = db.Column(db.String(255), nullable=True)

    start_date = db.Column(db.Date, nullable=True)
    end_date = db.Column(db.Date, nullable=True)

    # Relação com MacroStage
    macrostage = db.relationship("MacroStage", back_populates="stages")

    # Uma etapa possui várias tarefas
    tasks = db.relationship(
        "Task",
        back_populates="stage",
        order_by="Task.position, Task.id",
    )

    def __repr__(self) -> str:
        return f"<Stage id={self.id} name={self.name!r} macrostage_id={self.macrostage_id}>"


class Task(db.Model):
    """
    Representa uma Tarefa dentro de uma Etapa.

    Campos:
        - id        : identificador único
        - stage_id  : referência à etapa pai
        - name      : nome da tarefa
        - start_date: data de início (definida pelo usuário)
        - end_date  : data de fim (definida pelo usuário)
    """

    __tablename__ = "tasks"

    id = db.Column(db.Integer, primary_key=True)
    stage_id = db.Column(db.Integer, db.ForeignKey("stages.id"), nullable=True)
    macrostage_id = db.Column(db.Integer, db.ForeignKey("macrostages.id"), nullable=False)
    name = db.Column(db.String(255), nullable=False)

    start_date = db.Column(db.Date, nullable=True)
    end_date = db.Column(db.Date, nullable=True)
    position = db.Column(db.Integer, nullable=False, default=0)

    # Relação com Stage
    stage = db.relationship("Stage", back_populates="tasks")

    # Relação direta com MacroStage
    macrostage = db.relationship("MacroStage", back_populates="tasks")

    # Atualizações semanais relacionadas
    weekly_updates = db.relationship(
        "WeeklyUpdate",
        back_populates="task",
        cascade="all, delete-orphan",
        order_by="WeeklyUpdate.update_date.desc()",
    )

    def __repr__(self) -> str:
        return f"<Task id={self.id} name={self.name!r} stage_id={self.stage_id} macrostage_id={self.macrostage_id}>"


class WeeklyUpdate(db.Model):
    """
    Atualizações semanais associadas a uma tarefa.

    Campos:
        - id           : identificador único
        - task_id      : referência à tarefa
        - content      : descrição breve do andamento
        - update_date  : data associada à atualização
    """

    __tablename__ = "weekly_updates"

    id = db.Column(db.Integer, primary_key=True)
    task_id = db.Column(db.Integer, db.ForeignKey("tasks.id"), nullable=False)
    content = db.Column(db.Text, nullable=False)
    update_date = db.Column(db.Date, nullable=True)

    task = db.relationship("Task", back_populates="weekly_updates")

    def __repr__(self) -> str:
        return f"<WeeklyUpdate id={self.id} task_id={self.task_id}>"
