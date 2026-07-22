from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, TypedDict


PlannerCanvasView = Literal["hierarchical_horizontal", "cluster_mindmap"]
EpicPriority = Literal[1, 2, 3]

DEFAULT_RND_MUSCLE_SKILLS: tuple[str, ...] = (
    "Mobile App Development",
    "Software Development",
    "UI UX",
    "Dot Net",
    "Angular",
    "Python",
    "Azure",
    "Business Intelligence",
    "Quality Assurance",
    "Automation",
    "Technical Support",
)


@dataclass(frozen=True, slots=True)
class RndMuscleSkill:
    skill_id: str
    name: str
    is_default: bool = False


@dataclass(frozen=True, slots=True)
class RndMuscleTeam:
    team_id: str
    name: str
    color_hex: str
    skill_ids: tuple[str, ...] = ()
    resource_ids: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class RndMuscleResource:
    resource_id: str
    display_name: str
    initials: str
    email: str = ""
    team_id: str = ""
    skill_ids: tuple[str, ...] = ()
    resigned: bool = False
    resignation_date: str = ""


@dataclass(frozen=True, slots=True)
class RndMuscleEpic:
    epic_key: str
    epic_name: str
    project_key: str
    project_name: str
    priority: EpicPriority | None = None
    budgeted_hours: float = 0.0
    start_date: str = ""
    due_date: str = ""
    jira_url: str = ""


@dataclass(frozen=True, slots=True)
class RndMuscleBacklogItem:
    epic_key: str
    priority: EpicPriority | None
    budgeted_hours: float
    start_date: str
    due_date: str
    sort_order: int
    epic_name: str = ""
    project_key: str = ""
    project_name: str = ""


@dataclass(frozen=True, slots=True)
class RndMuscleEpicResourceMapping:
    epic_key: str
    resource_id: str
    allocation_hours: float = 0.0
    sort_order: int = 0
    created_at_utc: str = ""
    updated_at_utc: str = ""


@dataclass(frozen=True, slots=True)
class RndMuscleProjectTab:
    project_key: str
    project_name: str
    epic_count: int
    is_all_tab: bool = False


@dataclass(frozen=True, slots=True)
class RndMuscleQuickStats:
    resources_associated_with_epics: int = 0
    resources_not_yet_associated: int = 0
    selected_project_epic_count: int = 0
    high_priority_unassigned_epic_count: int = 0


@dataclass(frozen=True, slots=True)
class RndMusclePlannerState:
    active_project_key: str = "ALL"
    canvas_view: PlannerCanvasView = "hierarchical_horizontal"
    planner_epics: tuple[RndMuscleBacklogItem, ...] = ()
    backlog: tuple[RndMuscleBacklogItem, ...] = ()
    mappings: tuple[RndMuscleEpicResourceMapping, ...] = ()


@dataclass(frozen=True, slots=True)
class RndMuscleUtilizationPageState:
    report_name: str = "RnD Muscle Utilization"
    epics: tuple[RndMuscleEpic, ...] = ()
    resources: tuple[RndMuscleResource, ...] = ()
    teams: tuple[RndMuscleTeam, ...] = ()
    skills: tuple[RndMuscleSkill, ...] = field(default_factory=tuple)
    project_tabs: tuple[RndMuscleProjectTab, ...] = ()
    quick_stats: RndMuscleQuickStats = field(default_factory=RndMuscleQuickStats)
    planner: RndMusclePlannerState = field(default_factory=RndMusclePlannerState)


class RndMuscleTeamPayload(TypedDict, total=False):
    team_id: str
    name: str
    color_hex: str
    skill_ids: list[str]
    resource_ids: list[str]


class RndMuscleResourceSkillPayload(TypedDict, total=False):
    resource_id: str
    skill_ids: list[str]


class RndMusclePlannerMappingPayload(TypedDict, total=False):
    epic_key: str
    resource_ids: list[str]
    allocation_hours_by_resource_id: dict[str, float]
