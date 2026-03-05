from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List


TASK_TYPES = {
    "Click",
    "Move",
    "String",
    "Keypress",
    "Hotkey",
    "Condition",
    "Wait",
    "Variable",
    "Screenshot",
}


@dataclass
class Task:
    task_type: str
    params: Dict[str, Any]

    def validate(self) -> None:
        if self.task_type not in TASK_TYPES:
            raise ValueError(f"Unsupported task type: {self.task_type}")

    def summary(self) -> str:
        return f"{self.task_type}: {self.params}"


@dataclass
class MacroDocument:
    name: str = "Untitled Macro"
    tasks: List[Task] = field(default_factory=list)
    loop_count: int = 1
    inter_task_delay: float = 0.0
    variables: Dict[str, List[str]] = field(default_factory=dict)
    csv_sources: Dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "loop_count": self.loop_count,
            "variables": self.variables,
            "inter_task_delay": self.inter_task_delay,
            "csv_sources": self.csv_sources,
            "tasks": [
                {"task_type": task.task_type, "params": task.params}
                for task in self.tasks
            ],
        }

    @classmethod
    def from_dict(cls, payload: Dict[str, Any]) -> "MacroDocument":
        tasks = [Task(t["task_type"], t.get("params", {})) for t in payload.get("tasks", [])]
        for task in tasks:
            task.validate()
        return cls(
            name=payload.get("name", "Untitled Macro"),
            tasks=tasks,
            loop_count=int(payload.get("loop_count", 1)),
            inter_task_delay=float(payload.get("inter_task_delay", 0.0)),
            variables=payload.get("variables", {}),
            csv_sources=payload.get("csv_sources", {}),
        )
