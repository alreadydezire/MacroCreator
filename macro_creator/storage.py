from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Dict, List

from .models import MacroDocument


class MacroStorage:
    @staticmethod
    def save(path: str, document: MacroDocument) -> None:
        output = Path(path)
        output.write_text(json.dumps(document.to_dict(), indent=2), encoding="utf-8")

    @staticmethod
    def load(path: str) -> MacroDocument:
        input_path = Path(path)
        payload = json.loads(input_path.read_text(encoding="utf-8"))
        return MacroDocument.from_dict(payload)

    @staticmethod
    def load_csv_variables(path: str) -> Dict[str, List[str]]:
        values: Dict[str, List[str]] = {}
        with open(path, newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            if not reader.fieldnames:
                return values
            for field in reader.fieldnames:
                values[field] = []
            for row in reader:
                for field in reader.fieldnames:
                    values[field].append(row.get(field, ""))
        return values
