from macro_creator.engine import MacroEngine
from macro_creator.models import MacroDocument, Task


def test_document_roundtrip(tmp_path):
    from macro_creator.storage import MacroStorage

    doc = MacroDocument(
        name="A",
        loop_count=3,
        variables={"name": ["alice", "bob"]},
        tasks=[Task("String", {"text": "hello {{name}}"}), Task("Wait", {"seconds": "0.1"})],
    )
    out = tmp_path / "macro.json"
    MacroStorage.save(str(out), doc)
    loaded = MacroStorage.load(str(out))
    assert loaded.name == "A"
    assert loaded.loop_count == 3
    assert loaded.tasks[0].task_type == "String"


def test_inter_task_delay_roundtrip(tmp_path):
    from macro_creator.storage import MacroStorage

    doc = MacroDocument(name="D", inter_task_delay=0.75)
    out = tmp_path / "macro_delay.json"
    MacroStorage.save(str(out), doc)
    loaded = MacroStorage.load(str(out))
    assert loaded.inter_task_delay == 0.75


def test_variable_resolution():
    engine = MacroEngine()
    doc = MacroDocument(variables={"name": ["alice", "bob"]})
    assert engine._resolve_value("hi {{name}}", doc, 0) == "hi alice"
    assert engine._resolve_value("hi {{name}}", doc, 1) == "hi bob"


def test_document_roundtrip_csv_sources(tmp_path):
    from macro_creator.storage import MacroStorage

    doc = MacroDocument(name="B", csv_sources={"students.csv": "/tmp/students.csv"})
    out = tmp_path / "macro2.json"
    MacroStorage.save(str(out), doc)
    loaded = MacroStorage.load(str(out))
    assert loaded.csv_sources == {"students.csv": "/tmp/students.csv"}
