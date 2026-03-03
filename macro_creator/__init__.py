__all__ = ["MacroCreatorApp"]


def __getattr__(name: str):
    if name == "MacroCreatorApp":
        from .gui import MacroCreatorApp

        return MacroCreatorApp
    raise AttributeError(name)
