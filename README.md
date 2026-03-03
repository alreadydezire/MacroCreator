# MacroCreator (Tkinter + PyAutoGUI)

A refactored prototype of a desktop macro builder focused on maintainability and extension.

## Highlights
- Tkinter desktop UI with menus: **File / Edit / Macro / Other**
- Task builder/editor for:
  - Click
  - String
  - Keypress
  - Hotkey
  - Condition (pixel color wait)
  - Wait
  - Variable (iterates/imported CSV values)
  - Screenshot (saved automatically)
- Reorder tasks (Up/Down controls)
- Double-click task to edit
- Duplicate/Delete task actions
- Loop count control
- Run/Pause/Resume/Stop controls
- Stop and continue from same loop support
- Save/load reusable macro files (`.json`)
- CSV variable import

## Run
```bash
python3 run_macro_creator.py
```

## Notes
- Runtime execution uses `pyautogui`.
- In headless/dev environments where `pyautogui` is unavailable, UI still loads but macro execution actions may fail.
