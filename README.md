# CodeFormatter (VBA)

© 2025 Tibor Fellegi 
Licensed under the [MIT License](LICENSE)

**CodeFormatter** is a VBA class module designed to automatically format and indent code within your VBA project. It supports standard modules, class modules, and userforms.

---

## Important Notice

- This tool **modifies existing VBA code** by rewriting it with standardized formatting.
- **Use only on backup copies** of your files. Test the results thoroughly before applying changes to production code.
- The author is **not responsible for any data loss or code malfunction** resulting from the use of this tool.
- This module relies on `Scripting.Dictionary` (late binding), and therefore **requires Windows** to function.

---

## Usage

**CodeFormatter** is a *predeclared* class module (`PredeclaredId = True`), which means you do not need to instantiate it manually.

It exposes two public methods:

- `CodeFormatter.FormatModule([module  As Variant])`  
 Formats a specific module. You can pass either the module's name as a string or a VBComponent object (to format a module from another VBA project). If no argument is provided, the currently active module will be formatted by default.

- `CodeFormatter.FormatProject()`  
  Formats **all modules** in the current VBA project.

To use, open the **Immediate Window** (Ctrl+G) in the VBA editor and run either:

`CodeFormatter.FormatModule` or `CodeFormatter.FormatProject`
##  Customization
The formatting rules are predefined and not configurable through external settings.
However, you can customize the behavior by modifying the InitializeKeywords procedure. This method controls how spacing is applied above and below specific keywords.

##  Summary
- Automatically formats VBA modules for improved readability
- Simple API via Immediate Window
- Windows-only (due to Scripting.Dictionary)
- MIT Licensed

For questions, suggestions, or improvements, feel free to open an issue or pull request.

