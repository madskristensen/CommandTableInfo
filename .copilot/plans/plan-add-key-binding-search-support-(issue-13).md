# 🎯 Add key binding search support (Issue #13)

## Understanding
The user wants command filtering to include key bindings, aligned with issue #13, so commands can be found by shortcut text. The existing filter supports name, GUID, and ID, but not bindings.
## Assumptions
- Reusing the existing filter textbox is acceptable for shortcut search.
- Key binding strings from DTE are already available through `Command.Bindings` and can be normalized for matching.
- A lightweight shortcut capture mode (checkbox + key handling) is sufficient for issue #13 intent.
## Approach
I will update the filter UI to add a shortcut-search mode toggle and wire keyboard handling on the filter textbox so shortcut combinations can be entered as text (including two-stroke chords like `Ctrl+K, Ctrl+G`). In code-behind, I will extend filtering logic to include key binding matches using normalized comparisons, and make binding extraction null-safe to avoid runtime exceptions.

The main implementation will be in [src/ToolWindows/CommandTableExplorerControl.xaml](src/ToolWindows/CommandTableExplorerControl.xaml) for UI hooks and [src/ToolWindows/CommandTableExplorerControl.xaml.cs](src/ToolWindows/CommandTableExplorerControl.xaml.cs) for shortcut capture + filter logic.
## Key Files
- src/ToolWindows/CommandTableExplorerControl.xaml - add shortcut search mode checkbox and textbox key handler.
- src/ToolWindows/CommandTableExplorerControl.xaml.cs - implement key chord capture and binding-aware filter matching.
## Risks & Open Questions
- DTE binding text formatting can vary by scope/prefix and spacing; normalization will mitigate common differences.
- Capturing key chords in a textbox may differ for some system-reserved combinations.

**Progress**: 100% [██████████]

**Last Updated**: 2026-03-31 22:50:15

## 📝 Plan Steps
- ✅ **Update filter UI to include shortcut search mode and hook filter textbox key events.**
- ✅ **Add shortcut capture state and key-to-text formatting helpers in the control code-behind.**
- ✅ **Extend command filter logic to match command key bindings using normalized comparisons.**
- ✅ **Make binding extraction logic null-safe and consistent for both display and filtering.**
- ✅ **Build the solution to validate the changes compile cleanly.**

