# 🎯 Fix Inspect Mode Inconsistency (Issue #14)

## Understanding
The Inspect Mode in Command Explorer is unreliable: it sometimes shows the wrong command, sometimes nothing, and breaks when navigating away from VS. The root cause is that the `BeforeExecute` event handler is unsubscribed after each inspect and only re-subscribed through a delayed, racy code path in `RefreshAsync`.

## Assumptions
- The Ctrl+Shift+Click mechanism for inspecting should remain as the activation gesture
- The handler should stay subscribed as long as the Inspect Mode checkbox is checked
- We should not rely on WPF keyboard state which can be unreliable when VS loses/regains focus; instead, we can use a simpler approach: when inspect mode is on, always capture the command and set the filter text, without requiring modifier keys (or keep the modifier check but fix the subscription issues)
- Keeping the modifier key requirement is the intended design (the MessageBox says "Hold down Ctrl+Shift and execute any command")

## Approach
The core fix involves:
1. **Stop unsubscribing the event handler in `CommandEvents_BeforeExecute`** (line 157). The handler should stay subscribed as long as inspect mode is enabled. Unsubscribing-and-resubscribing creates timing races.
2. **Stop unsubscribing/resubscribing in `RefreshAsync`** (lines 118-123). This was the mechanism to re-subscribe after inspection, but it's racy due to the 300ms delay and text equality check. The subscription should only be managed by the checkbox handler.
3. **Keep the checkbox handler (`CheckBox_Checked`) as the sole place** that subscribes/unsubscribes the event.

Key files:
- [CommandTableExplorerControl.xaml.cs](src/ToolWindows/CommandTableExplorerControl.xaml.cs) — Contains all the inspect mode logic

## Key Files
- src/ToolWindows/CommandTableExplorerControl.xaml.cs - Contains the inspect mode event subscription and handler logic

## Risks & Open Questions
- Removing modifier key requirement would change expected behavior — keeping it
- The filter text change triggered by inspect will cause `RefreshAsync` to run, which previously managed re-subscription; we need to ensure removing that code doesn't break filtering

**Progress**: 100% [██████████]

**Last Updated**: 2026-03-31 21:46:33

## 📝 Plan Steps
- ✅ **Remove the event handler unsubscribe from `CommandEvents_BeforeExecute` so the handler stays active while inspect mode is checked**
- ✅ **Remove the event subscribe/unsubscribe logic from `RefreshAsync` since the checkbox handler is the sole manager of the subscription**
- ✅ **Build and verify the changes compile successfully**

