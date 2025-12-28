# Changelog

## [0.9.0] - 2025-12-28

### Highlights
- Simplified macros UX: 6 fixed macros (Macro 1…Macro 6).
  - Left-click: send text
  - Left-click + modifier (Alt/Ctrl): send text + Enter
  - Right-click: edit macro (persisted to settings)
- Improved WebView2 text injection robustness and fallbacks.
- Re-introduced global toggle hotkey (Ctrl+Shift+Space) for minimize/reactivate.

### Features
- Ensure `Properties.Settings.Default.Macros` contains six entries on startup.
- Inject `__wrokSend` helper into new and already-loaded documents.
- Targeted submit-button click fallback when Enter is requested.

### Fixes
- First-send failure fixed by injecting helper into current document.
- Restored visibility of the 6th macro item in the tray menu.
- Added JS retry + InputSimulator fallback for reliable text + Enter delivery.
- Separated hotkey handling: toggle hotkey vs macro hotkeys.

### Breaking changes / UX
- Removed per-macro subitems (Send, Send+Enter, New Macro).
- Double-click behavior removed (use modifier+click instead). Update docs/user notes.

### QA / Manual test steps
1. Open tray → Macros shows 6 items.
2. Left-click a macro → text inserted into active WebView2 input.
3. Left-click + Alt/Ctrl → text + submit.
4. Right-click macro → edit dialog; changes persist.
5. Press Ctrl+Shift+Space → app toggles minimize/reactivate.
6. Press Ctrl+1..Ctrl+5 and Ctrl+^ → corresponding macro sent.

### Rollback
- Revert changes to `RefreshMacrosMenu`, `InitializeMacrosMenu`, `SendTextToWebViewAsync`, `OnHandleCreated`, and `WndProc` to restore previous behavior.

### Suggested commit messages
- `feat(macros): simplify to 6 fixed macros and edit on right-click`
- `fix(webview): inject helper into existing doc and harden SendTextToWebViewAsync`
- `fix(hotkeys): register Ctrl+Shift+Space toggle and separate WndProc handling`