# 文档不加班.WpsAddin

This project is the WPS host adapter layer.

## Goals

- Keep 文档不加班 business logic in shared libraries.
- Implement WPS-specific startup, UI wiring, and host API bridges in this project.

## Next implementation steps

1. Replace the placeholder connection methods with WPS add-in lifecycle hooks.
2. Add a WPS document adapter that implements `IDocumentHostAdapter`.
3. Register ribbon/menu commands in WPS and map them to shared services.
4. Add installer logic for x86/x64 targets.
