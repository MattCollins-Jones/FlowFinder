# Flow Finder

Find, inspect and manage Power Automate cloud flows across solutions.

Flow Finder is a utility for XrmToolBox / Power Platform administrators that helps locate which solutions contain specific cloud flows and perform management tasks such as updating co-owners, modifying solution membership, filtering results and exporting flow inventories to CSV.

Highlights / Features

- Quickly find the solutions that contain a given cloud flow
- Update flow co-owners in bulk
- Add or remove flows from solutions (update solutions)
- Filter results by solution name
- Hide / show flows that are in managed solutions
- Export flow inventory to CSV for reporting or auditing

Requirements

- .NET Framework 4.8
- XrmToolBox (or any host that can load `FlowFinder.dll`)
- A Power Platform environment with sufficient privileges to view and update flows/co-owners


Usage

1. Start XrmToolBox and connect to a Power Platform environment with administrative privileges.
2. Open the Flow Finder tool.
3. Select flows and use the available actions to:
   - Update Co-owners
   - Update Solutions (add/remove from solutions)
   - Toggle visibility of flows in managed solutions in the list of flows
   - Export the displayed results to CSV

Typical CSV export fields

- Flow Name
- Description
- Solution Name(s)
- Co-Owners
- Triggering Source
- Triggering Table

License

- MIT (see `FlowFinder.nuspec` for package metadata)

Contributing

- Issues and pull requests are welcome. Open them on the project repository: https://github.com/MattCollins-Jones/FlowFinder

 Changelog (v1.2026.4.15)
- **New**: View Schema button — select any flow and click "View Schema" to inspect the full JSON definition (triggers, actions, connection references) in a dialog. Includes a Copy to Clipboard button.
- **Fix**: Co-owners added or removed were reappearing after a full reload. The principalobjectaccess query now filters out rows where `accessrightsmask = 0`, which previously caused revoked co-owners to persist.
- **Fix**: Loading animation froze while flows were loading. Grid data binding is now performed off the UI thread.
- **Fix**: GDI font handle leak — the bold font used for conditional row formatting is now properly disposed.
- **Fix**: Close button on dialogs was accessible during background operations, which could cause errors if clicked mid-load.
- **Improvement**: All labels, buttons and dialogs now consistently use "Co-Owners" capitalisation.
- **Improvement**: Large environments are now handled more reliably — Dataverse queries are paged (up to 5,000 records per page) and IN clauses are automatically chunked into batches of ≤500.
- **Improvement**: Solution and Co-Owner dialogs now load data asynchronously on open, keeping the UI responsive.


- **New**: Added clickable link to open a flow from "Link to flow" column 
- **New**: ICON UPDATE - Thanks to Cooky for the new icon design! Icons from Lizel Arina & kliwir art at flaticon.com
- **New**: Added current status of flow, active, inactive, suspended. Thanks to Amine Debba for this suggestion.
- **Fix**: Issue where conditional formatting disappeared after filtering the grid.

	

Changelog (v1.2025.12.22)

- **New**: "Triggering Entity" column renamed to "Triggering Table". - Thanks Cooky!
- **New**: "Show/Hide Managed" button now also filters the solutions dropdown. - Thanks Cooky! 
- **Fix**: Resolved an issue where management buttons became unresponsive after filtering the grid.- Thanks Andrew!
- **Improvement**: Enhanced exception handling with more detailed logging.

Changelog (Initial release)

- Initial release includes the ability to:
  - Find the solutions a flow is within
  - Update the Co-owners
  - Update the Solutions
  - Filter based on Solution
  - Export Flows to a CSV file
  - Hide/show flows in managed solutions

Contact / Links

- Repository: https://github.com/MattCollins-Jones/FlowFinder


