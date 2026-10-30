# Bookings to Media Delivery transfer

Google Apps Script in `MediaDeliverySync.gs`. `sendReadyBookingsToMediaDelivery()` transfers bookings marked ready into the Media Delivery workflow; `setupBookingsStatusDropdown()` configures the Ready/Sent dropdown. Read the source header and column mapping before changes; the README is only a short label.

- This is an upstream step for `../Media-Delivery`, followed by `../md-to-distribution` when editing can begin.
- Bookings column N may contain Drive smart chips. The source uses the Sheets Advanced Service because `RichTextValue` does not reliably expose a chip URL. Preserve this path and the instructions to enable Google Sheets API in the target Apps Script project.
- Integrate the menu item into the existing `onOpen()` in the Bookings project. Do not introduce a competing `onOpen()` or conflicting globals: Apps Script files share one namespace.
- Transfers and dropdown setup write to Google Sheets. Reading/developing locally does not require running those functions.
- No local automated test harness is documented. For a future functional change, validate the field mapping, link extraction, readiness filter and repeat-run status behavior on an appropriate test sheet.
- Live sheet configuration and deployment state are external; do not infer them from the local recording or last Git commit.

When working in `D:\Dev`, consult `../workspace-context/PROJECTS.md` for workspace routing and `../workspace-context/WORKSPACE_CONTEXT.md` for shared context if present. These workspace files are local context, not a dependency of the project. Update these notes when routing or field ownership changes.
