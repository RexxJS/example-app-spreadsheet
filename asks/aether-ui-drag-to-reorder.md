# Ask (aether-ui): drag within the app, to reorder (a spreadsheet's sheet tabs)

**Open**, 2026-10-15. Wanted by the spreadsheet's To do 13
(`docs/spreadsheet_demo_done_and_todo.md`): dragging a sheet's tab to
move the sheet, as in Excel.

## What the spreadsheet has today

`app/sheet/view/sheettabs.ts` draws the workbook's sheets as one button
each, in an `hstack` built with `each` over the sheet names. Moving a
sheet is the `sheet.move` command (`to=` a place, or `way=left|right`),
reached by the Move left and Move right buttons beside the tabs, and by
scripts through ADDRESS SHEET. It is undone like any change, and every formula
that names a sheet keeps reading it.

What is missing is only the gesture: press on a tab, drag it along the
row, let go between two others, and the sheet moves there.

## What aether-ui has

- `draggable(handle, path)`: a drag SOURCE carrying a file, OUT to another
  app (GtkDragSource over a GFile, an NSDraggingSession over an NSURL,
  OLE DoDragDrop offering CF_HDROP).
- `on_file_drop(cb)`: files dropped on the window FROM another app.

Neither carries something of the app's own from one of its widgets to
another; a file path is the wrong payload for "sheet Costs".

## What would do

One of two shapes; the first is smaller and enough here.

**A. A reorderable row.** `ui.reorderable(stack, cb)`: the children of an
`hstack` (or `vstack`) can be dragged among themselves; while one is
dragged the others show where it would land; let go and
`cb(from: int, to: int)` runs (child indexes, before the move). The
toolkit does not move the children itself: the page runs its command and
rebuilds the row from its state, as `each` already does, so undo and
scripts stay the one path.

**B. In-app drag and drop in general.** `ui.drag_text(handle, payload)`
makes a widget a drag source carrying a string of the app's own (not
offered to other apps); `ui.on_drop_text(handle, cb)` makes a widget a
target, `cb(payload, x, y)` with the drop point in its own coordinates.
A reorderable row is then a page's few lines over them; so are dragging
a chart to another sheet, or a cell range onto a panel.

## Each backend (aether-ui's rule: all of them, for real)

| Backend | A or B built on |
|---|---|
| GTK4 | `GtkDragSource` with a `GdkContentProvider` of a string, `GtkDropTarget` for `G_TYPE_STRING`; a private mime type keeps it in the app |
| AppKit | an `NSDraggingSource` with an `NSPasteboardItem` of a private type, `registerForDraggedTypes:` on the target |
| Win32 | OLE `DoDragDrop` with a registered private clipboard format (as `draggable` already uses OLE), or, for A alone, mouse capture within the window |
| UIKit | `UIDragInteraction` / `UIDropInteraction` with an `NSItemProvider` of a private type (`localObject` within the app) |

As `draggable` does, the drag starts on a move with the button held, not
on the press, so a tab can still be clicked to go to its sheet.

## Testing

The test driver already drags (`POST /widget/{id}/drag?x=&y=&dx=&dy=`,
which the spreadsheet's spec uses on the grid's column grips). For A, a
drag of one child onto another's place, ending in `cb(from, to)`; for B,
`POST /widget/{id}/drop_text?payload=` or a drag from one widget onto
another. The spreadsheet would add an `it` to `spec_sheet`: drag a tab
two places right, the sheet moved, a formula naming it still reading it,
Undo taking it back.

## Until then

Move left and Move right, and `sheet.move to=`, do the same thing; this
is the gesture only.
