# 0005: Each change is recorded at once, and the whole workbook saved later

## Context

Closing and reopening the browser must restore every edit (R33). Browsers
discard an IndexedDB transaction that starts while the page unloads, so a
save cannot wait for `pagehide` or `visibilitychange`: an edit must already
be on its way to storage when the browser closes. Saving the whole workbook
after every edit is too slow for the sheets of R45, where one snapshot
holds ten million cells.

Options considered:

- **Save the snapshot half a second after the last change.** Cheap, but an
  edit made just before the browser closes is lost.
- **Save the snapshot after every change.** Nothing is lost, but every
  Enter on a large sheet copies and stores the whole sheet.
- **Keep a few cells in `localStorage` when the page is hidden.** It is
  synchronous, but limited to a few megabytes and is a second store whose
  contents must be merged.
- **Record each command at once, and save the snapshot later.** A command
  is a few bytes, so its transaction starts at the moment of the edit and
  completes in milliseconds. The snapshot is still saved half a second
  after the last change, and replaces the commands it includes.

## Decision

Every command that the workbook executes is written at once to the
`changes` object store, numbered in order. The whole workbook is saved half
a second after the last change and when the page is hidden; that save
deletes, in the same transaction, the changes numbered up to the last one
recorded before it. Opening a workbook loads its last snapshot, or a new
workbook if it was never saved, and executes the recorded changes on it.

## Consequences

- A change is lost only if the browser closes in the few milliseconds of
  its own transaction.
- Commands are stored, so their shape is part of the storage format: a
  later change to a command must still read the commands stored before it.
- Replaying a command must give the same workbook as executing it did, so
  commands depend only on the workbook and their own fields, never on the
  time or the view.
- Undo (S05) is executed as commands too, and is recorded the same way.
