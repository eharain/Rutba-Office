# AI Tools

Rutba Office can let an AI app on the same computer make and read documents,
workbooks, presentations and PDFs. It does this through the Model Context
Protocol (MCP), which many AI apps support. The tools are off until you turn
them on.

## Off unless you turn them on

Open any Rutba Office app, open the app menu and choose **AI Tools**. There
are three settings:

| Setting | What it means |
| :-- | :-- |
| Off | The default. Every call is refused, with a sentence that says how to turn the tools on. |
| On for this session | The tools answer until the Rutba Office that turned them on closes. If Rutba Office stops unexpectedly, the session has ended, and the tools are off at the next start. |
| Always on | The tools answer until you turn them off. |

The setting is read again for every call, so turning the tools off stops the
next call, even in the middle of a conversation.

When the tools are first turned on, Rutba Office asks you to choose a folder
for them. If you cancel, they stay off. Rutba Office never makes a folder for
you. You can add more folders or remove any later.

## Only this computer

The AI app starts the tools itself, and talks to them through standard input
and output. Nothing listens on the network, so no other machine can connect.
An engine test fails if any file of the tools brings in a networking module or
opens a port.

## What they can do

- **List** the folders the tools may use, and the files in them.
- **Read** a Word document (headings, lists and tables), an Excel workbook (each
  sheet's values), a PowerPoint presentation (slide by slide), a PDF (page by
  page), or a text file.
- **Make a Word document** from a title, headings, paragraphs, bulleted and
  numbered lists and tables, in one of Rutba Office's themes.
- **Make an Excel workbook** from sheets of rows. A cell that starts with = is
  a formula.
- **Make a PowerPoint presentation** with a title slide, then a slide of points
  for each item.
- **Start from a template:** letters, a CV, a report, invoices, a budget, a
  timesheet, decks and more.
- **Make a PDF** of a Word document, laid out as Rutba Office prints it.
- **Make a new PDF from pages of others:** in any order, some pages left out,
  and pages turned.
- **Read the words in a PNG picture,** using Rutba Office's own text
  recognition on this computer.

The tools work only in the folders you chose. A path outside them is refused,
and so is a path that leads out through a link. They never overwrite or delete
a file: a file they make must have a name of its own. That leaves nothing to
undo and nothing to confirm.

## Seeing what they do

- **Badge.** While the tools are on, every window's title bar shows a badge,
  with a count when an AI app is connected.
- **Who is connected.** Click the badge to see each AI app: the name and
  version it gave, that it is on this computer, when it connected, and its
  last call.
- **Log.** AI Tools also lists what the tools have done: the tool, the files
  it named, and whether it worked. The log never holds a file's contents. It
  is kept in `ai-tools-log.jsonl` in Rutba Office's profile folder, and only
  the last 500 calls are kept.

## Connecting an AI app

AI Tools shows the lines to add to the AI app's list of MCP servers, with a
Copy button. They look like this. The paths are those of your own copy, and
the window has them right.

```json
{
  "mcpServers": {
    "rutba-office": {
      "command": "C:\\Users\\you\\AppData\\Local\\Programs\\Rutba Office\\Rutba Office.exe",
      "args": ["C:\\Users\\you\\AppData\\Local\\Programs\\Rutba Office\\resources\\app.asar\\main\\mcp\\node.js"],
      "env": { "ELECTRON_RUN_AS_NODE": "1" }
    }
  }
}
```

The AI app runs Rutba Office's own program as Node (`ELECTRON_RUN_AS_NODE`), so
no window opens and the copy of Rutba Office you are using is left alone.

## Not yet

- **Pictures other than PNG** for reading words. A JPEG has to be saved as a
  PNG first.
- **A PDF of a workbook or a presentation.** Rutba Office draws those in its
  windows, not here.
- **Video projects.**
- **Working in a document that is open in a window.**
- **Mail, Calendar and Contacts.** These are left out on purpose.
- **Connections from another machine.** Rutba Office will offer them only as a
  separate setting, off by default, that only you can turn on.
