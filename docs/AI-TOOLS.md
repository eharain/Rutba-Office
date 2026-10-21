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

## This computer only, or other machines too

AI Tools has a **Who can connect** setting:

| Setting | What it means |
| :-- | :-- |
| This computer only | The default. The AI app starts the tools itself, and talks to them through standard input and output. Nothing listens on the network, so no other machine can connect. |
| Other machines too | While the tools are on, Rutba Office also answers AI apps on other machines over the network, at the address and port AI Tools shows, to whoever has the key it shows. |

With other machines too:

- **The key.** AI Tools keeps the key hidden until you choose Show. Copy puts it
  on the clipboard. **New Key** makes a new one, and the old key, and every
  connection made with it, stops working at once.
- **What is refused.** A call with no key or the wrong key is refused, and an
  address that keeps trying is turned away for a minute. A call from a web
  page is refused, so a page open in your browser cannot reach the tools.
- **On and off.** Every call asks again whether the tools are on and the
  network is still chosen. Turning either off refuses the next call, and
  nothing listens any more.
- **Not encrypted.** The connection is plain HTTP. Use it on a network you
  trust, or reach it through a secure tunnel.
- **Seen.** Each AI app on the network is listed under Connected now with its
  address, and the log says where each call came from.

An AI app on another machine connects to an MCP server at an address. AI Tools
shows the lines to add, with the address and the key:

```json
{
  "mcpServers": {
    "rutba-office": {
      "type": "http",
      "url": "http://192.168.1.20:47613/mcp",
      "headers": { "Authorization": "Bearer <the key AI Tools shows>" }
    }
  }
}
```

An engine test fails if any file of the tools other than the network server
(`main/mcp/remote.js`) brings in a networking module or opens a port, or if
this computer only stops being the default.

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
