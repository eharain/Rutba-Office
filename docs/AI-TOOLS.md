# AI Tools

Rutba Office can let an AI app on the same computer make and read documents,
workbooks, presentations and PDFs, and use the mail accounts, calendars,
address books and databases you allow, each to the step you choose. It does this through the Model Context
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

## What the tools may reach

Beyond the folders, the tools can work with mail, calendars, contacts and
Data Files databases. Each mail account, archive, calendar, address book,
database and folder has its own setting, and everything is **Off** until you
choose. The setting is kept with the thing itself and shown in two places:
in its own app's settings, and in AI Tools under **What the tools may reach**.
A change in either place shows in the other at once.

| Step | What it allows |
| :-- | :-- |
| Off | The tools cannot see it. |
| Read only | The tools may read it, and change nothing. |
| Read and prepare | The tools may also prepare new things that stay on this computer: drafts, events, contacts, files and records. Nothing is sent. |
| Read, prepare and act | The tools may also ask to do something that leaves this computer or tells another person: send a message, send an invitation or answer one, add a contact to an address book kept on a server, or write to a table linked to a database engine. Each one is shown to you first, and done only when you press its button. |

Where each setting is:

| Thing | Its own settings | Steps it can have |
| :-- | :-- | :-- |
| Mail account | Mail, **Manage accounts**, beside the account | All four. A mailbox brought in from a file has no server to send through, so it stops at Read and prepare. |
| Archive opened in place | Mail, **Manage accounts**, under the accounts | Off or Read only. An archive is never written to. |
| Calendar | Calendar, **Accounts** | All four. |
| Address book | Contacts, **Accounts** | All four. |
| Data Files database | Data Files, **AI tools** in the side panel | All four, and for a protected database the role the tools work as. The setting is kept on this computer, not in the file. |
| Folder | AI Tools, under Folders | Off, Read only, or Read and prepare. Folders chosen before this version keep what they allowed: Read and prepare. |

Nothing the tools do deletes or moves anything, at any step. A refusal names
the thing and says where its setting is, for example: "The AI tools may not
do that with the mail account ana@example.com: it is set to Read only, and this
needs Read, prepare and act. Its setting is in Mail under Accounts, or in AI
Tools."

### The tools, and the step each needs

| Tool | Needs |
| :-- | :-- |
| `list_allowed`: what is not Off, and what each allows | |
| `search_mail`, `list_mail_folders`, `list_messages`, `read_message`, `read_attachment` | Read |
| `create_draft`: a draft kept in the account | Read and prepare |
| `send_message`: shown first, with Send, Keep as draft and Cancel | Read, prepare and act |
| `confirmation_answer`: what became of a request still waiting | |
| `list_events`, `read_event` | Read |
| `create_event`: in a calendar on this computer | Read and prepare |
| `create_event` in a calendar kept on a server, shown first | Read, prepare and act |
| `send_invitation`, `answer_invitation`: shown first | Read, prepare and act, on the calendar and on the mail account it goes from |
| `search_contacts`, `read_contact` | Read |
| `create_contact`: in the address book on this computer | Read and prepare |
| `create_contact` in an address book kept on a server, shown first | Read, prepare and act |
| `list_tables`, `read_records` | Read |
| `add_record`, `change_record`: never the design | Read and prepare |
| `add_record`, `change_record` in a table linked to a database engine, shown first | Read, prepare and act |

An attachment the tools add to a message comes only from a folder they may
read. A protected database is reached only as a role an Admin chose for the
tools in Data Files, and never as one that may change the design.

### Shown to you first

A message, an invitation or its answer, a contact for a server, or a change
for a database engine is shown in every Rutba Office window, with what it
would do: who it goes to, the subject, the words, the attachments. It has
three buttons:

- **Send** (or the action's own word) does it, and only then.
- **Keep as draft** keeps a message, an invitation or an answer as a draft in
  the mail account, and an event or a contact on this computer, and sends
  nothing. A change for a database engine has no draft.
- **Cancel** does nothing.

The AI app waits for the press for under a minute. If you have not pressed
anything by then, it is told the request is waiting, with its id, and can ask
again with `confirmation_answer`. A request nobody answers is withdrawn after
half an hour, and one still waiting when Rutba Office closes is withdrawn,
with nothing done either way.

### Rutba Office must be open

The tools an AI app starts on this computer read and make files themselves.
For mail, calendars, contacts and databases they ask the Rutba Office open in
the same profile, over a channel only that profile's tools can use: a named
pipe on this computer, with a key the running copy writes in the profile
folder for that run. Nothing listens on the network for it. With Rutba Office
closed, those tools say so; one that needs your press says that nothing was
done and nothing waits.

### A profile with a password

The tools reach nothing of a profile that has a password, its folders
included, unless that profile is open and unlocked in Rutba Office. With
**Always on** and Rutba Office closed, every call is refused.

## Seeing what they do

- **Badge.** While the tools are on, every window's title bar shows a badge,
  with a count when an AI app is connected.
- **Who is connected.** Click the badge to see each AI app: the name and
  version it gave, that it is on this computer, when it connected, and its
  last call.
- **Log.** AI Tools also lists what the tools have done: the tool, the files
  it named or the account, calendar, address book or database it used and the
  step that allowed it, and whether it worked. The log never holds what was
  read: no file's contents, no message, event, contact or record. It
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
      "args": ["C:\\Users\\you\\AppData\\Local\\Programs\\Rutba Office\\resources\\mcp\\rutba-office-mcp.cjs", "--profile", "default"],
      "env": { "ELECTRON_RUN_AS_NODE": "1" }
    }
  }
}
```

The AI app runs Rutba Office's own program as Node (`ELECTRON_RUN_AS_NODE`), so
no window opens and the copy of Rutba Office you are using is left alone.

`rutba-office-mcp.cjs` is a small file beside the app's archive. It loads the
tools from the archive, which Rutba Office's program reads when it runs as
Node. Lines copied from an earlier version named a file inside the archive,
`resources\app.asar\main\mcp\node.js`, and an AI app that checks the path
first said it was not there. Copy the lines again from AI Tools and put them
in place of the old ones.

`--profile` names the profile the lines were copied from, so the AI app works
with that profile's settings and folders. Each profile has its own AI Tools
settings. To use the tools in another profile, open AI Tools there and copy
its lines, under a name of their own in the AI app's list. A profile that no
longer exists is refused, and the AI app shows why.

## Not yet

- **Pictures other than PNG** for reading words. A JPEG has to be saved as a
  PNG first.
- **A PDF of a workbook or a presentation.** Rutba Office draws those in its
  windows, not here.
- **Video projects.**
- **Working in a document that is open in a window.**
- **Changing or answering what is already there** in a calendar or an address
  book, beyond answering an invitation. The tools add new events and contacts.
- **A choice per AI app.** Every AI app connected to a profile has the same
  settings.
