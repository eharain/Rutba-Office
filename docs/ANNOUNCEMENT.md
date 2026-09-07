# The announcement endpoint

`GET https://office.rutba.io/announcement`

This is the only request Rutba Office makes that is not an update check or a
server the person configured themselves. It exists for two reasons, and both are
stated in the application's About dialog:

1. **A notice board.** A release worth knowing about, a security note, a change
   to the licence. It appears as one line at the top of the launcher and is
   dismissed with one press.
2. **The only usage signal there is.** The server counts requests. Nothing
   identifies the copy making them, so what is counted is *"a copy of Rutba
   Office opened somewhere today"* — never *"this copy"* and never *"this
   person"*.

## What the client sends

```
GET /announcement?v=1.0.0&os=win32
Accept: application/json
User-Agent: RutbaOffice/1.0.0
```

| Sent | Why |
| :--- | :-- |
| `v` — the version | So a notice about a 1.0 problem is not shown to someone on 1.4 |
| `os` — `win32`, `darwin` or `linux` | So a Windows notice is not shown on Linux, and so we know which platforms to build for |

Nothing else. No cookie is sent or stored, no identifier is generated, no
identifier is persisted, and `credentials` is `omit` so the browser stack cannot
add one. There is no second request, no beacon, and no body.

**At most once every 20 hours per copy**, and only from the launcher window —
opening six documents does not make six requests.

## What the server should answer

`200` with a JSON object, or `204` / an empty object when there is nothing to
say. Anything else is treated as "nothing to say".

```json
{
  "id": "2026-09-r1",
  "kind": "release",
  "title": "Rutba Office 1.1 is out",
  "body": "Conversations in Mail, and a Markdown editor that will not rewrite your README.",
  "url": "https://office.rutba.io/releases/1.1",
  "linkLabel": "What changed",
  "at": "2026-09-14T09:00:00Z"
}
```

| Field | Required | Notes |
| :---- | :------: | :---- |
| `id` | no | The identity of this announcement. Dismissing it means it never comes back. **Change the `id` and it appears again**; leave it and it stays gone. Falls back to `title`, so an unchanged sentence is never shown twice. |
| `title` | **yes** | Up to 140 characters. Without it there is no announcement. |
| `body` | no | Up to 400 characters, shown on one line and clipped. |
| `kind` | no | `news` (default), `release`, or `security`. Only changes the icon and the colour. |
| `url` | no | **Must be `https` on a `rutba.io` host.** Anything else is dropped, so a compromised endpoint cannot send people elsewhere. |
| `linkLabel` | no | Up to 40 characters. Defaults to "Read more". |
| `at` | no | Informational. |

## What the client does with a failure

Every one of these means "no announcement today", and **none of them ever reach
the person as an error**:

- no network, DNS not answering, a firewall in the way
- a redirect to a captive portal or a login page
- `404`, `500`, or anything that is not `200`
- a body that is not JSON, or JSON that is not an object, or an object with no
  `title`
- an answer that takes longer than 4 seconds

When a check fails, the last announcement that was fetched successfully is still
shown if it has not been dismissed — so a notice does not vanish because the
office wifi is down.

## Turning it off

One checkbox in About. Off means the request is never made, and the notice board
is never shown. There is deliberately no setting that reports without showing
anything back: that would be telemetry with extra steps, and this suite does not
have any.

## Counting downloads

Separately, and needing nothing from the client:

```bash
node tools/downloads.js
```

Reads the GitHub releases API and reports per-release, per-platform download
counts. The update feed (`latest.yml`, `.blockmap`) is excluded, because every
installed copy fetches it on a timer and counting it would flatter the number by
an order of magnitude.
