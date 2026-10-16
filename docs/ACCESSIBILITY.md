# Accessibility conformance statement

Rutba Office 2.x, the desktop suite: Home and its ten apps (Documents,
Worksheets, Presentations, PDF, Mail, Calendar, Contacts, Pictures, Image and
Video), on Windows, macOS and Linux.

This statement says what is tested, against which criteria of the Web Content
Accessibility Guidelines 2.2 (WCAG 2.2, Levels A and AA, read for desktop
software as the W3C's WCAG2ICT note reads them), and what is known not to
meet them. It covers the suite's own windows: the ribbons, panes, dialogs,
menus, lists and status bars it draws. What a document holds (its colours,
its pictures' descriptions, its headings) is its author's; the Accessibility
Checker on each document app's Review tab helps with that.

It is a statement of what the checks prove, not a claim of full conformance.
The suite has not been evaluated by people using screen readers, switch
access or speech recognition, and that is the next step (see the end).

## How it is tested

Every gate runs four window checks written for this, beside the engine tests.
Each opens the apps on real files (a document with a title and headings, a
workbook, a deck, a PDF, a calendar with a busy day, contacts and a mailbox)
and measures the window rather than trusting the code.

- **a11y-audit** runs axe-core 4.14 (Deque's open rules engine, the rules
  tagged WCAG 2.0, 2.1 and 2.2 A and AA, and its best-practice rules) over
  every app's window, each of its ribbon tabs, its app menu, and the main
  dialogs and panes: Print and Find and Replace in Documents, Insert Link in
  Worksheets, Replace in Presentations, Find in PDF, a new contact, a new
  event and a new message. Any serious or critical finding fails the check.
  A finding WCAG itself sets aside is listed with its reason in the check
  (`EXCUSED`); one the suite does not yet meet is listed as known, printed on
  every run and named below (`KNOWN`). Neither is passed silently.
- **keyboard-pass** works each app from the keyboard alone: F6 and Shift+F6
  round the ribbon, the panes, the document and the status bar and back;
  Tab and Shift+Tab through a pane; the app menu opened with Enter, its items
  walked with the arrows and closed with Escape, the keyboard back on its
  button; a dialog opened by its shortcut taking the keyboard, keeping Tab
  inside it and closing with Escape, the keyboard back where it was. A step
  that leaves the keyboard on nothing fails.
- **screen-reader** reads Chromium's own accessibility tree, the tree Windows
  UI Automation is given, through the DevTools protocol, and checks what it
  says of each document: the Documents page and its headings, the active
  Worksheets cell's address and value as the arrow keys move it, the slides
  by number and title, Mail's messages by sender and subject, Calendar's days
  by date and its events by time and title, a PDF's pages and their words.
- **high-contrast** turns on forced colours in each window (Windows High
  Contrast, emulated through the DevTools protocol) and measures that the
  system's colours are used, the keyboard's ring shows, the chosen tab is told
  by its bar and weight, every pressed, chosen or current control is outlined,
  the ribbon's pictures are drawn and menus keep an edge.
- **dark-pass** (since 2.0) measures every piece of text on screen against
  what is under it in the dark theme, at two window sizes.
- **tests/a11y.test.js** (engine tests) holds the light theme's quieter greys
  and each app's colour as words to 4.5:1 on what they are drawn on, worked out
  from the style sheet itself, and the forced-colours rules to their places.

What is not tested yet: the suite with NVDA, JAWS, Narrator, VoiceOver or Orca
in a person's hands; Windows High Contrast themes on a real display (the check
emulates them); magnification and text spacing overrides; speech recognition;
and macOS and Linux accessibility interfaces, which Chromium feeds from the
same tree but which no check reads.

## Against WCAG 2.2

"Supports" means the checks above prove it for the suite's windows;
"partially" names what is missing; "not tested" means no check covers it.

| Criterion | Level | Status | What the checks show, and what is missing |
|---|---|---|---|
| 1.1.1 Non-text content | A | Partially supports | Every control drawn as a picture has a name (from its tip where it has no words); a picture inside a named control is hidden as part of it; a slide drawn as a picture is named by its title. Pictures in documents carry the description their author gave, or none. |
| 1.3.1 Info and relationships | A | Partially supports | Landmarks (the caption, the ribbon, the main area, each pane, the status bar); the document's name is the window's heading; Documents' headings are headings at their levels; fields are named by their labels; Mail's and Home's lists are grids of rows, Contacts' a list box. Worksheets tells only the active cell as a cell of a grid, not the grid's other cells; tables inside Word documents are not checked. |
| 1.3.2 Meaningful sequence | A | Not tested | |
| 1.3.4 Orientation, 1.3.5 Identify input purpose | AA | Not applicable / not tested | A desktop window; the few personal fields (Contacts) carry no autocomplete purpose. |
| 1.4.1 Use of colour | A | Partially supports | Chosen tabs are bold as well as coloured; unread mail is bold. In Calendar an event's calendar is told by its colour alone in the views; the event's own details name it. |
| 1.4.3 Contrast (minimum) | AA | Supports | The interface's words read 4.5:1 (large 3:1) by day (axe-core's colour contrast rule over every window and tab) and at night (dark-pass). Documents' own colours are their authors'. |
| 1.4.4 Resize text | AA | Not tested | Documents zoom; the interface follows the system's display scale but is not checked at 200%. |
| 1.4.10 Reflow | AA | Not tested | The ribbon folds its groups as the window narrows (checked by the ribbon checks); nothing checks 320 px. |
| 1.4.11 Non-text contrast | AA | Partially supports | Focus rings, chosen states and edges are checked in High Contrast; in the default themes borders and icons are not measured. |
| 1.4.12 Text spacing | AA | Not tested | |
| 1.4.13 Content on hover or focus | AA | Partially supports | The suite's tips appear after a rest, go with a key press or when the pointer leaves, and never hide what they describe; whether a tip stays while the pointer moves onto it is not checked. |
| 2.1.1 Keyboard | A | Partially supports | The ribbon (KeyTips, arrows), menus, dialogs, panes, lists, the sheet, the page and the slide strip are reached and worked from the keyboard (keyboard-pass, ribbon, dialogfocus). Known mouse-only parts: the ruler's markers in Documents (the Paragraph dialog sets the same indents), drawing and ink, Image's canvas tools, dragging clips on Video's timeline, and a click on an empty day in Calendar's month (Ctrl+N makes the same new event). |
| 2.1.2 No keyboard trap | A | Supports | A dialog keeps Tab inside it by design and Escape always closes it (dialogfocus, keyboard-pass). |
| 2.1.4 Character key shortcuts | A | Does not support | Mail's single-letter keys (j, k, r, s and others) act wherever the keyboard is outside a field, and cannot be turned off or remapped. |
| 2.4.1 Bypass blocks | A | Supports | F6 goes round the window's parts; the parts are landmarks. |
| 2.4.2 Page titled | A | Supports | Each window is titled with its document and app. |
| 2.4.3 Focus order | A | Partially supports | Tab through panes and dialogs is checked; the order across a whole window is not. |
| 2.4.6 Headings and labels | AA | Partially supports | Fields and panes are labelled (a11y-audit); labels are not reviewed by a person for meaning. |
| 2.4.7 Focus visible | AA | Supports | A ring shows on the ribbon's controls (ribbon), on what F6 and Tab reach (keyboard-pass captures) and in High Contrast (high-contrast). |
| 2.4.11 Focus not obscured (minimum) | AA | Not tested | |
| 2.5.3 Label in name | A | Partially supports | A button's tip begins with its visible name and its accessible name is taken from that start. Since 2.3.0 a picture that shows words (a slide in the strip, a theme in the Design gallery) stands beside its button rather than inside it, so its words are not taken for the button's label, which axe-core 4.14 checks in every window audited; not checked for every control. |
| 2.5.7 Dragging movements | AA | Partially supports | Panel widths take the arrow keys, a picked shape moves with the arrows, the ruler's indents have the Paragraph dialog; on Video's preview the Position and size fields do what the grips do, and Mark In and Mark Out set what the In and Out flags drag. Video's clips on the timeline and drawing have no alternative to dragging. |
| 2.5.8 Target size (minimum) | AA | Partially supports | Ribbon buttons, gallery arrows, sliders and the mini month are at least 24 px. Known short: the events in Calendar's month view are 16 px tall, so a busy day shows its events at all (full-size in the Day, Week and Agenda views and the day's "+N more" list). Mail's row check box and star are smaller, under WCAG's equivalent-control exception: Ctrl+click and Shift+click anywhere on the row check it (and a run of rows), and the ribbon's Star and the row's menu star it. |
| 3.1.1 Language of page | A | Supports | Each window says its language (English, Urdu, Arabic or Hindi) on its page. |
| 3.2.1, 3.2.2 On focus, on input | A | Not tested | |
| 3.3.1 Error identification, 3.3.2 Labels or instructions | A | Partially supports | Fields carry labels (a11y-audit); errors are said in words in notices, which are read out (4.1.3); not checked field by field. |
| 3.3.8 Accessible authentication (minimum) | AA | Not tested | Mail's and the calendars' sign-ins are not checked. |
| 4.1.2 Name, role, value | A | Partially supports | axe-core finds no serious or critical fault in any window checked; toggles say pressed, menus say expanded, lists and grids say selected. Custom drawings (charts, SmartArt, slides' shapes) are pictures with a name, not structures a screen reader can walk. |
| 4.1.3 Status messages | AA | Supports | Notices are announced through a live region that is on the page before they come; faults as alerts. Not yet heard through a screen reader. |

## Known not to meet

- **Calendar's month events are under the 24 px target size** (2.5.8). The
  a11y-audit check names them on every run.
- **Worksheets tells a screen reader the active cell only.** The other cells
  are not in the tree; reading a range means moving through it.
- **Some work is mouse-only** (2.1.1): Documents' ruler markers, drawing and
  ink in every app, Image's canvas tools, Video's timeline drags.
- **Calendar tells an event's calendar by colour alone** in its views (1.4.1).
- **Mail's single-letter shortcuts cannot be turned off** (2.1.4).
- **Presentations' shapes, charts and SmartArt are pictures** to a screen
  reader, named by their slide, not walked shape by shape.
- **Two disabled Presentations buttons** (Clear all formatting, Right-to-left
  text direction) are named by why they are off ("Select a text box first")
  instead of by what they do, until a text box is chosen.
- **Not evaluated with assistive technology by people.** The checks read what
  the technology is given; they do not listen to it.

## What comes next

A review with NVDA and Narrator on Windows and VoiceOver on macOS, by people
who use them every day; the whole Worksheets grid in the tree; keyboard ways
to draw and to move clips; and the month view's events at full size where a
day has room.

Problems with accessibility are welcome at
[office.rutba.io/contact](https://office.rutba.io/contact).
