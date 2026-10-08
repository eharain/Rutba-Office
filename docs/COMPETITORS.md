# The roadmap against the field — and what Image and Video should become

A review of Rutba Office 1.29.10 against the suites a person would weigh it
against, WPS Office first since that is the one asked about, then a plan for
the Image and Video apps so that someone who opens them is not left without
a tool. [ROADMAP.md](ROADMAP.md) sets the order of work; this file says
what the field has, where the suite stands apart, and what to add.

## How this was gathered

WPS's own home page (explore.wps.com) could not be fetched from the machine
this was written on, so WPS's claims come from its premium, Academy and
blog pages and from app-store listings, and should be read once against the
home page before anything is quoted from them. The other suites' facts come
from their release notes and documentation; prices are approximate and
change. Rutba's side comes from the code, GAPS.md and the release notes.

## The field

**WPS Office** (Kingsoft). Writer, Spreadsheet, Presentation and a PDF
editor in one window with tabs; a free tier with ads, limited templates and
1 GB of cloud; Pro and Pro+ (about $6 a month, or a lifetime licence) for
full PDF editing, OCR, no ads and 20 GB; an AI bundle on top (AI Writer, AI
Slides, Chat with PDF, document and picture translation, AI spell check,
AI photo tools). Also a screen recorder under Tools (region, window, webcam,
audio), screenshot-to-PDF with OCR, a scanner with handwriting extraction,
and background and watermark removal pitched as free. Its headline is
breadth in one download and PDF as a first-class document.

**Microsoft 365.** The reference the others are measured against: Word,
Excel, PowerPoint, Outlook, OneNote, OneDrive, Teams, and for pictures and
film the Photos app (generative erase, blur and replace background,
Designer's selective edits) and Clipchamp (multi-track timeline, auto
captions, green screen, screen and webcam recording, 1080p free, 4K paid).
Everything new is Copilot, on a subscription.

**Google Workspace.** Docs, Sheets, Slides, Gmail, Calendar, Meet, Drive,
Vids; browser-first, co-editing as the norm, Gemini throughout from the
Standard plan up (about $7 to $8 a user a month at the bottom).

**LibreOffice 25.8.** Writer, Calc, Impress, Draw, Base, Math; free and
offline; 25.8 brought faster loading, PDF 2.0 export, a viewer mode, Calc's
dynamic-array functions, redaction of images. Draw is its graphics tool; no
video, no mail.

**ONLYOFFICE Desktop Editors 9.x.** Document, spreadsheet, presentation and
PDF editors with a form builder; 9.0 added OCR of scanned PDFs through AI,
9.2 an AI agent that fills forms and drafts documents with the user's own
provider key or a local model. Free desktop; no image or video editor.

**SoftMaker Office 2024 / NX.** TextMaker, PlanMaker, Presentations; QR and
barcodes, embedded fonts, EPUB export, SQLite import; DeepL translation and
ChatGPT in the NX subscription.

**Zoho Workplace.** Writer, Sheet, Show, Mail, Notebook, Zia AI; a free tier,
paid from about $3 a user a month; Writer turns a document into a fillable
form, Sheet turns a picture of a table into cells.

## Where Rutba stands apart

These are the suite's own ground and should stay so:

- **Offline, no account, no ads, no telemetry, AGPL.** Every suite above but
  LibreOffice wants an account for its best features; WPS's free tier
  carries ads. Rutba's declared boundary — it makes no request it has not
  declared — is a feature, and the one the others cannot copy.
- **Nine apps on one shell**, with Mail, Calendar and Contacts in the same
  download as the editors. WPS, LibreOffice, ONLYOFFICE and SoftMaker have
  no mail client; Microsoft and Google have one in a separate app or tab.
- **Every age of file, in full.** Word 1.0 to 2003, Write and DOS Word,
  binary `.xls` and `.ppt`, encrypted CFB, OpenDocument and RTF, all read
  and most written. LibreOffice is the only other suite that comes close.
- **"Proven, not promised."** A button either works or is not there, and
  each release note names the tests that prove it. No competitor publishes
  this, and it is the suite's answer to "does it really open my file".

## What the field has that the roadmap did not mention

The review of GAPS.md against the suites above found these absent from the
register altogether. Each gets a decision here, inside the suite's boundary
— nothing leaves the computer unless the person asks it to.

1. **A PDF editor as a first-class app.** WPS, ONLYOFFICE and Microsoft all
   headline PDF: edit text and pictures in place, reorder and rotate pages,
   merge and split, fill and sign forms, redact, convert to and from Word,
   OCR a scan. Rutba reads and writes PDF (packages/pdf, export from every
   app) but has no PDF app. **Build it**, as the tenth app, on the existing
   package: pages first (view, reorder, rotate, merge, split, extract),
   then forms and signatures, then text editing in place, then redaction.
   This is the largest competitive gap in the suite.
2. **OCR, on this computer.** WPS and ONLYOFFICE sell it as a cloud or AI
   feature; a local engine of the Tesseract kind, shipped or downloaded once
   with the person's consent, gives scanned PDFs a text layer, a picture
   of a table its cells in Worksheets (Zoho's trick), and a screenshot its
   words. **Build it** after the PDF app's pages, since it serves PDF, Image
   and Worksheets at once.
3. **A screen and camera recorder.** WPS, Clipchamp and the OS tools have
   one; Electron's own capture gives region, window and screen with the
   microphone and the camera. **Build it** inside Video (below), where the
   recording lands on the timeline, and offer it from Presentations for
   narration.
4. **Templates.** Every suite opens with a gallery. Rutba's New menu offers
   blank documents. **Build** a local gallery (letter, CV, invoice, budget,
   schedule, pitch deck) written as real `.docx`, `.xlsx` and `.pptx` in the
   repository, so they are proven like everything else.
5. **Forms and content controls.** Fillable PDF forms (above), and Word's
   content controls and legacy form fields in Documents — read, drawn,
   filled, written. **Build** with the PDF app.
6. **Redaction, QR and barcodes, embedded fonts.** Small, well understood,
   each in a release of its own: redaction in Documents and PDF; Insert → QR
   code in all three apps; fonts embedded in `.docx` and `.pptx` as Word
   and PowerPoint write them.
7. **Assistance.** WPS, Microsoft, Google, Zoho, SoftMaker and ONLYOFFICE
   all now sell writing, summarising, translating and "chat with the file".
   The boundary says cloud services are not built, and that holds. What
   fits the boundary is what ONLYOFFICE 9.2 did: a door the person opens
   themselves to a model on this computer or to an endpoint they name, with
   nothing sent until they press the button, and the request shown. **Not
   before 1.31**, and only as that door; the owner decides whether the door
   exists at all.
8. **Cloud sync and co-editing.** Out of scope by the suite's declaration:
   there is no Rutba server. A file kept in a synced folder already syncs;
   CalDAV and CardDAV already keep calendars and contacts in step. Say so on
   the download page rather than build a half of it.

## Image: from a viewer's companion to an editor

Today Rutba Image has rotate, one flip, crop to a ratio, seven sliders
(brightness, contrast, saturation, hue, blur, sepia, greyscale) and export
to PNG, JPEG or WebP, on a non-destructive pipeline over the untouched
source. The pipeline already knows flip vertical, resize, invert and
annotations (rectangle, ellipse, arrow, text, pen) that no button offers.
Against Paint.NET, GIMP and Photopea, all free, which have layers,
selections, masks, healing and curves, and against WPS and Windows Photos,
which have one-click background removal, the app is a crop tool. The way
up, in three tiers, each shippable on its own:

**Tier 1 — what the engine already does, and the obvious next adjustments.**
Buttons for flip vertical, resize (by pixels or percent, keep ratio), invert
and the five annotation tools with colour and width; straighten by angle
with the crop following; auto-enhance; levels and curves; white balance
(temperature and tint); sharpen and noise reduction; vignette; red-eye;
exposure, highlights and shadows. Export gains quality, a size cap, strip
metadata, and copy to clipboard. From Pictures, a batch over the folder:
resize, convert, rename by pattern, rotate by EXIF, strip metadata — the one
thing office users do with pictures more than any other. HEIC, AVIF, TIFF
and SVG decoded in the app's own code, so the Linux thumbnail gap closes
too.

**Tier 2 — selections and layers.** Rectangle, ellipse, lasso and
magic-wand selections with feather, and every adjustment applied inside the
selection; a layers panel with opacity, blend modes, masks, text layers
(through the same text engine the deck uses) and shape layers; a healing
brush and a clone stamp; perspective correction. A project format of its
own, `.rimg`, a zip holding the source, the op list and the layers, with
PSD read (layers, masks, text where possible) and flattened PSD write.
Edit-in-place from Documents, Worksheets and Presentations: Picture Format →
Edit in Image opens the picture, Save puts it back where it was.

**Tier 3 — the advanced work, still on this computer.** Background removal
with a small segmentation model that ships with the app or is downloaded
once with consent and then runs offline; content-aware fill by patch
matching, which needs no model; RAW decoding (CR2, NEF, ARW, DNG) with
exposure and white-balance recovery; HDR merge and panorama stitch; a
recorded batch action, as Worksheets' Record Actions, run over a folder.
Each of these is a declared request at most once, and never a service.

## Video: from trim to a cutting room

Today Rutba Video opens one file, plays it, splits it, trims the ends,
removes a clip, previews at four speeds, and exports WebM at the speed the
film plays, by recording a canvas. Clipchamp, CapCut, Shotcut, OpenShot and
iMovie all give, free, a multi-track timeline, titles, transitions, green
screen and MP4 out; Clipchamp and CapCut give automatic captions. The way
up, again in three tiers:

**Tier 1 — a single-track editor that finishes a job.** Join several files
on one timeline; a still picture as a clip with a duration; per-clip speed
from the ribbon (the model already honours it); volume, mute, fade in and
out; crop and rotate; titles and lower thirds through the deck's text
engine; crossfade, wipe and slide transitions; a project file, `.rvid`,
that references its sources and autosaves as the documents do. Export to
MP4 (H.264 and AAC) through the browser's own encoder with a muxer of the
suite's own, faster than real time, with presets for 720p, 1080p and 4K
and a bitrate choice; WebM stays as the free-codec option. Pictures' "In
Video" and Presentations' Export to Video land on this timeline.

**Tier 2 — tracks, sound and words.** Two or more video tracks for
picture-in-picture and overlays, audio tracks for music and voice-over
recorded in the app; keyframes for position, scale, opacity and volume;
ducking of music under speech; noise reduction; captions imported and
exported as SRT and WebVTT, edited in a list, burnt in or kept as a track;
chroma key; colour correction (lift, gamma, gain, temperature) and `.cube`
LUTs; stabilisation by feature tracking. Automatic captions by a speech
model that runs on this computer after a one-time consented download, in
the languages the suite is localised to.

**Tier 3 — recording and the rest.** The screen and camera recorder from
the list above, recording straight onto the timeline; motion tracking for a
title that follows; proxies so a 4K film cuts smoothly on a laptop;
hardware encoding where the platform offers it; GIF and audio-only export;
export of a clip's frame as a picture into Image.

## Proof, for pictures and film

The same contract as the rest of the suite. Engine tests for each pipeline
operation against golden images with a tolerance, and for the timeline
model; window checks that press the ribbon and read the canvas and the
exported file (its size, duration, codec, a frame's pixels); a fixture set
of small pictures and clips of each format checked into the repository. A
release note for Image or Video names them as the others do.

## Where this sits in the order

The order in ROADMAP.md stands: the consolidation release first, then the
performance budget. After those, in parallel with localisation:

- **1.31**: Image Tier 1 and Video Tier 1, and the PDF app's pages.
- **1.32**: Image Tier 2, Video Tier 2, PDF forms and signatures, templates.
- **1.33**: OCR on this computer, the recorder, PDF text editing and
  redaction, QR codes and embedded fonts.
- **Then** Tier 3 of each, item by item, each a declared request at most
  once and never a service; and the assistance door only if the owner
  wants it.
