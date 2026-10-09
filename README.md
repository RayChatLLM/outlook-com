# outlook_com — releases

Public releases of the `outlook_com` RayChat plugin: classic Microsoft
Outlook automation over the Windows COM object model, implemented with
only the Python standard library (ctypes late-bound IDispatch — no
pywin32).

Subcommands: `setup`, `unread`, `reply`, `unanswered`, `mkfolder`,
`tag`, `text`, `summarize` (LLM-written inbox brief), and `search`
with `from:` / `to:` / `subject:` / `has:attachment` / `is:unread` /
`last:Nd` / `after:` / `before:` / `in:FOLDER` / `limit:` filters plus
quoted free text, compiled to a DASL predicate evaluated by Outlook.

Requirements: Windows with classic Outlook installed and running
(the "new Outlook" app has no COM interface), Python 3.11+.

## Install

Download the release zip and install it with RayChat's plugin manager:

```
/plugins install outlook_com-X.Y.Z.zip
```

or directly from the release URL:

```
/plugins install https://github.com/RayChatLLM/outlook-com/releases/download/vX.Y.Z/outlook_com-X.Y.Z.zip
```

The plugin also runs standalone without RayChat:

```
python -m outlook_com setup
python -m outlook_com search from:david has:attachment last:2w "project Alpha"
```

Each release lists the archive's SHA-256; source is maintained in a
private repository.
