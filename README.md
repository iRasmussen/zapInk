# zapInk

A small VBA macro for Microsoft PowerPoint for Mac. zapInk removes saved handwritten pen marks from a presentation, so you can keep an annotated PDF and reuse a clean PowerPoint source file.

## The problem

During a slide show, a presenter can write or draw on slides with PowerPoint's pen tools. PowerPoint can save those handwritten marks in the presentation. They can document what happened in a live lesson, and the annotated presentation can be exported as a PDF for students or other participants. But the reusable PowerPoint source is then left with those marks.

PowerPoint calls these handwritten objects **ink annotations**. zapInk removes the saved marks from the PowerPoint file; it does not create or alter the PDF.

## Workflow

Present → Write or draw on slides → Keep the saved annotations → Export an annotated PDF → Run zapInk on the source presentation → Reuse the clean PowerPoint file

The PDF keeps a record of the lesson; zapInk cleans the reusable PowerPoint source.

## Demo

Demo coming soon: keep the annotated PDF, clean the reusable PowerPoint source.

A planned 10–20 second, silent-friendly demo should show a slide show, a few pen marks, PowerPoint's prompt to keep or discard them, keeping the marks, the annotated presentation or PDF, running zapInk, and the same slide without the marks. No demo recording is included yet.

## Requirements

- macOS
- Microsoft PowerPoint for Mac desktop application
- Permission to run VBA macros

Developed and tested with the Danish-language version of Microsoft PowerPoint for Mac. The exact PowerPoint and macOS versions have not been recorded.

## Installation

The macro is supplied as plain VBA source. In PowerPoint's Visual Basic Editor, create or open a standard module and paste the contents of [`src/zapInk.bas`](src/zapInk.bas). The source was originally kept as a text clipping for copying into PowerPoint; importing this file as a module has not been verified.

The reported Danish menu route is **Tools (Værktøjer) → Macro (Makro) → Visual Basic Editor**. Menu labels and locations can vary by PowerPoint version; this route has not been independently re-verified here.

## Usage

1. Make a backup of the presentation.
2. Export or save the annotated PDF first if you need to preserve or share the marks.
3. Open the PowerPoint source presentation.
4. Run `zapInk` from PowerPoint's macro tools.
5. Inspect the slides and confirm the intended marks were removed.
6. For the first trial, save the cleaned presentation under a new name.

The macro deletes shapes. Whether PowerPoint can undo the operation has not been verified, so use a backup and **Save As** rather than relying on Undo.

## Compatibility

| Platform | PowerPoint | Language | Status |
| --- | --- | --- | --- |
| macOS | PowerPoint for Mac | Danish | Tested |
| macOS | PowerPoint for Mac | Other languages | Not tested |
| Windows | PowerPoint desktop | Any | Not tested |
| Any | PowerPoint for the web | Any | Not tested; VBA macros require the desktop application |

## Limitations

The supplied macro loops through every slide in `ActivePresentation` and deletes slide-level shapes whose `Type` equals `23`.

- It processes all slides in the active presentation.
- It does not recursively inspect grouped shapes; ink inside a group may remain.
- Whether every kind of PowerPoint ink object is represented by shape type 23 in all versions has not been tested.
- Notes and master/layout shapes are not traversed by this code.
- Other slide shapes whose type is not 23 are not targeted by this condition; broader effects have not been independently tested.
- Undo behavior has not been tested.

## Safety

- Make a backup first.
- Export or save the annotated version before removing anything.
- Test on a copy before using zapInk on an important presentation.

## Why “ink”?

“Ink” is Microsoft's term for freehand pen and drawing annotations created with PowerPoint's inking tools. The name zapInk refers to removing these saved handwritten objects from the reusable presentation.

## License

This project is available under the [MIT License](LICENSE).
