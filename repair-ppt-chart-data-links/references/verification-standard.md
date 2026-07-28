# Verification standard

## Structural

- Source SHA-256 unchanged; output is a valid OOXML ZIP.
- Slide and target chart counts unchanged.
- Each selected chart resolves to exactly one intended absolute workbook target; selected relationships contain no old target.
- Workbook, sheet, referenced range, `ptCount`, point indexes, and cache values agree.
- Style/color parts and unauthorized chart relationship parts are byte-identical.
- Temporary carrier worksheets and pasted helper chart shapes are absent from final deliverables.

## Desktop Office

- Every rebuilt workbook opens, saves, closes, and reopens in Excel.
- The PPT copy opens, saves, closes, and reopens in PowerPoint without a repair prompt.
- Every user-equivalent chart link check runs in a fresh PowerPoint session; do not reuse the session that opened or saved the deck.
- `ChartData.ActivateChartDataWindow` must succeed for every selected chart. This is the automated equivalent of right-clicking the chart and choosing Edit Data.
- `ChartData.Activate` must succeed for every selected chart, and `Workbook.FullName` must exactly match the rebuilt Excel workbook path.
- A closed-workbook mutation probe must succeed: byte-back up the recovered workbook, edit one referenced numeric cell, update/refresh an isolated PPT probe, save it, close it, reopen it in a fresh PowerPoint session, and verify the plotted series value changed. Restore the workbook byte-for-byte in `finally`; the delivery PPT is never the mutation target.
- `data_mappings.series_index` is an OOXML document-position index, not a guaranteed COM ordinal. Resolve the runtime `SeriesCollection` member by exact series name before reading values. A repeated series name（重复系列名称）, zero matches, or multiple matches is a blocker unless another uniquely named probe can be selected; never fall back to a guessed numeric ordinal.
- Slide count remains unchanged and target slides export to PNG.
- Run structural verification again after PowerPoint saves because Office may regenerate cache XML.
- Unattended COM runs suppress interactive alerts to avoid deadlock. Therefore `repair_prompt_observed=false` is not standalone proof; full success additionally requires `ActivateChartDataWindow`, `ChartData.Activate`, `Workbook.FullName`, mutation probe, post-save ZIP/XML, slide/chart-count, unauthorized-part, style/color, link-target, and cache-equality checks to pass.

## Visual

- Chart position and size are unchanged.
- Legend, labels, axes, series order, colors, fonts, and clipping remain acceptable.
- Updated period, number of plotted points, and visible labels agree with the workbook.
- Any unconfirmed mismatch between chart and page text is disclosed.

Full success requires all three sections. If Office is unavailable, report structural-only status, not full completion.

## Office coexistence and gate

- If any live PowerPoint process exists before the exclusive stage, stop with exit code 20 and the explicit close/save message. A threadless, handleless terminated-process shell may be ignored, but never attach to, save, close, or kill a live existing PowerPoint process.
- Record all Excel PIDs before automation. If chart activation returns a pre-existing Excel PID, stop without closing its workbook or process.
- Journal the exact PowerPoint and Excel PIDs created by each stage. Cleanup may target only those journaled PIDs; do not kill every PID that appeared after the baseline snapshot because it may belong to work the user opened during verification.
- Excel data windows may appear briefly during the user-equivalent test. Existing user workbooks must not be saved or closed.

## Portable bundle

- The master PPT is immutable; the launcher creates a current-machine linked copy.
- Every selected chart relationship must equal the absolute current location of its workbook in `data`.
- Behavioral testing uses a same-name workbook clone and an isolated PPT probe. It must pass Edit Data, `Workbook.FullName`, mutation, save and fresh-reopen checks without modifying the real package workbook.
- Copy the complete bundle to a second local directory, make the original directory unavailable, rerun the launcher, and require the same checks to pass.
- OneDrive, SharePoint, UNC/network paths and mapped network drives are out of scope and must be rejected rather than silently attempted.
