BRD and Solution JSON Extractor for WinPython Portable

PURPOSE
Select one BRD (.docx or text-based .pdf) and one solution slide deck (.pptx).
Click "Extract both to JSON". The program displays the two complete JSON
outputs and two Copilot prompts in separate tabs. Each prompt contains its
matching JSON and asks Copilot to create the relevant inventory. Use "Copy BRD
Copilot prompt" and "Copy solution Copilot prompt" separately. You can also
copy either raw JSON separately with its own button.

"Save single bundle ZIP" creates these three JSON files:
- brd_content.json: BRD text, tables, source references, and warnings.
- solution_content.json: slide text, tables, speaker notes, source references,
  and warnings.
- brd_solution_bundle.json: both extracted documents in one JSON object.

It also saves two ready-to-paste prompt files:
- brd_requirement_copilot_prompt.txt: instructions plus BRD JSON.
- solution_inventory_copilot_prompt.txt: instructions plus solution JSON.

RUNNING WITH YOUR PORTABLE WINPYTHON
1. Extract this ZIP to D:\etl\BRD_Solution_Inventory_WinPython
2. DOCX and PPTX need only standard Python plus Tkinter. To read PDF BRDs,
   install pypdf using your WinPython executable or your approved internal
   Python package source. If PyMuPDF / fitz is already available, it can be
   used as an alternative.

   "D:\Mapping Spec Tool\WinPython64-3.14.7.0slim\WPy64-31470\python\python.exe" -m pip install -r "D:\etl\BRD_Solution_Inventory_WinPython\requirements.txt"

3. Run the program:

   "D:\Mapping Spec Tool\WinPython64-3.14.7.0slim\WPy64-31470\python\python.exe" "D:\etl\BRD_Solution_Inventory_WinPython\app.py"

   Do not type the app.py path alone: Windows may open it in an editor.

COPILOT WORKFLOW
1. Review the BRD JSON and Solution JSON tabs for missing or unexpected text.
2. Click "Copy BRD Copilot prompt" and paste it into your approved Copilot
   conversation. This single clipboard item includes BRD instructions and
   the complete BRD JSON; request/response relates only to the BRD.
3. Click "Copy solution Copilot prompt" and paste it as a separate Copilot
   request. It includes solution instructions and only the solution JSON.
4. Review both inventories against the original files and source_refs.
   For very large documents the complete prompt may exceed your Copilot
   input allowance; check that the full paste was accepted. Nothing is
   silently truncated by the app.

NOTES
- The app processes files locally and does not upload them or call Copilot.
- DOCX references identify ordered blocks because DOCX page numbers change
  with rendering; PDF references identify pages and text blocks; PPTX
  references identify slide elements and notes.
- Scanned PDF pages, screenshots, and flattened diagrams do not have
  selectable text. Review images against the original or use approved OCR.
- Legacy .doc and .ppt files must first be converted to .docx and .pptx.
