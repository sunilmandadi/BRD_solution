"""Create self-contained Copilot prompts from locally extracted source JSON."""

import json


BRD_INSTRUCTIONS = """You are reviewing a Business Requirements Document for test planning.
Extract a complete REQUIREMENT INVENTORY using only the BRD JSON below.

Return one valid JSON object (no Markdown fences) with this structure:
{
  "source_file": "BRD filename",
  "requirements": [
    {
      "requirement_id": "REQ-0001",
      "requirement_text": "One independently verifiable requirement",
      "category": "functional | data | interface | reporting | security | nonfunctional | compliance | other",
      "priority": null,
      "acceptance_criteria": null,
      "related_systems": [],
      "source_refs": ["exact BRD source_ref from input"],
      "clarification_needed": null
    }
  ],
  "open_questions": []
}

Rules:
- Cover each explicit business requirement and constraint, including tables.
- Split distinct, independently verifiable requirements; merge true duplicates
  and keep all their source_refs.
- Do not convert background information or proposed solutions into requirements.
- Do not invent priority, acceptance criteria, rules, or missing details; use
  null or open_questions when the BRD does not say.
- Cite exact source_ref values present in this BRD JSON for every item.
- DOCX references are block locations; PDF references contain page numbers.
- Review the warnings field. Missing image/scanned text cannot be inferred.
- If none are explicit, return an empty requirements array.

BRD JSON:
"""


SOLUTION_INSTRUCTIONS = """You are reviewing solution design slides for test planning.
Extract a complete SOLUTION INVENTORY using only the solution JSON below.

Return one valid JSON object (no Markdown fences) with this structure:
{
  "source_file": "solution deck filename",
  "solutions": [
    {
      "solution_id": "SOL-0001",
      "component_or_flow": "Named component or process step",
      "description": "Explicit behavior or design decision",
      "interfaces": [],
      "data_elements": [],
      "controls": [],
      "source_refs": ["exact solution source_ref from input"],
      "clarification_needed": null
    }
  ],
  "open_questions": []
}

Rules:
- Include explicitly described components, process steps, integrations,
  data flows, controls, and material design decisions from all slides,
  tables, and speaker notes.
- Separate items that can be reviewed independently. Merge true duplicates
  and keep all source_refs.
- Do not infer architecture from slide titles or diagrams that have no
  extractable text. Do not invent interfaces, behavior, or controls.
- Cite exact source_ref values present in this solution JSON for each item.
- Review the warnings field. Images/screenshots require manual checking.
- If the deck gives too little detail, use clarification_needed or
  open_questions; if nothing is explicit, return an empty solutions array.

SOLUTION JSON:
"""


def build_copilot_prompt(source, payload):
    """Return one copyable prompt containing only the matching source JSON."""
    if source == "brd" and payload.get("source_type") == "BRD":
        instructions = BRD_INSTRUCTIONS
    elif source == "solution" and payload.get("source_type") == "Solution":
        instructions = SOLUTION_INSTRUCTIONS
    else:
        raise ValueError("Prompt type does not match the extracted source.")
    return instructions + json.dumps(payload, ensure_ascii=False, indent=2)
