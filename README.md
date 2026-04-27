# yyn-pm-prd

`yyn-pm-prd` is a Codex skill for turning a small feature request, product idea, or insight bundle into a complete Chinese product document package.

## Repository layout

```text
yyn-pm-prd-repo/
├─ README.md
├─ .gitignore
├─ install.ps1
├─ examples/
│  └─ pet-app-prd-example.md
└─ skill/
   ├─ SKILL.md
   ├─ agents/
   │  └─ openai.yaml
   ├─ references/
   │  ├─ prd-template.md
   │  ├─ journey-template.md
   │  ├─ flow-template.md
   │  ├─ sequence-template.md
   │  ├─ module-architecture-template.md
   │  ├─ prioritization-rules.md
   │  ├─ version-planning-template.md
   │  ├─ word-template.md
   │  └─ sample-pet-app-brief.md
   └─ scripts/
      └─ export_md_to_docx.py
```

## Install on another computer

```powershell
.\install.ps1
```

The skill will be installed to:

```text
C:\Users\<YourUser>\.codex\skills\yyn-pm-prd
```

## What the skill produces

- Chinese PRD
- Version planning
- P0 / P1 / P2 prioritization
- Mermaid user journey diagram
- Mermaid flowchart
- Mermaid sequence diagram
- Mermaid module architecture diagram
- Risks and open questions
- Markdown output
- Word `.docx` output

## Example output

See:

- `examples/pet-app-prd-example.md`

## Word export

This repository includes a first-pass Markdown-to-Word export script:

- `skill/scripts/export_md_to_docx.py`

If `python-docx` is not installed in your Python environment:

```powershell
pip install python-docx
```
