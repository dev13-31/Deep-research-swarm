# Deep Research Swarm

An autonomous research skill for Google Antigravity and agentic coding workflows. It breaks complex scientific and technical inquiries into focused subtopics, coordinates parallel sub-agents to survey literature, and compiles a verified technical dossier exported directly to HTML, Word (`.docx`), and GitHub Pages.

---

## What It Does

Standard single-prompt research often runs into context limits, hallucinated citations, and high-level summaries. Deep Research Swarm addresses these issues through structured multi-agent coordination:

* **Parallel Sub-Agent Delegation**: Breaks a research topic into investigative, explorative, and protocol domains, running specialized agents simultaneously with domain-specific search prompts.
* **Automated Crossref Verification**: Queries the Crossref REST API for every cited DOI, confirming titles, authors, and publication metadata before generating deliverables.
* **In-Text Citation Tracking**: Enforces statement-by-statement citation anchors (`[1]`, `[2]`) in the body text that link directly to a deduplicated bibliography.
* **Multi-Format Export Engine**: Translates Markdown into standalone HTML with protected LaTeX equations (MathJax), interactive popup tooltips, and Word documents with native OpenXML tables and internal bookmarks.
* **Dual-Format Image Rendering**: Inlines standalone vector SVGs for browser viewing and auto-rasterizes them to high-resolution PNGs via headless Edge/Chromium for Microsoft Word embedding.
* **GitHub Pages Ready**: Outputs `index.html` and `.nojekyll` directly into the export folder so reports can be served on GitHub Pages without extra build steps.
* **Living Document Maintenance**: Incorporates user follow-up questions into formal report addenda and re-compiles HTML and Word deliverables automatically.

---

## Directory Layout

```text
deep-research-swarm/
├── SKILL.md                 # Core agent runbook, workflow phases, and rules
├── README.md                # Skill documentation and usage instructions
├── LICENSE                  # MIT License
├── .gitignore               # Python and editor ignore rules
├── prompts/
│   ├── investigative.txt    # Prompt template for answering specific technical questions
│   ├── explorative.txt      # Prompt template for mapping fields, benchmarks, and gaps
│   └── protocol.txt         # Prompt template for experimental recipes and step-by-step methods
└── scripts/
    ├── export_report.py     # Python engine for Crossref audit, HTML/DOCX generation, and GitHub Pages
    └── requirements.txt     # Python dependencies for the export engine
```

---

## Installation

### For Google Antigravity

Clone or copy this repository into your personal Antigravity configuration directory:

```bash
# Windows
git clone https://github.com/dev13-31/deep-research-swarm.git "%USERPROFILE%\.gemini\config\skills\deep-research-swarm"

# Linux / macOS
git clone https://github.com/dev13-31/deep-research-swarm.git "$HOME/.gemini/config/skills/deep-research-swarm"
```

Once installed, Antigravity automatically detects the skill and activates it when you run deep research tasks or invoke the command:

```text
/deep-research-swarm <your research topic or question>
```

---

## The 6-Phase Swarm Pipeline

The skill operates through a strict six-phase workflow:

### Phase 1: Intake & Planning
The master agent analyzes the user's inquiry, identifies core themes, and proposes missing angles (such as interfacial physics, synthetic protocols, or catalytic benchmarks). It categorizes each line of inquiry into one of three query types:
* **Investigative**: Direct answers to unresolved mechanisms or literature debates.
* **Explorative**: Broad mapping of performance metrics, state-of-the-art results, and research gaps.
* **Protocol**: Step-by-step experimental procedures, precursor concentrations, and operational conditions.

### Phase 2: Parallel Delegation
The master agent spawns focused sub-agents concurrently using `invoke_subagent`. Each sub-agent receives a customized prompt template from `prompts/` and queries academic search backends (e.g., `paper-search` MCP server, Crossref, arXiv, Google Scholar) to extract quantitative data.

### Phase 3: Paywall Handling
If a sub-agent encounters paywalled literature, it alerts the master agent with the specific DOI and citation so the user can provide the document.

### Phase 4: Synthesis & Master Report
The master agent aggregates the sub-agent findings into a single canonical Markdown file (`master_report.md`). Synthesis follows strict rules:
* **In-Text Citation Mandate**: Every factual claim, activation energy, turnover frequency, and rate constant must be followed by a bracketed citation anchor (`[[N]](#ref-N)`).
* **Deduplicated Bibliography**: Reference lists are deduplicated and numbered sequentially from `[1]` to `[N]`.
* **LaTeX Hygiene**: Mathematical formulas remain pure LaTeX (`$ ... $` and `$$ ... $$`); markdown links and anchors are placed in the surrounding text to avoid MathJax parsing errors.
* **Visual Quality Assurance**: Vector diagrams use strict coordinate bounds ($x \ge 0, y \ge 0$) without negative offsets that cause canvas clipping.

### Phase 5: Export to HTML, Word, and GitHub Pages
The export script compiles the Markdown report into publication-ready formats:

```bash
pip install -r scripts/requirements.txt
python scripts/export_report.py "D:/path/to/master_report.md" "D:/path/to/output_dir"
```

The script runs an automated Crossref API check on all DOIs, inlines SVGs into HTML, generates high-resolution PNGs for Word, and outputs `index.html` and `.nojekyll` for web hosting.

### Phase 6: Interactive Follow-Ups & Living Document Maintenance
When users ask follow-up questions, the system triages them:
* **Tier 1 (Metric Lookup)**: Rapid lookups in the existing corpus.
* **Tier 2 (Thematic Expansion)**: Focused literature updates compiled into an **Addendum** in `master_report.md` and re-exported to HTML/DOCX.
* **Tier 3 (Systematic Pivot)**: Spawning a new swarm for deep investigation.

---

## Standalone Export Script Usage

`export_report.py` can also be run independently on any technical Markdown document containing LaTeX equations and citations:

```bash
python scripts/export_report.py <input_markdown_file> <output_directory>
```

### Script Requirements

Install dependencies via pip:

```bash
pip install -r scripts/requirements.txt
```

Core libraries:
* `markdown>=3.4` (Markdown to HTML parser)
* `requests>=2.28` (Crossref API verification)
* `python-docx>=0.8.11` (Microsoft Word generator)

Optional: Microsoft Edge or Chromium installed in standard locations allows automated headless rasterization of local SVG figures into PNGs for Word embedding.

---

## Deploying Reports to GitHub Pages

Every run of `export_report.py` automatically generates:
1. `index.html` (copy of your compiled HTML report)
2. `.nojekyll` (disables Jekyll processing)

To host your report:
1. Initialize a git repository in your report folder:
   ```bash
   git init
   git add .
   git commit -m "Initial report commit"
   ```
2. Create an empty repository on GitHub and push:
   ```bash
   git remote add origin https://github.com/<username>/<repository>.git
   git branch -M main
   git push -u origin main
   ```
3. In GitHub, open **Settings** > **Pages** > **Source**, choose **Deploy from a branch**, select `main` / `/ (root)`, and click **Save**. Your report will be live at `https://<username>.github.io/<repository>/`.

---

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
