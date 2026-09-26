---
name: deep-research-swarm
description: >-
  Use this skill when the user asks you to perform deep research on a topic, requiring a swarm of agents. It orchestrates sub-agents to perform PhD-level investigative, explorative, or protocol research, compiling a highly detailed master report with visuals and automatically exporting it to HTML/DOCX.
---

# Deep Research Swarm

Follow this runbook meticulously to execute the multi-agent research workflow.

## Phase 1: Intake & Planning
1. Ask the user for the primary **Topic** and specific **Subtopics**.
2. *Crucial*: Analyze the topic and suggest 2-3 additional important subtopics they may have missed. Ask if they want to include them.
3. Determine the **Query Type** for each subtopic:
   - **Investigative**: Finding answers to specific questions.
   - **Explorative**: Mapping present status, gaps, future directions, significance.
   - **Protocol/Method**: Finding step-by-step procedures.
4. Create the working directory using a safe name: `D:\gemini\deep research\<topic>\`

## Phase 2: Delegation
1. For each subtopic, invoke a sub-agent using `invoke_subagent`. 
2. Use the corresponding prompt template from `./prompts/` to provide instructions to the sub-agent. Replace `{SUBTOPIC}` with the actual topic.
3. **PhD-Level Depth & Information Density**: Remind the sub-agents to operate at a PhD level, capturing deep scientific working principles, institutional specs, and quantitative data. Demand high information density and zero fluff/unnecessary verbosity.
4. Instruct the sub-agents to use the `paper-search` MCP server to find and read papers, and `zotero` if requested.
5. Wait for the sub-agents to complete their research and return their reports.

## Phase 3: Paywall Handling
1. If a sub-agent reports that a crucial paper is paywalled and cannot be read, ask the user to download the PDF.
2. Ask the user for the folder path where they saved the PDF.
3. Read the PDF and pass the contents back to the sub-agent.

## Phase 4: Synthesis & Master Report
1. Evaluate the sub-agent reports.
2. Compile a single Markdown file `master_report.md` in `D:\gemini\deep research\<topic>\`.
3. **Strict Formatting & Verification Requirements**:
   - **Information Density**: Zero fluff. Be concise but highly technical. Do not be unnecessarily verbose.
   - **Working Principles**: A detailed technical section explaining the underlying physics/mechanism of the topic.
   - **Visuals & Image Quality Assurance (Visual QA)**:
     * Include standalone vector SVGs locally in `images/` illustrating mechanisms, band diagrams, and schematics.
     * *Mandatory Coordinate Budgeting*: Ensure all elements, text, and markers have coordinates strictly within viewBox bounds ($x \ge 0, y \ge 0$). Never use negative offsets that clip elements off-screen.
     * *Dual-Format Rendering for Word*: High-res `.png` equivalents must accompany `.svg` files (auto-rasterized via headless Edge/Chromium) so `master_report.docx` embeds full graphic figures without falling back to text placeholders.
   - **Interactive Tooltips**: *Mandatory*. For every technical term, you must use the syntax `[Term](URL "Brief 1-sentence definition")` so an instant popup card appears on hover.
   - **LaTeX Math & MathJax Hygiene**: Use standard `$$ ... $$` and `$ ... $` syntax. Strictly isolate mathematical formulas from Markdown link syntax or `#anchor` tags. In LaTeX, `#` is reserved for macro parameters; placing `#` in math mode causes MathJax rendering failures (`You can't use 'macro parameter character #' in math mode`). Citations for equations must always be placed cleanly in narrative prose outside math delimiters.
   - **In-Text Citation Mandate & Deduplicated Bibliography**:
     * *Fact-by-Fact In-Text Callouts*: Every statement of fact, numerical metric (TOF, AQY, rate, activation energy), mechanism, and experimental claim must be followed by a bracketed numerical citation: `[[N]](#ref-N)`. General author mentions without reference callouts are prohibited.
     * *Strict Deduplication*: The Bibliography must be strictly deduplicated, numbered sequentially from $[1]$ to $[N]$. 100% of bibliography entries must be cited in the text, and 100% of in-text citations must resolve to the bibliography.
     * *Interactive Bi-Directional Linking*: In HTML, in-text citations must render as interactive badges (`class="citation-ref"`) that smoothly jump to the target reference (`:target` highlighted). In DOCX, citations must resolve via native OpenXML internal bookmarks.
   - **Automated Bibliography & DOI Verification**: Every citation must contain a verified DOI or official publisher URL. DOIs must be programmatically cross-checked against Crossref to guarantee zero hallucinated citations.

## Phase 5: Export to HTML & DOCX (GitHub Pages Ready)
1. Install script dependencies: `pip install -r ./scripts/requirements.txt`
2. Run the export script: `python ./scripts/export_report.py "D:\gemini\deep research\<topic>\master_report.md" "D:\gemini\deep research\<topic>\"`
3. The export script will automatically:
   - Run Crossref verification on all bibliography items.
   - Embed high-resolution figures into both HTML and DOCX.
   - **GitHub Pages Readiness**: Auto-generate `index.html` (mirroring the master report) and `.nojekyll` in the root directory so the folder can be directly pushed to any GitHub repository and served instantly on GitHub Pages without build steps.
4. Confirm to the user that the verified `.html`, `.docx`, and GitHub Pages files (`index.html`, `.nojekyll`) are ready.

## Phase 6: Interactive Follow-Ups & Living Document Maintenance
When the user asks follow-up queries after the initial report is delivered:
1. **Query Triage**:
   - **Tier 1 (Targeted Parameter / Metric Lookup)**: Fast, direct lookup in existing synthesized corpus or single targeted web/Crossref search. Respond immediately without spinning up a heavy swarm.
   - **Tier 2 (Thematic Expansion / New Literature Subtopic)**: Perform focused investigation (or deploy a single specialist subagent), synthesize the new findings, and integrate them into `master_report.md` under an **Addendum** or dedicated chapter.
   - **Tier 3 (Major Pivot / New Systematic Investigation)**: Formulate an intake plan for a new sub-topic, deploy a specialized subagent swarm, and version the output.
2. **Living Document Maintenance**:
   - Do not let valuable follow-up analysis vanish into ephemeral chat history.
   - Automatically re-execute `export_report.py` after updating `master_report.md` so that `master_report.html` and `master_report.docx` remain 100% up-to-date.
3. **Proactive Steering**: Offer 2–3 actionable next steps (e.g. drafting proposal sections, generating presentation slides via `/paper-deck`, or exploring industrial scale-up).
