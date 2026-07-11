# Fantasy Grounds to Microsoft Word Compiler

An automated Python data extraction and document generation pipeline that compiles zipped Fantasy Grounds campaign module archive files (`.mod`) into highly structured, professional Microsoft Word blueprint workbooks (`.docx`).

The compiler uses a fully decoupled architecture, inspecting campaign metadata on the fly to dynamically load custom ruleset parsers, lock proportional asset aspect ratios via image stream filtering, and sequentially index variable appendices.

---

## 🏛️ System Architecture

The pipeline uses an abstract, data-driven factory design pattern. Rather than hardcoding structural components, the engine inspects the module environment at runtime and maps components to their corresponding assets dynamically.

```text
       [ Zipped FG Module File (.mod) ]
                       │
                       ▼
       ┌──────────────────────────────┐
       │     main.py (Core Engine)    │
       └───────────────┬──────────────┘
                       │
        (1) Inspects metadata via definition.xml
                       │
                       ▼
       ┌──────────────────────────────┐
       │ RulesetFactory (Dynamic)     │◄─── Evaluates Active Ruleset Environment
       └───────────────┬──────────────┘     (e.g., "FrontierSpace")
                       │
        (2) Dispatches rule-mapped files dynamically
                       │
       ┌───────────────┼───────────────┐
       ▼               ▼               ▼
┌──────────────┐┌──────────────┐┌──────────────┐
│  fs_npc_     ││  fs_item_    ││  fs_vehicle_ │──► Ruleset-specific Python Parsers
│  parser.py   ││  parser.py   ││  parser.py   │
└──────────────┘└──────────────┘└──────────────┘
       │               │               │
       ▼               ▼               ▼
┌──────────────┐┌──────────────┐┌──────────────┐
│     NPC      ││     Item     ││   Vehicle    │──► Unified Layout Renderers
│   Renderer   ││   Renderer   ││   Renderer   │
└──────┬───────┘└──────┬───────┘└──────┬───────┘
       │               │               │
       ▼               ▼               ▼
┌──────────────┐┌──────────────┐┌──────────────┐
│   FS_NPC_    ││   FS_Item_   ││ FS_Vehicle_  │──► Rule-Unique Word Templates (.docx)
│Template.docx ││Template.docx ││Template.docx │
└──────┬───────┘└──────┬───────┘└──────┬───────┘
       │               │               │
       └───────────────┼───────────────┘
                       │
                       ▼ (3) Sequentially Appends Appendix Loops
                       │
       [ Final Compiled Chronicle Workbook (.docx) ]

```

### 🗂️ Standardized Core Components

When a campaign engine ruleset requires specific data layouts or visual styles, its helper files must adopt the standardized naming prefix conventions:

* **Python Scripts:** Lowercase with matching rule identifier prefix (e.g., `fs_vehicle_parser.py`).
* **Word Templates:** Capitalized with matching rule identifier prefix (e.g., `FS_Vehicle_Template.docx`).

---

## ⚙️ Configuration Matrix (Execution Order)

The sequential rendering order of chapters and appendices is completely configurable through the `pipeline_order` matrix block located inside `main.py`.

```python
pipeline_order = [
    {"id": "story",      "is_appendix": False},
    {"id": "npc",        "is_appendix": True},
    {"id": "item",       "is_appendix": True},
    {"id": "vehicle",    "is_appendix": True},
    {"id": "starships",  "is_appendix": True},
    {"id": "quests",     "is_appendix": True},
    {"id": "encounters", "is_appendix": True},
    {"id": "parcels",    "is_appendix": True},
    {"id": "tables",     "is_appendix": True},
    {"id": "characters", "is_appendix": True},
    {"id": "skills",     "is_appendix": True},
    {"id": "benefits",   "is_appendix": True},
    {"id": "images",     "is_appendix": True}
]

```

### 🔀 Dynamic Alpha Appendix Indexing

* **Automated Data Skipping:** If a designated component (such as `starships` or `quests`) contains no data nodes inside the source file, the pipeline automatically skips the generation stage to prevent blank pages or empty references.
* **Sequential Indexing:** The engine dynamically maintains an internal letter tracking pointer. The rendering pipeline assigns letter sequences dynamically (e.g., `Appendix A`, `Appendix B`) *only* to components that actually contain data. This guarantees a flawless alphabetical sequence even when certain components are excluded.

---

## 🖼️ Proportional Layout Aspect Ratio Locking

To bypass Microsoft Word's default behaviors, which can force loose canvas images or custom WebP graphic assets into strict 1x1 square frame distortions, the compiler handles asset injection using memory streams and specific dimension math:

1. The module's binary image stream is read in-memory and processed using Pillow (`PIL`).
2. The engine extracts the true pixel dimensions and calculates the structural aspect ratio explicitly.
3. The stream is converted on the fly into a standard, universally recognized PNG buffer.
4. The template engine embeds the asset as an isolated sub-document snippet (`tpl.new_subdoc()`), providing both explicit width and height parameters. This forces Microsoft Word to honor the native layout proportions perfectly.

---

## 🚀 Usage Instructions

### 1. Prerequisites

Ensure you have Python installed along with your required environment libraries:

```bash
pip install python-docx docxtpl pillow

```

### 2. Workspace Layout Setup

Place your targets and blueprint templates directly inside your active workspace directory:

```text
C:/Users/.../FG2WordByPy/
│
├── main.py                    # Master Engine Control Line
├── ruleset_factory.py         # Dynamic Module Router
├── fg_parser.py               # Low-level XML Text Harvester
├── word_renderer.py           # Inline RichText Style Applier
│
├── npc_renderer.py            # NPC Sub-doc Injector
├── item_renderer.py           # Equipment Table Compiler
├── vehicle_renderer.py        # Vehicle Spec Sheet Generator
│
├── fs_vehicle_parser.py       # FrontierSpace Vehicle Data Mapper
│
├── FS_Story_Template.docx     # Primary Book Blueprint Layout
├── FS_NPC_Template.docx       # NPC Appendix Layout
├── FS_Item_Template.docx      # Equipment Appendix Layout
├── FS_Vehicle_Template.docx   # Vehicle Appendix Layout
└── FS_MtN.mod                 # Target Zipped FG Module Source File

```

### 3. Execution Command

To build your document workbook package, execute the master control line from your terminal:

```bash
python main.py

```

### 4. Output Execution Diagnostics

The compiler will unpack your module targets, run terminal logging diagnostics across all data branches, and write your finished file cleanly:

```text
--- Initiating Fantasy Grounds Module Extraction Pipeline ---
Target Archive: FS_MtN.mod
[SYSTEM] Module metadata verified. Active Ruleset: 'FrontierSpace'
--- Successfully loaded the Fantasy Grounds XML data! ---

Processing structured narrative hierarchy...
[DEBUG] Total Categories Found: 2
  -> Category: 'Uncategorized/Loose' (Contains 9 NPCs)
  -> Category: 'Foot Solders' (Contains 2 NPCs)
Injecting 2 sorted category tracks into blueprints...

Compiling 13 items into text layout tables...
[SUCCESS] Equipment appendix compiled into chronicle workbook.

--- [DEBUG START: VEHICLE PARSER] ---
[DEBUG] Detected 1 raw child records inside <vehicle> root element.
  [Vehicle Node 1] Processing Target: 'Hover Car'
[DEBUG END: VEHICLE PARSER] Finished processing. Total successfully packaged: 1 vehicles.
Compiling 1 vehicles into narrative spec sheets...
[SUCCESS] Vehicle appendix compiled into chronicle workbook.

--- SUCCESS ---
Generated structured module workbook: Mission_to_Nagol_Campaign.docx

```

```

```
