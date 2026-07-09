# FG2WordByPy
FG VTT to Word Utility
# FG2WordByPy: Fantasy Grounds VTT to Word Utility

A robust, template-first Python automation pipeline designed to extract campaign stories, maps, and non-player character (NPC) rosters directly from compiled Fantasy Grounds module files (`.mod`) and inject them cleanly into beautifully styled Microsoft Word documents (`.docx`).

This utility completely automates the layout process, mapping structural narrative hierarchies and building custom, grouped, and alphabetized character dossiers inside native Word graphic canvases.

## Key Features

- **Automated Metadata Validation**: Reads `definition.xml` directly from the zipped module to ensure system and ruleset alignment (configured for *FrontierSpace*).
- **Sequential Story Compilation**: Reconstructs the complete `refmanualindex` layout tree, parsing and sorting chapters, subchapters, and individual pages strictly by their internal numerical order tags.
- **Hierarchical NPC Grouping & Alphabetization**: Extracts primary character dossiers from the core database node, ignoring sub-list components like weapon pools and skill blocks. Automatically clusters unassigned/loose NPCs first, followed by folders arranged alphabetically, sorting the inner character profiles from A to Z.
- **Single-Pass Template Injection**: Avoids file corruption and layout drops by pushing data objects directly to Jinja loop layers, natively managing table row expansion and inline formatting canvas boxes.

---

## Technical Requirements

### System Requirements
- **Python**: Version `3.10` or higher (tested successfully on Python `3.14` environment chains).
- **Microsoft Word**: Needed to edit or view the layout blueprints (`.docx`) and output chronicles.

### Python Libraries & Packages
Run the following command in your terminal or PowerShell prompt to install all required engine extensions:

```bash
pip install python-docx docxtpl docxcompose
