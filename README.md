# Ghost Grid: Quantum Tic-Tac-Toe

**Rules adapted from Allan Goff's quantum tic-tac-toe.** Allan Goff, "Quantum
tic-tac-toe: A teaching metaphor for superposition in quantum mechanics,"
*American Journal of Physics* 74 (2006), pp. 962–973.
[DOI: 10.1119/1.2213635](https://doi.org/10.1119/1.2213635). That article is
behind an AAPT/AIP paywall, so it is not included in this repository.

**Developed for QCaMP.** These materials were built for QCaMP classroom use:
browser-playable games, slides, and printable handouts for teaching
superposition and entanglement through tic-tac-toe.

## What is here

| File | What it is |
| --- | --- |
| `Goff_TicTacToe_Offline.html` | The full game: superposition and entanglement. Play in any browser, no internet needed. Loops resolve one forced move at a time, so a class can follow each step. |
| `Superposition_Only_TicTacToe_Offline.html` | A simpler game: Goff's superposition rule (spooky marks may share a square), measured by rolling a die once the board fills. No entanglement loops. |
| `Superposition_Only_TicTacToe_StepByStep_Offline.html` | The same superposition-only game, with each forced move in a chain resolved by its own click instead of all at once. |
| `Superposition_Only_TicTacToe.ipynb` | Jupyter notebook that embeds the superposition-only game for offline classroom use. |
| `Superposition_Only_TicTacToe_Slides.pptx` | A 7-slide deck: superposition, qubits, entanglement, then the superposition-only game's rules. |
| `Quantum_TicTacToe_Goff_Grade6_Handout.docx` | One-page printable handout for the full Goff game: rules on top, a large numbered board below. |
| `Quantum_TicTacToe_Superposition_Only_Grade6_Handout.docx` | Printable handout for the superposition-only game. **Known issue:** its rule text does not yet match `Superposition_Only_TicTacToe_Offline.html` (it describes opposing-player choice with no dice, and no shared squares, instead of the die-roll measurement the game actually uses). Needs a rewrite before classroom use. |
| `Quantum Tic Tac Toe_ Two versions.pdf` | QCaMP's own two-page reference sheet describing a superposition-only version and an entanglement version, used while developing the games above. |
| `GhostGrid.png` | The Ghost Grid logo, used by the handout and slide generators. |
| `build-goff-full-handout.py` | Regenerates `Quantum_TicTacToe_Goff_Grade6_Handout.docx`. Requires `python-docx` and `Pillow`. |
| `build-superposition-slides.py` | Regenerates `Superposition_Only_TicTacToe_Slides.pptx`. Requires `python-pptx`, `Pillow`, and `lxml`. |
| `build-notebook.py` | Regenerates `Superposition_Only_TicTacToe.ipynb` from the offline HTML game. |

Each `build-*.py` script is the source of truth for the file it produces.
Edit the script and rerun it rather than hand-editing the `.docx`/`.pptx`/`.ipynb` output.

## Playing the games

Double-click any `*_Offline.html` file to open it in a browser. Everything
runs locally: no internet connection, account, or install is required.
