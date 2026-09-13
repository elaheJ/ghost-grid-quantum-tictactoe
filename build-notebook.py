"""Embed the separately named classroom game in a self-contained notebook."""
import ast
import json
from pathlib import Path

folder = Path(__file__).parent
notebook = json.loads((folder / 'Superposition_TicTacToe.before-goff.ipynb').read_text())
html = (folder / 'Superposition_Only_TicTacToe_Offline.html').read_text()

INTRO = """# GHOST GRID — Superposition Only (Offline)

This game uses Allan Goff's superposition rule without his loop-closing measurement. The board stays on the left and the playing pad on the right. Spooky marks may share a square. Each square holds at most two marks. This is a separate game from the full Goff version in this folder.

## Project and play offline

Double-click **Superposition_Only_TicTacToe_Offline.html** in this folder, then click **Full screen**. No internet, account, or Jupyter installation is needed for the HTML file.

To play in this notebook, use **locally installed Jupyter Notebook or JupyterLab** and run the next cell. The entire game is embedded. Trust your local notebook if Jupyter requests it. Google Colab requires internet. Use the HTML file for projection.

Choose **Classical** for one empty square or **Superposition** for two squares. Use the playing pad or keys **1–9** after clicking inside the game. The current player is shown in the badge at the top right and again above the playing pad. When every square is filled, select a pair and click **Roll die**."""

RULES = """## Rules and implementation conventions

- X and O alternate. A classical move places one real mark in a completely empty square. A superposition move places a matching pair of spooky marks in two different squares. Color and move number identify the pair.
- Spooky marks may share a square, following Goff. A square holds at most two marks. A real mark locks its square, so nothing further may be placed there.
- Placement ends when no square is empty. Measurement then begins.
- Players take turns measuring. A player may measure either player's pair. Select a half, then roll one six-sided die. Odd keeps A and even keeps B.
- Two pairs that share a square are entangled. They cannot both occupy that square. If the surviving half lands on the shared square, the other pair is forced into its far square. That forcing can cascade along a chain of shared squares. The game applies these forced collapses and lists them after the roll.
- If the surviving half leaves the shared square instead, the other pair keeps its superposition.
- Measure every pair before checking the outcome. Three real marks in a row wins. Spooky marks do not count toward a win. If no one has won and empty squares remain, place again and measure once the board is filled.
- A board of real marks with no winning line is a draw. There is no nine-move limit and no line-count scoring.

Goff leaves three details open in this setting. This version uses odd → A and even → B. Placement and measurement share one alternating turn order. If both players hold a winning line after all pairs are measured, the game is a draw. The game explains these conventions on screen.

**Undo** cancels a selection. It also reverses a placement or a measurement. **New game** resets the board. **Full screen** works best in the standalone HTML file. Everything, including the die, works offline.

Source: **QuantumTicTacToe.pdf** in this folder, Allan Goff, *American Journal of Physics* 74 (2006), pp. 962–973.

## Discussion

- Why do three spooky marks in a row not count as a win?
- Two pairs share a square. What does measuring one of them tell you about the other?
- A forced collapse can cascade through several pairs. What is the physical analogy?
- How do classical moves and superposition moves lead to different choices?"""

markdown = iter([INTRO, RULES])
for cell in notebook['cells']:
    if cell['cell_type'] == 'markdown':
        source = next(markdown)
    else:
        source = ''.join(cell['source'])
        tree = ast.parse(source)
        assignment = next(node for node in tree.body if isinstance(node, ast.Assign)
                          and any(isinstance(target, ast.Name) and target.id == 'GAME_HTML' for target in node.targets))
        lines = source.splitlines(keepends=True)
        lines[assignment.lineno-1:assignment.end_lineno] = ['GAME_HTML = ' + repr(html) + '\n']
        source = ''.join(lines).replace('Ghost Grid — offline game', 'Ghost Grid — superposition only')
        cell['outputs'] = []
        cell['execution_count'] = None
    cell['source'] = source.splitlines(keepends=True)

target = folder / 'Superposition_Only_TicTacToe.ipynb'
target.write_text(json.dumps(notebook, indent=1, ensure_ascii=False) + '\n')
print('Built:', target.name)
