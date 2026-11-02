"""Generate the audit notebook from the `# %%`-cell-delimited peer_monitor.py.
Single source -> both the runnable .py and the readable .ipynb (no drift)."""
import nbformat as nbf

SRC = "peer_monitor.py"
OUT = "Peer Performance Monitor.ipynb"

lines = open(SRC, encoding="utf-8").read().split("\n")
cells, cur, cur_type = [], [], "code"


def flush():
    global cur, cur_type
    text = "\n".join(cur).strip("\n")
    if text.strip():
        if cur_type == "markdown":
            md = "\n".join(l[2:] if l.startswith("# ") else (l[1:] if l == "#" else l)
                           for l in text.split("\n"))
            cells.append(nbf.v4.new_markdown_cell(md))
        else:
            cells.append(nbf.v4.new_code_cell(text))
    cur = []


for l in lines:
    if l.startswith("# %% [markdown]"):
        flush(); cur_type = "markdown"
    elif l.startswith("# %%"):
        flush(); cur_type = "code"
    else:
        cur.append(l)
flush()

# Drop the CLI `if __name__ == "__main__":` guard (argparse breaks in Jupyter) and add an
# explicit run() cell. The guard may be mid-cell, so truncate at it.
for i, c in enumerate(cells):
    if c.cell_type == "code" and "if __name__" in c.source:
        head = c.source.split("if __name__")[0].rstrip()
        cells[i] = nbf.v4.new_code_cell(head)
        cells.insert(i + 1, nbf.v4.new_code_cell(
            "# Run the full monthly pipeline end-to-end (writes both Excel files + dashboard "
            "JSON to ./out)\nresults = run(data_dir='Data', out_dir='out')"))
        break

nb = nbf.v4.new_notebook()
nb.cells = cells
nb.metadata = {"kernelspec": {"name": "python3", "display_name": "Python 3"},
               "language_info": {"name": "python"}}
nbf.write(nb, OUT)
print(f"wrote {OUT} with {len(cells)} cells")
