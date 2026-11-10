"""Tiny helper: build a notebook from a list of ("md"|"code", source) pairs."""
import nbformat


def build(cells, path):
    nb = nbformat.v4.new_notebook()
    for kind, src in cells:
        src = src.strip("\n")
        if kind == "md":
            nb.cells.append(nbformat.v4.new_markdown_cell(src))
        else:
            nb.cells.append(nbformat.v4.new_code_cell(src))
    nb.metadata["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
    nbformat.write(nb, path)
    print("wrote", path)
