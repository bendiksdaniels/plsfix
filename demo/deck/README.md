# PowerPoint pickers demo deck

`build_deck.py` (python-pptx) writes `pls,fix Demo Deck.pptx`: a small 16:9 walkthrough of
the Slide and Where pickers, shipped with every release alongside `pls,fix Demo Model.xlsx`.
The output file is tracked in git; rebuild it after changing the script.

Python is used here because no maintained Rust crate writes PPTX.

## Build

```bash
python3 -m venv demo/deck/.venv && demo/deck/.venv/bin/pip install -r demo/deck/requirements.txt
npm run demo:deck
```

python-pptx is pinned in `requirements.txt` (1.0.2); bump both together after testing a newer release.

`npm run demo:deck` just runs `demo/deck/.venv/bin/python demo/deck/build_deck.py`.
