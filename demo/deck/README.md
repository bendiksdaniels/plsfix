# PowerPoint pickers demo deck

`build_deck.py` (python-pptx) writes `pls,fix Demo Deck.pptx`: a small 16:9 walkthrough of
the Slide and Where pickers, shipped with every release alongside `pls,fix Demo Model.xlsx`.
The output file is tracked in git; rebuild it after changing the script.

Python is used here because no maintained Rust crate writes PPTX.

## Build

```bash
python3 -m venv demo/deck/.venv && demo/deck/.venv/bin/pip install python-pptx
npm run demo:deck
```

`npm run demo:deck` just runs `demo/deck/.venv/bin/python demo/deck/build_deck.py`.
