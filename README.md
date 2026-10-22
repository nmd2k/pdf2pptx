# pdf2pptx

Turn every page of a PDF into one full-bleed PowerPoint slide, rendered at high resolution.

## Install

Requires Python 3.9+.

With [uv](https://docs.astral.sh/uv/) (recommended):

```bash
uv tool install .
```

Or with pipx/pip:

```bash
pipx install .
# or
pip install .
```

Run it once without installing:

```bash
uvx --from . pdf2pptx slides.pdf
```

## Usage

```bash
pdf2pptx sample.pdf
```

This creates:

- `sample.pptx` — one 16:9 slide per page
- `sample_pages/` — rendered page images (300 DPI)

Common options:

```bash
pdf2pptx sample.pdf -o talk.pptx          # custom output file
pdf2pptx sample.pdf --dpi 600             # higher resolution
pdf2pptx sample.pdf --format png          # lossless images
pdf2pptx sample.pdf --image-dir pages/    # custom image folder
pdf2pptx sample.pdf --no-keep-images      # delete images after building
```

| Option | Default | Description |
| --- | --- | --- |
| `-o, --output` | `<pdf name>.pptx` | Output PowerPoint file |
| `--dpi` | `300` | Image resolution relative to slide width |
| `--format` | `jpeg` | `jpeg` or `png` |
| `--quality` | `95` | JPEG quality, 1-100 |
| `--image-dir` | `<pdf name>_pages` | Folder for rendered pages |
| `--no-keep-images` | off | Remove images after saving the deck |
| `--version` | | Show version |

## Known Issues

- The slide is always 13.333 in wide; height follows the first PDF page's aspect ratio.
- Close the `.pptx` in PowerPoint before regenerating it.
