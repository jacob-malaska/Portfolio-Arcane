# Portfolio Arcane
 2023 Industrial Design Portfolio - Jacob Malaska - All Rights Reserved

## Media loading

The page uses compressed WebP copies in `images/optimized/`. Original artwork stays
in `images/` and opens when a visitor enlarges an image. Images below the header
load lazily; below-the-fold videos start near the viewport and pause offscreen.
The top banner video uses immediate autoplay and `preload="auto"` so it starts
loading immediately. The browser controls exact network request priorities.

After adding or replacing an image in `index.html`, run:

```sh
python -m pip install Pillow
python scripts/optimize_images.py
```

The script preserves originals and adds image dimensions and loading attributes.
Change `MAX_EDGE` (currently 2400 pixels) or `QUALITY` (84) in the script to tune
size versus detail. Commit the generated files together with `index.html`.
Preview locally with `python -m http.server 8000`, then check desktop/mobile layouts,
carousel navigation, and click-to-enlarge before publishing through GitHub Pages.
