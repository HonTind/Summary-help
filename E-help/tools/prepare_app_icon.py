"""Convert the transparent source PNG to application icons; requires Pillow.

Run: python tools/prepare_app_icon.py
The full image is preserved; generated files contain no source metadata.
"""
from pathlib import Path

from PIL import Image


def main():
    assets = Path(__file__).resolve().parent.parent / "assets"
    with Image.open(assets / "app-icon-source.png") as source:
        if source.width != source.height:
            raise ValueError("Application icon source must be square; refusing to crop.")
        icon = source.convert("RGBA").resize((256, 256), Image.Resampling.LANCZOS)
        icon.info.clear()
        icon.save(assets / "app-icon.png", format="PNG")
        icon.save(assets / "app-icon.ico", format="ICO",
                  sizes=[(n, n) for n in (16, 24, 32, 48, 64, 128, 256)])


if __name__ == "__main__":
    main()
