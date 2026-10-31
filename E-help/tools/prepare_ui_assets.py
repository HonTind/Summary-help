"""Generate embedded transparent icons from vendored Feather 4.29.2 SVGs.

Developer dependencies: Node + sharp, and Pillow for the supplied app logo.
Pass --node-exe and --sharp-module when they are not on PATH. No network calls;
runtime uses only tkinter and the generated module. Do not edit generated data.
"""
import argparse
import base64
from io import BytesIO
from pathlib import Path
import shutil
import subprocess

from PIL import Image

PROJECT = Path(__file__).resolve().parent.parent
NAMES = {"workbench": "home", "history": "clock", "rules": "file-text",
         "help": "book-open", "feedback": "message-square", "about": "info",
         "offline": "settings", "calendar": "calendar", "previous": "chevron-left",
         "next": "chevron-right", "down": "chevron-down", "up": "chevron-up",
         "external": "arrow-up-right", "remove": "x"}
SIZES = (24, 30, 36, 42, 48, 60, 72, 96)
RENDERER = """const sharp=require(process.argv[1]);
let bytes=[]; process.stdin.on('data',c=>bytes.push(c));
process.stdin.on('end',()=>sharp(Buffer.concat(bytes),{density:Number(process.argv[2])*3})
 .resize(Number(process.argv[2]),Number(process.argv[2])).png().toBuffer()
 .then(b=>process.stdout.write(b)).catch(e=>{process.stderr.write(e.message);process.exit(1)}));"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--node-exe", default=shutil.which("node"))
    parser.add_argument("--sharp-module", default="sharp")
    options = parser.parse_args()
    if not options.node_exe:
        parser.error("Node executable is required.")
    icons = {}
    for size in SIZES:
        icons[size] = {}
        for tone, color in (("normal", "#667085"), ("active", "#0071E3")):
            icons[size][tone] = {}
            for name, source in NAMES.items():
                svg = (PROJECT / "assets/ui" / (source + ".svg")).read_text(encoding="utf-8")
                svg = svg.replace("currentColor", color)
                result = subprocess.run([options.node_exe, "-e", RENDERER, options.sharp_module, str(size)],
                                        input=svg.encode("utf-8"), capture_output=True, check=True)
                icons[size][tone][name] = base64.b64encode(result.stdout).decode("ascii")
    brands = {}
    with Image.open(PROJECT / "assets/app-icon.png") as source:
        for size in (32, 40, 48, 56, 64, 80, 96, 128):
            output = BytesIO()
            source.convert("RGBA").resize((size, size), Image.Resampling.LANCZOS).save(output, format="PNG")
            brands[size] = base64.b64encode(output.getvalue()).decode("ascii")
    content = ('"""Generated offline icon data: Feather 4.29.2 (MIT), plus supplied logo.\n'
               'Rebuild with tools/prepare_ui_assets.py; license in assets/ui/FEATHER-LICENSE.txt.\n'
               '"""\n\nICON_PNGS = ' + repr(icons) + '\n\nBRAND_PNGS = ' + repr(brands) + '\n'
               '\nPNG_ASSETS = dict(ICON_PNGS[24]["normal"], brand=BRAND_PNGS[32])\n')
    (PROJECT / "e_help/ui_assets.py").write_text(content, encoding="utf-8")
    print("Generated offline transparent icons at", SIZES, "pixels; logo variants included.")


if __name__ == "__main__":
    main()
