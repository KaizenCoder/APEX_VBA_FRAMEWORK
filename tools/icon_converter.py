from pathlib import Path
import cairosvg
from PIL import Image
import io

def svg_to_ico(svg_path: str, ico_path: str, sizes: list = [16, 32, 48, 64, 128, 256]):
    """Convertit un fichier SVG en ICO avec plusieurs tailles."""
    # Lecture du SVG
    svg_data = Path(svg_path).read_bytes()
    
    # Conversion en PNG pour chaque taille
    images = []
    for size in sizes:
        png_data = cairosvg.svg2png(
            bytestring=svg_data,
            output_width=size,
            output_height=size
        )
        img = Image.open(io.BytesIO(png_data))
        images.append(img)
    
    # Sauvegarde en ICO
    images[0].save(
        ico_path,
        format='ICO',
        sizes=[(size, size) for size in sizes],
        append_images=images[1:]
    )

if __name__ == "__main__":
    svg_to_ico(
        "assets/icon.svg",
        "assets/icon.ico",
        sizes=[16, 32, 48, 64, 128, 256]
    ) 