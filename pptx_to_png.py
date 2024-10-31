import argparse
from pathlib import Path
from pptx import Presentation
from PIL import Image
import io

def convert_pptx_to_png(
    source: str, 
    destination: str = None, 
    width: int = None, 
    height: int = None,
    log: bool = False
) -> None:
    """
    Convert slides in a PowerPoint file (.pptx) to PNG images.

    Args:
        source (str): Path to the source PowerPoint file.
        destination (str): Path to the destination folder for PNGs. Defaults to the same folder as the source.
        width (int): Desired width of the PNG. If specified, height is scaled to maintain aspect ratio unless height is also specified.
        height (int): Desired height of the PNG. If specified, width is scaled to maintain aspect ratio unless width is also specified.
        log (bool): If True, outputs log messages for each processed slide.
    """
    source_path = Path(source)
    if not source_path.is_file() or source_path.suffix.lower() != ".pptx":
        raise ValueError("Source must be a valid .pptx file.")
    
    destination_path = Path(destination) if destination else source_path.parent
    destination_path.mkdir(parents=True, exist_ok=True)
    
    if log:
        print(f"Processing file: {source_path.name}")
        print(f"Output directory: {destination_path}")

    prs = Presentation(source_path)
    for i, slide in enumerate(prs.slides):
        # Render slide as a PIL image using slide's image_bytes property
        image_stream = io.BytesIO(slide.shapes[0].image.blob)
        img = Image.open(image_stream)
        
        # Calculate dimensions
        if width and height:
            img = img.resize((width, height))
        elif width:
            aspect_ratio = img.height / img.width
            img = img.resize((width, int(width * aspect_ratio)))
        elif height:
            aspect_ratio = img.width / img.height
            img = img.resize((int(height * aspect_ratio), height))
        else:
            img = img.resize((1280, int(1280 * img.height / img.width)))

        # Save PNG
        output_name = f"{source_path.stem}-slide-{i+1:02}.png"
        output_path = destination_path / output_name
        img.save(output_path)

        if log:
            print(f"Created PNG: {output_path.name}")

def main() -> None:
    parser = argparse.ArgumentParser(description="Convert PowerPoint slides to PNG images.")
    parser.add_argument("source", "--s", "--source", help="Path to the source PowerPoint (.pptx) file.")
    parser.add_argument("--d", "--destination", help="Destination folder for PNG files. Defaults to source location.")
    parser.add_argument("--w", "--width", type=int, help="Fixed width for PNG. Height will scale to maintain aspect ratio if --height is not specified.")
    parser.add_argument("--h", "--height", type=int, help="Fixed height for PNG. Width will scale to maintain aspect ratio if --width is not specified.")
    parser.add_argument("--l", "--log", action="store_true", help="Enable logging of the file names being processed.")
    
    args = parser.parse_args()

    # Validate width and height
    if args.width is not None and args.width <= 0:
        parser.error("Width must be a positive integer.")
    if args.height is not None and args.height <= 0:
        parser.error("Height must be a positive integer.")

    try:
        convert_pptx_to_png(args.source, args.destination, args.width, args.height, args.log)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
