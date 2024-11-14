import argparse
from pathlib import Path
from pptx import Presentation  # Cross-platform way to access slide count
from PIL import Image, PngImagePlugin
import os


def add_timestamp_to_png(png_path: str, timestamp: float) -> None:
    """
    Adds a UNIX timestamp to a PNG file's metadata.

    Args:
        png_path (str): Path to the PNG file.
        timestamp (float): UNIX timestamp to add to the PNG metadata.
    """
    try:
        with Image.open(png_path) as img:
            meta = PngImagePlugin.PngInfo()
            meta.add_text("SourcePPTXTimestamp", str(timestamp))
            img.save(png_path, "PNG", pnginfo=meta)
        print(f"Timestamp added to PNG: {png_path}")
    except (IOError, OSError) as e:
        print(f"Error writing metadata to PNG: {png_path}. Error: {e}")


def get_timestamp_from_png(png_path: str) -> float:
    """
    Retrieves the UNIX timestamp from a PNG file's metadata.

    Args:
        png_path (str): Path to the PNG file.

    Returns:
        float: UNIX timestamp if found, else 0.0.
    """
    try:
        with Image.open(png_path) as img:
            timestamp_str = img.info.get("SourcePPTXTimestamp", "")
            return float(timestamp_str) if timestamp_str else 0.0
    except (IOError, OSError, ValueError) as e:
        print(f"Error reading metadata from PNG: {png_path}. Error: {e}")
        return 0.0


def get_slide_count(source: str) -> int:
    """
    Gets the number of slides in a PowerPoint file cross-platform without opening PowerPoint.

    Args:
        source (str): Path to the PowerPoint (.pptx) file.

    Returns:
        int: Number of slides in the presentation.
    """
    try:
        presentation = Presentation(source)
        return len(presentation.slides)
    except Exception as e:
        print(f"Error retrieving slide count: {e}")
        return 0


def needs_conversion(source: str, destination: str, num_slides: int) -> bool:
    """
    Pre-checks existing PNG files to see if conversion is necessary.

    Args:
        source (str): Path to the source PowerPoint file.
        destination (str): Directory where the PNG files will be saved.
        num_slides (int): Number of slides in the presentation.

    Returns:
        bool: True if any slide needs updating or is missing, False otherwise.
    """
    source_timestamp = os.path.getmtime(source)

    for i in range(1, num_slides + 1):
        output_name = f"{Path(source).stem}-slide-{i:02}.png"
        output_path = Path(destination) / output_name
        if not output_path.exists():
            print(f"PNG file missing for slide {i}: {output_name}")
            return True
        existing_timestamp = get_timestamp_from_png(output_path)
        if existing_timestamp < source_timestamp:
            print(f"PNG file outdated for slide {i}: {output_name}")
            return True

    print("All PNG files are up-to-date.")
    return False


def export_slides_as_png(
    source: str,
    destination: str,
    num_slides: int,
    width: int = None,
    height: int = None,
    log: bool = False,
) -> None:
    """
    Exports each slide in a PowerPoint file to individual PNG images with timestamp checking.
    """
    # Check if conversion is needed before opening PowerPoint
    if not needs_conversion(source, destination, num_slides):
        print("No slides need conversion. Exiting.")
        return

    print("Starting slide export process...")
    import comtypes.client  # Importing here to avoid loading if not needed

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    powerpoint.WindowState = 2  # Minimize the PowerPoint window
    presentation = powerpoint.Presentations.Open(source)

    source_timestamp = os.path.getmtime(source)

    for i, slide in enumerate(presentation.Slides, start=1):
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        aspect_ratio = slide_height / slide_width

        if width and height:
            ppt_width, ppt_height = width, height
        elif width:
            ppt_width = width
            ppt_height = int(width * aspect_ratio)
        elif height:
            ppt_height = height
            ppt_width = int(height / aspect_ratio)
        else:
            ppt_width = 1280  # Default width
            ppt_height = int(ppt_width * aspect_ratio)

        output_name = f"{Path(source).stem}-slide-{i:02}.png"
        output_path = str(Path(destination) / output_name)

        # Export slide as PNG
        print(f"Exporting slide {i} as PNG...")
        slide.Export(output_path, "PNG", ppt_width, ppt_height)
        add_timestamp_to_png(
            output_path, source_timestamp
        )  # Embed UNIX timestamp in PNG

        if log:
            print(f"Created PNG: {output_name}")

    presentation.Close()
    powerpoint.Quit()
    print("Slide export process completed.")


def main() -> None:
    """
    Main function to parse arguments and run the slide export process.
    """
    print("Initializing PNG conversion from PowerPoint slides...")

    parser = argparse.ArgumentParser(
        description="Convert PowerPoint slides to PNG images with timestamp checks."
    )
    parser.add_argument(
        "-s",
        "--source",
        required=True,
        help="Path to the source PowerPoint (.pptx) file.",
    )
    parser.add_argument(
        "-d",
        "--destination",
        help="Destination folder for PNG files. Defaults to source location.",
    )
    parser.add_argument(
        "-w",
        "--width",
        type=int,
        help="Fixed width for PNG. Height will scale to maintain aspect ratio if --height is not specified.",
    )
    parser.add_argument(
        "-ht",
        "--height",
        type=int,
        help="Fixed height for PNG. Width will scale to maintain aspect ratio if --width is not specified.",
    )
    parser.add_argument(
        "-l",
        "--log",
        action="store_true",
        help="Enable logging of the file names being processed.",
    )

    args = parser.parse_args()

    source = str(Path(args.source).resolve())
    destination = args.destination or Path(source).parent
    width = args.width
    height = args.height
    log = args.log

    Path(destination).mkdir(parents=True, exist_ok=True)

    # Get slide count cross-platform using python-pptx
    num_slides = get_slide_count(source)
    if num_slides == 0:
        print("Could not retrieve slide count. Exiting.")
        return

    export_slides_as_png(source, destination, num_slides, width, height, log)

    print("PNG conversion process completed.")


if __name__ == "__main__":
    main()
