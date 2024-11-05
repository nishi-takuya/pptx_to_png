import argparse
from pathlib import Path
import comtypes.client

def export_slides_as_png(source: str, destination: str, width: int = None, height: int = None, log: bool = False) -> None:
    """
    Exports each slide in a PowerPoint file to individual PNG images using PowerPoint's native export functionality.
    Determines each slide's aspect ratio and applies it to the output PNG size.

    Args:
        source (str): Path to the source PowerPoint file.
        destination (str): Directory where the PNG files will be saved.
        width (int, optional): Desired width of the PNG images. If specified, height is calculated based on each slide's aspect ratio unless height is also specified.
        height (int, optional): Desired height of the PNG images. If specified, width is calculated based on each slide's aspect ratio unless width is also specified.
        log (bool): If True, prints the names of each created PNG file.
    """
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    powerpoint.WindowState = 2  # 2 is the constant for ppWindowMinimized
    presentation = powerpoint.Presentations.Open(source)

    # Export each slide
    for i, slide in enumerate(presentation.Slides, start=1):
        # Get the slide dimensions in points (1 point = 1/72 inch)
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        aspect_ratio = slide_height / slide_width

        # Determine output dimensions based on specified width/height and aspect ratio
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

        # Export the slide as a PNG
        output_name = f"{Path(source).stem}-slide-{i:02}.png"
        output_path = str(Path(destination) / output_name)
        slide.Export(output_path, "PNG", ppt_width, ppt_height)

        if log:
            print(f"Created PNG: {output_name}")

    presentation.Close()
    powerpoint.Quit()

def main() -> None:
    """
    Main function to parse arguments and run the slide export process.
    Converts each slide in a PowerPoint presentation to a PNG image.
    """
    parser = argparse.ArgumentParser(description="Convert PowerPoint slides to PNG images.")
    parser.add_argument("-s", "--source", required=True, help="Path to the source PowerPoint (.pptx) file.")
    parser.add_argument("-d", "--destination", help="Destination folder for PNG files. Defaults to source location.")
    parser.add_argument("-w", "--width", type=int, help="Fixed width for PNG. Height will scale to maintain aspect ratio if --height is not specified.")
    parser.add_argument("-ht", "--height", type=int, help="Fixed height for PNG. Width will scale to maintain aspect ratio if --width is not specified.")
    parser.add_argument("-l", "--log", action="store_true", help="Enable logging of the file names being processed.")
    
    args = parser.parse_args()

    # Determine source, destination, width, height, and logging options
    source = str(Path(args.source).resolve())
    destination = args.destination or Path(source).parent
    width = args.width
    height = args.height
    log = args.log

    Path(destination).mkdir(parents=True, exist_ok=True)

    export_slides_as_png(source, destination, width, height, log)

if __name__ == "__main__":
    main()
