# pptx_to_png

A Python CLI tool for converting PowerPoint slides (`.pptx`) into individual PNG images with customizable dimensions.

## Features

- Convert all slides in a PowerPoint file to PNG images.
- Specify width, height, or both for output dimensions, preserving the slide’s aspect ratio.
- Output files to a specified folder or the default folder where the source `.pptx` is located.
- Optional logging to track each file processed.

## Requirements

- Python 3.6 or later
- `python-pptx` and `Pillow` libraries

Install the dependencies by running:

```sh
pip install -r requirements.txt
```

## Usage

The `pptx_to_png` tool is a command-line interface (CLI) program. Below are the instructions for usage.

### Basic Command

Convert all slides in a `.pptx` file to PNG images with a default width of 1280px (or resize as specified):

```sh
python pptx_to_png.py --source path/to/your/presentation.pptx
```

### Options

| Option          | Short | Description                                                                                    |
| --------------- | ----- | ---------------------------------------------------------------------------------------------- |
| `--source`      | `--s` | Path to the source PowerPoint (.pptx) file. **(Required)**                                     |
| `--destination` | `--d` | Directory where the PNG files will be saved. Defaults to the same folder as the source file.   |
| `--width`       | `--w` | Width of the output PNG images. Aspect ratio is preserved unless `--height` is also specified. |
| `--height`      | `--h` | Height of the output PNG images. Aspect ratio is preserved unless `--width` is also specified. |
| `--log`         | `--l` | Enables logging. Displays messages for each file processed and output file name.               |

### Examples

1. **Convert to PNGs with Default Width (1280px):**

   ```sh
   python pptx_to_png.py --s path/to/your/presentation.pptx
   ```

2. **Convert with Specified Width and Height:**

   ```sh
   python pptx_to_png.py --s path/to/your/presentation.pptx --w 1024 --h 768
   ```

3. **Convert with Logging Enabled:**

   ```sh
   python pptx_to_png.py --s path/to/your/presentation.pptx --l
   ```

### Output Format

The output PNG files are named according to the pattern:

```plaintext
{originalFileName}-slide-{slideIndex}.png
```

Where:

- `{originalFileName}` is the name of the source PowerPoint file without the extension.
- `{slideIndex}` is the slide number, zero-padded to two digits.

### License

This project is licensed under the MIT License.
