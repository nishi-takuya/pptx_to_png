# pptx_to_png.py

## 📖 Table of Contents

- [pptx\_to\_png.py](#pptx_to_pngpy)
  - [📖 Table of Contents](#-table-of-contents)
  - [📝 Description](#-description)
  - [✨ Features](#-features)
  - [⚙️ Installation Instructions](#️-installation-instructions)
  - [🚀 Usage Guide](#-usage-guide)
    - [Command-line Options and Parameters](#command-line-options-and-parameters)
  - [📁 Examples](#-examples)
  - [🔬 How It Works](#-how-it-works)
  - [✅ Prerequisites](#-prerequisites)
  - [📜 License](#-license)
  - [📬 Contact Information](#-contact-information)

## 📝 Description

A command-line tool designed to convert slides from a PowerPoint (.pptx) file into individual PNG images. The tool exports each slide while maintaining its aspect ratio, with options to set a custom width or height. Logging is available to track the output of generated PNG files.

## ✨ Features

- Converts PowerPoint slides to high-quality PNG images.
- Maintains slide aspect ratio during conversion.
- Option to specify output width and/or height.
- Saves PNG images in a user-defined directory.
- Command-line interface with detailed logging support.

## ⚙️ Installation Instructions

1. Ensure Python 3.9 or higher is installed.
2. Clone this repository or download the script.
3. Install any necessary dependencies (e.g., `pip install -r requirements.txt`).

## 🚀 Usage Guide

Run the script from the command line with the following syntax:

```sh
python pptx_to_png.py -s <path_to_pptx> -d <destination_folder> -w <width> -ht <height> -l
```

### Command-line Options and Parameters

- `-s, --source`: Path to the source PowerPoint (.pptx) file (required).
- `-d, --destination`: Directory where the PNG files will be saved. Defaults to the source file's directory.
- `-w, --width`: Desired width for the PNG images. The height will scale to maintain the aspect ratio if `--height` is not specified.
- `-ht, --height`: Desired height for the PNG images. The width will scale to maintain the aspect ratio if `--width` is not specified.
- `-l, --log`: Enable logging of the file names being processed.

## 📁 Examples

1. **Basic Conversion with Logging Enabled**:

    ```sh
    python pptx_to_png.py -s presentation.pptx -d ./slides_output -l
    ```

    This command converts all slides from `presentation.pptx` into PNG images, saves them in the `./slides_output` directory, and logs each file name.

2. **Conversion with a Fixed Width**:

    ```sh
    python pptx_to_png.py -s presentation.pptx -w 1920
    ```

    This command sets the width to 1920 pixels, scaling the height to maintain the aspect ratio. PNG files are saved in the source directory.

3. **Conversion with a Fixed Height and Logging Disabled**:

    ```sh
    python pptx_to_png.py -s presentation.pptx -ht 1080
    ```

    This command sets the height to 1080 pixels, scaling the width to maintain the aspect ratio. Logging is disabled by default.

4. **Conversion with Both Fixed Width and Height**:

    ```sh
    python pptx_to_png.py -s presentation.pptx -d ./output -w 1280 -ht 720
    ```

    This command sets the width to 1280 pixels and the height to 720 pixels, saving the output in the `./output` directory.

## 🔬 How It Works

1. The tool uses PowerPoint's COM interface to open the source `.pptx` file.
2. Each slide's dimensions are determined to calculate the aspect ratio.
3. The user-specified width or height (or both) is applied, and the slides are exported as PNGs while maintaining the aspect ratio.
4. PNG files are saved in the specified destination folder, and logging is enabled if requested.

## ✅ Prerequisites

- **Python 3.9 or higher**
- **Required Python package(s)**:
  - `comtypes==1.4.8` (For interacting with PowerPoint's COM interface)
- **Microsoft PowerPoint must be installed** on the machine running the script.

## 📜 License

This project is licensed under the MIT License. You can use, copy, modify, and distribute this software under the terms of the MIT License. See the [LICENSE](LICENSE.md) file for the full text.

## 📬 Contact Information

For any questions or support, please contact the [authors](authors.md).
