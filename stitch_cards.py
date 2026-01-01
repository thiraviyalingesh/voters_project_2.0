"""
Stitch individual card images vertically with folder selection GUI.
"""

from pathlib import Path
from PIL import Image
import tkinter as tk
from tkinter import filedialog


def select_folder():
    """Open folder selection dialog."""
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Select folder containing image subfolders")
    root.destroy()
    return folder


def stitch_images_vertically(image_folder: Path, output_path: Path):
    """Stitch all numbered images in a folder vertically."""

    # Find all image files
    image_files = []
    for ext in ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.JPEG', '*.PNG']:
        image_files.extend(image_folder.glob(ext))

    if not image_files:
        print(f"  No images found in {image_folder.name}")
        return False

    # Sort by numeric filename (1.jpg, 2.jpg, 3.jpg...)
    def get_number(f):
        try:
            return int(f.stem)
        except ValueError:
            return float('inf')

    image_files = sorted(image_files, key=get_number)

    # Load all images
    images = []
    for img_path in image_files:
        try:
            img = Image.open(img_path)
            images.append(img)
        except Exception as e:
            print(f"  Error loading {img_path.name}: {e}")

    if not images:
        return False

    # Calculate stitched dimensions
    max_width = max(img.width for img in images)
    total_height = sum(img.height for img in images)

    # Create stitched image
    stitched = Image.new('RGB', (max_width, total_height), (255, 255, 255))

    # Paste images vertically
    y_offset = 0
    for img in images:
        x_offset = (max_width - img.width) // 2
        stitched.paste(img, (x_offset, y_offset))
        y_offset += img.height

    # Save stitched image
    stitched.save(str(output_path), quality=95)
    print(f"  Stitched {len(images)} images -> {output_path.name}")

    for img in images:
        img.close()

    return True


def main():
    # Select input folder
    input_folder = select_folder()
    if not input_folder:
        print("No folder selected. Exiting.")
        return

    input_folder = Path(input_folder)
    print(f"Selected: {input_folder}")

    # Create output folder with _stitched suffix
    output_folder = input_folder.parent / f"{input_folder.name}_stitched"
    output_folder.mkdir(exist_ok=True)
    print(f"Output: {output_folder}\n")

    # Find all subfolders
    subfolders = sorted([f for f in input_folder.iterdir() if f.is_dir()])

    if not subfolders:
        print("No subfolders found!")
        return

    print(f"Found {len(subfolders)} subfolders\n")

    # Process each subfolder
    success = 0
    for subfolder in subfolders:
        print(f"Processing: {subfolder.name}")
        output_path = output_folder / f"{subfolder.name}.jpg"
        if stitch_images_vertically(subfolder, output_path):
            success += 1

    print(f"\nDone! Stitched {success}/{len(subfolders)} folders")
    print(f"Saved to: {output_folder}")


if __name__ == "__main__":
    main()
