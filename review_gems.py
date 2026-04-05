"""
Review captured gem screenshots and correct misclassifications.

Shows each gem image one by one. Press:
  Enter  — correct, keep as-is
  r/o/y/g/b/p/w — reclassify as red/orange/yellow/green/blue/purple/white
  f — Flame Gem
  s — Star Gem
  h — Hypercube
  n — Supernova Gem
  d — delete (not a gem, junk capture)
  q — quit review

Moves reclassified gems to the correct gem_library/<type>/ folder.
"""

import os
import shutil
import sys

import cv2
import numpy as np

LIBRARY_DIR = "gem_library"
UNKNOWN_DIR = "unknown_gems"

COLOR_KEYS = {
    ord("r"): "red",
    ord("o"): "orange",
    ord("y"): "yellow",
    ord("g"): "green",
    ord("b"): "blue",
    ord("p"): "purple",
    ord("w"): "white",
}

SPECIAL_KEYS = {
    ord("f"): "flame",
    ord("s"): "star",
    ord("h"): "hypercube",
    ord("n"): "supernova",
}

DELETE_KEY = ord("d")


def collect_images():
    """Collect all gem images from gem_library/ and unknown_gems/."""
    images = []

    # Gem library (organized by color)
    if os.path.isdir(LIBRARY_DIR):
        for color_name in sorted(os.listdir(LIBRARY_DIR)):
            color_dir = os.path.join(LIBRARY_DIR, color_name)
            if not os.path.isdir(color_dir):
                continue
            for filename in sorted(os.listdir(color_dir)):
                if filename.endswith(".png"):
                    images.append((os.path.join(color_dir, filename), color_name))

    # Unknown gems
    if os.path.isdir(UNKNOWN_DIR):
        for filename in sorted(os.listdir(UNKNOWN_DIR)):
            if filename.endswith(".png"):
                images.append((os.path.join(UNKNOWN_DIR, filename), "unknown"))

    return images


def main():
    images = collect_images()
    if not images:
        print("No gem images found. Run the bot first to capture gems.")
        sys.exit(0)

    print(f"=== Gem Review ({len(images)} images) ===\n")
    print("Keys:")
    print("  Enter       = correct, keep as-is")
    print("  r/o/y/g/b/p/w = classify as color (regular gem)")
    print("  color + f/s/h/n = color + special type (e.g. 'r' then 'f' = red_flame)")
    print("  f/s/h/n alone = special without color")
    print("  d = delete, q = quit")
    print()

    moved = 0
    deleted = 0
    confirmed = 0

    window_name = "Gem Review"

    for i, (path, current_label) in enumerate(images):
        img = cv2.imread(path)
        if img is None:
            continue

        # Scale up for visibility
        display = cv2.resize(img, (200, 200), interpolation=cv2.INTER_NEAREST)

        # Add label bar at the bottom showing current classification
        label_bar = np.zeros((40, 200, 3), dtype=np.uint8)
        cv2.putText(
            label_bar,
            f"{i + 1}/{len(images)}: {current_label}",
            (10, 28),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.7,
            (255, 255, 255),
            2,
        )
        display = np.vstack([display, label_bar])

        print(f"\n[{i + 1}/{len(images)}] {current_label} ({os.path.basename(path)})")

        cv2.imshow(window_name, display)
        key = cv2.waitKey(0) & 0xFF
        cv2.destroyAllWindows()

        if key == ord("q"):
            print("\nQuitting review.")
            break

        if key == 13 or key == 10:  # Enter
            confirmed += 1
            continue

        if key == DELETE_KEY:
            os.remove(path)
            deleted += 1
            print(f"  Deleted: {os.path.basename(path)}")
            continue

        # Build label: color key optionally followed by special type key
        new_label = None

        if key in COLOR_KEYS:
            color = COLOR_KEYS[key]
            # Show image again and wait for optional second key (special type or Enter)
            print(f"  Color: {color} — press f/s/h/n for special type, or Enter for regular")
            cv2.imshow(window_name, display)
            key2 = cv2.waitKey(0) & 0xFF
            cv2.destroyAllWindows()

            if key2 in SPECIAL_KEYS:
                new_label = f"{color}_{SPECIAL_KEYS[key2]}"
            else:
                new_label = color

        elif key in SPECIAL_KEYS:
            new_label = SPECIAL_KEYS[key]

        if new_label is None:
            continue

        if new_label == current_label:
            confirmed += 1
            continue

        # Move to new folder
        new_dir = os.path.join(LIBRARY_DIR, new_label)
        os.makedirs(new_dir, exist_ok=True)
        new_path = os.path.join(new_dir, os.path.basename(path))

        # Avoid overwriting
        if os.path.exists(new_path):
            name, ext = os.path.splitext(os.path.basename(path))
            new_path = os.path.join(new_dir, f"{name}_{new_label}{ext}")

        shutil.move(path, new_path)
        moved += 1
        print(f"  {current_label} -> {new_label}: {os.path.basename(path)}")

    cv2.destroyAllWindows()

    # Clean up empty directories
    for dir_path in [LIBRARY_DIR, UNKNOWN_DIR]:
        if os.path.isdir(dir_path):
            for sub in os.listdir(dir_path):
                sub_path = os.path.join(dir_path, sub)
                if os.path.isdir(sub_path) and not os.listdir(sub_path):
                    os.rmdir(sub_path)

    print(f"\nDone! Confirmed: {confirmed}, Moved: {moved}, Deleted: {deleted}")


if __name__ == "__main__":
    main()
