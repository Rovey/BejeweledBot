"""
Calibrate grid position for BejeweledBot.

Run this once to save the grid corner positions as percentages of the game
window. The bot will use these on every subsequent run, automatically
adjusting for window position and size.
"""

import json
import sys

import pyautogui


def main():
    print("=== BejeweledBot Grid Calibration ===\n")

    # Find the game window
    try:
        import pygetwindow as gw

        windows = gw.getWindowsWithTitle("Bejeweled 3")
        if not windows:
            print("ERROR: Bejeweled 3 window not found. Start the game first.")
            sys.exit(1)

        win = windows[0]
        if not win.isActive:
            win.activate()
        win_left, win_top = win.left, win.top
        win_w, win_h = win.width, win.height
        print(f"Found Bejeweled 3 window at ({win_left}, {win_top}, {win_w}x{win_h})\n")
    except ImportError:
        print("ERROR: pygetwindow is required. Install with: pip install pygetwindow")
        sys.exit(1)

    # Get corners from user
    input("Move your mouse to the TOP-LEFT corner of the game grid and press Enter.")
    tl = pyautogui.position()
    print(f"  Top-left: ({tl.x}, {tl.y})")

    input("Move your mouse to the BOTTOM-RIGHT corner of the game grid and press Enter.")
    br = pyautogui.position()
    print(f"  Bottom-right: ({br.x}, {br.y})")

    # Calculate percentages relative to window
    left_pct = (tl.x - win_left) / win_w
    top_pct = (tl.y - win_top) / win_h
    right_pct = (br.x - win_left) / win_w
    bottom_pct = (br.y - win_top) / win_h

    config = {
        "left_pct": round(left_pct, 4),
        "top_pct": round(top_pct, 4),
        "right_pct": round(right_pct, 4),
        "bottom_pct": round(bottom_pct, 4),
    }

    with open("grid_config.json", "w") as f:
        json.dump(config, f, indent=2)

    grid_w = br.x - tl.x
    grid_h = br.y - tl.y

    print(f"\nGrid size: {grid_w}x{grid_h} pixels")
    print(f"Grid position within window:")
    print(f"  Left:   {config['left_pct']:.1%}")
    print(f"  Top:    {config['top_pct']:.1%}")
    print(f"  Right:  {config['right_pct']:.1%}")
    print(f"  Bottom: {config['bottom_pct']:.1%}")
    print(f"\nSaved to grid_config.json")
    print("The bot will use these values automatically on next run.")


if __name__ == "__main__":
    main()
