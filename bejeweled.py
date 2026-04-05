"""
BejeweledBot - An automated bot that plays Bejeweled 3 (Steam).

Captures the game grid via screenshots, identifies gem colors using HSV analysis,
evaluates possible moves using a heuristic scoring system, and executes the best
move via mouse automation. Waits for board animations to settle before each move.
"""

import concurrent.futures
import json
import logging
import os
import time
from datetime import datetime

import cv2
import keyboard
import numpy as np
import pyautogui

GRID_SIZE = 8

# HSV hue ranges for gem classification (OpenCV: H=0-179)
# Red wraps around 0/179, handled separately in classify_hue
GEM_HUE_RANGES = [
    (8, 20, "orange"),
    (20, 35, "yellow"),
    (35, 85, "green"),
    (85, 125, "blue"),
    (125, 170, "purple"),
]

# HSV thresholds for pixel filtering
MIN_SATURATION = 80
MIN_VALUE = 80
WHITE_MAX_SAT = 50
WHITE_MIN_VAL = 180
MIN_COLOR_RATIO = 0.10  # At least 10% of center region must be colorful

# Board stability detection
STABILITY_THRESHOLD = 3.0  # Max mean pixel difference to consider board stable
STABILITY_DELAY = 0.15  # Seconds between stability checks
MAX_STABILITY_WAIT = 5.0  # Maximum seconds to wait for board to settle

# Minimum identified cells to trust a frame (out of 64)
MIN_IDENTIFIED_CELLS = 56

# If any single color covers more than this fraction of the grid, it's a popup/overlay
MAX_SINGLE_COLOR_RATIO = 0.55

# Minimum number of distinct colors for a valid game board
MIN_DISTINCT_COLORS = 4

# Single-letter abbreviations for log output
COLOR_ABBREV = {
    "blue": "B",
    "green": "G",
    "red": "R",
    "purple": "P",
    "orange": "O",
    "yellow": "Y",
    "white": "W",
}


def setup_logger():
    """Set up logging to both console and a timestamped log file."""
    os.makedirs("logs", exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = os.path.join("logs", f"playthrough_{timestamp}.log")

    logger = logging.getLogger("bejeweled")
    logger.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    )

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(message)s"))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info("Playthrough log: %s", log_file)
    return logger


def format_grid(color_grid):
    """Format the color grid as a compact string for logging."""
    lines = ["  " + " ".join(str(c) for c in range(GRID_SIZE))]
    for row_idx, row in enumerate(color_grid):
        cells = " ".join(COLOR_ABBREV.get(c, ".") for c in row)
        lines.append(f"{row_idx} {cells}")
    return "\n".join(lines)


def find_grid_from_window(logger):
    """Find the game grid using saved percentages and the game window position.

    Uses grid_config.json if available (created by calibrate.py), otherwise
    falls back to default percentages derived from Bejeweled 3's standard layout.
    Returns (top_left, bottom_right) in screen coordinates, or None if the
    game window is not found.
    """
    # Default percentages (Bejeweled 3 standard layout)
    config = {
        "left_pct": 0.329,
        "top_pct": 0.087,
        "right_pct": 0.980,
        "bottom_pct": 0.924,
    }

    # Load custom calibration if available
    config_path = os.path.join(os.path.dirname(__file__) or ".", "grid_config.json")
    if os.path.exists(config_path):
        with open(config_path) as f:
            saved = json.load(f)
            config.update(saved)
        logger.info("Loaded grid calibration from %s", config_path)
    else:
        logger.info("No grid_config.json found, using default percentages")
        logger.info("Run calibrate.py for a precise fit")

    # Find the game window
    try:
        import pygetwindow as gw

        windows = gw.getWindowsWithTitle("Bejeweled 3")
        if not windows:
            logger.info("Bejeweled 3 window not found")
            return None

        win = windows[0]
        if not win.isActive:
            win.activate()
            time.sleep(0.5)

        logger.info(
            "Found Bejeweled 3 window at (%d, %d, %dx%d)",
            win.left, win.top, win.width, win.height,
        )

        top_left = (
            int(win.left + win.width * config["left_pct"]),
            int(win.top + win.height * config["top_pct"]),
        )
        bottom_right = (
            int(win.left + win.width * config["right_pct"]),
            int(win.top + win.height * config["bottom_pct"]),
        )

        logger.info(
            "Grid from config: top-left=%s, bottom-right=%s (%dx%d)",
            top_left,
            bottom_right,
            bottom_right[0] - top_left[0],
            bottom_right[1] - top_left[1],
        )
        return top_left, bottom_right

    except ImportError:
        logger.info("pygetwindow not available")
        return None


def get_grid_coordinates():
    """Prompt the user to click the corners of the game grid."""
    input("Move your mouse to the top-left corner of the game grid and press Enter.")
    top_left = pyautogui.position()
    input("Move your mouse to the bottom-right corner of the game grid and press Enter.")
    bottom_right = pyautogui.position()
    return top_left, bottom_right


def capture_raw(top_left, bottom_right):
    """Capture a raw screenshot of the grid region."""
    screenshot = pyautogui.screenshot(
        region=(
            top_left[0],
            top_left[1],
            bottom_right[0] - top_left[0],
            bottom_right[1] - top_left[1],
        )
    )
    return cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)


def add_grid_overlay(grid_img):
    """Draw green grid lines on an image for visual feedback. Modifies in place."""
    cell_width = grid_img.shape[1] // GRID_SIZE
    cell_height = grid_img.shape[0] // GRID_SIZE
    for i in range(1, GRID_SIZE):
        cv2.line(
            grid_img,
            (i * cell_width, 0),
            (i * cell_width, grid_img.shape[0]),
            (0, 255, 0),
            2,
        )
        cv2.line(
            grid_img,
            (0, i * cell_height),
            (grid_img.shape[1], i * cell_height),
            (0, 255, 0),
            2,
        )
    return grid_img


def wait_for_stable_board(top_left, bottom_right, logger):
    """Wait until the board stops animating by comparing consecutive frames."""
    prev_frame = capture_raw(top_left, bottom_right)
    start = time.time()

    while time.time() - start < MAX_STABILITY_WAIT:
        time.sleep(STABILITY_DELAY)
        curr_frame = capture_raw(top_left, bottom_right)
        diff = np.mean(np.abs(curr_frame.astype(float) - prev_frame.astype(float)))

        if diff < STABILITY_THRESHOLD:
            logger.debug("Board stable (diff=%.2f)", diff)
            return True

        logger.debug("Board animating (diff=%.2f), waiting...", diff)
        prev_frame = curr_frame

    logger.warning("Board stability timeout after %.1fs", MAX_STABILITY_WAIT)
    return False


def classify_hue(hue):
    """Classify an HSV hue value into a gem color name."""
    if hue < 8 or hue > 170:
        return "red"
    for low, high, name in GEM_HUE_RANGES:
        if low <= hue < high:
            return name
    return ""


def identify_cell_color(grid_img, row, col):
    """Identify gem color using HSV analysis of the cell center. Returns (row, col, color_name)."""
    cell_width = grid_img.shape[1] // GRID_SIZE
    cell_height = grid_img.shape[0] // GRID_SIZE

    # Extract center 50% of cell to avoid edges and neighboring gems
    y_start = row * cell_height + cell_height // 4
    y_end = (row + 1) * cell_height - cell_height // 4
    x_start = col * cell_width + cell_width // 4
    x_end = (col + 1) * cell_width - cell_width // 4

    cell = grid_img[y_start:y_end, x_start:x_end]
    hsv = cv2.cvtColor(cell, cv2.COLOR_BGR2HSV)

    h, s, v = hsv[:, :, 0], hsv[:, :, 1], hsv[:, :, 2]
    total_pixels = h.size

    # Check for white first (high brightness, low saturation)
    white_mask = (s < WHITE_MAX_SAT) & (v > WHITE_MIN_VAL)
    if np.count_nonzero(white_mask) > total_pixels * MIN_COLOR_RATIO:
        return row, col, "white"

    # Filter for colorful, bright pixels
    color_mask = (s > MIN_SATURATION) & (v > MIN_VALUE)
    if np.count_nonzero(color_mask) < total_pixels * MIN_COLOR_RATIO:
        return row, col, ""  # Not enough colorful pixels - empty or animating

    # Classify by median hue of colorful pixels
    median_hue = int(np.median(h[color_mask]))
    return row, col, classify_hue(median_hue)


def build_color_grid(grid_img):
    """Identify colors for all cells in parallel. Returns an 8x8 color grid."""
    color_grid = [["" for _ in range(GRID_SIZE)] for _ in range(GRID_SIZE)]

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(identify_cell_color, grid_img, row, col)
            for row in range(GRID_SIZE)
            for col in range(GRID_SIZE)
        ]
        for future in concurrent.futures.as_completed(futures):
            row, col, color = future.result()
            color_grid[row][col] = color

    return color_grid


def is_valid_board(color_grid):
    """Check if the color grid looks like a real game board, not a popup or menu."""
    color_counts = {}
    for row in color_grid:
        for c in row:
            if c:
                color_counts[c] = color_counts.get(c, 0) + 1

    if not color_counts:
        return False, "no colors detected"

    # Check if one color dominates (popup/overlay)
    max_count = max(color_counts.values())
    total = GRID_SIZE * GRID_SIZE
    if max_count > total * MAX_SINGLE_COLOR_RATIO:
        dominant = max(color_counts, key=color_counts.get)
        return False, f"'{dominant}' covers {max_count}/{total} cells"

    # Check color diversity (a real board has at least 4 different gem colors)
    if len(color_counts) < MIN_DISTINCT_COLORS:
        return False, f"only {len(color_counts)} colors (need {MIN_DISTINCT_COLORS}+)"

    return True, ""


def evaluate_state(grid):
    """Score the grid based on consecutive matching gems (3+ in a row/column)."""
    score = 0

    # Check rows
    for row in range(GRID_SIZE):
        col = 0
        while col < GRID_SIZE:
            color = grid[row][col]
            if not color:
                col += 1
                continue
            run_length = 1
            while col + run_length < GRID_SIZE and grid[row][col + run_length] == color:
                run_length += 1
            if run_length >= 3:
                score += run_length
            col += run_length

    # Check columns
    for col in range(GRID_SIZE):
        row = 0
        while row < GRID_SIZE:
            color = grid[row][col]
            if not color:
                row += 1
                continue
            run_length = 1
            while row + run_length < GRID_SIZE and grid[row + run_length][col] == color:
                run_length += 1
            if run_length >= 3:
                score += run_length
            row += run_length

    return score


def evaluate_move(color_grid, row, col, direction):
    """Evaluate a move on a copy of the grid. Returns (score, move_tuple)."""
    grid = [r[:] for r in color_grid]

    if direction == "right":
        grid[row][col], grid[row][col + 1] = grid[row][col + 1], grid[row][col]
        return evaluate_state(grid), (row, col, row, col + 1)

    grid[row][col], grid[row + 1][col] = grid[row + 1][col], grid[row][col]
    return evaluate_state(grid), (row, col, row + 1, col)


def find_optimal_move(color_grid):
    """Find the move producing the highest score. Returns (move_tuple, score) or (None, 0)."""
    best_move = None
    best_score = 0

    moves = []
    for row in range(GRID_SIZE - 1, -1, -1):
        for col in range(GRID_SIZE):
            if not color_grid[row][col]:
                continue
            if col < GRID_SIZE - 1:
                moves.append((color_grid, row, col, "right"))
            if row < GRID_SIZE - 1:
                moves.append((color_grid, row, col, "down"))

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(evaluate_move, *args) for args in moves]
        for future in concurrent.futures.as_completed(futures):
            score, move = future.result()
            if score > best_score:
                best_score = score
                best_move = move

    return best_move, best_score


def perform_move(top_left, bottom_right, src_row, src_col, dest_row, dest_col):
    """Click the source and destination cells to perform a gem swap.

    Uses screen coordinates (not image pixels) to avoid DPI scaling mismatches.
    """
    cell_width = (bottom_right[0] - top_left[0]) / GRID_SIZE
    cell_height = (bottom_right[1] - top_left[1]) / GRID_SIZE

    from_pos = (
        int(top_left[0] + src_col * cell_width + cell_width / 2),
        int(top_left[1] + src_row * cell_height + cell_height / 2),
    )
    to_pos = (
        int(top_left[0] + dest_col * cell_width + cell_width / 2),
        int(top_left[1] + dest_row * cell_height + cell_height / 2),
    )

    pyautogui.click(from_pos)
    time.sleep(0.1)
    pyautogui.click(to_pos)


def main():
    logger = setup_logger()
    logger.info("BejeweledBot started (Bejeweled 3)")

    # Try auto-detection, fall back to manual calibration
    coords = find_grid_from_window(logger)
    if coords:
        top_left, bottom_right = coords
    else:
        logger.info("Auto-detect unavailable, using manual calibration")
        top_left, bottom_right = get_grid_coordinates()

    logger.info(
        "Grid coordinates: top-left=%s, bottom-right=%s", top_left, bottom_right
    )

    last_move = None
    move_count = 0
    game_number = 1
    game_moves = 0
    start_time = time.time()
    non_game_since = None  # When non-game screen was first detected

    logger.info("Game #%d started", game_number)

    while not keyboard.is_pressed("esc"):
        # Wait for animations to finish before scanning
        wait_for_stable_board(top_left, bottom_right, logger)

        raw_image = capture_raw(top_left, bottom_right)
        color_grid = build_color_grid(raw_image)

        # Show overlay for visual feedback
        display_image = add_grid_overlay(raw_image.copy())
        cv2.imshow("Grid Overlay", display_image)
        cv2.waitKey(1)

        identified = sum(1 for row in color_grid for c in row if c)
        logger.debug(
            "Grid state (move #%d):\n%s", move_count + 1, format_grid(color_grid)
        )
        logger.debug("Identified %d/%d cells", identified, GRID_SIZE * GRID_SIZE)

        # Skip if too many cells unidentified (board may still be settling)
        if identified < MIN_IDENTIFIED_CELLS:
            logger.debug(
                "Too few cells identified (%d/%d), skipping frame",
                identified,
                MIN_IDENTIFIED_CELLS,
            )
            continue

        # Check if this looks like a real game board
        valid, reason = is_valid_board(color_grid)
        if not valid:
            if non_game_since is None:
                non_game_since = time.time()
                logger.info("Non-game screen detected (%s), waiting...", reason)

            elapsed_non_game = time.time() - non_game_since

            # Brief interruptions (level transitions, bonus animations) clear
            # within a few seconds. Only pause if it persists 10+ seconds.
            if elapsed_non_game < 10:
                logger.debug(
                    "Non-game screen for %.1fs, waiting...", elapsed_non_game
                )
                time.sleep(1.0)
                continue

            # Persistent non-game screen = game over
            logger.info(
                "Game #%d ended with %d moves (non-game screen for %.1fs). "
                "Press Space to resume or Escape to quit.",
                game_number,
                game_moves,
                elapsed_non_game,
            )

            # Pause until user presses Space or Escape
            while True:
                if keyboard.is_pressed("space"):
                    time.sleep(0.3)  # Debounce
                    break
                if keyboard.is_pressed("esc"):
                    elapsed = time.time() - start_time
                    logger.info(
                        "Bot stopped. Total moves: %d across %d game(s), Duration: %.1fs",
                        move_count,
                        game_number,
                        elapsed,
                    )
                    cv2.destroyAllWindows()
                    return
                time.sleep(0.1)

            non_game_since = None
            game_number += 1
            game_moves = 0
            last_move = None
            logger.info("Resuming - Game #%d", game_number)

            # Re-detect grid in case window moved
            new_coords = find_grid_from_window(logger)
            if new_coords:
                top_left, bottom_right = new_coords
                logger.info(
                    "Updated grid: top-left=%s, bottom-right=%s",
                    top_left,
                    bottom_right,
                )
            continue

        # Board is valid — reset non-game timer
        non_game_since = None

        move, score = find_optimal_move(color_grid)

        if move:
            from_row, from_col, to_row, to_col = move

            # Stuck detection: compare with previous move before executing
            is_repeat = last_move == move
            if is_repeat:
                logger.info("Repeat move detected, adding delay")
                time.sleep(0.3)

            move_count += 1
            game_moves += 1
            src_color = color_grid[from_row][from_col]
            dest_color = color_grid[to_row][to_col]
            logger.info(
                "Move #%d: [%d,%d] (%s) -> [%d,%d] (%s) | score=%d%s",
                move_count,
                from_row,
                from_col,
                src_color,
                to_row,
                to_col,
                dest_color,
                score,
                " (repeat)" if is_repeat else "",
            )

            perform_move(top_left, bottom_right, from_row, from_col, to_row, to_col)
            last_move = move
        else:
            logger.debug("No valid move found this frame")

    elapsed = time.time() - start_time
    logger.info(
        "Bot stopped. Total moves: %d across %d game(s), Duration: %.1fs",
        move_count,
        game_number,
        elapsed,
    )
    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
