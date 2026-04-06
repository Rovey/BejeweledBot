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

try:
    import win32api
    import win32con
    import win32gui

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

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

# Special gem detection thresholds (calibrated from gem_library captures)
SPECIAL_BORDER_V_THRESHOLD = 130  # Regular max 104, special min 140 (raised from 120 to avoid hint glow)
HYPERCUBE_BORDER_V_THRESHOLD = 150  # Hypercube 166, hint glow ~125-145 (extra check to prevent false positives)
HYPERCUBE_BORDER_S_THRESHOLD = 90  # Hypercube 73, next lowest 102
FLAME_HUE_STD_THRESHOLD = 35  # Flame min 43, star max 30

# Bonus scores for moves involving special gems
HYPERCUBE_SWAP_BONUS = 1000  # Clears entire color (moderate: save for emergencies)
FLAME_MATCH_BONUS = 200  # 3x3 explosion (~8 extra gems)
STAR_MATCH_BONUS = 400  # Cross detonation (~14 extra gems)

# Single-letter abbreviations for log output
COLOR_ABBREV = {
    "blue": "B",
    "green": "G",
    "red": "R",
    "purple": "P",
    "orange": "O",
    "yellow": "Y",
    "white": "W",
    "hypercube": "H",
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
    """Format the color grid as a compact string for logging.

    Regular gems: uppercase (R, B, G, ...). Flame: lowercase (r, b, g, ...).
    Star: uppercase + * (R*, B*). Hypercube: H. Unknown: dot.
    """
    lines = ["   " + "  ".join(str(c) for c in range(GRID_SIZE))]
    for row_idx, row in enumerate(color_grid):
        cells = []
        for c in row:
            if not c:
                cells.append(" .")
            elif c == "hypercube":
                cells.append(" H")
            elif c.endswith("_flame"):
                cells.append(" " + COLOR_ABBREV.get(gem_base_color(c), "?").lower())
            elif c.endswith("_star"):
                cells.append(COLOR_ABBREV.get(gem_base_color(c), "?") + "*")
            else:
                cells.append(" " + COLOR_ABBREV.get(c, "?"))
        lines.append(f"{row_idx} " + " ".join(cells))
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
        "right_pct": 0.970,
        "bottom_pct": 0.914,
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

        hwnd = win._hWnd if HAS_WIN32 else None

        logger.info(
            "Grid from config: top-left=%s, bottom-right=%s (%dx%d)",
            top_left,
            bottom_right,
            bottom_right[0] - top_left[0],
            bottom_right[1] - top_left[1],
        )
        return top_left, bottom_right, hwnd

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


def gem_base_color(gem_name):
    """Extract base color from a gem name (e.g., 'red_flame' -> 'red')."""
    if not gem_name or gem_name == "hypercube":
        return gem_name
    return gem_name.split("_")[0]


def classify_hue(hue):
    """Classify an HSV hue value into a gem color name."""
    if hue < 8 or hue >= 170:
        return "red"
    for low, high, name in GEM_HUE_RANGES:
        if low <= hue < high:
            return name


def classify_special(cell_bgr):
    """Classify a gem cell as regular, flame, star, or hypercube.

    Analyzes the border region (outer 25% ring) of the cell:
    - Regular gems have dark borders (background)
    - Flame gems have bright, saturated orange/yellow fire aura
    - Star gems have bright, desaturated white/blue glow
    - Hypercube has bright border with low saturation (multicolor metallic)
    """
    hsv = cv2.cvtColor(cell_bgr, cv2.COLOR_BGR2HSV)
    h, s, v = hsv[:, :, 0], hsv[:, :, 1], hsv[:, :, 2]
    cell_h, cell_w = cell_bgr.shape[:2]

    # Border = outer ring (everything outside center 50%)
    border_mask = np.ones((cell_h, cell_w), dtype=bool)
    border_mask[cell_h // 4 : cell_h * 3 // 4, cell_w // 4 : cell_w * 3 // 4] = False

    border_v = float(np.mean(v[border_mask]))

    # Regular gems have dark borders
    if border_v < SPECIAL_BORDER_V_THRESHOLD:
        return "regular"

    border_s = float(np.mean(s[border_mask]))

    # Hypercube candidate: low border saturation AND very bright.
    # Extra check: the CENTER must also have multiple distinct colors (yellow/green/purple).
    # This prevents white flames (single color + low saturation glow) from being
    # misidentified as hypercubes.
    if border_s < HYPERCUBE_BORDER_S_THRESHOLD and border_v > HYPERCUBE_BORDER_V_THRESHOLD:
        center = hsv[cell_h // 4 : cell_h * 3 // 4, cell_w // 4 : cell_w * 3 // 4]
        ch, cs, cv = center[:, :, 0], center[:, :, 1], center[:, :, 2]
        bright_mask = (cs > 40) & (cv > 40)
        if np.count_nonzero(bright_mask) > 10:
            center_hue_std = float(np.std(ch[bright_mask]))
            if center_hue_std > 35:
                return "hypercube"
        # Low border saturation but single-color center = white flame/star

    # Flame vs Star: flames have high hue variance (warm fire colors),
    # stars have low hue variance (uniform white/blue glow)
    color_mask = (s > 60) & (v > 60)
    if np.count_nonzero(color_mask) > 10:
        hue_std = float(np.std(h[color_mask]))
        if hue_std > FLAME_HUE_STD_THRESHOLD:
            return "flame"

    return "star"


_unknown_gem_timestamps = {}  # Throttle: track last save time per cell


def identify_cell_color(grid_img, row, col):
    """Identify gem color and special type. Returns (row, col, gem_name).

    gem_name is one of: 'red', 'blue', ..., 'red_flame', 'blue_star', 'hypercube', or ''.
    """
    cell_width = grid_img.shape[1] // GRID_SIZE
    cell_height = grid_img.shape[0] // GRID_SIZE

    # Full cell for special gem detection
    full_cell = grid_img[
        row * cell_height : (row + 1) * cell_height,
        col * cell_width : (col + 1) * cell_width,
    ]

    # Check for special gem type (uses border glow analysis)
    special = classify_special(full_cell)

    # Hypercube has no base color — return immediately
    if special == "hypercube":
        return row, col, "hypercube"

    # Center 50% for base color classification
    center = full_cell[
        cell_height // 4 : cell_height - cell_height // 4,
        cell_width // 4 : cell_width - cell_width // 4,
    ]
    hsv = cv2.cvtColor(center, cv2.COLOR_BGR2HSV)

    h, s, v = hsv[:, :, 0], hsv[:, :, 1], hsv[:, :, 2]
    total_pixels = h.size

    # Check for white first (high brightness, low saturation)
    white_mask = (s < WHITE_MAX_SAT) & (v > WHITE_MIN_VAL)
    if np.count_nonzero(white_mask) > total_pixels * MIN_COLOR_RATIO:
        base = "white"
    else:
        # Filter for colorful, bright pixels
        color_mask = (s > MIN_SATURATION) & (v > MIN_VALUE)
        if np.count_nonzero(color_mask) < total_pixels * MIN_COLOR_RATIO:
            # Unknown — save screenshot for identification (throttled)
            cell_key = (row, col)
            now = time.time()
            if now - _unknown_gem_timestamps.get(cell_key, 0) > 5:
                _unknown_gem_timestamps[cell_key] = now
                unknown_dir = os.path.join(os.path.dirname(__file__) or ".", "unknown_gems")
                os.makedirs(unknown_dir, exist_ok=True)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                cv2.imwrite(os.path.join(unknown_dir, f"r{row}_c{col}_{ts}.png"), full_cell)
            return row, col, ""

        median_hue = int(np.median(h[color_mask]))
        base = classify_hue(median_hue)
        if not base:
            return row, col, ""

    # Combine base color with special type
    if special in ("flame", "star"):
        return row, col, f"{base}_{special}"
    return row, col, base


_library_saved = set()  # Track which (color, hue) combos we've already saved


def save_gem_library(grid_img, color_grid):
    """Save one example screenshot per unique gem type encountered.

    Saves to gem_library/<color>/ folders. Only saves new types not yet
    in the library, so it builds up over time without duplicates.
    """
    cell_w = grid_img.shape[1] // GRID_SIZE
    cell_h = grid_img.shape[0] // GRID_SIZE

    for row in range(GRID_SIZE):
        for col in range(GRID_SIZE):
            color = color_grid[row][col]
            if not color:
                continue

            # Only save one example per color (skip if we already have it)
            if color in _library_saved:
                continue
            _library_saved.add(color)

            color_dir = os.path.join(
                os.path.dirname(__file__) or ".", "gem_library", color
            )
            os.makedirs(color_dir, exist_ok=True)

            full_cell = grid_img[
                row * cell_h : (row + 1) * cell_h,
                col * cell_w : (col + 1) * cell_w,
            ]
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            cv2.imwrite(os.path.join(color_dir, f"{color}_{ts}.png"), full_cell)


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
                bc = gem_base_color(c)
                color_counts[bc] = color_counts.get(bc, 0) + 1

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


# Heuristic scoring weights (loosely based on Bejeweled 3 point values)
MATCH_WEIGHTS = {3: 50, 4: 100, 5: 500}  # 4=Flame Gem, 5=Hypercube
STAR_GEM_BONUS = 150  # L/T match creating a Star Gem
CASCADE_BASE_BONUS = 50  # Increases by 50 per cascade level
BOTTOM_ROW_BONUS = 2  # Small tiebreaker per row toward bottom


def find_matches(grid):
    """Find all matches on the grid. Returns list of (row, col, length, direction).

    Compares gems by base color (red_flame matches with red and red_star).
    Hypercubes don't form line matches. direction is 'h' or 'v'.
    """
    matches = []

    # Horizontal matches
    for row in range(GRID_SIZE):
        col = 0
        while col < GRID_SIZE:
            color = grid[row][col]
            bc = gem_base_color(color)
            if not bc or bc == "hypercube":
                col += 1
                continue
            run = 1
            while col + run < GRID_SIZE and gem_base_color(grid[row][col + run]) == bc:
                run += 1
            if run >= 3:
                matches.append((row, col, run, "h"))
            col += run

    # Vertical matches
    for col in range(GRID_SIZE):
        row = 0
        while row < GRID_SIZE:
            color = grid[row][col]
            bc = gem_base_color(color)
            if not bc or bc == "hypercube":
                row += 1
                continue
            run = 1
            while row + run < GRID_SIZE and gem_base_color(grid[row + run][col]) == bc:
                run += 1
            if run >= 3:
                matches.append((row, col, run, "v"))
            row += run

    return matches


def detect_star_gems(matches):
    """Count Star Gem formations (L/T/+ shapes where horizontal and vertical matches share a cell)."""
    h_cells = set()
    v_cells = set()
    for row, col, length, direction in matches:
        for i in range(length):
            if direction == "h":
                h_cells.add((row, col + i))
            else:
                v_cells.add((row + i, col))
    return len(h_cells & v_cells)


def clear_matches(grid, matches):
    """Clear matched gems from the grid (set to empty string)."""
    for row, col, length, direction in matches:
        for i in range(length):
            if direction == "h":
                grid[row][col + i] = ""
            else:
                grid[row + i][col] = ""


def apply_gravity(grid):
    """Drop gems down to fill empty spaces. Empty cells bubble to the top."""
    for col in range(GRID_SIZE):
        # Collect non-empty cells from bottom to top
        gems = [grid[row][col] for row in range(GRID_SIZE - 1, -1, -1) if grid[row][col]]
        # Fill column from bottom
        for row in range(GRID_SIZE - 1, -1, -1):
            idx = GRID_SIZE - 1 - row
            grid[row][col] = gems[idx] if idx < len(gems) else ""


def evaluate_state(grid, copy=True):
    """Score the grid using cascade simulation. Returns (score, resulting_grid).

    Heuristic scoring:
    - Match-3: 50 pts
    - Match-4 (Flame Gem): 100 pts
    - Match-5 (Hypercube): 500 pts
    - L/T intersection (Star Gem): +150 pts bonus
    - Cascades: +50, +100, +150... stacking bonus per level

    If copy=False, the input grid is modified in place (caller must provide a copy).
    """
    if copy:
        grid = [r[:] for r in grid]
    total_score = 0
    cascade_level = 0

    while True:
        matches = find_matches(grid)
        if not matches:
            break

        # Score each match by length
        match_score = 0
        for _, _, length, _ in matches:
            capped = min(length, 5)
            match_score += MATCH_WEIGHTS.get(capped, MATCH_WEIGHTS[5])

        # Star Gem bonus for L/T intersections
        star_count = detect_star_gems(matches)
        match_score += star_count * STAR_GEM_BONUS

        # Cascade bonus (increases each level)
        cascade_bonus = cascade_level * CASCADE_BASE_BONUS
        total_score += match_score + cascade_bonus

        # Simulate: clear matches and apply gravity
        clear_matches(grid, matches)
        apply_gravity(grid)
        cascade_level += 1

    return total_score, grid


def _score_swap(grid, row, col, direction):
    """Score a swap on a grid (modified in place). Returns (score, resulting_grid).

    Handles cascade simulation and special gem bonuses. Only applies flame/star
    bonuses when the special gem is actually part of a match.
    The caller must provide a copy — this function modifies grid in place.
    """
    if direction == "right":
        src, dst = grid[row][col], grid[row][col + 1]
        grid[row][col], grid[row][col + 1] = dst, src
        # After swap: src is at col+1, dst is at col
        src_pos, dst_pos = (row, col + 1), (row, col)
    else:
        src, dst = grid[row][col], grid[row + 1][col]
        grid[row][col], grid[row + 1][col] = dst, src
        # After swap: src is at row+1, dst is at row
        src_pos, dst_pos = (row + 1, col), (row, col)

    # Find initial matches to check which gems are involved
    initial_matches = find_matches(grid)
    matched_cells = set()
    for r, c, length, d in initial_matches:
        for i in range(length):
            if d == "h":
                matched_cells.add((r, c + i))
            else:
                matched_cells.add((r + i, c))

    score, grid = evaluate_state(grid, copy=False)

    # Special gem bonuses (only if the special gem is part of a match)
    if src == "hypercube" or dst == "hypercube":
        score += HYPERCUBE_SWAP_BONUS
    elif score > 0:
        for gem, pos in ((src, src_pos), (dst, dst_pos)):
            if gem and pos in matched_cells:
                if "_flame" in gem:
                    score += FLAME_MATCH_BONUS
                elif "_star" in gem:
                    score += STAR_MATCH_BONUS

    return score, grid


def simulate_move(grid, row, col, direction):
    """Apply a swap on a grid copy, run cascades, return (score, resulting_grid).

    The resulting grid has empty spaces where gems were cleared (pessimistic:
    no new gems fall from above, only existing gems drop via gravity).
    """
    return _score_swap([r[:] for r in grid], row, col, direction)


def best_next_score(grid):
    """Find the best single-move score on a board state (for look-ahead)."""
    best = 0
    for row in range(GRID_SIZE):
        for col in range(GRID_SIZE):
            if not gem_base_color(grid[row][col]):
                continue
            if col < GRID_SIZE - 1 and grid[row][col + 1]:
                s, _ = _score_swap([r[:] for r in grid], row, col, "right")
                if s > best:
                    best = s
            if row < GRID_SIZE - 1 and grid[row + 1][col]:
                s, _ = _score_swap([r[:] for r in grid], row, col, "down")
                if s > best:
                    best = s
    return best


def evaluate_move(color_grid, row, col, direction):
    """Evaluate a move with 2-step look-ahead. Returns (score, move_tuple).

    Step 1: simulate the move and its cascades, get score and resulting board.
    Step 2: on the resulting board (empty spaces stay empty), find the best
    possible follow-up move and add its score (discounted).
    """
    if direction == "right":
        move = (row, col, row, col + 1)
    else:
        move = (row, col, row + 1, col)

    # Step 1: evaluate this move
    step1_score, resulting_grid = simulate_move(color_grid, row, col, direction)

    # In Bejeweled, a swap is only valid if it creates an immediate match.
    # If step1 scores 0 (no match), the game rejects the move entirely.
    if step1_score == 0:
        return 0, move

    # Step 2: best follow-up move on the resulting board (discounted by ~33%).
    # The board already has empty cells (no new gems simulated), so step-2
    # scores are naturally deflated — a mild discount avoids double-penalizing.
    step2_score = best_next_score(resulting_grid)
    total = step1_score + step2_score * 2 // 3

    # Tiebreaker: prefer moves lower on the board (more cascade potential)
    total += row * BOTTOM_ROW_BONUS
    return total, move


def find_optimal_move(color_grid, failed_moves=None):
    """Find the move producing the highest score. Returns (move_tuple, score) or (None, 0).

    failed_moves: set of move tuples to skip (moves that have been tried
    repeatedly without the board changing, likely involving special gems).
    """
    if failed_moves is None:
        failed_moves = set()

    best_move = None
    best_score = 0

    moves = []
    for row in range(GRID_SIZE):
        for col in range(GRID_SIZE):
            if not color_grid[row][col]:
                continue
            if col < GRID_SIZE - 1:
                move_tuple = (row, col, row, col + 1)
                if move_tuple not in failed_moves:
                    moves.append((color_grid, row, col, "right"))
            if row < GRID_SIZE - 1:
                move_tuple = (row, col, row + 1, col)
                if move_tuple not in failed_moves:
                    moves.append((color_grid, row, col, "down"))

    # Evaluate sequentially — Python's GIL prevents ThreadPoolExecutor from
    # achieving real parallelism on CPU-bound work like move scoring.
    for args in moves:
        score, move = evaluate_move(*args)
        if score > best_score:
            best_score = score
            best_move = move

    return best_move, best_score


def _send_click(hwnd, screen_x, screen_y):
    """Send a mouse click to a window without moving the physical mouse."""
    client_x, client_y = win32gui.ScreenToClient(hwnd, (screen_x, screen_y))
    lparam = win32api.MAKELONG(client_x, client_y)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)


def perform_move(top_left, bottom_right, src_row, src_col, dest_row, dest_col, hwnd=None):
    """Click the source and destination cells to perform a gem swap.

    Uses screen coordinates (not image pixels) to avoid DPI scaling mismatches.
    If hwnd is provided, sends clicks via SendMessage (doesn't move the mouse).
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

    if hwnd:
        _send_click(hwnd, *from_pos)
        time.sleep(0.1)
        _send_click(hwnd, *to_pos)
    else:
        pyautogui.click(from_pos)
        time.sleep(0.1)
        pyautogui.click(to_pos)


def main():
    logger = setup_logger()
    logger.info("BejeweledBot started (Bejeweled 3)")

    # Try auto-detection, fall back to manual calibration
    coords = find_grid_from_window(logger)
    if coords:
        top_left, bottom_right, hwnd = coords
    else:
        logger.info("Auto-detect unavailable, using manual calibration")
        top_left, bottom_right = get_grid_coordinates()
        hwnd = None

    if hwnd:
        logger.info("Using SendMessage for clicks (mouse stays free)")
    else:
        logger.info("Using pyautogui for clicks (mouse will be controlled)")

    logger.info(
        "Grid coordinates: top-left=%s, bottom-right=%s", top_left, bottom_right
    )

    last_move = None
    move_count = 0
    game_number = 1
    game_moves = 0
    start_time = time.time()
    non_game_since = None  # When non-game screen was first detected
    move_history = []  # Track recent moves to detect stuck loops
    failed_moves = set()  # Moves to skip (tried repeatedly without effect)
    prev_grid_state = None  # Track board state to clear blacklist on change

    logger.info("Game #%d started", game_number)

    while not keyboard.is_pressed("esc"):
        # Wait for animations to finish before scanning
        wait_for_stable_board(top_left, bottom_right, logger)

        raw_image = capture_raw(top_left, bottom_right)
        color_grid = build_color_grid(raw_image)

        # Build gem screenshot library (saves one example per color type)
        save_gem_library(raw_image, color_grid)

        # P: snapshot all 64 cells for manual review (catches special gems)
        if keyboard.is_pressed("p"):
            snap_dir = os.path.join(os.path.dirname(__file__) or ".", "gem_library", "snapshot")
            os.makedirs(snap_dir, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            cell_w = raw_image.shape[1] // GRID_SIZE
            cell_h = raw_image.shape[0] // GRID_SIZE
            for r in range(GRID_SIZE):
                for c in range(GRID_SIZE):
                    cell_img = raw_image[r * cell_h : (r + 1) * cell_h, c * cell_w : (c + 1) * cell_w]
                    color = color_grid[r][c] or "empty"
                    cv2.imwrite(os.path.join(snap_dir, f"{ts}_r{r}_c{c}_{color}.png"), cell_img)
            logger.info("Snapshot: saved 64 cells to gem_library/snapshot/")
            time.sleep(0.5)  # Debounce

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
            move_history.clear()
            failed_moves.clear()
            logger.info("Resuming - Game #%d", game_number)

            # Re-detect grid in case window moved
            new_coords = find_grid_from_window(logger)
            if new_coords:
                top_left, bottom_right, hwnd = new_coords
                logger.info(
                    "Updated grid: top-left=%s, bottom-right=%s",
                    top_left,
                    bottom_right,
                )
            continue

        # Board is valid — reset non-game timer
        non_game_since = None

        # Clear blacklist only on significant board changes (4+ cells = real cascade).
        # Minor flickers (1-2 cells from hint glow) should NOT reset the blacklist.
        grid_state = tuple(tuple(row) for row in color_grid)
        if grid_state != prev_grid_state:
            if prev_grid_state:
                changes = sum(
                    a != b
                    for row_a, row_b in zip(grid_state, prev_grid_state)
                    for a, b in zip(row_a, row_b)
                )
                if changes >= 4:
                    if failed_moves:
                        logger.debug(
                            "Board changed (%d cells), clearing %d blacklisted moves",
                            changes,
                            len(failed_moves),
                        )
                    failed_moves.clear()
                    move_history.clear()
            prev_grid_state = grid_state

        move, score = find_optimal_move(color_grid, failed_moves)

        if move:
            # --- Double-scan validation ---
            # Wait briefly then re-capture to ensure no animation is playing.
            # Without this delay, we might screenshot mid-animation and see
            # a false-stable board that matches the first scan by coincidence.
            time.sleep(0.7)
            raw_image2 = capture_raw(top_left, bottom_right)
            color_grid2 = build_color_grid(raw_image2)
            grid_state2 = tuple(tuple(row) for row in color_grid2)

            if grid_state2 != grid_state:
                changes = sum(
                    a != b
                    for row_a, row_b in zip(grid_state, grid_state2)
                    for a, b in zip(row_a, row_b)
                )
                if changes >= 4:
                    logger.debug(
                        "Board changed during planning (%d cells), re-scanning",
                        changes,
                    )
                    continue

            from_row, from_col, to_row, to_col = move

            # Track move history for stuck loop detection
            move_history.append(move)
            if len(move_history) > 10:
                move_history.pop(0)

            # Count how many times this move appears in recent history
            repeat_count = move_history.count(move)

            if repeat_count >= 3:
                # This move has been tried 3+ times recently without effect.
                # Blacklist ALL moves involving these cells (not just this swap
                # direction) to force the bot to try a different area of the board.
                r1, c1, r2, c2 = move
                for cell_r, cell_c in ((r1, c1), (r2, c2)):
                    for dr, dc in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                        nr, nc = cell_r + dr, cell_c + dc
                        if 0 <= nr < GRID_SIZE and 0 <= nc < GRID_SIZE:
                            if dr == 0:  # horizontal
                                failed_moves.add((cell_r, min(cell_c, nc), cell_r, max(cell_c, nc)))
                            else:  # vertical
                                failed_moves.add((min(cell_r, nr), cell_c, max(cell_r, nr), cell_c))
                logger.info(
                    "Move [%d,%d]->[%d,%d] stuck %d times, blacklisting area (%d moves blocked)",
                    *move, repeat_count, len(failed_moves),
                )
                move, score = find_optimal_move(color_grid, failed_moves)
                if not move:
                    logger.info("No alternative moves, clearing blacklist")
                    failed_moves.clear()
                    move_history.clear()
                    time.sleep(0.5)
                    continue

                from_row, from_col, to_row, to_col = move
                # Track replacement move too so IT can also be blacklisted
                move_history.append(move)
                if len(move_history) > 10:
                    move_history.pop(0)

            move_count += 1
            game_moves += 1
            src_color = color_grid[from_row][from_col]
            dest_color = color_grid[to_row][to_col]
            logger.info(
                "Move #%d: [%d,%d] (%s) -> [%d,%d] (%s) | score=%d",
                move_count,
                from_row,
                from_col,
                src_color,
                to_row,
                to_col,
                dest_color,
                score,
            )

            perform_move(top_left, bottom_right, from_row, from_col, to_row, to_col, hwnd)
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
