# UI Layout Redesign Spec

## Goal
Redesign the GUI to a horizontal desktop-friendly layout with top config bar and 50/50 split below.

## Layout Structure

### Top Bar (compact, 2 rows)
- Row 1: API Key input + show/hide toggle
- Row 2: Keep folder path input + browse button
- No cards, minimal padding, spans full width

### Bottom Left (50% width)
- Start Scan button (full width, prominent)
- Progress bar + status text
- Current file info (name, size, date, counter)
- Action buttons: [保留] [移到temp] / [跳过] [退出] (2x2 grid)
- Scrollable if content exceeds height

### Bottom Right (50% width)
- AI Summary area (top 50%) - large text area for readability
- Log area (bottom 50%) - scrollable log output
- Both in card-style containers with section labels

## Style
- Apple-inspired: #F5F5F7 background, #FFFFFF cards, #0071E3 accent
- Full-screen (zoomed) on launch
- Segoe UI font, clean spacing
- Card-based sections with subtle borders

## Technical
- Use `tk.PanedWindow` for the left/right split
- Use `tk.PanedWindow` for the right summary/log split
- Mouse wheel binding fixed for left pane scrolling
- All existing functionality preserved
