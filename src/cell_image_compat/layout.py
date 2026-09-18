"""V1 image placement: merged bounds and normalized DrawingML markers."""

EMU_PER_PIXEL = 9525
DEFAULT_MARGIN = 2


def merged_bounds(sheet, col, row, ns, cell_position):
    for node in sheet.findall("main:mergeCells/main:mergeCell", ns):
        first, _, last = node.attrib["ref"].partition(":")
        left, top = cell_position(first)
        right, bottom = cell_position(last or first)
        if left <= col <= right and top <= row <= bottom:
            return left, top, right, bottom
    return col, row, col, row


def region_geometry(sheet, bounds, cell_size):
    left, top, right, bottom = bounds
    return (
        sum(cell_size(sheet, col, top)[0] for col in range(left, right + 1)),
        sum(cell_size(sheet, left, row)[1] for row in range(top, bottom + 1)),
    )


def offset_marker(sheet, col, row, x, y, cell_size):
    """Normalize offsets across columns/rows (including merged regions)."""
    while col < 16383 and x >= cell_size(sheet, col, row)[0]:
        x -= cell_size(sheet, col, row)[0]
        col += 1
    while row < 1048575 and y >= cell_size(sheet, col, row)[1]:
        y -= cell_size(sheet, col, row)[1]
        row += 1
    return (("col", col), ("colOff", x), ("row", row), ("rowOff", y))
