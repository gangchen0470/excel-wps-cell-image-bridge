"""Excel/WPS cell-image compatibility core."""

from .core import convert_floating_to_native, convert_native, inspect_workbook, make_compatible_copy

__all__ = ["inspect_workbook", "make_compatible_copy", "convert_native", "convert_floating_to_native"]
