"""Media tools: add image, table, chart."""

from typing import Any, Dict, List, Optional

from ..core.connection import (
    get_powerpoint, get_presentation, get_slide, inches_to_points, points_to_inches
)
from ..core.errors import ToolError


def add_image_sync(
    slide_index: int,
    image_path: str,
    left_inches: float = 1.0,
    top_inches: float = 1.0,
    width_inches: Optional[float] = None,
    height_inches: Optional[float] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add an image to a slide.

    Args:
        image_path: Absolute path to image file (PNG, JPG, BMP, etc.).
        width_inches/height_inches: Size. If only one given, aspect ratio is maintained.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    import os
    if not os.path.exists(image_path):
        raise ToolError(f"Image file not found: {image_path}")

    # Add picture: LinkToFile=False, SaveWithDocument=True
    shape = slide.Shapes.AddPicture(
        image_path,
        False,  # LinkToFile
        True,   # SaveWithDocument
        inches_to_points(left_inches),
        inches_to_points(top_inches),
        inches_to_points(width_inches) if width_inches else -1,
        inches_to_points(height_inches) if height_inches else -1,
    )

    # If only one dimension given, lock aspect ratio
    if width_inches and not height_inches:
        shape.LockAspectRatio = True
        shape.Width = inches_to_points(width_inches)
    elif height_inches and not width_inches:
        shape.LockAspectRatio = True
        shape.Height = inches_to_points(height_inches)

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "position": {
            "left_inches": round(points_to_inches(shape.Left), 2),
            "top_inches": round(points_to_inches(shape.Top), 2),
            "width_inches": round(points_to_inches(shape.Width), 2),
            "height_inches": round(points_to_inches(shape.Height), 2),
        },
        "message": f"Image added to slide {slide_index}.",
    }


def add_table_sync(
    slide_index: int,
    rows: int,
    cols: int,
    data: Optional[List[List[str]]] = None,
    left_inches: float = 1.0,
    top_inches: float = 2.0,
    width_inches: float = 8.0,
    height_inches: float = 3.0,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add a table to a slide.

    Args:
        rows: Number of rows.
        cols: Number of columns.
        data: 2D list of cell values. First row is header.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    shape = slide.Shapes.AddTable(
        rows, cols,
        inches_to_points(left_inches),
        inches_to_points(top_inches),
        inches_to_points(width_inches),
        inches_to_points(height_inches),
    )

    table = shape.Table

    # Fill data if provided
    if data:
        for r_idx, row in enumerate(data):
            for c_idx, val in enumerate(row):
                if r_idx < rows and c_idx < cols:
                    table.Cell(r_idx + 1, c_idx + 1).Shape.TextFrame.TextRange.Text = str(val)

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "rows": rows,
        "cols": cols,
        "message": f"Table ({rows}x{cols}) added to slide {slide_index}.",
    }


def add_chart_sync(
    slide_index: int,
    chart_type: str = "column",
    data: Optional[List[List]] = None,
    left_inches: float = 1.0,
    top_inches: float = 1.5,
    width_inches: float = 8.0,
    height_inches: float = 5.0,
    title: Optional[str] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add a chart to a slide.

    Args:
        chart_type: "column", "bar", "line", "pie", "area", "scatter".
        data: 2D list: first row = categories, subsequent rows = [series_name, val1, val2, ...].
        title: Chart title.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    chart_types = {
        "column": 51,       # xlColumnClustered
        "bar": 57,          # xlBarClustered
        "line": 4,          # xlLine
        "pie": 5,           # xlPie
        "area": 1,          # xlArea
        "scatter": -4169,   # xlXYScatter
    }

    xl_type = chart_types.get(chart_type.lower())
    if xl_type is None:
        raise ToolError(
            f"Invalid chart type '{chart_type}'. Use: {', '.join(chart_types.keys())}"
        )

    # Close any open chart data grids before adding a new chart
    try:
        for i in range(1, pres.Slides.Count + 1):
            for j in range(1, pres.Slides(i).Shapes.Count + 1):
                shp = pres.Slides(i).Shapes(j)
                if shp.HasChart:
                    try:
                        shp.Chart.ChartData.Workbook.Close(False)
                    except Exception:
                        pass
    except Exception:
        pass

    shape = slide.Shapes.AddChart2(
        -1,  # Style: default
        xl_type,
        inches_to_points(left_inches),
        inches_to_points(top_inches),
        inches_to_points(width_inches),
        inches_to_points(height_inches),
    )

    chart = shape.Chart

    # Set data if provided
    if data and len(data) > 1:
        try:
            wb = chart.ChartData.Workbook
            ws = wb.Worksheets(1)
            # Clear existing data
            ws.Cells.Clear()
            # Write data
            for r_idx, row in enumerate(data):
                for c_idx, val in enumerate(row):
                    ws.Cells(r_idx + 1, c_idx + 1).Value = val
            # Set data range
            num_rows = len(data)
            num_cols = max(len(row) for row in data)
            chart.SetSourceData(ws.Range(ws.Cells(1, 1), ws.Cells(num_rows, num_cols)))
            wb.Close(False)
        except Exception as e:
            raise ToolError(f"Failed to set chart data: {e}") from e

    # Set title
    if title:
        chart.HasTitle = True
        chart.ChartTitle.Text = title

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "chart_type": chart_type,
        "title": title,
        "message": f"Chart ({chart_type}) added to slide {slide_index}.",
    }
