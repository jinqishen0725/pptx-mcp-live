"""PowerPoint COM connection manager."""

import pythoncom
import win32com.client as win32

from .errors import ToolError


def get_powerpoint():
    """Get the running PowerPoint application instance.

    Returns:
        PowerPoint.Application COM object.

    Raises:
        ToolError: If PowerPoint is not running.
    """
    pythoncom.CoInitialize()
    try:
        return win32.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        raise ToolError("Could not connect to PowerPoint. Is it running?") from e


def get_presentation(app, name=None):
    """Get a presentation by name, or the active one.

    Args:
        app: PowerPoint.Application COM object.
        name: Presentation filename (e.g., "slides.pptx"). None = active.

    Returns:
        Presentation COM object.

    Raises:
        ToolError: If presentation not found.
    """
    if name is None:
        try:
            return app.ActivePresentation
        except Exception as e:
            raise ToolError("No active presentation. Open a file in PowerPoint.") from e

    for i in range(1, app.Presentations.Count + 1):
        pres = app.Presentations(i)
        if pres.Name == name:
            return pres

    raise ToolError(f"Presentation '{name}' not found. Is it open?")


def get_slide(pres, slide_index):
    """Get a slide by 1-based index.

    Args:
        pres: Presentation COM object.
        slide_index: 1-based slide index.

    Returns:
        Slide COM object.

    Raises:
        ToolError: If slide index is out of range.
    """
    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ToolError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})."
        )
    return pres.Slides(slide_index)


def get_shape(slide, shape_ref):
    """Get a shape by name or 1-based index.

    Args:
        slide: Slide COM object.
        shape_ref: Shape name (str) or 1-based index (int).

    Returns:
        Shape COM object.

    Raises:
        ToolError: If shape not found.
    """
    if isinstance(shape_ref, int):
        if shape_ref < 1 or shape_ref > slide.Shapes.Count:
            raise ToolError(
                f"Shape index {shape_ref} out of range (1-{slide.Shapes.Count})."
            )
        return slide.Shapes(shape_ref)

    # Search by name
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.Name == shape_ref:
            return shape

    raise ToolError(
        f"Shape '{shape_ref}' not found on slide. "
        f"Available: {[slide.Shapes(i).Name for i in range(1, slide.Shapes.Count + 1)]}"
    )


def inches_to_points(inches):
    """Convert inches to points (1 inch = 72 points)."""
    return float(inches) * 72


def points_to_inches(points):
    """Convert points to inches."""
    return float(points) / 72
