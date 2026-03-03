"""Layout tools: move, resize, align, distribute, arrange, group, rotate."""

from typing import Any, Dict, List, Optional

from ..core.connection import (
    get_powerpoint, get_presentation, get_slide, get_shape,
    inches_to_points, points_to_inches
)
from ..core.errors import ToolError


def move_shape_sync(
    slide_index: int,
    shape_ref,
    left_inches: Optional[float] = None,
    top_inches: Optional[float] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Move a shape to a specific position."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    if left_inches is not None:
        shape.Left = inches_to_points(left_inches)
    if top_inches is not None:
        shape.Top = inches_to_points(top_inches)

    return {
        "success": True,
        "shape_name": shape.Name,
        "left_inches": round(points_to_inches(shape.Left), 2),
        "top_inches": round(points_to_inches(shape.Top), 2),
        "message": f"Shape '{shape.Name}' moved to ({round(points_to_inches(shape.Left), 2)}, {round(points_to_inches(shape.Top), 2)}) inches.",
    }


def resize_shape_sync(
    slide_index: int,
    shape_ref,
    width_inches: Optional[float] = None,
    height_inches: Optional[float] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Resize a shape."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    if width_inches is not None:
        shape.Width = inches_to_points(width_inches)
    if height_inches is not None:
        shape.Height = inches_to_points(height_inches)

    return {
        "success": True,
        "shape_name": shape.Name,
        "width_inches": round(points_to_inches(shape.Width), 2),
        "height_inches": round(points_to_inches(shape.Height), 2),
        "message": f"Shape '{shape.Name}' resized to ({round(points_to_inches(shape.Width), 2)} x {round(points_to_inches(shape.Height), 2)}) inches.",
    }


def rotate_shape_sync(
    slide_index: int,
    shape_ref,
    degrees: float,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Rotate a shape by degrees."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    shape.Rotation = degrees

    return {
        "success": True,
        "shape_name": shape.Name,
        "rotation": degrees,
        "message": f"Shape '{shape.Name}' rotated to {degrees} degrees.",
    }


def arrange_shape_sync(
    slide_index: int,
    shape_ref,
    action: str,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Change z-order of a shape.

    Args:
        action: "bring_to_front", "send_to_back", "bring_forward", "send_backward".
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    action_map = {
        "bring_to_front": 0,   # msoBringToFront
        "send_to_back": 1,     # msoSendToBack
        "bring_forward": 2,    # msoBringForward
        "send_backward": 3,    # msoSendBackward
    }

    val = action_map.get(action.lower())
    if val is None:
        raise ToolError(
            f"Invalid action '{action}'. Use: bring_to_front, send_to_back, bring_forward, send_backward."
        )

    shape.ZOrder(val)

    return {
        "success": True,
        "shape_name": shape.Name,
        "action": action,
        "message": f"Shape '{shape.Name}' z-order changed: {action}.",
    }


def align_shapes_sync(
    slide_index: int,
    shape_names: List[str],
    alignment: str,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Align multiple shapes.

    Args:
        shape_names: List of shape names to align.
        alignment: "left", "center", "right", "top", "middle", "bottom".
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    # Select shapes
    shape_indices = []
    for name in shape_names:
        shape = get_shape(slide, name)
        shape_indices.append(shape.Name)

    # Select all shapes
    slide_range = slide.Shapes.Range(shape_indices)

    align_map = {
        "left": 0,      # msoAlignLefts
        "center": 1,    # msoAlignCenters
        "right": 2,     # msoAlignRights
        "top": 3,       # msoAlignTops
        "middle": 4,    # msoAlignMiddles
        "bottom": 5,    # msoAlignBottoms
    }

    val = align_map.get(alignment.lower())
    if val is None:
        raise ToolError(
            f"Invalid alignment '{alignment}'. Use: left, center, right, top, middle, bottom."
        )

    slide_range.Align(val, False)  # False = relative to each other, not slide

    return {
        "success": True,
        "shapes": shape_names,
        "alignment": alignment,
        "message": f"Aligned {len(shape_names)} shapes: {alignment}.",
    }


def distribute_shapes_sync(
    slide_index: int,
    shape_names: List[str],
    direction: str,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Distribute shapes evenly.

    Args:
        shape_names: List of shape names.
        direction: "horizontal" or "vertical".
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    shape_indices = []
    for name in shape_names:
        shape = get_shape(slide, name)
        shape_indices.append(shape.Name)

    slide_range = slide.Shapes.Range(shape_indices)

    dist_map = {
        "horizontal": 0,  # msoDistributeHorizontally
        "vertical": 1,    # msoDistributeVertically
    }

    val = dist_map.get(direction.lower())
    if val is None:
        raise ToolError(f"Invalid direction '{direction}'. Use: horizontal, vertical.")

    slide_range.Distribute(val, False)

    return {
        "success": True,
        "shapes": shape_names,
        "direction": direction,
        "message": f"Distributed {len(shape_names)} shapes: {direction}.",
    }


def group_shapes_sync(
    slide_index: int,
    shape_names: List[str],
    ungroup: bool = False,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Group or ungroup shapes.

    Args:
        shape_names: List of shape names.
        ungroup: True to ungroup, False to group.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    if ungroup:
        shape = get_shape(slide, shape_names[0])
        if shape.Type != 6:  # msoGroup
            raise ToolError(f"Shape '{shape_names[0]}' is not a group.")
        shape.Ungroup()
        return {
            "success": True,
            "action": "ungrouped",
            "shape": shape_names[0],
            "message": f"Shape '{shape_names[0]}' ungrouped.",
        }
    else:
        shape_indices = []
        for name in shape_names:
            shape = get_shape(slide, name)
            shape_indices.append(shape.Name)

        slide_range = slide.Shapes.Range(shape_indices)
        group = slide_range.Group()

        return {
            "success": True,
            "action": "grouped",
            "shapes": shape_names,
            "group_name": group.Name,
            "message": f"Grouped {len(shape_names)} shapes into '{group.Name}'.",
        }
