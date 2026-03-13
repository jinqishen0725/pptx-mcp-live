"""Layout tools: move, resize, align, distribute, arrange, group, rotate, delete, add shapes."""

from typing import Any, Dict, List, Optional

from ..core.connection import (
    get_powerpoint, get_presentation, get_slide, get_shape,
    inches_to_points, points_to_inches
)
from ..core.errors import ToolError


def _unlock_placeholder(shape):
    """Unlock a placeholder shape so it can be freely moved/resized.

    Placeholders inherit position/size from the slide layout master.
    Setting LockAnchor and forcing the shape type allows overriding.
    """
    try:
        # ppPlaceholder = 14
        if shape.Type == 14:
            # Cut the link to the layout master by converting to a freeform
            tf = shape.TextFrame
            # Disable auto-size so manual sizing sticks
            try:
                tf.AutoSize = 0  # ppAutoSizeNone
            except Exception:
                pass
            try:
                tf.WordWrap = -1  # msoTrue
            except Exception:
                pass
    except Exception:
        pass


def move_shape_sync(
    slide_index: int,
    shape_ref,
    left_inches: Optional[float] = None,
    top_inches: Optional[float] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Move a shape to a specific position. Works on all shape types including placeholders."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    _unlock_placeholder(shape)

    if left_inches is not None:
        shape.Left = inches_to_points(left_inches)
    if top_inches is not None:
        shape.Top = inches_to_points(top_inches)

    # Verify the position was actually applied
    actual_left = round(points_to_inches(shape.Left), 2)
    actual_top = round(points_to_inches(shape.Top), 2)
    requested_left = round(left_inches, 2) if left_inches is not None else actual_left
    requested_top = round(top_inches, 2) if top_inches is not None else actual_top

    warnings = []
    if abs(actual_left - requested_left) > 0.05:
        warnings.append(f"Left position: requested {requested_left} but got {actual_left} (shape may be layout-locked)")
    if abs(actual_top - requested_top) > 0.05:
        warnings.append(f"Top position: requested {requested_top} but got {actual_top} (shape may be layout-locked)")

    result = {
        "success": True,
        "shape_name": shape.Name,
        "left_inches": actual_left,
        "top_inches": actual_top,
        "message": f"Shape '{shape.Name}' moved to ({actual_left}, {actual_top}) inches.",
    }
    if warnings:
        result["warnings"] = warnings
    return result


def resize_shape_sync(
    slide_index: int,
    shape_ref,
    width_inches: Optional[float] = None,
    height_inches: Optional[float] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Resize a shape. Works on all shape types including placeholders."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    _unlock_placeholder(shape)

    if width_inches is not None:
        shape.Width = inches_to_points(width_inches)
    if height_inches is not None:
        shape.Height = inches_to_points(height_inches)

    actual_w = round(points_to_inches(shape.Width), 2)
    actual_h = round(points_to_inches(shape.Height), 2)

    warnings = []
    if width_inches is not None and abs(actual_w - round(width_inches, 2)) > 0.05:
        warnings.append(f"Width: requested {round(width_inches, 2)} but got {actual_w}")
    if height_inches is not None and abs(actual_h - round(height_inches, 2)) > 0.05:
        warnings.append(f"Height: requested {round(height_inches, 2)} but got {actual_h}")

    result = {
        "success": True,
        "shape_name": shape.Name,
        "width_inches": actual_w,
        "height_inches": actual_h,
        "message": f"Shape '{shape.Name}' resized to ({actual_w} x {actual_h}) inches.",
    }
    if warnings:
        result["warnings"] = warnings
    return result


def delete_shape_sync(
    slide_index: int,
    shape_ref,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Delete a shape from a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)
    name = shape.Name

    shape.Delete()

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "deleted_shape": name,
        "remaining_shapes": slide.Shapes.Count,
        "message": f"Shape '{name}' deleted from slide {slide_index}.",
    }


def add_shape_sync(
    slide_index: int,
    shape_type: str,
    left_inches: float = 1.0,
    top_inches: float = 1.0,
    width_inches: float = 2.0,
    height_inches: float = 1.0,
    text: Optional[str] = None,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add an auto shape to a slide.

    Args:
        shape_type: Shape type name. Options: rectangle, rounded_rectangle,
            oval, triangle, diamond, pentagon, hexagon, right_arrow,
            left_arrow, up_arrow, down_arrow, star_5, star_4, heart,
            lightning_bolt, cross, plus, chevron, notched_right_arrow.
        text: Optional text to place inside the shape.
        fill_color: Fill color as #RRGGBB hex.
        line_color: Border color as #RRGGBB hex.
    """
    shape_map = {
        "rectangle": 1,            # msoShapeRectangle
        "rounded_rectangle": 5,    # msoShapeRoundedRectangle
        "oval": 9,                 # msoShapeOval
        "triangle": 7,             # msoShapeIsoscelesTriangle
        "diamond": 4,              # msoShapeDiamond
        "pentagon": 56,            # msoShapePentagon (regular)
        "hexagon": 10,             # msoShapeHexagon
        "right_arrow": 33,         # msoShapeRightArrow
        "left_arrow": 34,          # msoShapeLeftArrow
        "up_arrow": 35,            # msoShapeUpArrow
        "down_arrow": 36,          # msoShapeDownArrow
        "star_5": 92,              # msoShape5pointStar
        "star_4": 91,              # msoShape4pointStar
        "heart": 21,               # msoShapeHeart
        "lightning_bolt": 22,      # msoShapeLightningBolt
        "cross": 11,               # msoShapeCross
        "plus": 11,                # alias for cross
        "chevron": 52,             # msoShapeChevron
        "notched_right_arrow": 50, # msoShapeNotchedRightArrow
        "right_triangle": 6,       # msoShapeRightTriangle
        "parallelogram": 2,        # msoShapeParallelogram
        "trapezoid": 3,            # msoShapeTrapezoid
        "donut": 18,               # msoShapeDonut
        "no_symbol": 19,           # msoShapeNoSymbol
        "block_arc": 20,           # msoShapeBlockArc
        "left_right_arrow": 37,    # msoShapeLeftRightArrow
        "up_down_arrow": 38,       # msoShapeUpDownArrow
        "bent_arrow": 41,          # msoShapeBentArrow
        "u_turn_arrow": 42,        # msoShapeUTurnArrow
        "striped_right_arrow": 49, # msoShapeStripedRightArrow
        "curved_right_arrow": 54,  # not standard, approximate
        "circular_arrow": 60,      # msoShapeCircularArrow
    }

    type_id = shape_map.get(shape_type.lower())
    if type_id is None:
        raise ToolError(
            f"Unknown shape type '{shape_type}'. "
            f"Available: {', '.join(sorted(shape_map.keys()))}"
        )

    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    shape = slide.Shapes.AddShape(
        type_id,
        inches_to_points(left_inches),
        inches_to_points(top_inches),
        inches_to_points(width_inches),
        inches_to_points(height_inches),
    )

    if text is not None:
        shape.TextFrame.TextRange.Text = text

    if fill_color is not None:
        from .formatters import _hex_to_rgb_int
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = _hex_to_rgb_int(fill_color)

    if line_color is not None:
        from .formatters import _hex_to_rgb_int
        shape.Line.Visible = True
        shape.Line.ForeColor.RGB = _hex_to_rgb_int(line_color)

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "shape_type": shape_type,
        "position": {
            "left_inches": left_inches,
            "top_inches": top_inches,
            "width_inches": width_inches,
            "height_inches": height_inches,
        },
        "message": f"Shape '{shape.Name}' ({shape_type}) added to slide {slide_index}.",
    }


def set_slide_layout_sync(
    slide_index: int,
    layout_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Change the layout of an existing slide.

    Args:
        slide_index: 1-based slide number.
        layout_index: 1-based layout index. Use list_slide_layouts to see available layouts.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    try:
        layout = pres.SlideMaster.CustomLayouts(layout_index)
    except Exception:
        raise ToolError(
            f"Layout index {layout_index} not found. "
            f"Available: 1-{pres.SlideMaster.CustomLayouts.Count}. "
            f"Use list_slide_layouts to see names."
        )

    slide.CustomLayout = layout

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "layout_index": layout_index,
        "layout_name": layout.Name,
        "message": f"Slide {slide_index} layout changed to '{layout.Name}'.",
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
