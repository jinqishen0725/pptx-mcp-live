"""Export tools: export slide image, capture for feedback, export PDF."""

import base64
import os
import tempfile
from typing import Any, Dict, Optional

from ..core.connection import get_powerpoint, get_presentation, get_slide
from ..core.errors import ToolError


def export_slide_image_sync(
    slide_index: int,
    output_path: Optional[str] = None,
    width: int = 1920,
    height: int = 1080,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Export a single slide as a PNG image file.

    Args:
        slide_index: 1-based slide index.
        output_path: Path to save the image. None = temp file.
        width: Image width in pixels.
        height: Image height in pixels.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    if output_path is None:
        output_path = os.path.join(
            tempfile.gettempdir(),
            f"slide_{slide_index}_{pres.Name.replace('.pptx', '')}.png"
        )

    slide.Export(output_path, "PNG", width, height)

    if not os.path.exists(output_path):
        raise ToolError(f"Failed to export slide {slide_index}.")

    file_size = os.path.getsize(output_path)

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "output_path": output_path,
        "width": width,
        "height": height,
        "file_size_bytes": file_size,
        "message": f"Slide {slide_index} exported to {output_path}",
    }


def capture_slide_sync(
    slide_index: int,
    width: int = 1280,
    height: int = 720,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Capture a slide as base64 PNG for agent visual feedback.

    Returns the image inline so the agent can see and evaluate the slide.

    Args:
        slide_index: 1-based slide index.
        width: Image width in pixels.
        height: Image height in pixels.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    # Export to temp file
    temp_path = os.path.join(
        tempfile.gettempdir(),
        f"pptx_mcp_capture_{slide_index}.png"
    )

    slide.Export(temp_path, "PNG", width, height)

    if not os.path.exists(temp_path):
        raise ToolError(f"Failed to capture slide {slide_index}.")

    # Read and encode as base64
    with open(temp_path, "rb") as f:
        image_data = base64.b64encode(f.read()).decode("utf-8")

    file_size = os.path.getsize(temp_path)

    # Clean up temp file
    try:
        os.remove(temp_path)
    except Exception:
        pass

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "width": width,
        "height": height,
        "image_base64": image_data,
        "file_size_bytes": file_size,
        "message": f"Slide {slide_index} captured ({width}x{height}).",
    }


def export_pdf_sync(
    output_path: str,
    slide_index: Optional[int] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Export presentation to PDF.

    Args:
        output_path: Path to save the PDF.
        slide_index: Export only this slide. None = all slides.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)

    # ppFixedFormatTypePDF = 2
    try:
        pres.ExportAsFixedFormat(
            output_path,
            2,  # ppFixedFormatTypePDF
            Intent=1,  # ppFixedFormatIntentScreen
        )
    except Exception as e:
        # Fallback: try SaveAs with PDF filter
        try:
            pres.SaveAs(output_path, 32)  # ppSaveAsPDF = 32
        except Exception as e2:
            raise ToolError(f"Failed to export PDF: {e2}") from e2

    if not os.path.exists(output_path):
        raise ToolError("Failed to export PDF.")

    file_size = os.path.getsize(output_path)

    return {
        "success": True,
        "presentation": pres.Name,
        "output_path": output_path,
        "slide_index": slide_index,
        "file_size_bytes": file_size,
        "message": f"Exported to {output_path}",
    }
