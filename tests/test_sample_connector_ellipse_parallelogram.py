"""Regression test for orthogonal connector endpoints (ellipse ↔ parallelogram)."""

from pathlib import Path

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.model.intermediate import ConnectorElement, ShapeElement


def _approx(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol


def test_sample_drawio_ellipse_side_connector_goes_down_from_ellipse():
    """
    In sample.drawio, the orthogonal connector between the parallelogram and the ellipse ('sample')
    should attach to the ellipse from the bottom side (i.e., the segment adjacent to the ellipse is vertical).
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"

    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    elements = loader.extract_elements(diagrams[0])

    shapes = {e.id: e for e in elements if isinstance(e, ShapeElement)}
    connectors = [e for e in elements if isinstance(e, ConnectorElement)]

    # This edge connects parallelogram (id ...-10) and ellipse (id ...-3) in sample.drawio.
    edge_id = "GStdcLXKth4fSFfuQepI-11"
    conn = next(c for c in connectors if c.id == edge_id)

    target = shapes[conn.target_id]
    assert "ellipse" in (target.shape_type or "").lower()

    # Endpoint should be on the bottom of the ellipse.
    end_x, end_y = conn.points[-1]
    assert _approx(end_y, target.y + target.h)

    # The segment adjacent to the ellipse should be vertical (x stays the same) and "downward" from the ellipse.
    prev_x, prev_y = conn.points[-2]
    assert _approx(prev_x, end_x)
    assert prev_y > end_y


