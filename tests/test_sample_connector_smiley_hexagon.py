"""Regression test for orthogonal connector routing (smiley → hexagon)."""

from pathlib import Path

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.model.intermediate import ConnectorElement, ShapeElement


def _approx(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol


def test_sample_drawio_smiley_to_hexagon_respects_entry_side_left():
    """
    In sample.drawio, the orthogonal connector from the smiley to the hexagon declares entryX=0,entryY=0.5.

    Therefore, the segment adjacent to the hexagon must be horizontal (approaching from the left),
    even when draw.io doesn't store explicit waypoints for the connector.
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"

    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    elements = loader.extract_elements(diagrams[0])

    shapes = {e.id: e for e in elements if isinstance(e, ShapeElement)}
    connectors = [e for e in elements if isinstance(e, ConnectorElement)]

    edge_id = "fYnb-Lad83hC8_SQXFiI-27"  # smiley -> hexagon
    conn = next(c for c in connectors if c.id == edge_id)

    target = shapes[conn.target_id]
    assert "hexagon" in (target.shape_type or "").lower()
    assert conn.edge_style == "orthogonal"

    # Should have at least one bend for an orthogonal connector.
    assert len(conn.points) >= 3

    end_x, end_y = conn.points[-1]
    prev_x, prev_y = conn.points[-2]

    # Segment adjacent to the target must be horizontal (y stays the same).
    assert _approx(prev_y, end_y)

    # And it must approach the left edge from the left side (prev_x < end_x).
    assert prev_x < end_x


