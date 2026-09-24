"""
core/template_registry.py
Central registry mapping template IDs to builder classes.
Add new templates here without touching any other file.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

import config
from utils.logging_utils import get_logger

logger = get_logger(__name__)


# ---------------------------------------------------------------------------
# Registry definition
# ---------------------------------------------------------------------------

_REGISTRY: Dict[str, Dict[str, Any]] = {
    "dark_navy": {
        "id":          "dark_navy",
        "name":        "Template 1",
        "description": "Dark Navy",
        "icon":        "🌑",
    },
    # "techm": {
    #     "id":          "techm",
    #     "name":        "TechM Corporate",
    #     "description": "Tech Mahindra corporate branding template",
    #     "icon":        "🏢",
    # },
    "white_blue": {
        "id":          "white_blue",
        "name":        "Template 2",
        "description": "White & Blue",
        "icon":        "☁️",
    },
    "techm_v3": {
        "id":          "techm_v3",
        "name":        "Template 3",
        "description": "TechM",
        "icon":        "✨",
    },
    "template1": {
        "id":          "template1",
        "name":        "Template 4",
        "description": "Professional",
        "icon":        "🟤",
    },
    "hld_qbr": {
        "id":          "hld_qbr",
        "name":        "HLD QBR",
        "description": "UPS Healthcare",
        "icon":        "🏥",
    },
}


def list_templates() -> List[Dict[str, str]]:
    """Return list of template metadata dicts (id, name, description, icon)."""
    return list(_REGISTRY.values())


def get_builder(template_id: str, layout_manager=None):
    """
    Instantiate and return the appropriate builder for the given template_id.

    Args:
        template_id:    One of 'dark_navy', 'techm', 'white_blue', 'techm_v3', 'template1', 'hld_qbr'
        layout_manager: Required only for dark_navy (existing LayoutManager instance)

    Returns:
        Builder instance with a .build(plan) -> bytes method
    """
    tid = template_id.lower().strip()

    if tid == "template1":
        from core.builders.template1_builder import Template1Builder
        return Template1Builder(template_path=config.TEMPLATE1_FILE)

    if tid == "hld_qbr":
        from core.builders.hld_qbr_builder import HLDQBRBuilder
        return HLDQBRBuilder(template_path=config.HLD_QBR_TEMPLATE_FILE)

    if tid == "techm_v3":
        from core.builders.techm_v3_builder import TechMV3Builder
        return TechMV3Builder(template_path=config.TECHM_V3_TEMPLATE_FILE)

    if tid == "techm":
        from core.builders.techm_builder import TechMBuilder
        return TechMBuilder(template_path=config.TECHM_TEMPLATE_FILE)

    if tid == "white_blue":
        from core.builders.white_blue_builder import WhiteBlueBuilder
        return WhiteBlueBuilder()

    if tid == "dark_navy" or tid not in _REGISTRY:
        if tid not in _REGISTRY:
            logger.warning("Unknown template_id '%s', falling back to dark_navy", template_id)
        from core.builders.dark_navy_builder import DarkNavyBuilder
        if layout_manager is None:
            from core.layout_manager import LayoutManager
            layout_manager = LayoutManager(
                template_path=config.TEMPLATE_FILE,
                metadata_path=config.LAYOUT_METADATA_FILE,
            )
        return DarkNavyBuilder(layout_manager)

    # Should never reach here due to fallback above
    from core.builders.dark_navy_builder import DarkNavyBuilder
    return DarkNavyBuilder(layout_manager)
