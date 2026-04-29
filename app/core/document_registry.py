from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ModuleRegistryItem:
    key: str
    sheet_name: str
    document_type: str
    concept: str
    template_keys: tuple[str, ...]


MODULES: dict[str, ModuleRegistryItem] = {
    "RD_APR": ModuleRegistryItem("RD_APR", "RD_APR", "RD", "APR", ("rd_apr_cero", "rd_apr_negativo", "rd_apr_positivo")),
    "RD_DGE": ModuleRegistryItem("RD_DGE", "RD_DGE", "RD", "DGE", ("rd_dge_cero", "rd_dge_negativo", "rd_dge_positivo")),
    "RM_APR": ModuleRegistryItem("RM_APR", "RM_APR", "RM", "APR", ("rm_apr_176_1", "rm_apr_178_1")),
    "RM_DGE": ModuleRegistryItem("RM_DGE", "RM_DGE", "RM", "DGE", ("rm_dge_176_1", "rm_dge_178_1")),
}


TEMPLATE_LABELS: dict[str, str] = {
    "rd_apr_cero": "RD APR cero",
    "rd_apr_negativo": "RD APR negativo",
    "rd_apr_positivo": "RD APR positivo",
    "rd_dge_cero": "RD DGE cero",
    "rd_dge_negativo": "RD DGE negativo",
    "rd_dge_positivo": "RD DGE positivo",
    "rm_apr_176_1": "RM APR 176-1",
    "rm_apr_178_1": "RM APR 178-1",
    "rm_dge_176_1": "RM DGE 176-1",
    "rm_dge_178_1": "RM DGE 178-1",
}
