#!/usr/bin/env python3
"""
Module d'analyse avancée des clashes — groupement, enrichissement IFC, export BCF.
Utilisé par app.py pour construire une interface type Solibri Model Checker.
"""

from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Optional

import ifcopenshell
import ifcopenshell.util.element as element_util
import ifcopenshell.util.placement as placement_util

from clash_detection import _severity, _clash_type_label


# ── Classification par discipline ─────────────────────────────────────────────

DISCIPLINE_MAP: dict[str, str] = {}

_DISCIPLINES = {
    "Structure": [
        "IfcWall", "IfcWallStandardCase", "IfcSlab", "IfcSlabStandardCase",
        "IfcColumn", "IfcColumnStandardCase", "IfcBeam", "IfcBeamStandardCase",
        "IfcFooting", "IfcFoundation", "IfcPile", "IfcMember",
    ],
    "Architecture": [
        "IfcDoor", "IfcDoorStandardCase", "IfcWindow", "IfcWindowStandardCase",
        "IfcStair", "IfcStairFlight", "IfcRamp", "IfcRampFlight",
        "IfcRailing", "IfcCurtainWall", "IfcPlate", "IfcCovering",
    ],
    "CVC": [
        "IfcDuctSegment", "IfcDuctFitting", "IfcAirTerminal",
        "IfcAirToAirHeatRecovery", "IfcFan", "IfcUnitaryEquipment",
        "IfcHeatExchanger",
    ],
    "Plomberie": [
        "IfcPipeSegment", "IfcPipeFitting", "IfcSanitaryTerminal",
        "IfcValve", "IfcFlowMovingDevice", "IfcFlowStorageDevice",
        "IfcFlowTreatmentDevice",
    ],
    "Électricité": [
        "IfcCableCarrierSegment", "IfcCableCarrierFitting",
        "IfcCableSegment", "IfcCableFitting", "IfcElectricDistributionBoard",
        "IfcElectricAppliance", "IfcLamp", "IfcLightFixture",
        "IfcProtectiveDevice", "IfcSwitchingDevice",
    ],
}

for _disc, _classes in _DISCIPLINES.items():
    for _c in _classes:
        DISCIPLINE_MAP[_c] = _disc


def discipline_of(ifc_class: str) -> str:
    return DISCIPLINE_MAP.get(ifc_class, "Autre")


# ── Modèle de données ─────────────────────────────────────────────────────────

@dataclass
class ClashRecord:
    """Enrichi au-delà du format ifcclash natif."""
    id: str
    set_name: str
    mode: str
    clash_type: str            # Collision/Traversée/Protrusion/Dégagement
    severity: str              # Critique/Majeur/Mineur
    a_class: str
    a_name: str
    a_guid: str
    b_class: str
    b_name: str
    b_guid: str
    a_discipline: str
    b_discipline: str
    distance: float
    p1: list[float]
    p2: list[float]
    # Métadonnées utilisateur
    status: str = "open"       # open/wip/resolved/ignored
    comment: str = ""
    group_id: Optional[int] = None

    def center(self) -> list[float]:
        return [
            (self.p1[0] + self.p2[0]) / 2,
            (self.p1[1] + self.p2[1]) / 2,
            (self.p1[2] + self.p2[2]) / 2,
        ]

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "set_name": self.set_name,
            "mode": self.mode,
            "type": self.clash_type,
            "severity": self.severity,
            "a_class": self.a_class,
            "a_name": self.a_name,
            "a_guid": self.a_guid,
            "b_class": self.b_class,
            "b_name": self.b_name,
            "b_guid": self.b_guid,
            "a_discipline": self.a_discipline,
            "b_discipline": self.b_discipline,
            "distance": self.distance,
            "p1": self.p1,
            "p2": self.p2,
            "center": self.center(),
            "status": self.status,
            "comment": self.comment,
            "group_id": self.group_id,
        }


def build_records(clash_sets: list[dict]) -> list[ClashRecord]:
    """Convertit la sortie ifcclash native en liste de ClashRecord enrichis."""
    out: list[ClashRecord] = []
    for cs in clash_sets:
        set_name = cs.get("name", "?")
        mode = cs.get("mode", "?")
        for cid, c in (cs.get("clashes") or {}).items():
            a_class = c.get("a_ifc_class", "?")
            b_class = c.get("b_ifc_class", "?")
            p1 = list(c.get("p1") or [0, 0, 0])
            p2 = list(c.get("p2") or p1)
            out.append(ClashRecord(
                id=cid,
                set_name=set_name,
                mode=mode,
                clash_type=_clash_type_label(c),
                severity=_severity(c),
                a_class=a_class,
                a_name=c.get("a_name", "?") or "?",
                a_guid=c.get("a_global_id", ""),
                b_class=b_class,
                b_name=c.get("b_name", "?") or "?",
                b_guid=c.get("b_global_id", ""),
                a_discipline=discipline_of(a_class),
                b_discipline=discipline_of(b_class),
                distance=float(c.get("distance", 0) or 0),
                p1=p1,
                p2=p2,
            ))
    return out


# ── Smart grouping — cluster par proximité ────────────────────────────────────

def smart_group(records: list[ClashRecord], max_distance: float = 1.0) -> list[ClashRecord]:
    """
    Groupe les clashes proches spatialement (union-find par distance euclidienne
    entre centres). Complexité O(n²), OK jusqu'à quelques milliers de clashes.
    """
    n = len(records)
    if n == 0:
        return records

    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int):
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[ri] = rj

    centers = [r.center() for r in records]
    max_sq = max_distance ** 2

    for i in range(n):
        ci = centers[i]
        for j in range(i + 1, n):
            cj = centers[j]
            dx = ci[0] - cj[0]
            dy = ci[1] - cj[1]
            dz = ci[2] - cj[2]
            if dx*dx + dy*dy + dz*dz <= max_sq:
                union(i, j)

    # Numérotation contigue
    root_to_gid: dict[int, int] = {}
    for i in range(n):
        root = find(i)
        if root not in root_to_gid:
            root_to_gid[root] = len(root_to_gid) + 1
        records[i].group_id = root_to_gid[root]

    return records


# ── Inspecteur d'élément IFC ──────────────────────────────────────────────────

def inspect_element(ifc: ifcopenshell.file, guid: str) -> dict[str, Any]:
    """Retourne toutes les propriétés d'un élément IFC pour le panneau détail."""
    try:
        entity = ifc.by_guid(guid)
    except RuntimeError:
        return {"error": f"Élément introuvable : {guid}"}

    info: dict[str, Any] = {
        "class": entity.is_a(),
        "name": entity.Name or "",
        "description": entity.Description or "",
        "guid": guid,
        "tag": getattr(entity, "Tag", None),
    }

    # Type
    try:
        type_entity = element_util.get_type(entity)
        if type_entity:
            info["type_name"] = type_entity.Name
            info["type_class"] = type_entity.is_a()
    except Exception:
        pass

    # Hiérarchie spatiale (containment)
    try:
        container = element_util.get_container(entity)
        if container:
            info["container_class"] = container.is_a()
            info["container_name"] = container.Name or ""
    except Exception:
        pass

    # Matériaux
    try:
        materials = element_util.get_materials(entity)
        info["materials"] = [m.Name for m in materials if hasattr(m, "Name") and m.Name]
    except Exception:
        info["materials"] = []

    # Property sets
    try:
        psets = element_util.get_psets(entity)
        info["psets"] = {
            name: {k: str(v) for k, v in props.items() if k != "id"}
            for name, props in psets.items()
        }
    except Exception:
        info["psets"] = {}

    # Placement absolu
    try:
        if hasattr(entity, "ObjectPlacement") and entity.ObjectPlacement:
            matrix = placement_util.get_local_placement(entity.ObjectPlacement)
            info["location"] = [float(matrix[0][3]), float(matrix[1][3]), float(matrix[2][3])]
    except Exception:
        pass

    return info


# ── Matrices d'analyse ────────────────────────────────────────────────────────

def discipline_matrix(records: list[ClashRecord]) -> dict[tuple[str, str], int]:
    """Matrice discipline A × discipline B."""
    m: dict[tuple[str, str], int] = defaultdict(int)
    for r in records:
        key = tuple(sorted([r.a_discipline, r.b_discipline]))
        m[key] += 1
    return dict(m)


def class_matrix(records: list[ClashRecord], top_k: int = 15) -> dict[tuple[str, str], int]:
    """Matrice des paires de classes IFC les plus fréquentes."""
    m: dict[tuple[str, str], int] = defaultdict(int)
    for r in records:
        key = tuple(sorted([r.a_class, r.b_class]))
        m[key] += 1
    # Top K
    sorted_items = sorted(m.items(), key=lambda kv: kv[1], reverse=True)[:top_k]
    return dict(sorted_items)


def group_summary(records: list[ClashRecord]) -> list[dict]:
    """Résumé par groupe (cluster)."""
    groups: dict[int, list[ClashRecord]] = defaultdict(list)
    for r in records:
        if r.group_id is not None:
            groups[r.group_id].append(r)

    out = []
    for gid, recs in sorted(groups.items()):
        center = [0.0, 0.0, 0.0]
        for r in recs:
            c = r.center()
            center[0] += c[0]; center[1] += c[1]; center[2] += c[2]
        n = len(recs)
        center = [x / n for x in center]

        severities = [r.severity for r in recs]
        types = [r.clash_type for r in recs]
        out.append({
            "group_id": gid,
            "count": n,
            "critical": sum(1 for s in severities if s == "Critique"),
            "major": sum(1 for s in severities if s == "Majeur"),
            "minor": sum(1 for s in severities if s == "Mineur"),
            "types": dict(_freq(types)),
            "center": center,
            "class_pairs": list({tuple(sorted([r.a_class, r.b_class])) for r in recs})[:3],
        })
    return out


def _freq(items: list) -> dict:
    d: dict = defaultdict(int)
    for i in items:
        d[i] += 1
    return d


# ── Export BCF simplifié (BCF 2.1 XML) ────────────────────────────────────────

def export_bcf_zip(
    records: list[ClashRecord],
    output_path: str,
    project_name: str = "IFC Clash Report",
) -> str:
    """
    Génère un fichier BCF 2.1 minimal (sans screenshots) depuis les records.
    Plus léger que l'export ifcclash natif et indépendant du Clasher en mémoire.
    """
    import uuid
    import zipfile
    from datetime import datetime
    from xml.sax.saxutils import escape

    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    project_id = str(uuid.uuid4())

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # bcf.version
        zf.writestr(
            "bcf.version",
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<Version VersionId="2.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
            '  <DetailedVersion>2.1</DetailedVersion>\n</Version>\n',
        )
        # project.bcfp
        zf.writestr(
            "project.bcfp",
            f'<?xml version="1.0" encoding="UTF-8"?>\n'
            f'<ProjectExtension xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
            f'  <Project ProjectId="{project_id}"><Name>{escape(project_name)}</Name></Project>\n'
            f'  <ExtensionSchema></ExtensionSchema>\n</ProjectExtension>\n',
        )

        for r in records:
            topic_id = str(uuid.uuid4())
            title = f"[{r.severity}] {r.clash_type} — {r.a_class}/{r.a_name} vs {r.b_class}/{r.b_name}"

            markup_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Markup xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Topic Guid="{topic_id}" TopicType="Clash" TopicStatus="{r.status.capitalize()}">
    <Title>{escape(title)}</Title>
    <Priority>{r.severity}</Priority>
    <CreationDate>{now}</CreationDate>
    <CreationAuthor>ifcclash</CreationAuthor>
    <Description>{escape(r.comment or f'Distance: {r.distance:.4f} m · Set: {r.set_name}')}</Description>
  </Topic>
</Markup>
"""
            zf.writestr(f"{topic_id}/markup.bcf", markup_xml)

            # Viewpoint avec position
            center = r.center()
            viewpoint_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<VisualizationInfo xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" Guid="{topic_id}">
  <Components>
    <Selection>
      <Component IfcGuid="{r.a_guid}"/>
      <Component IfcGuid="{r.b_guid}"/>
    </Selection>
  </Components>
  <PerspectiveCamera>
    <CameraViewPoint><X>{center[0]+5}</X><Y>{center[1]+5}</Y><Z>{center[2]+3}</Z></CameraViewPoint>
    <CameraDirection><X>-1</X><Y>-1</Y><Z>-0.5</Z></CameraDirection>
    <CameraUpVector><X>0</X><Y>0</Y><Z>1</Z></CameraUpVector>
    <FieldOfView>60</FieldOfView>
  </PerspectiveCamera>
</VisualizationInfo>
"""
            zf.writestr(f"{topic_id}/viewpoint.bcfv", viewpoint_xml)

    return output_path
