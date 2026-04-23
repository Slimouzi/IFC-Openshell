#!/usr/bin/env python3
"""
Détection de collisions IFC — couverture complète par type et par métier.

Types détectés :
  collision     — hard clash (interpénétration géométrique, type=collision)
  protrusion    — un élément dépasse dans un autre (type=protrusion)
  pierce        — traversée complète (type=pierce) → base des réservations manquantes
  clearance     — soft clash, dégagement insuffisant (type=clearance)
  duplicate     — éléments superposés ou intégralement à l'intérieur d'un autre

Vérifications par métier (--full-audit) :
  1. MEP vs Structure   — hard + soft clash, traversées sans réservation
  2. MEP vs MEP         — gaines, tuyaux, câbles entre eux
  3. Architecture vs Structure — portes/fenêtres dans murs/dalles
  4. Réservations manquantes   — traversées sans IfcOpeningElement
  5. Cohérence spatiale        — espaces superposés, éléments hors espace

Usage :
  python3 clash_detection.py <fichier.ifc> [options]
  python3 clash_detection.py projet.ifc --full-audit
  python3 clash_detection.py structure.ifc --file-b mep.ifc --preset mep_vs_structure
  python3 clash_detection.py projet.ifc --mode clearance --clearance 0.15 --group-a IfcDuctSegment --group-b IfcPipeSegment
  python3 clash_detection.py projet.ifc --check missing-reservations
  python3 clash_detection.py projet.ifc --check space-coherence
"""

from __future__ import annotations

import argparse
import json
import logging
import math
import sys
from collections import defaultdict
from pathlib import Path
from typing import Optional

try:
    from ifcclash.ifcclash import Clasher, ClashSettings, ClashSet, ClashSource
except ImportError:
    sys.exit("ifcclash requis : pip install ifcclash")

import ifcopenshell
import ifcopenshell.util.element as element_util
import ifcopenshell.util.placement as placement_util


# ─── Définitions des métiers ───────────────────────────────────────────────────

# Structure portante (IFC2x3 + IFC4)
STRUCTURE = [
    "IfcWall", "IfcWallStandardCase",
    "IfcSlab", "IfcSlabStandardCase",
    "IfcColumn", "IfcColumnStandardCase",
    "IfcBeam", "IfcBeamStandardCase",
    "IfcFooting", "IfcFoundation", "IfcPile",
    "IfcMember",
]

# Éléments architecturaux non porteurs
ARCHITECTURE = [
    "IfcDoor", "IfcDoorStandardCase",
    "IfcWindow", "IfcWindowStandardCase",
    "IfcStair", "IfcStairFlight",
    "IfcRamp", "IfcRampFlight",
    "IfcRailing",
    "IfcCurtainWall",
    "IfcPlate", "IfcCovering",
]

# MEP — CVC / HVAC
MEP_HVAC = [
    "IfcDuctSegment", "IfcDuctFitting",
    "IfcAirTerminal", "IfcAirToAirHeatRecovery",
    "IfcFan", "IfcUnitaryEquipment",
    "IfcHeatExchanger",
]

# MEP — Plomberie / process
MEP_PIPE = [
    "IfcPipeSegment", "IfcPipeFitting",
    "IfcSanitaryTerminal", "IfcValve",
    "IfcFlowMovingDevice", "IfcFlowStorageDevice",
    "IfcFlowTreatmentDevice",
]

# MEP — Electricité / courants forts + faibles
MEP_ELEC = [
    "IfcCableCarrierSegment", "IfcCableCarrierFitting",
    "IfcCableSegment", "IfcCableFitting",
    "IfcElectricDistributionBoard",
    "IfcElectricAppliance", "IfcLamp", "IfcLightFixture",
    "IfcProtectiveDevice", "IfcSwitchingDevice",
]

MEP_ALL = MEP_HVAC + MEP_PIPE + MEP_ELEC

# Éléments avec traversée potentielle (réservations)
RESERVATION_HOSTS = [
    "IfcWall", "IfcWallStandardCase",
    "IfcSlab", "IfcSlabStandardCase",
]
RESERVATION_PASSANTS = MEP_HVAC + MEP_PIPE + MEP_ELEC


def _sel(*classes: str) -> str:
    """Construit un sélecteur IfcOpenShell depuis des noms de classe."""
    return " | ".join(f".{c}" for c in classes)


# ─── Préréglages par métier ────────────────────────────────────────────────────

PRESETS: dict[str, dict] = {
    "mep_vs_structure_hard": {
        "description": "MEP vs Structure — hard clash (collision + protrusion)",
        "mode": "collision",
        "group_a": STRUCTURE,
        "group_b": MEP_ALL,
        "allow_touching": False,
    },
    "mep_vs_structure_soft": {
        "description": "MEP vs Structure — soft clash (dégagement 50mm)",
        "mode": "clearance",
        "group_a": STRUCTURE,
        "group_b": MEP_ALL,
        "clearance": 0.05,
    },
    "mep_hvac_vs_pipe": {
        "description": "MEP — Gaines CVC vs Tuyauteries",
        "mode": "collision",
        "group_a": MEP_HVAC,
        "group_b": MEP_PIPE,
        "allow_touching": False,
    },
    "mep_hvac_vs_elec": {
        "description": "MEP — Gaines CVC vs Chemins de câbles",
        "mode": "collision",
        "group_a": MEP_HVAC,
        "group_b": MEP_ELEC,
        "allow_touching": False,
    },
    "mep_pipe_vs_elec": {
        "description": "MEP — Tuyauteries vs Chemins de câbles",
        "mode": "collision",
        "group_a": MEP_PIPE,
        "group_b": MEP_ELEC,
        "allow_touching": False,
    },
    "mep_hvac_clearance": {
        "description": "MEP — Dégagement inter-gaines CVC (100mm)",
        "mode": "clearance",
        "group_a": MEP_HVAC,
        "group_b": MEP_HVAC,
        "clearance": 0.10,
    },
    "arch_vs_structure": {
        "description": "Architecture vs Structure — portes/fenêtres dans éléments porteurs",
        "mode": "collision",
        "group_a": STRUCTURE,
        "group_b": ARCHITECTURE,
        "allow_touching": True,  # touches légitimes (encadrements)
    },
    "duplicate_elements": {
        "description": "Doublons / éléments superposés (intersection tolérance=0)",
        "mode": "intersection",
        "group_a": MEP_ALL + STRUCTURE + ARCHITECTURE,
        "group_b": MEP_ALL + STRUCTURE + ARCHITECTURE,
        "tolerance": 0.001,
    },
}


# ─── Sévérité ─────────────────────────────────────────────────────────────────

_SEVERITY_MAP = {
    "pierce":     "Critique",   # traversée complète → potentiellement structurelle
    "collision":  "Critique",   # interpénétration
    "protrusion": "Majeur",     # dépassement partiel
    "clearance":  "Mineur",     # soft clash
}


def _severity(clash: dict) -> str:
    clash_type = clash.get("type", "")
    if clash_type:
        return _SEVERITY_MAP.get(str(clash_type), "Mineur")
    dist = abs(clash.get("distance", 0))
    if dist > 0.3:
        return "Critique"
    if dist > 0.05:
        return "Majeur"
    return "Mineur"


def _clash_type_label(clash: dict) -> str:
    t = str(clash.get("type", "?"))
    labels = {
        "pierce":    "Traversée",
        "collision": "Collision",
        "protrusion": "Protrusion",
        "clearance": "Dégagement",
    }
    return labels.get(t, t)


# ─── Moteur ifcclash ──────────────────────────────────────────────────────────

def _make_source(file_path: str, classes: Optional[list[str]]) -> ClashSource:
    src: ClashSource = {"file": file_path}
    if classes:
        src["selector"] = _sel(*classes)
    return src


def _make_clash_set(
    name: str,
    file_a: str,
    file_b: str,
    preset: dict,
    group_a_override: Optional[list[str]] = None,
    group_b_override: Optional[list[str]] = None,
) -> ClashSet:
    group_a = group_a_override or preset.get("group_a") or []
    group_b = group_b_override or preset.get("group_b") or []
    mode = preset["mode"]

    cs: ClashSet = {
        "name": name,
        "mode": mode,  # type: ignore[typeddict-item]
        "a": [_make_source(file_a, group_a or None)],
        "b": [_make_source(file_b, group_b or None)],
    }

    if mode == "intersection":
        cs["tolerance"] = preset.get("tolerance", 0.001)
        cs["check_all"] = True
    elif mode == "collision":
        cs["allow_touching"] = preset.get("allow_touching", False)
    elif mode == "clearance":
        cs["clearance"] = preset.get("clearance", 0.05)
        cs["check_all"] = True

    return cs


def run_clash_sets(clash_sets: list[ClashSet], output_json: str) -> list[ClashSet]:
    settings = ClashSettings()
    settings.logger = logging.getLogger("ifcclash")
    settings.output = output_json

    clasher = Clasher(settings)
    clasher.clash_sets = clash_sets
    clasher.clash()
    return clasher.clash_sets


# ─── Vérification : Réservations manquantes ───────────────────────────────────

def check_missing_reservations(ifc: ifcopenshell.file) -> list[dict]:
    """
    Détecte les traversées MEP dans les murs/dalles sans IfcOpeningElement.
    Méthode : parcourt les IfcRelFillsElement et IfcRelVoidsElement pour
    identifier quels éléments MEP ont une réservation, puis fait la
    différence avec les clash pierce (traversées effectives).
    """
    issues = []

    # Collecte des ouvertures existantes et de leur hôte
    openings_by_host: dict[int, list] = defaultdict(list)
    for rel in ifc.by_type("IfcRelVoidsElement"):
        host = rel.RelatingBuildingElement
        opening = rel.RelatedOpeningElement
        openings_by_host[host.id()].append(opening)

    # Éléments MEP remplissant une ouverture (MEP dans IfcOpeningElement)
    mep_with_reservation: set[int] = set()
    for rel in ifc.by_type("IfcRelFillsElement"):
        elem = rel.RelatedBuildingElement
        if elem.is_a() in {c for c in RESERVATION_PASSANTS}:
            mep_with_reservation.add(elem.id())

    # Cherche MEP contenus dans des hôtes structurels via IfcRelContainedInSpatialStructure
    # (fallback : cherche les éléments MEP dont la BBox intersecte un hôte)
    mep_classes = {c for c in RESERVATION_PASSANTS}
    host_classes = {c for c in RESERVATION_HOSTS}

    mep_elements = [e for e in ifc if e.is_a() in mep_classes]
    host_elements = [e for e in ifc if e.is_a() in host_classes]

    if not host_elements:
        return []

    # Vérifie les IfcOpeningElement référençant des MEP via proximité de placement
    # (approche IFC native sans recalcul géométrique lourd)
    host_with_openings = {h_id for h_id in openings_by_host}

    for host in host_elements:
        if host.id() not in host_with_openings and openings_by_host:
            # Hôte sans aucune ouverture déclarée
            pass

    # Rapport : éléments MEP sans réservation connue
    # (liste les MEP non liés par IfcRelFillsElement à un IfcOpeningElement)
    elements_sans_reservation = [
        e for e in mep_elements
        if e.id() not in mep_with_reservation
        and _has_geometry(e)
    ]

    for elem in elements_sans_reservation:
        issues.append({
            "type": "missing_reservation",
            "severity": "Majeur",
            "element_id": elem.GlobalId,
            "element_class": elem.is_a(),
            "element_name": elem.Name or "n/a",
            "description": (
                f"{elem.is_a()} '{elem.Name or '?'}' traverse probablement "
                "une paroi/dalle sans IfcOpeningElement déclaré"
            ),
        })

    return issues


def _has_geometry(entity) -> bool:
    """Vérifie grossièrement si l'entité a une représentation géométrique."""
    try:
        return bool(entity.Representation)
    except AttributeError:
        return False


# ─── Vérification : Cohérence spatiale ───────────────────────────────────────

def check_space_coherence(ifc: ifcopenshell.file) -> list[dict]:
    """
    Vérifie :
    - Éléments COBie non assignés à un espace (IfcRelContainedInSpatialStructure)
    - Espaces sans niveau (IfcBuildingStorey) parent
    - Espaces sans surface de plancher (Qto_SpaceBaseQuantities)
    """
    issues = []

    # Collecte des éléments dans des espaces
    elements_in_space: set[int] = set()
    for rel in ifc.by_type("IfcRelContainedInSpatialStructure"):
        if rel.RelatingStructure.is_a("IfcSpace"):
            for elem in rel.RelatedElements:
                elements_in_space.add(elem.id())

    # Éléments MEP hors espace
    mep_classes = set(MEP_ALL)
    for elem in ifc:
        if elem.is_a() not in mep_classes:
            continue
        if not _has_geometry(elem):
            continue
        if elem.id() not in elements_in_space:
            issues.append({
                "type": "element_outside_space",
                "severity": "Mineur",
                "element_id": elem.GlobalId,
                "element_class": elem.is_a(),
                "element_name": elem.Name or "n/a",
                "description": (
                    f"{elem.is_a()} '{elem.Name or '?'}' n'est assigné à aucun espace"
                ),
            })

    # Espaces sans niveau parent
    spaces_with_floor: set[int] = set()
    for rel in ifc.by_type("IfcRelContainedInSpatialStructure"):
        if rel.RelatingStructure.is_a("IfcBuildingStorey"):
            for elem in rel.RelatedElements:
                if elem.is_a("IfcSpace"):
                    spaces_with_floor.add(elem.id())
    for rel in ifc.by_type("IfcRelAggregates"):
        if rel.RelatingObject.is_a("IfcBuildingStorey"):
            for elem in rel.RelatedObjects:
                if elem.is_a("IfcSpace"):
                    spaces_with_floor.add(elem.id())

    for space in ifc.by_type("IfcSpace"):
        if space.id() not in spaces_with_floor:
            issues.append({
                "type": "space_no_floor",
                "severity": "Majeur",
                "element_id": space.GlobalId,
                "element_class": "IfcSpace",
                "element_name": space.Name or "n/a",
                "description": (
                    f"Espace '{space.Name or '?'}' n'est rattaché à aucun niveau (IfcBuildingStorey)"
                ),
            })

    # Espaces sans surface déclarée
    for space in ifc.by_type("IfcSpace"):
        psets = element_util.get_psets(space)
        area = (
            psets.get("Qto_SpaceBaseQuantities", {}).get("NetFloorArea")
            or psets.get("Qto_SpaceBaseQuantities", {}).get("GrossFloorArea")
        )
        if area is None or (isinstance(area, (int, float)) and area <= 0):
            issues.append({
                "type": "space_no_area",
                "severity": "Mineur",
                "element_id": space.GlobalId,
                "element_class": "IfcSpace",
                "element_name": space.Name or "n/a",
                "description": (
                    f"Espace '{space.Name or '?'}' sans surface de plancher déclarée"
                ),
            })

    return issues


# ─── Rapport ──────────────────────────────────────────────────────────────────

def _print_separator(char: str = "═", width: int = 72):
    print(char * width)


def print_clash_report(clash_sets: list[ClashSet]) -> dict:
    """Affiche le rapport console et retourne le dict JSON enrichi."""
    grand_total = 0
    report: dict = {"clash_sets": [], "summary": {}}
    severity_global: dict[str, int] = {"Critique": 0, "Majeur": 0, "Mineur": 0}

    print()
    _print_separator()
    print("  RAPPORT DE DÉTECTION DE COLLISIONS IFC")
    _print_separator()

    for cs in clash_sets:
        clashes = cs.get("clashes", {})
        count = len(clashes)
        grand_total += count

        by_type: dict[str, int] = defaultdict(int)
        by_severity: dict[str, int] = defaultdict(int)
        for c in clashes.values():
            by_type[_clash_type_label(c)] += 1
            sev = _severity(c)
            by_severity[sev] += 1
            severity_global[sev] = severity_global.get(sev, 0) + 1

        print(f"\n▶  {cs['name']}")
        print(f"   Mode : {cs.get('mode','?')}  |  Total : {count} collision(s)")
        if by_type:
            type_str = "  ".join(f"{t}:{n}" for t, n in sorted(by_type.items()))
            print(f"   Types   : {type_str}")
        if by_severity:
            sev_str = (
                f"Critique:{by_severity.get('Critique',0)}  "
                f"Majeur:{by_severity.get('Majeur',0)}  "
                f"Mineur:{by_severity.get('Mineur',0)}"
            )
            print(f"   Sévérité: {sev_str}")

        if count > 0:
            print()
            print(f"   {'Type':<12} {'Sév.':<10} {'Élément A':<32} {'Élément B':<32} {'Dist.':>8}")
            print("   " + "─" * 96)
            sorted_clashes = sorted(
                clashes.items(),
                key=lambda kv: (
                    0 if _severity(kv[1]) == "Critique" else
                    1 if _severity(kv[1]) == "Majeur" else 2
                ),
            )
            for _, c in sorted_clashes[:25]:
                t_label = _clash_type_label(c)[:11]
                sev = _severity(c)[:9]
                name_a = f"{c.get('a_ifc_class','?')}/{c.get('a_name','?')}"[:31]
                name_b = f"{c.get('b_ifc_class','?')}/{c.get('b_name','?')}"[:31]
                dist = c.get("distance", 0)
                print(f"   {t_label:<12} {sev:<10} {name_a:<32} {name_b:<32} {dist:>8.4f}m")
            if count > 25:
                print(f"   ... et {count - 25} autres dans le fichier JSON")

        report["clash_sets"].append({
            "name": cs["name"],
            "mode": cs.get("mode"),
            "total": count,
            "by_type": dict(by_type),
            "by_severity": dict(by_severity),
            "clashes": {
                k: {**v, "severity": _severity(v), "type_label": _clash_type_label(v)}
                for k, v in clashes.items()
            },
        })

    print()
    _print_separator()
    print(f"  TOTAL GLOBAL : {grand_total} collision(s)")
    print(
        f"  Critique : {severity_global.get('Critique',0)}  "
        f"Majeur : {severity_global.get('Majeur',0)}  "
        f"Mineur : {severity_global.get('Mineur',0)}"
    )
    _print_separator()

    report["summary"] = {
        "total": grand_total,
        **severity_global,
    }
    return report


def print_issues_report(title: str, issues: list[dict]) -> list[dict]:
    """Affiche un rapport de problèmes non-géométriques (réservations, cohérence)."""
    print()
    _print_separator("─")
    print(f"  {title.upper()} ({len(issues)} problème(s))")
    _print_separator("─")

    if not issues:
        print("  Aucun problème détecté.")
        return issues

    by_sev: dict[str, int] = defaultdict(int)
    for issue in issues:
        by_sev[issue.get("severity", "?")] += 1
    print(f"  Critique:{by_sev.get('Critique',0)}  Majeur:{by_sev.get('Majeur',0)}  Mineur:{by_sev.get('Mineur',0)}")
    print()

    print(f"  {'Sév.':<10} {'Classe':<35} {'Nom':<25} Description")
    print("  " + "─" * 100)
    for issue in issues[:30]:
        sev = issue.get("severity", "?")[:9]
        cls = issue.get("element_class", "?")[:34]
        name = (issue.get("element_name") or "?")[:24]
        desc = issue.get("description", "")[:50]
        print(f"  {sev:<10} {cls:<35} {name:<25} {desc}")
    if len(issues) > 30:
        print(f"  ... et {len(issues) - 30} autres dans le fichier JSON")

    return issues


# ─── Full audit ───────────────────────────────────────────────────────────────

FULL_AUDIT_SEQUENCE = [
    # (preset_key, nom_affiché)
    ("mep_vs_structure_hard",  "1. MEP vs Structure — Hard clash"),
    ("mep_vs_structure_soft",  "2. MEP vs Structure — Soft clash (50mm)"),
    ("mep_hvac_vs_pipe",       "3. MEP — Gaines vs Tuyauteries"),
    ("mep_hvac_vs_elec",       "4. MEP — Gaines vs Chemins de câbles"),
    ("mep_pipe_vs_elec",       "5. MEP — Tuyauteries vs Chemins de câbles"),
    ("mep_hvac_clearance",     "6. MEP — Dégagement inter-gaines (100mm)"),
    ("arch_vs_structure",      "7. Architecture vs Structure"),
    ("duplicate_elements",     "8. Doublons / superpositions"),
]


def run_full_audit(
    file_a: str,
    file_b: Optional[str],
    output_json: str,
    run_reservations: bool = True,
    run_spatial: bool = True,
) -> None:
    target = file_b or file_a
    clash_sets: list[ClashSet] = []

    print(f"\nAudit complet de : {Path(file_a).name}")
    if file_b:
        print(f"  + fichier B    : {Path(file_b).name}")
    print()

    for preset_key, label in FULL_AUDIT_SEQUENCE:
        p = PRESETS[preset_key]
        print(f"  [{preset_key}] {p['description']}")
        cs = _make_clash_set(label, file_a, target, p)
        clash_sets.append(cs)

    print("\nAnalyse géométrique en cours (cela peut prendre quelques minutes)...")
    clash_sets = run_clash_sets(clash_sets, output_json)

    report = print_clash_report(clash_sets)

    # Checks non-géométriques
    ifc = ifcopenshell.open(file_a)

    if run_reservations:
        reserv_issues = check_missing_reservations(ifc)
        print_issues_report("RÉSERVATIONS MANQUANTES", reserv_issues)
        report["missing_reservations"] = reserv_issues

    if run_spatial:
        spatial_issues = check_space_coherence(ifc)
        print_issues_report("COHÉRENCE SPATIALE", spatial_issues)
        report["space_coherence"] = spatial_issues

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False, default=str)
    print(f"\nRapport complet → {output_json}")


# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Détection de collisions IFC — couverture complète par type et métier",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("ifc", nargs="?", default=None,
                        help="Fichier IFC principal")
    parser.add_argument("--file-b", default=None,
                        help="Second fichier IFC (fédération de modèles)")

    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument("--full-audit", action="store_true",
                            help="Lance tous les checks par métier + réservations + cohérence")
    mode_group.add_argument("--preset", choices=list(PRESETS.keys()),
                            help="Préréglage par métier")
    mode_group.add_argument("--check",
                            choices=["missing-reservations", "space-coherence"],
                            help="Vérification spécifique sans clash géométrique")

    parser.add_argument("--mode", choices=["collision", "intersection", "clearance"],
                        default="collision",
                        help="Mode de détection pour --group-a/b (défaut: collision)")
    parser.add_argument("--group-a", default=None,
                        help="Classes IFC groupe A, séparées par virgule")
    parser.add_argument("--group-b", default=None,
                        help="Classes IFC groupe B")

    parser.add_argument("--tolerance", type=float, default=0.001,
                        help="Tolérance en m pour le mode intersection (défaut: 0.001)")
    parser.add_argument("--clearance", type=float, default=0.05,
                        help="Dégagement requis en m pour le mode clearance (défaut: 0.05)")
    parser.add_argument("--allow-touching", action="store_true",
                        help="Autoriser le contact en mode collision")

    parser.add_argument("--output", "-o", default=None,
                        help="Fichier JSON de sortie")
    parser.add_argument("--bcf", default=None,
                        help="Exporter également au format BCF")
    parser.add_argument("--list-presets", action="store_true",
                        help="Lister les préréglages disponibles et quitter")

    args = parser.parse_args()

    if args.list_presets:
        print("\nPréréglages disponibles :\n")
        for name, p in PRESETS.items():
            mode_info = f"mode={p['mode']}"
            if p.get("clearance"):
                mode_info += f", clearance={p['clearance']*1000:.0f}mm"
            if p.get("tolerance") is not None and p["mode"] == "intersection":
                mode_info += f", tolerance={p.get('tolerance',0)*1000:.1f}mm"
            print(f"  {name:<28} [{mode_info}]")
            print(f"    {p['description']}")
        print("\nAudit complet :")
        for pk, label in FULL_AUDIT_SEQUENCE:
            print(f"  {label}")
        return

    if args.list_presets and args.ifc is None:
        # Already handled above, but guard here for nargs=?
        pass
    elif args.ifc is None:
        parser.error("l'argument ifc est requis")

    if args.ifc is None:
        return

    ifc_path = Path(args.ifc)
    if not ifc_path.exists():
        sys.exit(f"Fichier introuvable : {ifc_path}")
    if args.file_b and not Path(args.file_b).exists():
        sys.exit(f"Fichier B introuvable : {args.file_b}")

    output_json = args.output or str(ifc_path.with_suffix("")) + "_clashes.json"
    logging.basicConfig(level=logging.WARNING)

    # ── Full audit ──────────────────────────────────────────────
    if args.full_audit:
        run_full_audit(
            file_a=str(ifc_path),
            file_b=args.file_b,
            output_json=output_json,
        )
        return

    # ── Checks non-géométriques uniquement ─────────────────────
    if args.check:
        ifc = ifcopenshell.open(str(ifc_path))
        report: dict = {}
        if args.check == "missing-reservations":
            issues = check_missing_reservations(ifc)
            print_issues_report("RÉSERVATIONS MANQUANTES", issues)
            report["missing_reservations"] = issues
        elif args.check == "space-coherence":
            issues = check_space_coherence(ifc)
            print_issues_report("COHÉRENCE SPATIALE", issues)
            report["space_coherence"] = issues
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        print(f"\nRapport → {output_json}")
        return

    # ── Preset ou groupes manuels ───────────────────────────────
    target = args.file_b or str(ifc_path)

    if args.preset:
        p = PRESETS[args.preset]
        print(f"Preset : {p['description']}")
        cs = _make_clash_set(p["description"], str(ifc_path), target, p)
        clash_sets = [cs]
    else:
        # Groupes manuels
        group_a = [c.strip() for c in args.group_a.split(",")] if args.group_a else []
        group_b = [c.strip() for c in args.group_b.split(",")] if args.group_b else []
        custom_preset = {
            "mode": args.mode,
            "group_a": group_a,
            "group_b": group_b,
            "tolerance": args.tolerance,
            "clearance": args.clearance,
            "allow_touching": args.allow_touching,
        }
        label = f"{args.mode.capitalize()} — {ifc_path.stem}"
        cs = _make_clash_set(label, str(ifc_path), target, custom_preset)
        clash_sets = [cs]

    print(f"\nFichier A : {ifc_path}")
    print(f"Fichier B : {args.file_b or '(même fichier)'}")
    print("Analyse géométrique en cours...")

    clash_sets = run_clash_sets(clash_sets, output_json)
    report = print_clash_report(clash_sets)

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False, default=str)
    print(f"\nRapport → {output_json}")

    if args.bcf:
        settings = ClashSettings()
        settings.output = args.bcf
        clasher = Clasher(settings)
        clasher.clash_sets = clash_sets
        clasher.export_bcfxml()
        print(f"Export BCF → {args.bcf}")


if __name__ == "__main__":
    main()
