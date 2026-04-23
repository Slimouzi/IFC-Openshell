#!/usr/bin/env python3
"""
Inspection rapide d'un fichier IFC — résumé du contenu, hiérarchie spatiale,
statistiques des classes et propriétés COBie disponibles.

Usage : python3 ifc_inspect.py <fichier.ifc>
"""

import sys
import argparse
from pathlib import Path
from collections import Counter

import ifcopenshell
import ifcopenshell.util.fm as fm
import ifcopenshell.util.element as element_util


def print_section(title: str):
    print(f"\n{'─' * 60}")
    print(f"  {title}")
    print("─" * 60)


def inspect(ifc_path: str) -> None:
    ifc = ifcopenshell.open(ifc_path)

    # ── En-tête ────────────────────────────────────────────────
    print_section("INFORMATIONS GÉNÉRALES")
    project = ifc.by_type("IfcProject")
    site = ifc.by_type("IfcSite")
    building = ifc.by_type("IfcBuilding")

    print(f"  Schéma IFC  : {ifc.schema}")
    print(f"  Projet      : {project[0].Name if project else 'n/a'}")
    print(f"  Site        : {site[0].Name if site else 'n/a'}")
    print(f"  Bâtiment    : {building[0].Name if building else 'n/a'}")
    print(f"  Phase       : {getattr(project[0], 'Phase', 'n/a') if project else 'n/a'}")

    # ── Hiérarchie spatiale ────────────────────────────────────
    print_section("HIÉRARCHIE SPATIALE")
    storeys = ifc.by_type("IfcBuildingStorey")
    for storey in sorted(storeys, key=lambda s: s.Elevation or 0):
        spaces = [
            rel.RelatedElements
            for rel in ifc.by_type("IfcRelContainedInSpatialStructure")
            if rel.RelatingStructure == storey
        ]
        space_list = [e for group in spaces for e in group if e.is_a("IfcSpace")]
        elev = f"{storey.Elevation:.2f}m" if storey.Elevation is not None else "?"
        print(f"  {storey.Name or 'Niveau ?':30} elev={elev:<10} {len(space_list)} espace(s)")

    # ── Statistiques des classes ───────────────────────────────
    print_section("STATISTIQUES DES CLASSES IFC")
    counter: Counter = Counter()
    for entity in ifc:
        counter[entity.is_a()] += 1

    # Top 20
    for cls, count in counter.most_common(20):
        bar = "█" * min(count // 10, 40)
        print(f"  {cls:<45} {count:>6}  {bar}")

    total = sum(counter.values())
    print(f"\n  Total : {total} entités, {len(counter)} classes distinctes")

    # ── Éléments COBie ─────────────────────────────────────────
    print_section("ÉLÉMENTS COBie")
    types = list(fm.get_cobie_types(ifc))
    components = list(fm.get_cobie_components(ifc))
    spaces = ifc.by_type("IfcSpace")
    storeys_list = ifc.by_type("IfcBuildingStorey")
    systems = ifc.by_type("IfcSystem")
    zones = ifc.by_type("IfcZone")

    print(f"  Types (IfcElementType)    : {len(types)}")
    print(f"  Composants (IfcElement)   : {len(components)}")
    print(f"  Espaces (IfcSpace)        : {len(spaces)}")
    print(f"  Niveaux (IfcStorey)       : {len(storeys_list)}")
    print(f"  Systèmes (IfcSystem)      : {len(systems)}")
    print(f"  Zones (IfcZone)           : {len(zones)}")

    # ── Property Sets disponibles ──────────────────────────────
    print_section("PROPERTY SETS LES PLUS FRÉQUENTS")
    pset_counter: Counter = Counter()
    for entity in list(types) + list(components):
        psets = element_util.get_psets(entity, should_inherit=False)
        for pset_name in psets:
            pset_counter[pset_name] += 1

    for pset_name, count in pset_counter.most_common(15):
        print(f"  {pset_name:<50} {count:>5} élément(s)")

    # ── Taux de remplissage COBie ──────────────────────────────
    print_section("TAUX DE REMPLISSAGE COBie (aperçu)")
    cobie_fields = {
        "COBie_Type": ["Category", "AssetType", "Manufacturer"],
        "COBie_Component": ["SerialNumber", "TagNumber", "AssetIdentifier"],
        "Pset_ManufacturerTypeInformation": ["Manufacturer", "ModelLabel"],
    }
    all_entities = list(types) + list(components)
    for pset_name, fields in cobie_fields.items():
        filled = {f: 0 for f in fields}
        for entity in all_entities:
            psets = element_util.get_psets(entity)
            pset_data = psets.get(pset_name, {})
            for field in fields:
                val = pset_data.get(field)
                if val and str(val) not in ("n/a", "None", ""):
                    filled[field] += 1
        if all_entities:
            print(f"\n  {pset_name}")
            for field, cnt in filled.items():
                pct = 100 * cnt // len(all_entities)
                bar = "█" * (pct // 5)
                print(f"    {field:<40} {pct:>3}%  {bar}")


def main():
    parser = argparse.ArgumentParser(description="Inspection rapide d'un fichier IFC")
    parser.add_argument("ifc", help="Chemin vers le fichier IFC")
    args = parser.parse_args()

    path = Path(args.ifc)
    if not path.exists():
        sys.exit(f"Fichier introuvable : {path}")

    print(f"\nInspection de : {path}")
    inspect(str(path))


if __name__ == "__main__":
    main()
