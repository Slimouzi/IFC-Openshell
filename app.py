#!/usr/bin/env python3
"""
Interface Streamlit — Export COBie depuis un fichier IFC.
Lancement : streamlit run app.py
"""

import io
import os
import tempfile
from pathlib import Path

import streamlit as st

# ── Config page ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="COBie Export — IFC OpenShell",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ COBie Export — IFC OpenShell")
st.caption("Génère un fichier COBie 2.4 (Excel ou CSV) depuis un fichier IFC")

# ── Imports lazys (évite crash si pas installé) ────────────────────────────────
@st.cache_resource
def _check_deps():
    missing = []
    try:
        import ifcopenshell  # noqa: F401
    except ImportError:
        missing.append("ifcopenshell")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        missing.append("openpyxl")
    return missing

missing = _check_deps()
if missing:
    st.error(f"Dépendances manquantes : `pip install {' '.join(missing)}`")
    st.stop()

import ifcopenshell  # noqa: E402
import ifcopenshell.util.fm as fm  # noqa: E402
import ifcopenshell.util.element as element_util  # noqa: E402
import openpyxl  # noqa: E402

# Import des fonctions du module cobie_export
from cobie_export import (  # noqa: E402
    SHEETS,
    export_cobie,
)


# ── Sélection du fichier IFC ───────────────────────────────────────────────────
st.header("1. Fichier IFC")

data_dir = Path(__file__).parent / "data"
local_ifcs = sorted(data_dir.glob("*.ifc"))

col1, col2 = st.columns([1, 1])

with col1:
    uploaded = st.file_uploader(
        "Déposer un fichier IFC",
        type=["ifc"],
        help="Le fichier est traité localement, rien n'est envoyé sur un serveur externe.",
    )

with col2:
    local_choice = None
    if local_ifcs:
        options = ["— choisir —"] + [f.name for f in local_ifcs]
        choice = st.selectbox("Ou utiliser un fichier du dossier `data/`", options)
        if choice != "— choisir —":
            local_choice = data_dir / choice
    else:
        st.info("Aucun fichier IFC dans `data/`. Déposez-en un via l'uploader ou copiez-le dans `data/`.")


# Résolution du fichier source
ifc_path: Path | None = None
tmp_file = None

if uploaded is not None:
    tmp_file = tempfile.NamedTemporaryFile(suffix=".ifc", delete=False)
    tmp_file.write(uploaded.read())
    tmp_file.flush()
    ifc_path = Path(tmp_file.name)
    st.success(f"Fichier chargé : **{uploaded.name}** ({uploaded.size / 1024:.0f} Ko)")
elif local_choice is not None:
    ifc_path = local_choice
    st.success(f"Fichier sélectionné : **{local_choice.name}**")


# ── Chargement et aperçu ───────────────────────────────────────────────────────
if ifc_path:
    @st.cache_data(show_spinner="Chargement du fichier IFC...")
    def load_ifc(path: str):
        return ifcopenshell.open(path)

    ifc = load_ifc(str(ifc_path))

    # Métadonnées rapides
    project = ifc.by_type("IfcProject")
    building = ifc.by_type("IfcBuilding")
    storeys = ifc.by_type("IfcBuildingStorey")
    spaces = ifc.by_type("IfcSpace")
    types_cobie = list(fm.get_cobie_types(ifc))
    components_cobie = list(fm.get_cobie_components(ifc))

    st.header("2. Aperçu du modèle")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Schéma", ifc.schema)
    m2.metric("Niveaux", len(storeys))
    m3.metric("Espaces", len(spaces))
    m4.metric("Types COBie", len(types_cobie))
    m5.metric("Composants COBie", len(components_cobie))
    m6.metric("Projet", project[0].Name if project else "—")

    # ── Options d'export ───────────────────────────────────────────────────────
    st.header("3. Options d'export")

    col_fmt, col_sheets = st.columns([1, 2])

    with col_fmt:
        fmt = st.radio(
            "Format de sortie",
            ["Excel (.xlsx)", "CSV (une feuille)"],
            index=0,
        )

    with col_sheets:
        sheet_names = [name for name, _ in SHEETS]
        selected_sheets = st.multiselect(
            "Feuilles COBie à inclure",
            sheet_names,
            default=sheet_names,
        )

    if fmt == "CSV (une feuille)":
        csv_sheet = st.selectbox(
            "Feuille à exporter en CSV",
            selected_sheets if selected_sheets else sheet_names,
        )

    # ── Génération ─────────────────────────────────────────────────────────────
    st.header("4. Génération")

    if st.button("⚙️ Générer le COBie", type="primary", use_container_width=True):
        if not selected_sheets:
            st.warning("Sélectionnez au moins une feuille.")
        else:
            with st.spinner("Export en cours..."):
                # Génération Excel en mémoire
                from cobie_export import (
                    build_contact_sheet, build_facility_sheet, build_floor_sheet,
                    build_space_sheet, build_type_sheet, build_component_sheet,
                    build_system_sheet, build_zone_sheet, build_attribute_sheet,
                )

                builder_map = {
                    "Contact":   build_contact_sheet,
                    "Facility":  build_facility_sheet,
                    "Floor":     build_floor_sheet,
                    "Space":     build_space_sheet,
                    "Type":      build_type_sheet,
                    "Component": build_component_sheet,
                    "System":    build_system_sheet,
                    "Zone":      build_zone_sheet,
                    "Attribute": build_attribute_sheet,
                }

                wb = openpyxl.Workbook()
                wb.remove(wb.active)
                sheet_row_counts = {}

                for sheet_name in selected_sheets:
                    builder = builder_map.get(sheet_name)
                    if builder:
                        ws = wb.create_sheet(sheet_name)
                        builder(ws, ifc)
                        # Ajustement largeur
                        for col in ws.columns:
                            max_len = max(
                                (len(str(cell.value or "")) for cell in col), default=10
                            )
                            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
                        sheet_row_counts[sheet_name] = ws.max_row - 1  # hors header

                # Sauvegarde en mémoire
                xlsx_buffer = io.BytesIO()
                wb.save(xlsx_buffer)
                xlsx_buffer.seek(0)

                st.success("Export terminé !")

                # Résumé
                st.subheader("Résumé")
                cols = st.columns(min(len(sheet_row_counts), 5))
                for i, (sname, count) in enumerate(sheet_row_counts.items()):
                    cols[i % 5].metric(sname, f"{count} lignes")

                # ── Téléchargement ─────────────────────────────────────────────
                st.subheader("Téléchargement")

                stem = (uploaded.name if uploaded else ifc_path.name).replace(".ifc", "")

                if fmt == "Excel (.xlsx)":
                    st.download_button(
                        label="⬇️ Télécharger le fichier Excel COBie",
                        data=xlsx_buffer,
                        file_name=f"{stem}_cobie.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                else:  # CSV
                    # Extraction de la feuille choisie en CSV
                    import csv

                    ws_csv = wb[csv_sheet] if csv_sheet in wb.sheetnames else wb.active
                    csv_buffer = io.StringIO()
                    writer = csv.writer(csv_buffer, delimiter=";")
                    for row in ws_csv.iter_rows(values_only=True):
                        writer.writerow([v if v is not None else "" for v in row])
                    csv_bytes = csv_buffer.getvalue().encode("utf-8-sig")

                    st.download_button(
                        label=f"⬇️ Télécharger {csv_sheet}.csv",
                        data=csv_bytes,
                        file_name=f"{stem}_cobie_{csv_sheet}.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

                    # Aperçu tableau
                    st.subheader(f"Aperçu — {csv_sheet} (25 premières lignes)")
                    import pandas as pd
                    ws_prev = wb[csv_sheet]
                    rows = list(ws_prev.iter_rows(values_only=True))
                    if rows:
                        df = pd.DataFrame(rows[1:26], columns=rows[0])
                        st.dataframe(df, use_container_width=True)

    # ── Nettoyage fichier temp ─────────────────────────────────────────────────
    if tmp_file:
        try:
            os.unlink(tmp_file.name)
        except Exception:
            pass

else:
    st.info("👆 Chargez ou sélectionnez un fichier IFC pour commencer.")
