from __future__ import annotations

import hashlib
import re
from pathlib import Path

import pandas as pd
import streamlit as st

from converter import (
    SCHEDULE_COLUMNS,
    SUMMARY_COLUMNS,
    convert_acd_loan,
    parse_acd_loan_file,
    parse_acd_loan_paste,
    to_clipboard_tsv,
    to_csv_bytes,
    to_xlsx_bytes,
    validate_pennylane_loan,
)


st.set_page_config(
    page_title="Échéanciers ACD → Pennylane",
    page_icon="🏦",
    layout="wide",
)

st.markdown(
    """
    <style>
      .block-container {padding-top: 2rem; padding-bottom: 3rem;}
      [data-testid="stMetric"] {background: #f5f8fb; border: 1px solid #dce5ed; padding: .8rem; border-radius: .6rem;}
      div[data-testid="stAlert"] {border-radius: .6rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Convertisseur d'échéanciers d'emprunt ACD vers Pennylane")
st.caption(
    "Transforme un échéancier ACD dans la structure du classeur Pennylane : paramètres en haut et échéancier détaillé à partir de la ligne 5."
)

conversion_tab, clipboard_tab, help_tab = st.tabs(
    ["1 — Conversion", "2 — Copier-coller", "3 — Mode d'emploi"]
)

edited_summary: pd.DataFrame | None = None
edited_schedule: pd.DataFrame | None = None
blocking_errors = True
output_stem = "echeancier_emprunt_pennylane"

with conversion_tab:
    input_method = st.radio(
        "Source ACD",
        ["Importer un CSV ou XLSX", "Copier-coller le tableau"],
        horizontal=True,
    )

    parsed = None
    raw_key = b""
    if input_method == "Importer un CSV ou XLSX":
        uploaded = st.file_uploader(
            "Déposez l'échéancier exporté depuis ACD",
            type=["csv", "txt", "tsv", "asc", "xlsx", "xlsm"],
            help="Le fichier doit contenir les colonnes Date, Annuité, Intérêts, Capital remb. et Restant Dû.",
        )
        if uploaded is not None:
            raw_key = uploaded.getvalue()
            output_stem = re.sub(r"[^A-Za-z0-9_-]+", "_", Path(uploaded.name).stem).strip("_") or output_stem
            try:
                parsed = parse_acd_loan_file(uploaded.name, raw_key)
            except Exception as exc:
                st.error(f"Lecture impossible : {exc}")
    else:
        pasted = st.text_area(
            "Collez ici le tableau ACD, en conservant la ligne d'en-tête",
            height=280,
            placeholder="N°\tDate\tAnnuité\tTx. Int.\tIntérêts\tTx. Assur.\tAssurance\tTx. Com.\tCommission\tCapital remb.\tAutre\tRestant Dû",
        )
        if pasted.strip():
            raw_key = pasted.encode("utf-8")
            try:
                parsed = parse_acd_loan_paste(pasted)
            except Exception as exc:
                st.error(f"Lecture impossible : {exc}")

    if parsed is None:
        st.info("Importez un fichier ACD ou collez l'échéancier pour commencer.")
    else:
        try:
            st.success(
                f"{len(parsed.data)} ligne(s) ACD détectée(s) — source {parsed.source_type}, en-tête trouvé ligne {parsed.header_row}."
            )
            with st.expander("Vérifier les données ACD lues", expanded=False):
                st.dataframe(parsed.data, hide_index=True, use_container_width=True)

            with st.expander("Paramètres de conversion", expanded=True):
                first, second, third = st.columns(3)
                with first:
                    payment_label = st.selectbox(
                        "Calcul de l'échéance Pennylane",
                        [
                            "Annuité ACD + assurance + frais — recommandé",
                            "Somme des composantes détaillées",
                        ],
                        help=(
                            "Le premier choix conserve l'annuité ACD, puis ajoute l'assurance, la commission et les autres frais. "
                            "Le second recalcule tout depuis intérêts + capital + assurance + frais."
                        ),
                    )
                with second:
                    keep_zero = st.checkbox(
                        "Conserver les échéances à zéro",
                        value=True,
                        help="Recommandé : ces lignes représentent généralement un différé ou un moratoire et conservent la chronologie du prêt.",
                    )
                with third:
                    loan_name = st.text_input(
                        "Nom ou numéro de l'emprunt",
                        value=output_stem,
                        help="Utilisé uniquement pour nommer le fichier téléchargé.",
                    )

            result = convert_acd_loan(
                parsed,
                payment_method="components" if payment_label.startswith("Somme") else "annuity_plus_fees",
                keep_zero_installments=keep_zero,
            )
            for warning in result.source_warnings:
                st.warning(warning)
            if result.initial_rows_excluded:
                st.info(
                    f"La ligne ACD n°0 a été reconnue comme solde initial et retirée de l'échéancier. Capital détecté : "
                    f"{result.summary.loc[0, 'Capital (€)']:,.2f} €".replace(",", " ")
                )
            if result.zero_installments and keep_zero:
                st.info(
                    f"{result.zero_installments} échéance(s) à zéro conservée(s) pour respecter les périodes de différé ou de moratoire."
                )

            editor_settings = "|".join([payment_label, str(keep_zero)]).encode("utf-8")
            editor_hash = hashlib.sha1(raw_key + editor_settings).hexdigest()[:12]

            st.subheader("1. Paramètres de l'emprunt")
            st.caption(
                f"Périodicité détectée : {result.frequency_label}. Le taux d'intérêt affiché est annualisé à partir du taux périodique médian ACD."
            )
            summary_config = {
                "Capital (€)": st.column_config.NumberColumn(format="%.2f", min_value=0.01),
                "Taux d’intérêt (%)": st.column_config.NumberColumn(format="%.6f", min_value=0.0),
                "Taux d’assurance (%)": st.column_config.NumberColumn(format="%.6f", min_value=0.0),
                "Montant total d’assurance (€)": st.column_config.NumberColumn(format="%.2f", min_value=0.0),
                "Pourcentage des autres frais (%)": st.column_config.NumberColumn(format="%.6f", min_value=0.0),
                "Montant total des autres frais (€)": st.column_config.NumberColumn(format="%.2f", min_value=0.0),
                "Nombre d’échéances": st.column_config.NumberColumn(format="%d", min_value=1, step=1),
            }
            edited_summary = st.data_editor(
                result.summary,
                key=f"summary_{editor_hash}",
                column_config=summary_config,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
            ).reindex(columns=SUMMARY_COLUMNS)

            st.subheader("2. Échéancier à importer")
            schedule_config = {
                "Date": st.column_config.DateColumn(format="DD/MM/YYYY", required=True),
                **{
                    column: st.column_config.NumberColumn(format="%.2f", required=True)
                    for column in SCHEDULE_COLUMNS[1:]
                },
            }
            edited_schedule = st.data_editor(
                result.schedule,
                key=f"schedule_{editor_hash}",
                column_config=schedule_config,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                height=520,
            ).reindex(columns=SCHEDULE_COLUMNS)

            diagnostics = validate_pennylane_loan(edited_summary, edited_schedule)
            error_count = int((diagnostics["Niveau"] == "Erreur").sum()) if not diagnostics.empty else 0
            warning_count = int((diagnostics["Niveau"] == "Avertissement").sum()) if not diagnostics.empty else 0
            blocking_errors = error_count > 0

            capital = pd.to_numeric(edited_summary["Capital (€)"], errors="coerce").iloc[0]
            total_interest = pd.to_numeric(edited_schedule["Intérêt (€)"], errors="coerce").fillna(0).sum()
            final_balance = pd.to_numeric(edited_schedule["Solde (€)"], errors="coerce").iloc[-1]
            metric_1, metric_2, metric_3, metric_4 = st.columns(4)
            metric_1.metric("Capital initial", f"{capital:,.2f} €".replace(",", " "))
            metric_2.metric("Échéances", len(edited_schedule))
            metric_3.metric("Intérêts totaux", f"{total_interest:,.2f} €".replace(",", " "))
            metric_4.metric("Solde final", f"{final_balance:,.2f} €".replace(",", " "))

            if diagnostics.empty:
                st.success("Contrôles terminés : l'échéancier est cohérent.")
            else:
                if error_count:
                    st.error("Corrigez les erreurs avant de télécharger le fichier Pennylane.")
                else:
                    st.warning("Le fichier peut être téléchargé, mais les avertissements doivent être vérifiés.")
                st.dataframe(diagnostics, hide_index=True, use_container_width=True)

            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", loan_name).strip("_") or "echeancier_emprunt"
            xlsx_content = to_xlsx_bytes(edited_summary, edited_schedule)
            csv_content = to_csv_bytes(edited_summary, edited_schedule)
            left, right = st.columns(2)
            with left:
                st.download_button(
                    "Télécharger le XLSX Pennylane — recommandé",
                    data=xlsx_content,
                    file_name=f"{safe_name}_pennylane.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    disabled=blocking_errors,
                    use_container_width=True,
                )
            with right:
                st.download_button(
                    "Télécharger la version CSV",
                    data=csv_content,
                    file_name=f"{safe_name}_pennylane.csv",
                    mime="text/csv",
                    disabled=blocking_errors,
                    use_container_width=True,
                )

        except Exception as exc:
            st.error(f"Conversion impossible : {exc}")

with clipboard_tab:
    st.subheader("Copier-coller dans le modèle Pennylane")
    if edited_summary is None or edited_schedule is None:
        st.info("Importez ou collez d'abord un échéancier ACD dans l'onglet Conversion.")
    elif blocking_errors:
        st.error("Le copier-coller sera disponible après correction des erreurs.")
    else:
        st.markdown(
            """
            1. Cliquez sur l'icône **Copier** du bloc ci-dessous.
            2. Ouvrez un classeur Excel vide ou le modèle Pennylane.
            3. Sélectionnez la cellule **A1**, puis collez avec **Ctrl+V**.
            4. Enregistrez au format `.xlsx`, puis importez le fichier dans Pennylane.

            Les deux lignes vides entre les paramètres et l'échéancier sont déjà incluses.
            """
        )
        clipboard_text = to_clipboard_tsv(edited_summary, edited_schedule)
        st.code(clipboard_text, language=None)
        st.download_button(
            "Télécharger aussi la version tabulée (.txt)",
            data=clipboard_text.encode("utf-8-sig"),
            file_name="echeancier_pennylane_copier_coller.txt",
            mime="text/tab-separated-values",
        )

with help_tab:
    st.subheader("Structure produite pour Pennylane")
    st.markdown(
        """
        Le classeur reproduit la structure du modèle fourni :

        - ligne 1 : intitulés des paramètres de l'emprunt ;
        - ligne 2 : capital, taux, frais et nombre d'échéances ;
        - lignes 3 et 4 : lignes vides ;
        - ligne 5 : en-têtes de l'échéancier ;
        - à partir de la ligne 6 : échéances détaillées.
        """
    )

    st.subheader("Règles de conversion ACD")
    rules = pd.DataFrame(
        [
            ["Capital", "Restant Dû de la ligne n°0", "La ligne n°0 est un solde initial, pas une échéance."],
            ["Taux d'intérêt", "Médiane de Tx. Int.", "Annualisé selon la périodicité détectée ; en mensuel, taux × 12."],
            ["Taux d'assurance", "Total Assurance / Capital", "Pourcentage total sur la durée, comme dans le modèle Pennylane fourni."],
            ["Autres frais", "Commission + Autre", "Regroupés dans une seule colonne Pennylane."],
            ["Intérêt", "Intérêts", "Montant de chaque période."],
            ["Amortissement", "Capital remb.", "Capital remboursé à chaque échéance."],
            ["Échéance", "Annuité + assurance + frais", "Option recommandée ; un recalcul détaillé est également proposé."],
            ["Solde", "Restant Dû", "Contrôlé à chaque ligne : solde précédent - capital remboursé."],
            ["Différés", "Lignes entièrement à zéro", "Conservés par défaut pour respecter la chronologie du prêt."],
            ["Nombre d'échéances", "Nombre de lignes après la ligne n°0", "Inclut les échéances nulles conservées."],
        ],
        columns=["Champ Pennylane", "Source ACD", "Traitement"],
    )
    st.dataframe(rules, hide_index=True, use_container_width=True)

    st.info(
        "Dans l'exemple transmis, le taux ACD d'environ 0,066674 % est un taux mensuel. L'application l'annualise à environ 0,800089 %."
    )

