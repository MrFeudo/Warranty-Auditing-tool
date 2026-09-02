# -*- coding: utf-8 -*-
"""
Warranty Audit Portal
Streamlit app para digitalizar auditorías de garantías.

Instalación:
    py -m pip install streamlit pandas openpyxl xlsxwriter

Ejecución:
    streamlit run warranty_audit_portal.py

Notas de cálculo:
- Claim document checklist I suma 58 puntos.
- Claim old parts checklist II suma 42 puntos.
- Total auditoría = 100 puntos.
- No aplica = máxima puntuación del apartado.
- Campañas son informativas y no suman ni restan.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =============================================================================
# CONFIGURACIÓN DE PUNTUACIÓN
# =============================================================================

@dataclass(frozen=True)
class AuditOption:
    label: str
    points: Optional[int]
    status: str


@dataclass(frozen=True)
class AuditCheck:
    block: str
    key: str
    label: str
    max_points: int
    options: Tuple[AuditOption, ...]
    guidance: str


PENDING = AuditOption("Pendiente de revisar", None, "Pendiente")


def options_0_5_7() -> Tuple[AuditOption, ...]:
    return (
        PENDING,
        AuditOption("OK / conforme", 7, "OK"),
        AuditOption("Parcial / incompleto", 5, "Parcial"),
        AuditOption("NOK / falta o no conforme", 0, "NOK"),
        AuditOption("No aplica", 7, "N/A"),
    )


def options_0_6() -> Tuple[AuditOption, ...]:
    return (
        PENDING,
        AuditOption("OK / conforme", 6, "OK"),
        AuditOption("NOK / no conforme", 0, "NOK"),
        AuditOption("No aplica", 6, "N/A"),
    )


def options_date_0_3_6() -> Tuple[AuditOption, ...]:
    return (
        PENDING,
        AuditOption("0 a 5 días laborables", 6, "OK"),
        AuditOption("5 a 15 días laborables", 3, "Parcial"),
        AuditOption("Más de 15 días laborables", 0, "NOK"),
        AuditOption("No aplica", 6, "N/A"),
    )


def options_old_binary_0_7() -> Tuple[AuditOption, ...]:
    return (
        PENDING,
        AuditOption("OK / correcta", 7, "OK"),
        AuditOption("NOK / incorrecta", 0, "NOK"),
        AuditOption("No aplica", 7, "N/A"),
    )


def options_campaigns() -> Tuple[AuditOption, ...]:
    return (
        AuditOption("Pendiente de revisar", None, "Pendiente"),
        AuditOption("OK", None, "OK"),
        AuditOption("NOK / revisar", None, "NOK"),
        AuditOption("No aplica", None, "N/A"),
    )


DOCUMENT_CHECKS: List[AuditCheck] = [
    AuditCheck(
        "Claim document checklist I", "doc_or", "OR", 7, options_0_5_7(),
        "Comprobar si existe la orden de reparación/reclamación, si la documentación es completa y si la firma está correcta.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_parts_order", "Pedido piezas / albarán", 7, options_0_5_7(),
        "Comprobar pedido/albarán de piezas y que la documentación sea completa, razonable y conforme.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_previous_or", "OR previa", 7, options_0_5_7(),
        "Solo para reclamaciones de repuestos o reparaciones anteriores. Si no procede, marcar No aplica.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_evidence", "Evidencias", 7, options_0_5_7(),
        "Comprobar si faltan evidencias adjuntas o si las evidencias no son correctas/suficientes.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_causal_part", "Pieza causa correcta", 6, options_0_6(),
        "Comprobar si la pieza principal dañada es correcta, razonable y coherente con la avería.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_labor", "Mano de obra", 6, options_0_6(),
        "Comprobar si los tiempos de mano de obra son correctos, no repetitivos y ajustados al estándar.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_aux_material", "Material auxiliar", 6, options_0_6(),
        "Comprobar que el material auxiliar/consumible cumpla normativa, campo correcto y cantidad razonable.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_dates", "Fecha/hora envío Claim", 6, options_date_0_3_6(),
        "La claim debe enviarse dentro de plazo tras completar la reparación.",
    ),
    AuditCheck(
        "Claim document checklist I", "doc_vin", "VIN", 6, options_0_6(),
        "Comprobar que VIN, kilometraje, fecha de reparación, tipo de reclamación y datos del vehículo sean correctos.",
    ),
]

OLD_PARTS_CHECKS: List[AuditCheck] = [
    AuditCheck(
        "Claim old parts checklist II", "old_management", "Gestión piezas viejas", 7, options_0_5_7(),
        "Las piezas viejas deben estar ordenadas, localizables y disponibles durante la auditoría.",
    ),
    AuditCheck(
        "Claim old parts checklist II", "old_label", "Etiquetado pieza vieja", 7, options_0_5_7(),
        "La etiqueta debe incluir datos básicos del vehículo, claim, referencia, causa del daño, etc.",
    ),
    AuditCheck(
        "Claim old parts checklist II", "old_causal_part", "Pieza causa", 7, options_old_binary_0_7(),
        "Comprobar que la pieza antigua es la causa real, consistente con modelo/vehículo y fecha de producción.",
    ),
    AuditCheck(
        "Claim old parts checklist II", "old_failure_info", "Info tipo fallo pieza causa", 7, options_old_binary_0_7(),
        "Comprobar que la información del fallo sea coherente con la pieza y las fotos/evidencias.",
    ),
    AuditCheck(
        "Claim old parts checklist II", "old_destruction", "Destrucción pieza vieja", 7, options_0_5_7(),
        "Las piezas viejas deben destruirse siguiendo el proceso y sin posibilidad de reutilización.",
    ),
    AuditCheck(
        "Claim old parts checklist II", "old_destruction_certificate", "Certificado destrucción piezas viejas", 7, options_0_5_7(),
        "El certificado/informe de destrucción debe existir, subirse y archivarse en plazo.",
    ),
]

CAMPAIGN_CHECKS: List[AuditCheck] = [
    AuditCheck(
        "Campañas informativas", "info_campaign_check", "Comprobación campañas", 0, options_campaigns(),
        "Campo informativo. No afecta al porcentaje de éxito.",
    ),
    AuditCheck(
        "Campañas informativas", "info_pending_campaigns", "Campañas pendientes", 0, options_campaigns(),
        "Campo informativo. Indicar si existen campañas pendientes o si se ha revisado correctamente.",
    ),
]

ALL_SCORING_CHECKS = DOCUMENT_CHECKS + OLD_PARTS_CHECKS
ALL_CHECKS = DOCUMENT_CHECKS + OLD_PARTS_CHECKS + CAMPAIGN_CHECKS
MAX_DOCUMENT_POINTS = sum(check.max_points for check in DOCUMENT_CHECKS)  # 58
MAX_OLD_PARTS_POINTS = sum(check.max_points for check in OLD_PARTS_CHECKS)  # 42
MAX_TOTAL_POINTS = MAX_DOCUMENT_POINTS + MAX_OLD_PARTS_POINTS  # 100


# =============================================================================
# UTILIDADES
# =============================================================================

def option_labels(check: AuditCheck) -> List[str]:
    labels = []
    for option in check.options:
        if option.points is None:
            labels.append(option.label)
        else:
            labels.append(f"{option.label} ({option.points}/{check.max_points})")
    return labels


def option_from_label(check: AuditCheck, selected_label: str) -> AuditOption:
    for option in check.options:
        prefix = option.label
        if selected_label == prefix or selected_label.startswith(prefix + " ("):
            return option
    return PENDING


def option_from_points(check: AuditCheck, points: Any) -> AuditOption:
    """Prefill desde un Excel existente, si trae puntuaciones numéricas."""
    try:
        if pd.isna(points) or str(points).strip() == "":
            return PENDING
        numeric_points = int(float(points))
    except Exception:
        return PENDING

    candidates = [option for option in check.options if option.points == numeric_points]
    if candidates:
        return candidates[0]
    return PENDING


def safe_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def new_evaluation(check: AuditCheck, prefill_points: Any = None) -> Dict[str, Any]:
    option = option_from_points(check, prefill_points)
    return {
        "status": option.status,
        "label": option.label,
        "points": option.points,
        "max_points": check.max_points,
        "comment": "",
    }


def empty_claim_record(claim_no: str) -> Dict[str, Any]:
    return {
        "claim_no": claim_no,
        "dealer": "",
        "vin": "",
        "model": "",
        "amount": "",
        "repair_date": "",
        "submission_date": "",
        "general_comment": "",
        "evaluations": {check.key: new_evaluation(check) for check in ALL_CHECKS},
    }


def calculate_claim_score(claim: Dict[str, Any]) -> Dict[str, Any]:
    doc_points = 0
    old_points = 0
    pending = []
    lost_by_area = []

    for check in DOCUMENT_CHECKS:
        evaluation = claim["evaluations"].get(check.key, {})
        points = evaluation.get("points")
        if points is None:
            pending.append(check.label)
            continue
        doc_points += int(points)
        lost = check.max_points - int(points)
        if lost > 0:
            lost_by_area.append((check.block, check.label, lost, check.max_points, evaluation.get("status", "")))

    for check in OLD_PARTS_CHECKS:
        evaluation = claim["evaluations"].get(check.key, {})
        points = evaluation.get("points")
        if points is None:
            pending.append(check.label)
            continue
        old_points += int(points)
        lost = check.max_points - int(points)
        if lost > 0:
            lost_by_area.append((check.block, check.label, lost, check.max_points, evaluation.get("status", "")))

    total_points = doc_points + old_points
    completed = len(pending) == 0

    return {
        "doc_points": doc_points,
        "old_points": old_points,
        "total_points": total_points,
        "success_percent": total_points,  # como el máximo fijo es 100
        "pending": pending,
        "completed": completed,
        "lost_by_area": lost_by_area,
    }


def calculate_audit_score(claims: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    if not claims:
        return {
            "claims": 0,
            "completed_claims": 0,
            "doc_points": 0,
            "old_points": 0,
            "total_points": 0,
            "max_points": 0,
            "success_percent": 0,
        }

    completed_claims = 0
    doc_points = 0
    old_points = 0
    total_points = 0

    for claim in claims.values():
        score = calculate_claim_score(claim)
        if score["completed"]:
            completed_claims += 1
        doc_points += score["doc_points"]
        old_points += score["old_points"]
        total_points += score["total_points"]

    max_points = len(claims) * MAX_TOTAL_POINTS
    percent = (total_points / max_points * 100) if max_points else 0

    return {
        "claims": len(claims),
        "completed_claims": completed_claims,
        "doc_points": doc_points,
        "old_points": old_points,
        "total_points": total_points,
        "max_points": max_points,
        "success_percent": percent,
    }


# =============================================================================
# IMPORTACIÓN DE EXCEL
# =============================================================================

def read_claims_from_uploaded_excel(uploaded_file) -> Dict[str, Dict[str, Any]]:
    """
    Lee la checklist oficial si encuentra las hojas:
    - Claim document checklist I
    - Claim old parts checklist II

    También soporta un Excel simple con una columna Claim No. / Claim / claim_number.
    """
    xls = pd.ExcelFile(uploaded_file)
    claims: Dict[str, Dict[str, Any]] = {}

    if "Claim document checklist I" in xls.sheet_names:
        doc_df = pd.read_excel(uploaded_file, sheet_name="Claim document checklist I", header=None)
        # En la plantilla: fila 3 humana = índice 2, datos desde fila 4 = índice 3.
        for row_idx in range(3, len(doc_df)):
            claim_no = safe_str(doc_df.iloc[row_idx, 1] if doc_df.shape[1] > 1 else "")
            if not claim_no or claim_no.lower() == "nan":
                continue
            claim = claims.setdefault(claim_no, empty_claim_record(claim_no))

            doc_columns = {
                "doc_or": 2,
                "doc_parts_order": 3,
                "doc_previous_or": 4,
                "doc_evidence": 5,
                "doc_causal_part": 6,
                "doc_labor": 7,
                "doc_aux_material": 8,
                "doc_dates": 9,
                "doc_vin": 10,
            }
            check_by_key = {check.key: check for check in DOCUMENT_CHECKS}
            for key, col_idx in doc_columns.items():
                if col_idx < doc_df.shape[1]:
                    claim["evaluations"][key] = new_evaluation(check_by_key[key], doc_df.iloc[row_idx, col_idx])

            if doc_df.shape[1] > 15:
                claim["general_comment"] = safe_str(doc_df.iloc[row_idx, 15])
            if doc_df.shape[1] > 11:
                campaign_check = safe_str(doc_df.iloc[row_idx, 11])
                if campaign_check:
                    claim["evaluations"]["info_campaign_check"]["comment"] = campaign_check
            if doc_df.shape[1] > 12:
                pending_campaigns = safe_str(doc_df.iloc[row_idx, 12])
                if pending_campaigns:
                    claim["evaluations"]["info_pending_campaigns"]["comment"] = pending_campaigns

    if "Claim old parts checklist II" in xls.sheet_names:
        old_df = pd.read_excel(uploaded_file, sheet_name="Claim old parts checklist II", header=None)
        for row_idx in range(3, len(old_df)):
            claim_no = safe_str(old_df.iloc[row_idx, 1] if old_df.shape[1] > 1 else "")
            if not claim_no or claim_no.lower() == "nan":
                continue
            claim = claims.setdefault(claim_no, empty_claim_record(claim_no))

            old_columns = {
                "old_management": 2,
                "old_label": 3,
                "old_causal_part": 4,
                "old_failure_info": 5,
                "old_destruction": 6,
                "old_destruction_certificate": 7,
            }
            check_by_key = {check.key: check for check in OLD_PARTS_CHECKS}
            for key, col_idx in old_columns.items():
                if col_idx < old_df.shape[1]:
                    claim["evaluations"][key] = new_evaluation(check_by_key[key], old_df.iloc[row_idx, col_idx])

            if old_df.shape[1] > 10:
                old_comment = safe_str(old_df.iloc[row_idx, 10])
                if old_comment:
                    existing = claim.get("general_comment", "")
                    claim["general_comment"] = (existing + "\n" + old_comment).strip() if existing else old_comment

    # Fallback para Excel simple de HQ.
    if not claims:
        df = pd.read_excel(uploaded_file, sheet_name=xls.sheet_names[0])
        normalized_cols = {str(col).strip().lower(): col for col in df.columns}
        possible_claim_cols = ["claim no.", "claim no", "claim", "claim_number", "claim number", "garantía", "garantia"]
        claim_col = next((normalized_cols[col] for col in possible_claim_cols if col in normalized_cols), None)
        if claim_col is None:
            raise ValueError("No encuentro una columna de Claim No. en el Excel.")

        for _, row in df.iterrows():
            claim_no = safe_str(row.get(claim_col, ""))
            if not claim_no:
                continue
            claim = claims.setdefault(claim_no, empty_claim_record(claim_no))
            for target, possible_names in {
                "dealer": ["dealer", "concesionario", "service dealer"],
                "vin": ["vin", "chassis", "bastidor"],
                "model": ["model", "modelo"],
                "amount": ["amount", "importe", "coste", "total"],
                "repair_date": ["repair date", "fecha reparacion", "fecha reparación"],
                "submission_date": ["submission date", "fecha envio", "fecha envío"],
            }.items():
                source_col = next((normalized_cols[name] for name in possible_names if name in normalized_cols), None)
                if source_col is not None:
                    claim[target] = safe_str(row.get(source_col, ""))

    return claims


# =============================================================================
# EXPORTACIÓN E INFORME
# =============================================================================

def build_summary_dataframe(claims: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for claim in claims.values():
        score = calculate_claim_score(claim)
        rows.append({
            "Claim No.": claim["claim_no"],
            "Dealer": claim.get("dealer", ""),
            "VIN": claim.get("vin", ""),
            "Modelo": claim.get("model", ""),
            "Puntos documentación": score["doc_points"],
            "Puntos piezas viejas": score["old_points"],
            "Resultado /100": score["total_points"],
            "% éxito": score["success_percent"],
            "Estado": "Completada" if score["completed"] else "Pendiente",
            "Pendientes": " | ".join(score["pending"]),
            "Comentarios": claim.get("general_comment", ""),
        })
    return pd.DataFrame(rows)


def build_detail_dataframe(claims: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    check_by_key = {check.key: check for check in ALL_CHECKS}
    for claim in claims.values():
        for key, evaluation in claim["evaluations"].items():
            check = check_by_key[key]
            rows.append({
                "Claim No.": claim["claim_no"],
                "Bloque": check.block,
                "Apartado": check.label,
                "Estado": evaluation.get("status", ""),
                "Puntos": evaluation.get("points"),
                "Máximo": check.max_points,
                "Pérdida": "" if evaluation.get("points") is None else check.max_points - int(evaluation.get("points") or 0),
                "Comentario apartado": evaluation.get("comment", ""),
                "Criterio": check.guidance,
            })
    return pd.DataFrame(rows)


def build_improvement_dataframe(claims: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for claim in claims.values():
        score = calculate_claim_score(claim)
        for block, area, lost, max_points, status in score["lost_by_area"]:
            rows.append({
                "Claim No.": claim["claim_no"],
                "Bloque": block,
                "Área de mejora": area,
                "Estado": status,
                "Puntos perdidos": lost,
                "Máximo apartado": max_points,
                "Comentario claim": claim.get("general_comment", ""),
            })
    return pd.DataFrame(rows)


def export_excel(claims: Dict[str, Dict[str, Any]], audit_name: str, dealer: str, auditor: str) -> bytes:
    output = BytesIO()
    summary_df = build_summary_dataframe(claims)
    detail_df = build_detail_dataframe(claims)
    improvement_df = build_improvement_dataframe(claims)
    audit_score = calculate_audit_score(claims)

    cover_df = pd.DataFrame([
        {"Campo": "Auditoría", "Valor": audit_name},
        {"Campo": "Dealer", "Valor": dealer},
        {"Campo": "Auditor", "Valor": auditor},
        {"Campo": "Fecha exportación", "Valor": datetime.now().strftime("%Y-%m-%d %H:%M")},
        {"Campo": "Claims", "Valor": audit_score["claims"]},
        {"Campo": "Claims completadas", "Valor": audit_score["completed_claims"]},
        {"Campo": "Resultado", "Valor": f"{audit_score['total_points']}/{audit_score['max_points']}"},
        {"Campo": "% éxito", "Valor": f"{audit_score['success_percent']:.1f}%"},
    ])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        cover_df.to_excel(writer, index=False, sheet_name="Resumen auditoría")
        summary_df.to_excel(writer, index=False, sheet_name="Claims")
        detail_df.to_excel(writer, index=False, sheet_name="Detalle apartados")
        improvement_df.to_excel(writer, index=False, sheet_name="Áreas de mejora")

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        percent_format = workbook.add_format({"num_format": "0.0%"})
        integer_format = workbook.add_format({"num_format": "0"})
        wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, None, header_format)
            worksheet.autofilter(0, 0, 0, 20)
            worksheet.set_column(0, 0, 18)
            worksheet.set_column(1, 4, 24)
            worksheet.set_column(5, 20, 18)
            if sheet_name in ["Detalle apartados", "Áreas de mejora", "Claims"]:
                worksheet.set_column(8, 10, 45, wrap_format)
                worksheet.set_column(9, 9, 60, wrap_format)

        if "Claims" in writer.sheets:
            ws = writer.sheets["Claims"]
            ws.set_column(4, 7, 18, integer_format)
            ws.set_column(10, 10, 60, wrap_format)

    return output.getvalue()


def generate_text_report(claims: Dict[str, Dict[str, Any]], audit_name: str, dealer: str, auditor: str) -> str:
    audit_score = calculate_audit_score(claims)
    improvement_df = build_improvement_dataframe(claims)
    summary_df = build_summary_dataframe(claims)

    if improvement_df.empty:
        top_areas = pd.DataFrame(columns=["Área de mejora", "Puntos perdidos"])
    else:
        top_areas = (
            improvement_df.groupby("Área de mejora", as_index=False)["Puntos perdidos"]
            .sum()
            .sort_values("Puntos perdidos", ascending=False)
            .head(5)
        )

    critical_claims = summary_df.sort_values("Resultado /100", ascending=True).head(5)

    lines = []
    lines.append(f"Informe de auditoría: {audit_name or 'Sin nombre'}")
    lines.append(f"Dealer: {dealer or 'No informado'}")
    lines.append(f"Auditor: {auditor or 'No informado'}")
    lines.append("")
    lines.append("Resumen ejecutivo")
    lines.append(
        f"Se han revisado {audit_score['claims']} garantías, con {audit_score['completed_claims']} completadas. "
        f"El resultado global de la auditoría es {audit_score['success_percent']:.1f}% "
        f"({audit_score['total_points']}/{audit_score['max_points']} puntos)."
    )
    lines.append("")
    lines.append("Resultado por bloque")
    lines.append(f"- Documentación de claim: {audit_score['doc_points']}/{audit_score['claims'] * MAX_DOCUMENT_POINTS} puntos.")
    lines.append(f"- Piezas viejas: {audit_score['old_points']}/{audit_score['claims'] * MAX_OLD_PARTS_POINTS} puntos.")
    lines.append("- Campañas: revisión informativa, sin impacto en puntuación.")
    lines.append("")
    lines.append("Principales áreas de mejora")
    if top_areas.empty:
        lines.append("No se han detectado desviaciones puntuables.")
    else:
        for _, row in top_areas.iterrows():
            lines.append(f"- {row['Área de mejora']}: {int(row['Puntos perdidos'])} puntos perdidos.")
    lines.append("")
    lines.append("Claims con menor puntuación")
    for _, row in critical_claims.iterrows():
        lines.append(f"- {row['Claim No.']}: {row['Resultado /100']}/100. {safe_str(row['Comentarios'])}")
    lines.append("")
    lines.append("Conclusión")
    lines.append(
        "Se recomienda focalizar el plan de mejora en las áreas con mayor pérdida de puntos, "
        "reforzando la calidad documental, la justificación técnica y la trazabilidad de piezas viejas cuando aplique."
    )
    return "\n".join(lines)


# =============================================================================
# INTERFAZ STREAMLIT
# =============================================================================

def init_state():
    if "claims" not in st.session_state:
        st.session_state.claims = {}
    if "selected_claim" not in st.session_state:
        st.session_state.selected_claim = None


def render_check_editor(claim: Dict[str, Any], checks: List[AuditCheck]):
    for check in checks:
        evaluation = claim["evaluations"][check.key]
        labels = option_labels(check)

        # Buscar índice actual.
        current_label = evaluation.get("label", PENDING.label)
        current_index = 0
        for idx, label in enumerate(labels):
            if label.startswith(current_label):
                current_index = idx
                break

        with st.container(border=True):
            cols = st.columns([2.2, 1.2, 2.8])
            with cols[0]:
                st.markdown(f"**{check.label}**")
                st.caption(f"Máximo: {check.max_points} puntos · {check.guidance}")
            with cols[1]:
                selected = st.selectbox(
                    "Evaluación",
                    labels,
                    index=current_index,
                    key=f"select_{claim['claim_no']}_{check.key}",
                    label_visibility="collapsed",
                )
                option = option_from_label(check, selected)
                evaluation["status"] = option.status
                evaluation["label"] = option.label
                evaluation["points"] = option.points
                evaluation["max_points"] = check.max_points

                if option.points is None:
                    st.warning("Pendiente")
                else:
                    st.metric("Puntos", f"{option.points}/{check.max_points}")
            with cols[2]:
                evaluation["comment"] = st.text_area(
                    "Comentario del apartado",
                    value=evaluation.get("comment", ""),
                    key=f"comment_{claim['claim_no']}_{check.key}",
                    height=90,
                )


def render_campaign_editor(claim: Dict[str, Any]):
    st.info("Las campañas son orientativas/informativas. No suman ni restan en el porcentaje de éxito.")
    render_check_editor(claim, CAMPAIGN_CHECKS)


def main():
    st.set_page_config(page_title="Warranty Audit Portal", layout="wide")
    init_state()

    st.title("Warranty Audit Portal")
    st.caption("Auditoría online de garantías: documentación + piezas viejas = 100 puntos. Campañas solo informativo.")

    with st.sidebar:
        st.header("Auditoría")
        audit_name = st.text_input("Nombre auditoría", value="Auditoría garantías")
        dealer = st.text_input("Dealer", value="")
        auditor = st.text_input("Auditor", value="")

        st.divider()
        uploaded_file = st.file_uploader("Subir checklist o lista HQ", type=["xlsx", "xlsm", "xls"])
        if uploaded_file is not None and st.button("Cargar claims", type="primary"):
            try:
                st.session_state.claims = read_claims_from_uploaded_excel(uploaded_file)
                if st.session_state.claims:
                    st.session_state.selected_claim = next(iter(st.session_state.claims.keys()))
                    st.success(f"Cargadas {len(st.session_state.claims)} claims.")
                else:
                    st.warning("No se encontraron claims en el archivo.")
            except Exception as exc:
                st.error(f"No se pudo cargar el archivo: {exc}")

        st.divider()
        manual_claim = st.text_input("Añadir claim manual")
        if st.button("Añadir claim") and manual_claim.strip():
            claim_no = manual_claim.strip()
            st.session_state.claims.setdefault(claim_no, empty_claim_record(claim_no))
            st.session_state.selected_claim = claim_no
            st.success(f"Claim {claim_no} añadida.")

        st.divider()
        st.subheader("Regla de puntuación")
        st.write(f"Documentación: **{MAX_DOCUMENT_POINTS}**")
        st.write(f"Piezas viejas: **{MAX_OLD_PARTS_POINTS}**")
        st.write(f"Total: **{MAX_TOTAL_POINTS}**")
        st.caption("No aplica = máximo del apartado. Campañas = informativo.")

    claims: Dict[str, Dict[str, Any]] = st.session_state.claims

    if not claims:
        st.info("Sube la checklist de auditoría o añade una claim manual para empezar.")
        st.stop()

    audit_score = calculate_audit_score(claims)
    kpi_cols = st.columns(5)
    kpi_cols[0].metric("Claims", audit_score["claims"])
    kpi_cols[1].metric("Completadas", audit_score["completed_claims"])
    kpi_cols[2].metric("Documentación", f"{audit_score['doc_points']}/{audit_score['claims'] * MAX_DOCUMENT_POINTS}")
    kpi_cols[3].metric("Piezas viejas", f"{audit_score['old_points']}/{audit_score['claims'] * MAX_OLD_PARTS_POINTS}")
    kpi_cols[4].metric("Éxito global", f"{audit_score['success_percent']:.1f}%")

    st.divider()

    left, right = st.columns([1.1, 2.4])

    with left:
        st.subheader("Claims")
        summary_df = build_summary_dataframe(claims)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        claim_options = list(claims.keys())
        if st.session_state.selected_claim not in claim_options:
            st.session_state.selected_claim = claim_options[0]

        selected_claim = st.selectbox(
            "Seleccionar claim",
            claim_options,
            index=claim_options.index(st.session_state.selected_claim),
        )
        st.session_state.selected_claim = selected_claim

        st.download_button(
            "Exportar auditoría a Excel",
            data=export_excel(claims, audit_name, dealer, auditor),
            file_name=f"audit_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with right:
        claim = claims[st.session_state.selected_claim]
        score = calculate_claim_score(claim)

        st.subheader(f"Revisión claim {claim['claim_no']}")
        header_cols = st.columns(4)
        claim["dealer"] = header_cols[0].text_input("Dealer claim", value=claim.get("dealer", ""), key=f"dealer_{claim['claim_no']}")
        claim["vin"] = header_cols[1].text_input("VIN", value=claim.get("vin", ""), key=f"vin_{claim['claim_no']}")
        claim["model"] = header_cols[2].text_input("Modelo", value=claim.get("model", ""), key=f"model_{claim['claim_no']}")
        claim["amount"] = header_cols[3].text_input("Importe", value=claim.get("amount", ""), key=f"amount_{claim['claim_no']}")

        score_cols = st.columns(4)
        score_cols[0].metric("Documentación", f"{score['doc_points']}/{MAX_DOCUMENT_POINTS}")
        score_cols[1].metric("Piezas viejas", f"{score['old_points']}/{MAX_OLD_PARTS_POINTS}")
        score_cols[2].metric("Resultado claim", f"{score['total_points']}/100")
        score_cols[3].metric("Estado", "Completada" if score["completed"] else "Pendiente")

        if score["pending"]:
            st.warning("Apartados pendientes: " + ", ".join(score["pending"]))

        tabs = st.tabs(["I. Documentación", "II. Piezas viejas", "Campañas", "Comentarios", "Informe"])

        with tabs[0]:
            render_check_editor(claim, DOCUMENT_CHECKS)

        with tabs[1]:
            render_check_editor(claim, OLD_PARTS_CHECKS)

        with tabs[2]:
            render_campaign_editor(claim)

        with tabs[3]:
            claim["general_comment"] = st.text_area(
                "Comentarios generales de la claim",
                value=claim.get("general_comment", ""),
                key=f"general_comment_{claim['claim_no']}",
                height=180,
            )

            if st.button("Generar comentario base desde desviaciones", key=f"generate_comment_{claim['claim_no']}"):
                pieces = []
                for block, area, lost, max_points, status in score["lost_by_area"]:
                    pieces.append(f"{area}: desviación detectada ({status}), {lost}/{max_points} puntos perdidos.")
                claim["general_comment"] = "\n".join(pieces)
                st.rerun()

        with tabs[4]:
            report = generate_text_report(claims, audit_name, dealer, auditor)
            st.text_area("Informe generado", value=report, height=420)
            st.download_button(
                "Descargar informe .txt",
                data=report.encode("utf-8"),
                file_name=f"audit_report_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain",
            )


if __name__ == "__main__":
    main()
