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
- Permite guardar y recargar auditorías de trabajo en JSON.
"""

from __future__ import annotations

import json
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
# GUARDAR / CARGAR AUDITORÍA DE TRABAJO
# =============================================================================

def normalize_loaded_evaluation(check: AuditCheck, raw_evaluation: Any) -> Dict[str, Any]:
    """Normaliza una evaluación cargada desde JSON para que siempre sea compatible."""
    base = new_evaluation(check)

    if not isinstance(raw_evaluation, dict):
        return base

    label = safe_str(raw_evaluation.get("label", ""))
    status = safe_str(raw_evaluation.get("status", ""))
    raw_points = raw_evaluation.get("points", None)

    option = None
    if label:
        option = option_from_label(check, label)
    elif raw_points is not None:
        option = option_from_points(check, raw_points)

    if option is None:
        option = PENDING

    # Si venía como Pendiente con points None, respetamos pendiente.
    points = option.points
    if raw_points is None and status.lower() == "pendiente":
        points = None

    return {
        "status": status or option.status,
        "label": label or option.label,
        "points": points,
        "max_points": check.max_points,
        "comment": safe_str(raw_evaluation.get("comment", "")),
    }


def normalize_loaded_claim(raw_claim: Any, fallback_claim_no: str = "") -> Optional[Dict[str, Any]]:
    """Convierte una claim guardada en JSON a la estructura interna actual."""
    if not isinstance(raw_claim, dict):
        return None

    claim_no = safe_str(raw_claim.get("claim_no", fallback_claim_no))
    if not claim_no:
        return None

    claim = empty_claim_record(claim_no)

    for field in [
        "dealer",
        "vin",
        "model",
        "amount",
        "repair_date",
        "submission_date",
        "general_comment",
    ]:
        claim[field] = safe_str(raw_claim.get(field, claim.get(field, "")))

    raw_evaluations = raw_claim.get("evaluations", {})
    if isinstance(raw_evaluations, dict):
        for check in ALL_CHECKS:
            if check.key in raw_evaluations:
                claim["evaluations"][check.key] = normalize_loaded_evaluation(check, raw_evaluations[check.key])

    return claim


def serialize_audit_workfile(
    claims: Dict[str, Dict[str, Any]],
    audit_name: str,
    dealer: str,
    auditor: str,
) -> bytes:
    """Genera un archivo JSON descargable para poder continuar la auditoría otro día."""
    payload = {
        "file_type": "warranty_audit_workfile",
        "version": 1,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "audit": {
            "audit_name": audit_name or "",
            "dealer": dealer or "",
            "auditor": auditor or "",
        },
        "claims": claims,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def load_audit_workfile(uploaded_file) -> Tuple[Dict[str, Dict[str, Any]], str, str, str]:
    """Carga una auditoría de trabajo guardada previamente como JSON."""
    content = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8")
    payload = json.loads(content)

    if not isinstance(payload, dict) or payload.get("file_type") != "warranty_audit_workfile":
        raise ValueError("El archivo no parece una auditoría de trabajo generada por esta app.")

    audit = payload.get("audit", {}) if isinstance(payload.get("audit", {}), dict) else {}
    raw_claims = payload.get("claims", {})

    claims: Dict[str, Dict[str, Any]] = {}

    if isinstance(raw_claims, dict):
        iterable = raw_claims.items()
    elif isinstance(raw_claims, list):
        iterable = ((safe_str(item.get("claim_no", "")) if isinstance(item, dict) else "", item) for item in raw_claims)
    else:
        iterable = []

    for fallback_claim_no, raw_claim in iterable:
        claim = normalize_loaded_claim(raw_claim, fallback_claim_no)
        if claim is not None:
            claims[claim["claim_no"]] = claim

    if not claims:
        raise ValueError("El archivo se pudo abrir, pero no contiene claims válidas.")

    return (
        claims,
        safe_str(audit.get("audit_name", "Auditoría garantías")) or "Auditoría garantías",
        safe_str(audit.get("dealer", "")),
        safe_str(audit.get("auditor", "")),
    )

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
    """
    Exportación analítica en Excel usando xlsxwriter.
    Requiere en requirements.txt: xlsxwriter
    """
    output = BytesIO()
    summary_df = build_summary_dataframe(claims)
    summary_export_df = summary_df.copy()
    if "% éxito" in summary_export_df.columns:
        summary_export_df["% éxito"] = pd.to_numeric(summary_export_df["% éxito"], errors="coerce") / 100
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
        summary_export_df.to_excel(writer, index=False, sheet_name="Claims")
        detail_df.to_excel(writer, index=False, sheet_name="Detalle apartados")
        improvement_df.to_excel(writer, index=False, sheet_name="Áreas de mejora")

        workbook = writer.book
        header_format = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        body_format = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
        percent_format = workbook.add_format({"num_format": "0.0%", "border": 1, "valign": "top"})
        integer_format = workbook.add_format({"num_format": "0", "border": 1, "valign": "top"})

        sheet_widths = {
            "Resumen auditoría": [22, 55],
            "Claims": [18, 24, 22, 20, 20, 14, 14, 16, 45, 60],
            "Detalle apartados": [18, 26, 30, 26, 18, 14, 14, 14, 45, 80],
            "Áreas de mejora": [18, 28, 32, 18, 18, 18, 60],
        }

        for sheet_name, worksheet in writer.sheets.items():
            df = {
                "Resumen auditoría": cover_df,
                "Claims": summary_export_df,
                "Detalle apartados": detail_df,
                "Áreas de mejora": improvement_df,
            }[sheet_name]

            rows, cols = df.shape
            worksheet.freeze_panes(1, 0)
            if rows >= 0 and cols > 0:
                worksheet.autofilter(0, 0, rows, cols - 1)

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            widths = sheet_widths.get(sheet_name, [])
            for col_num in range(cols):
                width = widths[col_num] if col_num < len(widths) else 18
                worksheet.set_column(col_num, col_num, width, body_format)

            if sheet_name == "Claims" and rows > 0:
                worksheet.set_column(3, 5, 16, integer_format)
                worksheet.set_column(6, 6, 14, percent_format)
                worksheet.conditional_format(1, 6, rows, 6, {
                    "type": "3_color_scale",
                    "min_color": "#F4CCCC",
                    "mid_color": "#FFF2CC",
                    "max_color": "#E2F0D9",
                })

    return output.getvalue()


def export_report_card_excel(claims: Dict[str, Dict[str, Any]], audit_name: str, dealer: str, auditor: str) -> bytes:
    """
    Genera un Excel tipo "boletín de notas" usando xlsxwriter.

    Hojas:
    - Boletín de notas
    - content
    - Claim document checklist I
    - Claim old parts checklist II
    - Improvement of claim issues III
    - Evaluation content

    Regla de puntuación:
    - Documentación = 58 puntos
    - Piezas viejas = 42 puntos
    - Total = 100 puntos
    - No aplica = máxima puntuación del apartado
    - Campañas = informativo, sin sumar ni restar
    """
    output = BytesIO()
    workbook = None

    try:
        import xlsxwriter  # noqa: F401
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})

        # ------------------------------------------------------------------
        # Formatos
        # ------------------------------------------------------------------
        fmt_title = workbook.add_format({
            "bold": True,
            "font_size": 14,
            "font_color": "#1F4E78",
            "valign": "vcenter",
            "text_wrap": True,
        })
        fmt_header = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        fmt_body = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
        fmt_bold = workbook.add_format({"bold": True, "border": 1, "valign": "top", "text_wrap": True})
        fmt_int = workbook.add_format({"border": 1, "valign": "top", "num_format": "0"})
        fmt_percent = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "0.0%", "bold": True})
        fmt_formula_percent = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "0.0%"})
        fmt_ok = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "bg_color": "#E2F0D9", "bold": True})
        fmt_mid = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "bg_color": "#FFF2CC", "bold": True})
        fmt_bad = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "bg_color": "#F4CCCC", "bold": True})
        fmt_gray = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "bg_color": "#E7E6E6"})

        def write_row(ws, row_idx, values, fmt=fmt_body):
            for col_idx, value in enumerate(values):
                ws.write(row_idx, col_idx, value, fmt)

        def write_headers(ws, row_idx, headers):
            ws.set_row(row_idx, 30)
            for col_idx, value in enumerate(headers):
                ws.write(row_idx, col_idx, value, fmt_header)

        def set_widths(ws, widths):
            for col_idx, width in enumerate(widths):
                ws.set_column(col_idx, col_idx, width)

        def result_label(points):
            try:
                p = float(points)
            except Exception:
                return "Pendiente"
            if p >= 90:
                return "Excelente"
            if p >= 80:
                return "Correcto"
            if p >= 60:
                return "Mejorable"
            return "Crítico"

        def result_format(points):
            try:
                p = float(points)
            except Exception:
                return fmt_gray
            if p >= 80:
                return fmt_ok
            if p >= 60:
                return fmt_mid
            return fmt_bad

        def evaluation_value(claim, check):
            evaluation = claim["evaluations"].get(check.key, {})
            points = evaluation.get("points")
            return "" if points is None else points

        def evaluation_status(claim, check):
            evaluation = claim["evaluations"].get(check.key, {})
            return evaluation.get("status", "")

        def evaluation_comment(claim, check):
            evaluation = claim["evaluations"].get(check.key, {})
            return evaluation.get("comment", "")

        # ------------------------------------------------------------------
        # Boletín de notas
        # ------------------------------------------------------------------
        grade_ws = workbook.add_worksheet("Boletín de notas")
        grade_ws.merge_range("A1:J1", "Boletín de notas - Warranty Audit", fmt_title)
        grade_headers = [
            "Claim No.", "Dealer", "VIN", "Documentación /58", "Piezas viejas /42",
            "Total /100", "% éxito", "Resultado", "Pendientes", "Comentarios",
        ]
        write_headers(grade_ws, 2, grade_headers)
        set_widths(grade_ws, [20, 24, 24, 18, 18, 14, 14, 16, 48, 60])
        grade_ws.freeze_panes(3, 0)

        row_idx = 3
        for claim in claims.values():
            score = calculate_claim_score(claim)
            total_points = score["total_points"]
            row_values = [
                claim["claim_no"],
                claim.get("dealer", dealer or ""),
                claim.get("vin", ""),
                score["doc_points"],
                score["old_points"],
                total_points,
                score["success_percent"] / 100,
                result_label(total_points),
                " | ".join(score["pending"]),
                claim.get("general_comment", ""),
            ]
            write_row(grade_ws, row_idx, row_values, fmt_body)
            grade_ws.write_number(row_idx, 3, score["doc_points"], fmt_int)
            grade_ws.write_number(row_idx, 4, score["old_points"], fmt_int)
            grade_ws.write_number(row_idx, 5, total_points, result_format(total_points))
            grade_ws.write_number(row_idx, 6, score["success_percent"] / 100, fmt_percent)
            grade_ws.write(row_idx, 7, result_label(total_points), result_format(total_points))
            row_idx += 1
        if row_idx > 3:
            grade_ws.autofilter(2, 0, row_idx - 1, len(grade_headers) - 1)
            grade_ws.conditional_format(3, 6, row_idx - 1, 6, {
                "type": "3_color_scale",
                "min_color": "#F4CCCC",
                "mid_color": "#FFF2CC",
                "max_color": "#E2F0D9",
            })

        # ------------------------------------------------------------------
        # content
        # ------------------------------------------------------------------
        content_ws = workbook.add_worksheet("content")
        content_ws.merge_range("A1:D1", "Warranty audit report card / Boletín de auditoría", fmt_title)
        content_rows = [
            ["Audit name", audit_name or ""],
            ["Dealer", dealer or ""],
            ["Auditor", auditor or ""],
            ["Export date", datetime.now().strftime("%Y-%m-%d %H:%M")],
            ["Scoring rule", "Claim document checklist I = 58 / Claim old parts checklist II = 42 / Total = 100"],
            ["N/A rule", "No aplica = maximum score of the section"],
            ["Campaigns", "Informative only. No points added or deducted."],
        ]
        set_widths(content_ws, [24, 90, 18, 18])
        for idx, row in enumerate(content_rows, start=2):
            content_ws.write(idx, 0, row[0], fmt_bold)
            content_ws.write(idx, 1, row[1], fmt_body)
        audit_score = calculate_audit_score(claims)
        content_ws.write(10, 0, "Global result", fmt_bold)
        content_ws.write(10, 1, f"{audit_score['success_percent']:.1f}% ({audit_score['total_points']}/{audit_score['max_points']})", result_format(audit_score["success_percent"]))

        # ------------------------------------------------------------------
        # Claim document checklist I
        # ------------------------------------------------------------------
        doc_ws = workbook.add_worksheet("Claim document checklist I")
        doc_headers = [
            "No.", "Claim No.",
            *[check.label for check in DOCUMENT_CHECKS],
            "Comprobación campañas", "Campañas pendientes",
            "Total documentación", "Resultado documentación %", "Comentarios",
        ]
        write_headers(doc_ws, 0, doc_headers)
        set_widths(doc_ws, [8, 20, 12, 18, 14, 12, 14, 12, 16, 18, 12, 24, 24, 18, 18, 50])
        doc_ws.freeze_panes(1, 2)
        for idx, claim in enumerate(claims.values(), start=1):
            excel_row = idx + 1
            row = [idx, claim["claim_no"]]
            row.extend(evaluation_value(claim, check) for check in DOCUMENT_CHECKS)
            row.append(claim["evaluations"].get("info_campaign_check", {}).get("comment", ""))
            row.append(claim["evaluations"].get("info_pending_campaigns", {}).get("comment", ""))
            row.append(f"=SUM(C{excel_row}:K{excel_row})")
            row.append(f"=L{excel_row}/{MAX_DOCUMENT_POINTS}")
            row.append(claim.get("general_comment", ""))
            write_row(doc_ws, idx, row, fmt_body)
            for c in range(2, 11):
                if isinstance(row[c], (int, float)):
                    doc_ws.write_number(idx, c, row[c], fmt_int)
            doc_ws.write_formula(idx, 13, f"=SUM(C{excel_row}:K{excel_row})", fmt_int)
            doc_ws.write_formula(idx, 14, f"=N{excel_row}/{MAX_DOCUMENT_POINTS}", fmt_formula_percent)
        if len(claims) > 0:
            doc_ws.autofilter(0, 0, len(claims), len(doc_headers) - 1)
            doc_ws.conditional_format(1, 14, len(claims), 14, {
                "type": "3_color_scale",
                "min_color": "#F4CCCC",
                "mid_color": "#FFF2CC",
                "max_color": "#E2F0D9",
            })

        # ------------------------------------------------------------------
        # Claim old parts checklist II
        # ------------------------------------------------------------------
        old_ws = workbook.add_worksheet("Claim old parts checklist II")
        old_headers = [
            "No.", "Claim No.",
            *[check.label for check in OLD_PARTS_CHECKS],
            "Total piezas viejas", "Resultado piezas viejas %", "Comentarios",
        ]
        write_headers(old_ws, 0, old_headers)
        set_widths(old_ws, [8, 20, 18, 20, 16, 24, 20, 26, 18, 18, 50])
        old_ws.freeze_panes(1, 2)
        for idx, claim in enumerate(claims.values(), start=1):
            excel_row = idx + 1
            row = [idx, claim["claim_no"]]
            row.extend(evaluation_value(claim, check) for check in OLD_PARTS_CHECKS)
            row.append(f"=SUM(C{excel_row}:H{excel_row})")
            row.append(f"=I{excel_row}/{MAX_OLD_PARTS_POINTS}")
            row.append(claim.get("general_comment", ""))
            write_row(old_ws, idx, row, fmt_body)
            for c in range(2, 8):
                if isinstance(row[c], (int, float)):
                    old_ws.write_number(idx, c, row[c], fmt_int)
            old_ws.write_formula(idx, 8, f"=SUM(C{excel_row}:H{excel_row})", fmt_int)
            old_ws.write_formula(idx, 9, f"=I{excel_row}/{MAX_OLD_PARTS_POINTS}", fmt_formula_percent)
        if len(claims) > 0:
            old_ws.autofilter(0, 0, len(claims), len(old_headers) - 1)
            old_ws.conditional_format(1, 9, len(claims), 9, {
                "type": "3_color_scale",
                "min_color": "#F4CCCC",
                "mid_color": "#FFF2CC",
                "max_color": "#E2F0D9",
            })

        # ------------------------------------------------------------------
        # Improvement of claim issues III
        # ------------------------------------------------------------------
        imp_ws = workbook.add_worksheet("Improvement of claim issues III")
        imp_headers = ["Claim No.", "Parameter / auditable area", "Exception / comments", "Observations", "Countermeasure"]
        write_headers(imp_ws, 0, imp_headers)
        set_widths(imp_ws, [20, 32, 32, 60, 60])
        imp_ws.freeze_panes(1, 0)
        imp_row = 1
        for claim in claims.values():
            score = calculate_claim_score(claim)
            for block, area, lost, max_points, status in score["lost_by_area"]:
                comments = []
                check = next((item for item in ALL_SCORING_CHECKS if item.block == block and item.label == area), None)
                if check is not None:
                    section_comment = evaluation_comment(claim, check)
                    if section_comment:
                        comments.append(section_comment)
                if claim.get("general_comment", ""):
                    comments.append(claim.get("general_comment", ""))
                write_row(imp_ws, imp_row, [
                    claim["claim_no"],
                    area,
                    f"{status}: {lost}/{max_points} puntos perdidos",
                    "\n".join(comments),
                    "Reforzar el cumplimiento del criterio y revisar la documentación antes del envío de la claim.",
                ], fmt_body)
                imp_row += 1
        if imp_row == 1:
            write_row(imp_ws, imp_row, ["", "Sin desviaciones puntuables", "", "", ""], fmt_body)
            imp_row += 1
        imp_ws.autofilter(0, 0, imp_row - 1, len(imp_headers) - 1)

        # ------------------------------------------------------------------
        # Evaluation content
        # ------------------------------------------------------------------
        eval_ws = workbook.add_worksheet("Evaluation content")
        eval_headers = ["Checklist", "Apartado", "Máximo", "Opciones", "Criterio"]
        write_headers(eval_ws, 0, eval_headers)
        set_widths(eval_ws, [28, 34, 12, 42, 85])
        eval_ws.freeze_panes(1, 0)
        for row_idx, check in enumerate(ALL_CHECKS, start=1):
            options_text = " | ".join(
                f"{option.status}: {'-' if option.points is None else option.points}"
                for option in check.options
            )
            write_row(eval_ws, row_idx, [check.block, check.label, check.max_points, options_text, check.guidance], fmt_body)
        eval_ws.autofilter(0, 0, len(ALL_CHECKS), len(eval_headers) - 1)

        workbook.close()
        workbook = None
        return output.getvalue()

    finally:
        if workbook is not None:
            workbook.close()

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
    if "audit_name" not in st.session_state:
        st.session_state.audit_name = "Auditoría garantías"
    if "audit_dealer" not in st.session_state:
        st.session_state.audit_dealer = ""
    if "audit_auditor" not in st.session_state:
        st.session_state.audit_auditor = ""


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


def display_value(value: Any, fallback: str = "No informado") -> str:
    value = safe_str(value)
    return value if value else fallback


def render_claim_quick_card(claim: Dict[str, Any], default_dealer: str = ""):
    """Mantiene internamente el dealer si viene informado, pero no muestra campos extra.

    La lista de HQ suele traer solo el número de claim, así que VIN/importe/modelo
    no se muestran para mantener la revisión rápida y limpia.
    """
    if not claim.get("dealer") and default_dealer:
        claim["dealer"] = default_dealer

def main():
    st.set_page_config(page_title="Warranty Audit Portal", layout="wide")
    init_state()

    st.title("Warranty Audit Portal")
    st.caption("Auditoría online de garantías: documentación + piezas viejas = 100 puntos. Campañas solo informativo.")

    with st.sidebar:
        st.header("Auditoría")

        st.subheader("Guardar / cargar progreso")
        workfile = st.file_uploader(
            "Cargar auditoría guardada (.json)",
            type=["json"],
            key="workfile_upload",
            help="Carga un archivo generado con el botón 'Guardar auditoría de trabajo'.",
        )
        if workfile is not None and st.button("Cargar auditoría guardada", type="primary"):
            try:
                loaded_claims, loaded_audit_name, loaded_dealer, loaded_auditor = load_audit_workfile(workfile)
                st.session_state.claims = loaded_claims
                st.session_state.selected_claim = next(iter(loaded_claims.keys()))
                st.session_state.audit_name = loaded_audit_name
                st.session_state.audit_dealer = loaded_dealer
                st.session_state.audit_auditor = loaded_auditor
                st.success(f"Auditoría cargada: {len(loaded_claims)} claims.")
                st.rerun()
            except Exception as exc:
                st.error(f"No se pudo cargar la auditoría guardada: {exc}")

        if st.session_state.claims:
            st.download_button(
                "Guardar auditoría de trabajo (.json)",
                data=serialize_audit_workfile(
                    st.session_state.claims,
                    st.session_state.audit_name,
                    st.session_state.audit_dealer,
                    st.session_state.audit_auditor,
                ),
                file_name=f"audit_workfile_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                help="Guarda todo el progreso: claims, estados, N/A, comentarios y datos de cabecera.",
            )

        st.divider()
        audit_name = st.text_input("Nombre auditoría", key="audit_name")
        dealer = st.text_input("Dealer", key="audit_dealer")
        auditor = st.text_input("Auditor", key="audit_auditor")

        st.divider()
        uploaded_file = st.file_uploader("Subir checklist o lista HQ", type=["xlsx", "xlsm", "xls"], key="claims_upload")
        if uploaded_file is not None and st.button("Cargar claims", type="secondary"):
            try:
                st.session_state.claims = read_claims_from_uploaded_excel(uploaded_file)
                if st.session_state.claims:
                    # Si la checklist oficial no trae dealer por claim, usamos el dealer general
                    # indicado en la barra lateral para que no quede la ficha en blanco.
                    default_dealer = safe_str(st.session_state.get("audit_dealer", ""))
                    if default_dealer:
                        for loaded_claim in st.session_state.claims.values():
                            if not safe_str(loaded_claim.get("dealer", "")):
                                loaded_claim["dealer"] = default_dealer

                    st.session_state.selected_claim = next(iter(st.session_state.claims.keys()))
                    st.success(f"Cargadas {len(st.session_state.claims)} claims.")
                    st.rerun()
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
            st.rerun()

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
        compact_columns = [
            "Claim No.",
            "Puntos documentación",
            "Puntos piezas viejas",
            "Resultado /100",
            "Estado",
        ]
        st.dataframe(
            summary_df[compact_columns],
            use_container_width=True,
            hide_index=True,
        )

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

        st.download_button(
            "Exportar boletín de notas Excel",
            data=export_report_card_excel(claims, audit_name, dealer, auditor),
            file_name=f"audit_boletin_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Genera un Excel con hojas tipo plantilla: documentación, piezas viejas, áreas de mejora, criterios y boletín por claim.",
        )

    with right:
        claim = claims[st.session_state.selected_claim]
        score = calculate_claim_score(claim)

        st.subheader(f"Revisión claim {claim['claim_no']}")
        render_claim_quick_card(claim, dealer or "")

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

