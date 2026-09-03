# -*- coding: utf-8 -*-
"""
Warranty Internal Audit Tool
Streamlit app interna para digitalizar auditorías de garantías.

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
- El comentario general puede generarse desde las observaciones de cada apartado.
- Permite completar ficha por claim: dealer, VIN, modelo e importe.
- Dealer se selecciona desde listado de dealers activos.
- Los archivos exportados usan el nombre base Dealer_fecha_auditor.
- Permite adjuntar fotos de piezas viejas y certificado de destrucción por claim.
- Añade estado de trabajo por claim y validación de cierre antes de exportar.
"""

from __future__ import annotations

import base64
import json
import re
import zipfile
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

OLD_PART_PHOTO_FILE_TYPES = ["jpg", "jpeg", "png", "webp"]
DESTRUCTION_CERTIFICATE_FILE_TYPES = ["jpg", "jpeg", "png", "webp", "pdf"]

CLAIM_WORKFLOW_STATUSES: List[str] = [
    "Pendiente",
    "En revisión",
    "Completada",
    "Requiere aclaración",
    "Cerrada",
]

AUDIT_MODES: List[str] = [
    "NSC / Auditor interno",
    "Dealer / Autoauditoría (fase futura)",
]

ACTIVE_DEALERS: List[str] = [
    "ACAI MOTOR MÁLAGA",
    "ALFAVISA BILBAO",
    "ALIMOTOR ELCHE",
    "ANFERPA SEGOVIA",
    "AUTO YALDE CALAHORRA",
    "AUTO YALDE LOGROÑO",
    "AUTOCAM MOTOR VILAFRANCA",
    "AUTOCAM VILANOVA",
    "AUTOCYL PALENCIA",
    "AUTOCYL VALLADOLID",
    "AUTOVIDAL PALMA DE MALLORCA",
    "AVANTI GRANADA",
    "AXIS MOTORS",
    "BLENDIO LAREDO",
    "BLENDIO LUGO",
    "BLENDIO OURENSE",
    "BLENDIO OVIEDO",
    "BLENDIO SANTANDER",
    "BLENDIO TORRELAVEGA",
    "BORJAMOTOR ALICANTE",
    "CERVERA AVILA",
    "CERVERA SALAMANCA",
    "CHINARES GUADALAJARA",
    "DILOAUTOJAEN",
    "DUMOSA BENAVENTE",
    "ESLAUTO LEON",
    "FIMALAGA MÁLAGA",
    "FIMALAGA MARBELLA",
    "GRUP BASOLS IGUALADA",
    "GRUPO JULIAN BURGOS",
    "GRUPO NIETO MÁLAGA",
    "GRUPO NIETO MARBELLA",
    "HIMASA SEDAVÍ",
    "JEMOYA SORIA",
    "JOVERAUTO MELILLA",
    "LASACAR MIRANDA DE EBRO",
    "LASACAR VITORIA",
    "LEPAS AUTOCAM VILANOVA",
    "LEPAS AUTOVIVO SANT BOI",
    "LEPAS BASOLS IGUALADA",
    "LEPAS BASOLS VIC",
    "LEPAS GAMBOA MAJADAHONDA",
    "LEPAS JULIÁN BURGOS",
    "LEPAS MONECAR SAGUNTO",
    "LEPAS PREMIER VITORIA",
    "LEPAS RAFAEL AFONSO LAS PALMAS",
    "LEPAS RESNOVA CORUÑA",
    "LEPAS RESNOVA VIGO",
    "LEPAS TECNOTARRACO TARRAGONA",
    "LEPAS TUMASA HUESCA",
    "LEPAS VALLESCAR SABADELL",
    "LEPAS VALLESCAR TERRASSA",
    "LEPAS ZEN MOTOR GIPUZKOA",
    "LEPAS ZEN MOTOR ZARAGOZA",
    "M AUTOMOCIÓN ALCALÁ",
    "M AUTOMOCIÓN BCN (GRAN VÍA)",
    "M AUTOMOCIÓN BCN GUAYAQUIL",
    "M AUTOMOCIÓN CASTELLÓN",
    "M AUTOMOCIÓN GERONA",
    "M AUTOMOCIÓN MATARÓ",
    "M TECNIK ALCALÁ DE HENARES",
    "M TECNIK BARCELONA MAQUINISTA",
    "M TECNIK CASTELLÓN",
    "M TECNIK FIGUERES",
    "M TECNIK GERONA",
    "M TECNIK MATARÓ",
    "M TECNIK VINAROZ",
    "MARTIN LIZAGA TERUEL",
    "MAS AUTO LEGANÉS",
    "MAVEN BADAJOZ",
    "MAVEN CÁCERES",
    "MAVEN DON BENITO",
    "MAVEN MÉRIDA",
    "MAVEN PLASENCIA",
    "MOLL MOTOR DENIA",
    "MOLL MOTOR GANDIA",
    "MOLL VALENCIA",
    "MONECAR CUENCA",
    "MOTOR NACIENTE LEGANÉS",
    "MOVINSUR GRANADA",
    "MOVINSUR JAÉN",
    "MOVINSUR MOTRIL",
    "MY CARS CÓRDOBA",
    "NOVACAR BCN SANT BOI",
    "PALAUSA ZAMORA",
    "PROCHERY ALBACETE",
    "PROCHERY CARTAGENA",
    "PROCHERY MURCIA",
    "PRUNA CAR GO GRANOLLERS",
    "RAFAEL AFONSO AGUIMES",
    "RAFAEL AFONSO LANZAROTE",
    "RAFAEL AFONSO LAS PALMAS",
    "RAFAEL AFONSO TENERIFE",
    "RESNOVA MOTOR CORUÑA",
    "RESNOVA MOTOR GIJÓN",
    "RESNOVA MOTOR NARÓN",
    "RESNOVA MOTOR OVIEDO",
    "RESNOVA MOTOR SANTIAGO",
    "RESNOVA MOTOR VIGO",
    "SEGRE LLEIDA",
    "SEGRE MOTORS LERIDA",
    "SERTECAUTO PONFERRADA",
    "SYRSA ALGECIRAS",
    "SYRSA ALMERIA",
    "SYRSA EJIDO",
    "SYRSA HUELVA",
    "SYRSA SEVILLA",
    "TALAUTO CAZALEGAS",
    "TALAUTO TOLEDO",
    "TALLERES CHINARES",
    "TECNOTARRACO TARRAGONA",
    "TERRY MOBILITY JERÉZ",
    "TRADECAR GAMBOA ALCORCÓN",
    "TRADECAR GAMBOA MADRID",
    "TRADECAR GAMBOA MAJADAHONDA",
    "TRADECAR GAMBOA RIVAS",
    "TUMASA HUESCA",
    "TUMASA MONZÓN",
    "UNIONE ALCAZAR DE SAN JUAN",
    "UNIONE CIUDAD REAL",
    "VALLESCAR SABADELL",
    "VALLESCAR TERRASSA",
    "VIAN ALCORCÓN",
    "VIAN AUTOMOBILE VILLALBA",
    "VIAN MÓSTOLES",
    "VIAN NAVARRA",
    "ZEN MOTOR OLABERRIA",
    "ZEN MOTOR PAMPLONA",
    "ZEN MOTOR SAN SEBASTIÁN",
    "ZEN MOTOR ZARAGOZA",
]


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


def sanitize_for_filename(value: Any, fallback: str = "item") -> str:
    """Limpia nombres para usarlos en ZIP/descargas sin romper rutas."""
    text = safe_str(value) or fallback
    text = re.sub(r"[^\w\-. ]+", "_", text, flags=re.UNICODE).strip()
    text = re.sub(r"\s+", "_", text)
    return text or fallback


def get_dealer_options(current_value: str = "") -> List[str]:
    """Lista de dealers para selectbox, conservando valores cargados de JSON antiguos."""
    current_value = safe_str(current_value)
    options = [""] + ACTIVE_DEALERS.copy()

    if current_value and current_value not in options:
        options.insert(1, current_value)

    return options


def format_dealer_option(value: str) -> str:
    return "Selecciona dealer" if not safe_str(value) else value


def build_audit_file_basename(dealer: str, auditor: str, when: Optional[datetime] = None) -> str:
    """Base común para todos los archivos exportados: Dealer_fecha_auditor."""
    when = when or datetime.now()
    dealer_part = sanitize_for_filename(dealer, "Dealer")
    date_part = when.strftime("%Y%m%d")
    auditor_part = sanitize_for_filename(auditor, "Auditor")
    return f"{dealer_part}_{date_part}_{auditor_part}"


def apply_default_dealer_to_blank_claims(claims: Dict[str, Dict[str, Any]], dealer: str) -> None:
    """Aplica el dealer general a claims sin dealer propio, sin pisar valores ya informados."""
    dealer = safe_str(dealer)
    if not dealer:
        return

    for claim in claims.values():
        if not safe_str(claim.get("dealer", "")):
            claim["dealer"] = dealer


def normalize_workflow_status(value: Any) -> str:
    """Normaliza el estado manual de trabajo de una claim."""
    status = safe_str(value)
    return status if status in CLAIM_WORKFLOW_STATUSES else "Pendiente"


def get_claim_workflow_status(claim: Dict[str, Any]) -> str:
    """Estado visible de la claim: manual si existe, si no derivado de la puntuación."""
    manual_status = normalize_workflow_status(claim.get("workflow_status", "Pendiente"))

    # Si el usuario no ha marcado nada manualmente y la claim está completa, mostramos Completada.
    if manual_status == "Pendiente":
        try:
            if calculate_claim_score(claim).get("completed"):
                return "Completada"
        except Exception:
            pass

    return manual_status


def set_claim_workflow_status(claim: Dict[str, Any], status: str) -> None:
    """Actualiza el estado manual de la claim."""
    claim["workflow_status"] = normalize_workflow_status(status)


def uploaded_file_to_attachment(uploaded_file) -> Dict[str, Any]:
    """Convierte un UploadedFile de Streamlit en un registro serializable en JSON."""
    data = uploaded_file.getvalue()
    return {
        "name": safe_str(getattr(uploaded_file, "name", "archivo")) or "archivo",
        "type": safe_str(getattr(uploaded_file, "type", "")),
        "size": len(data),
        "uploaded_at": datetime.now().isoformat(timespec="seconds"),
        "data_base64": base64.b64encode(data).decode("ascii"),
    }


def normalize_loaded_attachment(raw_attachment: Any) -> Optional[Dict[str, Any]]:
    """Normaliza adjuntos cargados desde JSON y evita registros corruptos."""
    if not isinstance(raw_attachment, dict):
        return None

    name = safe_str(raw_attachment.get("name", ""))
    data_base64 = safe_str(raw_attachment.get("data_base64", ""))

    if not name and not data_base64:
        return None

    try:
        size = int(raw_attachment.get("size", 0) or 0)
    except Exception:
        size = 0

    return {
        "name": name or "archivo",
        "type": safe_str(raw_attachment.get("type", "")),
        "size": size,
        "uploaded_at": safe_str(raw_attachment.get("uploaded_at", "")),
        "data_base64": data_base64,
    }


def normalize_claim_attachments(raw_attachments: Any) -> Dict[str, Any]:
    """Estructura estable de adjuntos por claim."""
    attachments = {
        "old_parts_photos": [],
        "destruction_certificate": None,
    }

    if not isinstance(raw_attachments, dict):
        return attachments

    raw_photos = raw_attachments.get("old_parts_photos", [])
    if isinstance(raw_photos, list):
        attachments["old_parts_photos"] = [
            item for item in (normalize_loaded_attachment(photo) for photo in raw_photos)
            if item is not None
        ]

    certificate = normalize_loaded_attachment(raw_attachments.get("destruction_certificate"))
    attachments["destruction_certificate"] = certificate

    return attachments


def get_claim_attachments(claim: Dict[str, Any]) -> Dict[str, Any]:
    """Garantiza que la claim tenga estructura de adjuntos aunque venga de versiones antiguas."""
    claim["attachments"] = normalize_claim_attachments(claim.get("attachments", {}))
    return claim["attachments"]


def attachment_bytes(attachment: Dict[str, Any]) -> bytes:
    data_base64 = safe_str(attachment.get("data_base64", ""))
    if not data_base64:
        return b""

    try:
        return base64.b64decode(data_base64.encode("ascii"))
    except Exception:
        return b""


def attachment_names_summary(attachments: List[Dict[str, Any]]) -> str:
    if not attachments:
        return ""
    names = [safe_str(item.get("name", "archivo")) or "archivo" for item in attachments]
    return f"{len(names)} archivo(s): " + " | ".join(names)


def old_parts_photos_summary(claim: Dict[str, Any]) -> str:
    attachments = get_claim_attachments(claim)
    return attachment_names_summary(attachments.get("old_parts_photos", []))


def destruction_certificate_summary(claim: Dict[str, Any]) -> str:
    attachments = get_claim_attachments(claim)
    certificate = attachments.get("destruction_certificate")
    if not certificate:
        return ""
    return safe_str(certificate.get("name", "Certificado adjunto")) or "Certificado adjunto"


def build_attachments_dataframe(claims: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for claim in claims.values():
        attachments = get_claim_attachments(claim)
        for index, photo in enumerate(attachments.get("old_parts_photos", []), start=1):
            rows.append({
                "Claim No.": claim.get("claim_no", ""),
                "Tipo adjunto": "Foto pieza vieja",
                "Nº": index,
                "Archivo": safe_str(photo.get("name", "")),
                "Formato": safe_str(photo.get("type", "")),
                "Tamaño bytes": photo.get("size", 0),
                "Subido en": safe_str(photo.get("uploaded_at", "")),
            })
        certificate = attachments.get("destruction_certificate")
        if certificate:
            rows.append({
                "Claim No.": claim.get("claim_no", ""),
                "Tipo adjunto": "Certificado destrucción",
                "Nº": 1,
                "Archivo": safe_str(certificate.get("name", "")),
                "Formato": safe_str(certificate.get("type", "")),
                "Tamaño bytes": certificate.get("size", 0),
                "Subido en": safe_str(certificate.get("uploaded_at", "")),
            })
    return pd.DataFrame(rows)


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
        "workflow_status": "Pendiente",
        "nsc_comment": "",
        "general_comment": "",
        "evaluations": {check.key: new_evaluation(check) for check in ALL_CHECKS},
        "attachments": {
            "old_parts_photos": [],
            "destruction_certificate": None,
        },
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
        "nsc_comment",
        "general_comment",
    ]:
        claim[field] = safe_str(raw_claim.get(field, claim.get(field, "")))

    claim["workflow_status"] = normalize_workflow_status(raw_claim.get("workflow_status", claim.get("workflow_status", "Pendiente")))

    raw_evaluations = raw_claim.get("evaluations", {})
    if isinstance(raw_evaluations, dict):
        for check in ALL_CHECKS:
            if check.key in raw_evaluations:
                claim["evaluations"][check.key] = normalize_loaded_evaluation(check, raw_evaluations[check.key])

    claim["attachments"] = normalize_claim_attachments(raw_claim.get("attachments", {}))

    return claim


def serialize_audit_workfile(
    claims: Dict[str, Dict[str, Any]],
    audit_name: str,
    dealer: str,
    auditor: str,
    audit_mode: str = "",
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
            "audit_mode": (audit_mode or (st.session_state.get("audit_mode", AUDIT_MODES[0]) if "st" in globals() else AUDIT_MODES[0])),
        },
        "claims": claims,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def load_audit_workfile(uploaded_file) -> Tuple[Dict[str, Dict[str, Any]], str, str, str, str]:
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

    loaded_mode = safe_str(audit.get("audit_mode", AUDIT_MODES[0])) or AUDIT_MODES[0]
    if loaded_mode not in AUDIT_MODES:
        loaded_mode = AUDIT_MODES[0]

    return (
        claims,
        safe_str(audit.get("audit_name", "Auditoría garantías")) or "Auditoría garantías",
        safe_str(audit.get("dealer", "")),
        safe_str(audit.get("auditor", "")),
        loaded_mode,
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
            "Modelo": claim.get("model", ""),
            "Importe": claim.get("amount", ""),
            "Puntos documentación": score["doc_points"],
            "Puntos piezas viejas": score["old_points"],
            "Resultado /100": score["total_points"],
            "% éxito": score["success_percent"],
            "Estado": get_claim_workflow_status(claim),
            "Pendientes": " | ".join(score["pending"]),
            "Comentario NSC": claim.get("nsc_comment", ""),
            "Comentarios": claim.get("general_comment", ""),
            "Fotos piezas viejas": old_parts_photos_summary(claim),
            "Certificado destrucción": destruction_certificate_summary(claim),
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
                "Dealer": claim.get("dealer", ""),
                "VIN": claim.get("vin", ""),
                "Modelo": claim.get("model", ""),
                "Importe": claim.get("amount", ""),
                "Estado trabajo claim": get_claim_workflow_status(claim),
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
                "Dealer": claim.get("dealer", ""),
                "VIN": claim.get("vin", ""),
                "Modelo": claim.get("model", ""),
                "Importe": claim.get("amount", ""),
                "Bloque": block,
                "Área de mejora": area,
                "Estado": status,
                "Puntos perdidos": lost,
                "Máximo apartado": max_points,
                "Comentario NSC": claim.get("nsc_comment", ""),
                "Comentario claim": claim.get("general_comment", ""),
            })
    return pd.DataFrame(rows)



def old_parts_evidence_required(claim: Dict[str, Any]) -> bool:
    """Decide si hay que pedir evidencias de piezas viejas para la claim.

    Si todos los apartados de piezas viejas están en No aplica o Pendiente, no forzamos
    fotos/certificado. Si el auditor empieza a puntuar piezas viejas como OK/Parcial/NOK
    o ya ha subido adjuntos, activamos la validación.
    """
    attachments = get_claim_attachments(claim)
    if attachments.get("old_parts_photos") or attachments.get("destruction_certificate"):
        return True

    for check in OLD_PARTS_CHECKS:
        evaluation = claim.get("evaluations", {}).get(check.key, {})
        status = safe_str(evaluation.get("status", ""))
        if status not in ("", "Pendiente", "N/A"):
            return True

    return False


def destruction_certificate_required(claim: Dict[str, Any]) -> bool:
    evaluation = claim.get("evaluations", {}).get("old_destruction_certificate", {})
    status = safe_str(evaluation.get("status", ""))
    return status not in ("", "Pendiente", "N/A")


def build_closing_validation_dataframe(claims: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
    """Checklist de cierre: avisa de pendientes antes de dar la auditoría por cerrada."""
    rows = []

    for claim in claims.values():
        score = calculate_claim_score(claim)
        claim_no = claim.get("claim_no", "")
        workflow_status = get_claim_workflow_status(claim)

        if score["pending"]:
            rows.append({
                "Claim No.": claim_no,
                "Tipo aviso": "Puntuación pendiente",
                "Detalle": "Apartados sin revisar: " + " | ".join(score["pending"]),
                "Severidad": "Alta",
                "Estado claim": workflow_status,
            })

        if score["lost_by_area"] and not safe_str(claim.get("general_comment", "")):
            rows.append({
                "Claim No.": claim_no,
                "Tipo aviso": "Comentario general vacío",
                "Detalle": "Hay puntos perdidos, pero no hay comentario general generado/escrito.",
                "Severidad": "Media",
                "Estado claim": workflow_status,
            })

        attachments = get_claim_attachments(claim)
        photos_count = len(attachments.get("old_parts_photos", []))
        if old_parts_evidence_required(claim) and photos_count < 3:
            rows.append({
                "Claim No.": claim_no,
                "Tipo aviso": "Fotos piezas viejas",
                "Detalle": f"Hay {photos_count} foto(s) adjunta(s). Recomendado mínimo: 3.",
                "Severidad": "Media",
                "Estado claim": workflow_status,
            })

        if destruction_certificate_required(claim) and not attachments.get("destruction_certificate"):
            rows.append({
                "Claim No.": claim_no,
                "Tipo aviso": "Certificado destrucción",
                "Detalle": "El apartado de certificado no está marcado como N/A/Pendiente, pero no hay archivo adjunto.",
                "Severidad": "Media",
                "Estado claim": workflow_status,
            })

        if workflow_status in ("Pendiente", "En revisión") and score["completed"]:
            rows.append({
                "Claim No.": claim_no,
                "Tipo aviso": "Estado de trabajo",
                "Detalle": "La puntuación está completa, pero la claim no está marcada como Completada/Cerrada.",
                "Severidad": "Baja",
                "Estado claim": workflow_status,
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
    attachments_df = build_attachments_dataframe(claims)
    closing_df = build_closing_validation_dataframe(claims)
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
        attachments_df.to_excel(writer, index=False, sheet_name="Adjuntos")
        closing_df.to_excel(writer, index=False, sheet_name="Validación cierre")

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
            "Claims": [18, 24, 22, 18, 14, 20, 20, 14, 14, 16, 45, 50, 60, 42, 34],
            "Detalle apartados": [18, 24, 22, 18, 14, 18, 28, 30, 26, 18, 14, 14, 45, 80],
            "Áreas de mejora": [18, 24, 22, 18, 14, 28, 32, 18, 18, 18, 50, 60],
            "Adjuntos": [18, 24, 8, 42, 24, 16, 22],
            "Validación cierre": [18, 26, 80, 14, 18],
        }

        for sheet_name, worksheet in writer.sheets.items():
            df = {
                "Resumen auditoría": cover_df,
                "Claims": summary_export_df,
                "Detalle apartados": detail_df,
                "Áreas de mejora": improvement_df,
                "Adjuntos": attachments_df,
                "Validación cierre": closing_df,
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
                for points_column in ["Puntos documentación", "Puntos piezas viejas", "Resultado /100"]:
                    if points_column in df.columns:
                        col_idx = df.columns.get_loc(points_column)
                        worksheet.set_column(col_idx, col_idx, 16, integer_format)
                if "% éxito" in df.columns:
                    percent_col = df.columns.get_loc("% éxito")
                    worksheet.set_column(percent_col, percent_col, 14, percent_format)
                    worksheet.conditional_format(1, percent_col, rows, percent_col, {
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
            "Claim No.", "Dealer", "VIN", "Modelo", "Importe",
            "Documentación /58", "Piezas viejas /42", "Total /100", "% éxito",
            "Resultado", "Estado trabajo", "Pendientes", "Comentario NSC", "Comentarios", "Fotos piezas viejas", "Certificado destrucción",
        ]
        write_headers(grade_ws, 2, grade_headers)
        set_widths(grade_ws, [20, 24, 24, 18, 14, 18, 18, 14, 14, 16, 18, 48, 50, 60, 42, 34])
        grade_ws.freeze_panes(3, 0)

        row_idx = 3
        for claim in claims.values():
            score = calculate_claim_score(claim)
            total_points = score["total_points"]
            row_values = [
                claim["claim_no"],
                claim.get("dealer", dealer or ""),
                claim.get("vin", ""),
                claim.get("model", ""),
                claim.get("amount", ""),
                score["doc_points"],
                score["old_points"],
                total_points,
                score["success_percent"] / 100,
                result_label(total_points),
                get_claim_workflow_status(claim),
                " | ".join(score["pending"]),
                claim.get("nsc_comment", ""),
                claim.get("general_comment", ""),
                old_parts_photos_summary(claim),
                destruction_certificate_summary(claim),
            ]
            write_row(grade_ws, row_idx, row_values, fmt_body)
            grade_ws.write_number(row_idx, 5, score["doc_points"], fmt_int)
            grade_ws.write_number(row_idx, 6, score["old_points"], fmt_int)
            grade_ws.write_number(row_idx, 7, total_points, result_format(total_points))
            grade_ws.write_number(row_idx, 8, score["success_percent"] / 100, fmt_percent)
            grade_ws.write(row_idx, 9, result_label(total_points), result_format(total_points))
            row_idx += 1
        if row_idx > 3:
            grade_ws.autofilter(2, 0, row_idx - 1, len(grade_headers) - 1)
            grade_ws.conditional_format(3, 8, row_idx - 1, 8, {
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
            "Fotos piezas viejas", "Certificado destrucción",
        ]
        write_headers(old_ws, 0, old_headers)
        set_widths(old_ws, [8, 20, 18, 20, 16, 24, 20, 26, 18, 18, 50, 42, 34])
        old_ws.freeze_panes(1, 2)
        for idx, claim in enumerate(claims.values(), start=1):
            excel_row = idx + 1
            row = [idx, claim["claim_no"]]
            row.extend(evaluation_value(claim, check) for check in OLD_PARTS_CHECKS)
            row.append(f"=SUM(C{excel_row}:H{excel_row})")
            row.append(f"=I{excel_row}/{MAX_OLD_PARTS_POINTS}")
            row.append(claim.get("general_comment", ""))
            row.append(old_parts_photos_summary(claim))
            row.append(destruction_certificate_summary(claim))
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
        # Adjuntos piezas viejas
        # ------------------------------------------------------------------
        attachments_ws = workbook.add_worksheet("Adjuntos")
        attachments_headers = ["Claim No.", "Tipo adjunto", "Nº", "Archivo", "Formato", "Tamaño bytes", "Subido en"]
        write_headers(attachments_ws, 0, attachments_headers)
        set_widths(attachments_ws, [20, 26, 8, 42, 24, 16, 22])
        attachments_ws.freeze_panes(1, 0)
        attachments_df = build_attachments_dataframe(claims)
        if attachments_df.empty:
            write_row(attachments_ws, 1, ["", "Sin adjuntos", "", "", "", "", ""], fmt_body)
            attachments_ws.autofilter(0, 0, 1, len(attachments_headers) - 1)
        else:
            for row_idx, (_, attachment_row) in enumerate(attachments_df.iterrows(), start=1):
                write_row(attachments_ws, row_idx, [attachment_row.get(header, "") for header in attachments_headers], fmt_body)
            attachments_ws.autofilter(0, 0, len(attachments_df), len(attachments_headers) - 1)

        # ------------------------------------------------------------------
        # Validación de cierre
        # ------------------------------------------------------------------
        closing_ws = workbook.add_worksheet("Validación cierre")
        closing_headers = ["Claim No.", "Tipo aviso", "Detalle", "Severidad", "Estado claim"]
        write_headers(closing_ws, 0, closing_headers)
        set_widths(closing_ws, [20, 28, 85, 14, 20])
        closing_ws.freeze_panes(1, 0)
        closing_df = build_closing_validation_dataframe(claims)
        if closing_df.empty:
            write_row(closing_ws, 1, ["", "Sin avisos", "La auditoría no tiene avisos de cierre.", "OK", ""], fmt_body)
            closing_ws.autofilter(0, 0, 1, len(closing_headers) - 1)
        else:
            for row_idx, (_, closing_row) in enumerate(closing_df.iterrows(), start=1):
                write_row(closing_ws, row_idx, [closing_row.get(header, "") for header in closing_headers], fmt_body)
            closing_ws.autofilter(0, 0, len(closing_df), len(closing_headers) - 1)

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
    closing_df = build_closing_validation_dataframe(claims)
    lines.append("Validación de cierre")
    if closing_df.empty:
        lines.append("No hay avisos de cierre pendientes.")
    else:
        high = len(closing_df[closing_df["Severidad"] == "Alta"]) if "Severidad" in closing_df.columns else 0
        medium = len(closing_df[closing_df["Severidad"] == "Media"]) if "Severidad" in closing_df.columns else 0
        low = len(closing_df[closing_df["Severidad"] == "Baja"]) if "Severidad" in closing_df.columns else 0
        lines.append(f"Avisos detectados: {len(closing_df)} (alta: {high}, media: {medium}, baja: {low}).")
        for _, row in closing_df.head(8).iterrows():
            lines.append(f"- {row['Claim No.']} · {row['Tipo aviso']}: {row['Detalle']}")
    lines.append("")
    lines.append("Conclusión")
    lines.append(
        "Se recomienda focalizar el plan de mejora en las áreas con mayor pérdida de puntos, "
        "reforzando la calidad documental, la justificación técnica y la trazabilidad de piezas viejas cuando aplique."
    )
    return "\n".join(lines)


def export_audit_package_zip(claims: Dict[str, Dict[str, Any]], audit_name: str, dealer: str, auditor: str) -> bytes:
    """Exporta un paquete completo con Excel, boletín, informe, JSON de trabajo y adjuntos reales."""
    output = BytesIO()
    base_name = build_audit_file_basename(dealer, auditor)

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(
            f"{base_name}/{base_name}.json",
            serialize_audit_workfile(claims, audit_name, dealer, auditor, st.session_state.get("audit_mode", AUDIT_MODES[0])),
        )
        zip_file.writestr(
            f"{base_name}/{base_name}_analitico.xlsx",
            export_excel(claims, audit_name, dealer, auditor),
        )
        zip_file.writestr(
            f"{base_name}/{base_name}_boletin.xlsx",
            export_report_card_excel(claims, audit_name, dealer, auditor),
        )
        zip_file.writestr(
            f"{base_name}/{base_name}_informe.txt",
            generate_text_report(claims, audit_name, dealer, auditor).encode("utf-8"),
        )

        for claim in claims.values():
            claim_folder = sanitize_for_filename(claim.get("claim_no", "claim"), "claim")
            attachments = get_claim_attachments(claim)

            for index, photo in enumerate(attachments.get("old_parts_photos", []), start=1):
                filename = sanitize_for_filename(photo.get("name", f"foto_{index}"), f"foto_{index}")
                zip_file.writestr(
                    f"{base_name}/evidencias/{claim_folder}/piezas_viejas/fotos/{index:02d}_{filename}",
                    attachment_bytes(photo),
                )

            certificate = attachments.get("destruction_certificate")
            if certificate:
                filename = sanitize_for_filename(certificate.get("name", "certificado_destruccion"), "certificado_destruccion")
                zip_file.writestr(
                    f"{base_name}/evidencias/{claim_folder}/piezas_viejas/certificado_destruccion/{filename}",
                    attachment_bytes(certificate),
                )

    return output.getvalue()


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
    if "audit_mode" not in st.session_state:
        st.session_state.audit_mode = AUDIT_MODES[0]


def build_general_comment_from_observations(claim: Dict[str, Any], include_campaigns: bool = True) -> str:
    """Construye el comentario general usando solo las observaciones escritas en cada apartado.

    No inventa desviaciones por puntuación: simplemente recopila los comentarios manuales
    que el auditor ha escrito en Documentación, Piezas viejas y, si procede, Campañas.
    """
    checks = ALL_CHECKS if include_campaigns else ALL_SCORING_CHECKS
    lines = []

    for check in checks:
        evaluation = claim.get("evaluations", {}).get(check.key, {})
        comment = safe_str(evaluation.get("comment", ""))

        if not comment:
            continue

        # Evita duplicar el título si el usuario ya ha empezado el comentario con "OR: ...".
        normalized_comment = comment.lower().lstrip()
        normalized_label = check.label.lower().strip()
        if normalized_comment.startswith(normalized_label + ":"):
            lines.append(comment)
        else:
            lines.append(f"{check.label}: {comment}")

    return "\n".join(lines).strip()


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


def sync_claim_meta_field(claim: Dict[str, Any], field: str, label: str, default_value: str = "") -> str:
    claim_no = claim["claim_no"]
    key = f"claim_meta_{claim_no}_{field}"

    if not safe_str(claim.get(field, "")) and default_value:
        claim[field] = default_value

    if key not in st.session_state:
        st.session_state[key] = safe_str(claim.get(field, ""))

    value = st.text_input(label, key=key)
    claim[field] = safe_str(value)
    return claim[field]


def sync_claim_dealer_field(claim: Dict[str, Any], default_dealer: str = "") -> str:
    """Dealer por claim mediante desplegable de dealers activos."""
    claim_no = claim["claim_no"]
    key = f"claim_meta_{claim_no}_dealer"

    if not safe_str(claim.get("dealer", "")) and default_dealer:
        claim["dealer"] = default_dealer

    current_value = safe_str(claim.get("dealer", ""))
    options = get_dealer_options(current_value or default_dealer)

    if key not in st.session_state or st.session_state[key] not in options:
        st.session_state[key] = current_value if current_value in options else ""

    if not safe_str(st.session_state.get(key, "")) and current_value:
        st.session_state[key] = current_value

    value = st.selectbox(
        "Dealer",
        options,
        key=key,
        format_func=format_dealer_option,
    )
    claim["dealer"] = safe_str(value)
    return claim["dealer"]


def render_claim_quick_card(claim: Dict[str, Any], default_dealer: str = ""):
    """Ficha editable por claim: útil para futura integración en plataforma mayor."""
    st.caption("Ficha de la garantía")
    cols = st.columns(4)
    with cols[0]:
        sync_claim_dealer_field(claim, default_dealer)
    with cols[1]:
        sync_claim_meta_field(claim, "vin", "VIN")
    with cols[2]:
        sync_claim_meta_field(claim, "model", "Modelo")
    with cols[3]:
        sync_claim_meta_field(claim, "amount", "Importe")


def render_old_parts_attachments(claim: Dict[str, Any]):
    """Adjuntos específicos de la pestaña de piezas viejas."""
    attachments = get_claim_attachments(claim)
    claim_no = claim["claim_no"]
    version_key = f"attachment_uploader_version_{claim_no}"
    if version_key not in st.session_state:
        st.session_state[version_key] = 0
    version = st.session_state[version_key]

    st.markdown("### Adjuntos de piezas viejas")
    st.caption(
        "Adjunta al menos 3 fotos de piezas viejas y, aparte, el certificado de destrucción "
        "si aplica. Los adjuntos se guardan en el JSON de trabajo y en el paquete ZIP, incluidos PDF."
    )

    photo_files = st.file_uploader(
        "Fotos de piezas viejas — mínimo 3 archivos",
        type=OLD_PART_PHOTO_FILE_TYPES,
        accept_multiple_files=True,
        key=f"old_parts_photos_{claim_no}_{version}",
        help="Formatos admitidos: JPG, JPEG, PNG y WEBP.",
    )

    if photo_files:
        attachments["old_parts_photos"] = [uploaded_file_to_attachment(file) for file in photo_files]

    photos = attachments.get("old_parts_photos", [])
    if len(photos) >= 3:
        st.success(f"Fotos de piezas viejas adjuntas: {len(photos)} archivo(s).")
    elif len(photos) > 0:
        st.warning(f"Fotos de piezas viejas adjuntas: {len(photos)} archivo(s). Recomendado mínimo: 3.")
    else:
        st.info("Todavía no hay fotos de piezas viejas adjuntas para esta claim.")

    if photos:
        with st.expander("Ver / descargar fotos adjuntas"):
            for index, photo in enumerate(photos, start=1):
                col_name, col_download = st.columns([3, 1])
                with col_name:
                    st.write(f"{index}. {safe_str(photo.get('name', 'foto'))} · {photo.get('size', 0)} bytes")
                with col_download:
                    st.download_button(
                        "Descargar",
                        data=attachment_bytes(photo),
                        file_name=safe_str(photo.get("name", f"foto_{index}")) or f"foto_{index}",
                        mime=safe_str(photo.get("type", "application/octet-stream")) or "application/octet-stream",
                        key=f"download_photo_{claim_no}_{index}_{safe_str(photo.get('name', 'foto'))}",
                    )

        if st.button("Eliminar fotos adjuntas", key=f"clear_photos_{claim_no}"):
            attachments["old_parts_photos"] = []
            st.session_state[version_key] += 1
            st.rerun()

    certificate_file = st.file_uploader(
        "Certificado de destrucción — foto o PDF",
        type=DESTRUCTION_CERTIFICATE_FILE_TYPES,
        accept_multiple_files=False,
        key=f"destruction_certificate_{claim_no}_{version}",
        help="Formatos admitidos: JPG, JPEG, PNG, WEBP y PDF.",
    )

    if certificate_file is not None:
        attachments["destruction_certificate"] = uploaded_file_to_attachment(certificate_file)

    certificate = attachments.get("destruction_certificate")
    if certificate:
        cert_cols = st.columns([3, 1, 1])
        with cert_cols[0]:
            st.success(
                f"Certificado adjunto: {safe_str(certificate.get('name', 'certificado'))} "
                f"· {certificate.get('size', 0)} bytes"
            )
        with cert_cols[1]:
            st.download_button(
                "Descargar certificado",
                data=attachment_bytes(certificate),
                file_name=safe_str(certificate.get("name", "certificado_destruccion")) or "certificado_destruccion",
                mime=safe_str(certificate.get("type", "application/octet-stream")) or "application/octet-stream",
                key=f"download_certificate_{claim_no}",
            )
        with cert_cols[2]:
            if st.button("Eliminar certificado", key=f"clear_certificate_{claim_no}"):
                attachments["destruction_certificate"] = None
                st.session_state[version_key] += 1
                st.rerun()
    else:
        st.info("Todavía no hay certificado de destrucción adjunto para esta claim.")

    claim["attachments"] = attachments


def sync_uploaded_attachments_from_session_state(claims: Dict[str, Dict[str, Any]]) -> None:
    """
    Sincroniza adjuntos ya subidos con st.session_state antes de construir descargas.

    En Streamlit el script se ejecuta de arriba a abajo. Los botones de descarga están
    por encima del uploader de adjuntos, así que, si el usuario subía un certificado PDF
    y descargaba el ZIP en esa misma pantalla, el ZIP podía construirse con el estado
    anterior. Esta función lee directamente los file_uploader ya presentes en
    session_state y los persiste en la claim antes de generar JSON/Excel/ZIP.
    """
    if not claims:
        return

    for claim in claims.values():
        claim_no = safe_str(claim.get("claim_no", ""))
        if not claim_no:
            continue

        attachments = get_claim_attachments(claim)
        version = st.session_state.get(f"attachment_uploader_version_{claim_no}", 0)

        photos_key = f"old_parts_photos_{claim_no}_{version}"
        photo_files = st.session_state.get(photos_key)
        if photo_files:
            if not isinstance(photo_files, list):
                photo_files = [photo_files]
            attachments["old_parts_photos"] = [
                uploaded_file_to_attachment(file)
                for file in photo_files
                if file is not None
            ]

        certificate_key = f"destruction_certificate_{claim_no}_{version}"
        certificate_file = st.session_state.get(certificate_key)
        if isinstance(certificate_file, list):
            certificate_file = certificate_file[0] if certificate_file else None
        if certificate_file is not None:
            attachments["destruction_certificate"] = uploaded_file_to_attachment(certificate_file)

        claim["attachments"] = attachments



def sync_editor_state_from_session_state(claims: Dict[str, Dict[str, Any]]) -> None:
    """Sincroniza widgets de edición antes de construir descargas.

    Los botones de descarga están en la columna izquierda y se renderizan antes que
    el editor de la claim. Esta función evita que una exportación se lleve valores
    de la ejecución anterior si acabas de cambiar un selector o escribir una
    observación.
    """
    if not claims:
        return

    check_by_key = {check.key: check for check in ALL_CHECKS}

    for claim in claims.values():
        claim_no = safe_str(claim.get("claim_no", ""))
        if not claim_no:
            continue

        dealer_key = f"claim_meta_{claim_no}_dealer"
        if dealer_key in st.session_state:
            claim["dealer"] = safe_str(st.session_state.get(dealer_key, ""))

        for field in ["vin", "model", "amount", "repair_date", "submission_date"]:
            meta_key = f"claim_meta_{claim_no}_{field}"
            if meta_key in st.session_state:
                claim[field] = safe_str(st.session_state.get(meta_key, ""))

        workflow_key = f"claim_workflow_status_{claim_no}"
        if workflow_key in st.session_state:
            claim["workflow_status"] = normalize_workflow_status(st.session_state.get(workflow_key, "Pendiente"))

        nsc_comment_key = f"nsc_comment_{claim_no}"
        if nsc_comment_key in st.session_state:
            claim["nsc_comment"] = safe_str(st.session_state.get(nsc_comment_key, ""))

        general_comment_key = f"general_comment_{claim_no}"
        if general_comment_key in st.session_state:
            claim["general_comment"] = safe_str(st.session_state.get(general_comment_key, ""))

        for check_key, check in check_by_key.items():
            evaluation = claim.get("evaluations", {}).setdefault(check_key, new_evaluation(check))

            select_key = f"select_{claim_no}_{check_key}"
            if select_key in st.session_state:
                option = option_from_label(check, safe_str(st.session_state.get(select_key, "")))
                evaluation["status"] = option.status
                evaluation["label"] = option.label
                evaluation["points"] = option.points
                evaluation["max_points"] = check.max_points

            comment_key = f"comment_{claim_no}_{check_key}"
            if comment_key in st.session_state:
                evaluation["comment"] = safe_str(st.session_state.get(comment_key, ""))



def main():
    st.set_page_config(page_title="Warranty Internal Audit Tool", layout="wide")
    init_state()
    sync_uploaded_attachments_from_session_state(st.session_state.claims)
    sync_editor_state_from_session_state(st.session_state.claims)

    st.title("Warranty Internal Audit Tool")
    st.caption("Herramienta interna de auditoría de garantías: documentación + piezas viejas = 100 puntos. Campañas solo informativo. Ficha por claim, adjuntos, estados y validación de cierre.")

    with st.sidebar:
        st.header("Auditoría")

        st.selectbox(
            "Modo de trabajo",
            AUDIT_MODES,
            key="audit_mode",
            help="De momento es funcionalmente interno. El modo dealer queda como base para una futura plataforma/Lovable.",
        )

        st.subheader("Guardar / cargar progreso")
        workfile = st.file_uploader(
            "Cargar auditoría guardada (.json)",
            type=["json"],
            key="workfile_upload",
            help="Carga un archivo generado con el botón 'Guardar auditoría de trabajo'.",
        )
        if workfile is not None and st.button("Cargar auditoría guardada", type="primary"):
            try:
                loaded_claims, loaded_audit_name, loaded_dealer, loaded_auditor, loaded_mode = load_audit_workfile(workfile)
                st.session_state.claims = loaded_claims
                st.session_state.selected_claim = next(iter(loaded_claims.keys()))
                st.session_state.audit_name = loaded_audit_name
                st.session_state.audit_dealer = loaded_dealer
                st.session_state.audit_auditor = loaded_auditor
                st.session_state.audit_mode = loaded_mode
                st.success(f"Auditoría cargada: {len(loaded_claims)} claims.")
                st.rerun()
            except Exception as exc:
                st.error(f"No se pudo cargar la auditoría guardada: {exc}")

        if st.session_state.claims:
            st.download_button(
                "Descargar auditoría de trabajo (.json)",
                data=serialize_audit_workfile(
                    st.session_state.claims,
                    st.session_state.audit_name,
                    st.session_state.audit_dealer,
                    st.session_state.audit_auditor,
                    st.session_state.get("audit_mode", AUDIT_MODES[0]),
                ),
                file_name=f"{build_audit_file_basename(st.session_state.get('audit_dealer', ''), st.session_state.get('audit_auditor', ''))}.json",
                mime="application/json",
                help="Descarga el archivo de progreso para poder reabrir la auditoría más adelante.",
                key="download_workfile_sidebar",
            )
            st.caption("Este JSON es el archivo de trabajo editable. El ZIP también incluye una copia.")
        else:
            st.caption("Carga claims para poder descargar el JSON de trabajo.")

        st.divider()
        audit_name = st.text_input("Nombre auditoría", key="audit_name")

        dealer_values = get_dealer_options(st.session_state.get("audit_dealer", ""))
        dealer = st.selectbox(
            "Dealer",
            dealer_values,
            key="audit_dealer",
            format_func=format_dealer_option,
            help="Listado de dealers activos. Si cargas un JSON antiguo con otro dealer, se conservará como opción temporal.",
        )

        auditor = st.text_input("Auditor", key="audit_auditor")
        export_base_name = build_audit_file_basename(dealer, auditor)
        st.caption(f"Nombre base de archivos: `{export_base_name}`")

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
            new_claim = empty_claim_record(claim_no)
            if safe_str(st.session_state.get("audit_dealer", "")):
                new_claim["dealer"] = safe_str(st.session_state.audit_dealer)
            st.session_state.claims.setdefault(claim_no, new_claim)
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
    apply_default_dealer_to_blank_claims(claims, dealer)
    sync_uploaded_attachments_from_session_state(claims)
    sync_editor_state_from_session_state(claims)

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

        if st.button("Eliminar claim seleccionada", help="Solo elimina la claim de esta auditoría en curso."):
            claim_to_delete = st.session_state.selected_claim
            if claim_to_delete in claims:
                del claims[claim_to_delete]
                st.session_state.selected_claim = next(iter(claims.keys()), None)
                st.rerun()

        st.download_button(
            "💾 Guardar progreso editable (.json)",
            data=serialize_audit_workfile(claims, audit_name, dealer, auditor, st.session_state.get("audit_mode", AUDIT_MODES[0])),
            file_name=f"{build_audit_file_basename(st.session_state.get('audit_dealer', ''), st.session_state.get('audit_auditor', ''))}.json",
            mime="application/json",
            help="Este es el archivo que debes subir en 'Cargar auditoría guardada' para continuar editando otro día.",
            key="download_workfile_main",
        )

        st.download_button(
            "Exportar auditoría a Excel",
            data=export_excel(claims, audit_name, dealer, auditor),
            file_name=f"{export_base_name}_analitico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "Exportar boletín de notas Excel",
            data=export_report_card_excel(claims, audit_name, dealer, auditor),
            file_name=f"{export_base_name}_boletin.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Genera un Excel con hojas tipo plantilla: documentación, piezas viejas, áreas de mejora, criterios y boletín por claim.",
        )

        st.download_button(
            "Exportar paquete completo (.zip)",
            data=export_audit_package_zip(claims, audit_name, dealer, auditor),
            file_name=f"{export_base_name}.zip",
            mime="application/zip",
            help="Incluye Excel analítico, boletín, informe TXT, JSON de trabajo y los archivos adjuntos reales por claim.",
        )

        closing_df = build_closing_validation_dataframe(claims)
        if closing_df.empty:
            st.success("Validación de cierre: OK")
        else:
            st.warning(f"Validación de cierre: {len(closing_df)} aviso(s)")

    with right:
        claim = claims[st.session_state.selected_claim]
        score = calculate_claim_score(claim)

        st.subheader(f"Revisión claim {claim['claim_no']}")
        render_claim_quick_card(claim, dealer or "")

        score_cols = st.columns(4)
        score_cols[0].metric("Documentación", f"{score['doc_points']}/{MAX_DOCUMENT_POINTS}")
        score_cols[1].metric("Piezas viejas", f"{score['old_points']}/{MAX_OLD_PARTS_POINTS}")
        score_cols[2].metric("Resultado claim", f"{score['total_points']}/100")
        score_cols[3].metric("Estado", get_claim_workflow_status(claim))

        workflow_key = f"claim_workflow_status_{claim['claim_no']}"
        pending_workflow_key = f"pending_workflow_status_{claim['claim_no']}"
        current_workflow_status = get_claim_workflow_status(claim)
        claim["workflow_status"] = current_workflow_status

        # Los botones de estado escriben en una clave temporal para no modificar
        # directamente el valor de un widget después de crearlo en la misma ejecución.
        if pending_workflow_key in st.session_state:
            current_workflow_status = normalize_workflow_status(st.session_state[pending_workflow_key])
            claim["workflow_status"] = current_workflow_status
            st.session_state[workflow_key] = current_workflow_status
            del st.session_state[pending_workflow_key]

        if workflow_key not in st.session_state or st.session_state[workflow_key] not in CLAIM_WORKFLOW_STATUSES:
            st.session_state[workflow_key] = current_workflow_status

        workflow_cols = st.columns([1.4, 1, 1, 1.2])
        with workflow_cols[0]:
            selected_workflow_status = st.selectbox(
                "Estado de trabajo",
                CLAIM_WORKFLOW_STATUSES,
                index=CLAIM_WORKFLOW_STATUSES.index(st.session_state[workflow_key]),
                key=workflow_key,
            )
            claim["workflow_status"] = normalize_workflow_status(selected_workflow_status)
        with workflow_cols[1]:
            if st.button("Marcar completada", use_container_width=True, key=f"mark_completed_{claim['claim_no']}"):
                claim["workflow_status"] = "Completada"
                st.session_state[pending_workflow_key] = "Completada"
                st.rerun()
        with workflow_cols[2]:
            if st.button("Requiere aclaración", use_container_width=True, key=f"mark_clarification_{claim['claim_no']}"):
                claim["workflow_status"] = "Requiere aclaración"
                st.session_state[pending_workflow_key] = "Requiere aclaración"
                st.rerun()
        with workflow_cols[3]:
            if st.button("Cerrar claim", use_container_width=True, key=f"mark_closed_{claim['claim_no']}"):
                claim["workflow_status"] = "Cerrada"
                st.session_state[pending_workflow_key] = "Cerrada"
                st.rerun()

        nsc_comment_key = f"nsc_comment_{claim['claim_no']}"
        if nsc_comment_key not in st.session_state:
            st.session_state[nsc_comment_key] = claim.get("nsc_comment", "")
        claim["nsc_comment"] = st.text_input(
            "Comentario interno / seguimiento rápido",
            key=nsc_comment_key,
            placeholder="Ej.: revisar fotos, pedir aclaración, OK para cierre...",
        )

        if score["pending"]:
            st.warning("Apartados pendientes: " + ", ".join(score["pending"]))

        section_names = ["I. Documentación", "II. Piezas viejas", "Campañas", "Comentarios", "Informe"]

        if "active_audit_section" not in st.session_state:
            st.session_state.active_audit_section = section_names[0]

        if st.session_state.active_audit_section not in section_names:
            st.session_state.active_audit_section = section_names[0]

        st.radio(
            "Sección de revisión",
            section_names,
            index=section_names.index(st.session_state.active_audit_section),
            horizontal=True,
            key="active_audit_section",
            label_visibility="collapsed",
        )

        selected_section = st.session_state.active_audit_section

        if selected_section == "I. Documentación":
            render_check_editor(claim, DOCUMENT_CHECKS)

        elif selected_section == "II. Piezas viejas":
            render_check_editor(claim, OLD_PARTS_CHECKS)
            render_old_parts_attachments(claim)

        elif selected_section == "Campañas":
            render_campaign_editor(claim)

        elif selected_section == "Comentarios":
            st.caption(
                "El comentario general se puede generar automáticamente a partir de las "
                "observaciones que hayas escrito en cada apartado. No inventa nada por la nota."
            )

            general_comment_key = f"general_comment_{claim['claim_no']}"
            if general_comment_key not in st.session_state:
                st.session_state[general_comment_key] = claim.get("general_comment", "")

            if st.button("Generar desde observaciones de apartados", key=f"generate_comment_{claim['claim_no']}"):
                generated_comment = build_general_comment_from_observations(claim, include_campaigns=True)

                if generated_comment:
                    claim["general_comment"] = generated_comment
                    st.session_state[general_comment_key] = generated_comment
                    st.success("Comentario general generado desde las observaciones de los apartados.")
                else:
                    st.warning("No hay observaciones escritas en los apartados para generar el comentario general.")

            claim["general_comment"] = st.text_area(
                "Comentarios generales de la claim",
                key=general_comment_key,
                height=180,
            )

        elif selected_section == "Informe":
            report = generate_text_report(claims, audit_name, dealer, auditor)
            st.text_area("Informe generado", value=report, height=420)
            st.download_button(
                "Descargar informe .txt",
                data=report.encode("utf-8"),
                file_name=f"{export_base_name}_informe.txt",
                mime="text/plain",
            )

            st.markdown("### Validación de cierre")
            closing_df = build_closing_validation_dataframe(claims)
            if closing_df.empty:
                st.success("Sin avisos de cierre. La auditoría está limpia para exportar/cerrar.")
            else:
                st.warning(f"Hay {len(closing_df)} aviso(s) antes de cerrar la auditoría.")
                st.dataframe(closing_df, use_container_width=True, hide_index=True)

        st.divider()
        section_index = section_names.index(st.session_state.active_audit_section)
        section_nav_cols = st.columns([1, 1, 2])

        with section_nav_cols[0]:
            if st.button(
                "← Apartado anterior",
                disabled=section_index == 0,
                use_container_width=True,
                key=f"previous_section_{claim['claim_no']}",
            ):
                st.session_state.active_audit_section = section_names[section_index - 1]
                st.rerun()

        with section_nav_cols[1]:
            if st.button(
                "Siguiente apartado →",
                type="primary",
                disabled=section_index >= len(section_names) - 1,
                use_container_width=True,
                key=f"next_section_{claim['claim_no']}",
            ):
                st.session_state.active_audit_section = section_names[section_index + 1]
                st.rerun()

        with section_nav_cols[2]:
            st.caption(f"Apartado {section_index + 1} de {len(section_names)} · {st.session_state.active_audit_section}")

        st.divider()
        current_index = claim_options.index(st.session_state.selected_claim)
        nav_cols = st.columns([1, 1, 2])

        with nav_cols[0]:
            if st.button("← Anterior claim", disabled=current_index == 0, use_container_width=True):
                st.session_state.selected_claim = claim_options[current_index - 1]
                st.session_state.active_audit_section = section_names[0]
                st.rerun()

        with nav_cols[1]:
            if st.button("Siguiente claim →", disabled=current_index >= len(claim_options) - 1, use_container_width=True):
                st.session_state.selected_claim = claim_options[current_index + 1]
                st.session_state.active_audit_section = section_names[0]
                st.rerun()

        with nav_cols[2]:
            st.caption(f"Claim {current_index + 1} de {len(claim_options)}")


if __name__ == "__main__":
    main()
