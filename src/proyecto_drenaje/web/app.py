from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

from flask import Flask, render_template
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel
from openpyxl.utils.exceptions import InvalidFileException


PROJECT_ROOT = Path(__file__).resolve().parents[3]
SOURCE_FILE = PROJECT_ROOT / "Maquinaria.xlsx"
LOGISTICS_SOURCE_FILE = PROJECT_ROOT / "SeguimientoLogistica.xlsx"

app = Flask(__name__)


class ReportDataError(RuntimeError):
    """Raised when the source workbook cannot provide report-ready data."""


@dataclass(frozen=True)
class ActivityCost:
    activity_id: int
    activity: str
    partida: str
    budget_usd: str
    executed_usd: str


@dataclass(frozen=True)
class LogisticsRecord:
    partida: str
    fund: str
    order_number: str
    supplier: str
    total_usd: str


@dataclass(frozen=True)
class LogisticsSummary:
    project: str
    records: int
    priced_records: int
    total_budget_usd: str
    total_executed_usd: str
    activity_costs: list[ActivityCost]
    detail_records: list[LogisticsRecord]
    detail_total_usd: str


@dataclass(frozen=True)
class StopRecord:
    fund: str
    machine: str
    provider: str
    reason: str


@dataclass(frozen=True)
class StopReason:
    label: str
    count: int
    width: int


@dataclass(frozen=True)
class QualityMetric:
    label: str
    count: int
    percent: float
    display: str
    width: int


@dataclass(frozen=True)
class ReportData:
    start_date: str
    end_date: str
    daily_date: str
    cost_total: str
    cost_usd: str
    hours: str
    parts: int
    hours_per_part: str
    meters: str
    cost_per_meter: str
    paid: str
    paid_percent: str
    paid_width: int
    pending: str
    pending_percent: str
    hours_with_tariff_rate: str
    blended_rate: str
    unpriced_hours: str
    zero_hour_parts: int
    vendor: str
    vendor_share: str
    scheduled: int
    operational: int
    nonoperational: int
    availability: str
    availability_width: float
    mosqueta_total: int
    mosqueta_operational: int
    stop_reasons: list[StopReason]
    stop_records: list[StopRecord]
    quality_metrics: list[QualityMetric]
    missing_fund: int
    zero_hour_percent: str
    purchase_orders: str
    purchase_order_status: str
    has_formula_errors: bool


def _number(worksheet: Any, cell: str) -> float:
    value = worksheet[cell].value
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ReportDataError(
            f"Se esperaba un valor numérico en {worksheet.title}!{cell}; "
            f"se encontró {value!r}."
        )
    if not math.isfinite(value):
        raise ReportDataError(
            f"El valor de {worksheet.title}!{cell} no es un número finito."
        )
    return float(value)


def _text(worksheet: Any, cell: str) -> str:
    value = worksheet[cell].value
    if value is None:
        return ""
    return str(value).strip()


def _format_number(value: float, decimals: int = 2) -> str:
    formatted = f"{value:,.{decimals}f}"
    return formatted.replace(",", "_").replace(".", ",").replace("_", ".")


def _as_date(value: Any, epoch: datetime) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        converted = from_excel(value, epoch)
        if isinstance(converted, datetime):
            return converted.date()
        if isinstance(converted, date):
            return converted
    raise ReportDataError(f"No se pudo interpretar la fecha del libro: {value!r}.")


def _format_date(value: Any, epoch: datetime) -> str:
    return _as_date(value, epoch).strftime("%d/%m/%Y")


def _normalize_label(value: Any) -> str:
    decomposed = unicodedata.normalize("NFKD", str(value or "").casefold())
    plain = "".join(
        character for character in decomposed
        if not unicodedata.combining(character)
    )
    return " ".join(re.findall(r"[a-z0-9]+", plain))


def _display_value(value: Any) -> str:
    if value is None:
        return "—"
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return text or "—"


def load_logistics_summary(
    source_file: Path = LOGISTICS_SOURCE_FILE,
    machinery_file: Path = SOURCE_FILE,
) -> LogisticsSummary:
    if not source_file.is_file():
        raise ReportDataError(f"No se encontró la base logística: {source_file}")

    try:
        workbook = load_workbook(source_file, read_only=True, data_only=True)
    except (BadZipFile, InvalidFileException, OSError) as error:
        raise ReportDataError(
            f"No se pudo abrir la base logística: {source_file}"
        ) from error

    try:
        if "BD_LOGISTICO" not in workbook.sheetnames:
            raise ReportDataError(
                "Falta la hoja BD_LOGISTICO en SeguimientoLogistica.xlsx."
            )

        worksheet = workbook["BD_LOGISTICO"]
        headers = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))
        columns = {
            str(value).strip().upper(): index
            for index, value in enumerate(headers)
            if value is not None
        }
        required_columns = {
            "PROYECTO", "PARTIDA", "FUNDO", "NRO ORDEN", "PROVEEDOR", "TOTAL $"
        }
        if not required_columns.issubset(columns):
            raise ReportDataError(
                "BD_LOGISTICO debe incluir PROYECTO, PARTIDA, FUNDO, NRO ORDEN, "
                "PROVEEDOR y TOTAL $."
            )

        records = 0
        priced_records = 0
        project_labels: set[str] = set()
        partida_totals: dict[str, float] = {}
        detail_records: list[LogisticsRecord] = []
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            project_value = row[columns["PROYECTO"]]
            project = " ".join(str(project_value or "").split()).upper()
            if project != "DRENAJE" and not project.startswith("DRENAJE "):
                continue

            records += 1
            project_labels.add(project)
            partida_value = " ".join(str(row[columns["PARTIDA"]] or "").split())
            partida = partida_value or "(Sin partida)"
            partida_totals.setdefault(partida, 0.0)
            amount = row[columns["TOTAL $"]]
            if amount is None:
                amount_display = "—"
            else:
                if isinstance(amount, bool) or not isinstance(amount, (int, float)):
                    raise ReportDataError(
                        f"El costo en dólares del registro {records} no es numérico."
                    )
                if not math.isfinite(amount):
                    raise ReportDataError(
                        f"El costo en dólares del registro {records} no es finito."
                    )
                amount_number = float(amount)
                partida_totals[partida] += amount_number
                priced_records += 1
                amount_display = f"US$ {_format_number(amount_number)}"

            detail_records.append(
                LogisticsRecord(
                    partida=partida,
                    fund=_display_value(row[columns["FUNDO"]]),
                    order_number=_display_value(row[columns["NRO ORDEN"]]),
                    supplier=_display_value(row[columns["PROVEEDOR"]]),
                    total_usd=amount_display,
                )
            )

        if not records:
            raise ReportDataError(
                "No se encontraron registros del proyecto DRENAJE en BD_LOGISTICO."
            )

    finally:
        workbook.close()

    if not machinery_file.is_file():
        raise ReportDataError(f"No se encontr\u00f3 el libro de maquinaria: {machinery_file}")
    try:
        machinery_workbook = load_workbook(
            machinery_file, read_only=True, data_only=True
        )
    except (BadZipFile, InvalidFileException, OSError) as error:
        raise ReportDataError(
            f"No se pudo abrir el libro de maquinaria: {machinery_file}"
        ) from error

    try:
        sheet_name = "Gantt_Drenaje integral E01"
        if sheet_name not in machinery_workbook.sheetnames:
            raise ReportDataError(f"Falta la hoja {sheet_name} en Maquinaria.xlsx.")

        gantt = machinery_workbook[sheet_name]
        header_row = 0
        gantt_columns: dict[str, int] = {}
        for row_number, row in enumerate(
            gantt.iter_rows(min_row=1, max_row=min(10, gantt.max_row), values_only=True),
            start=1,
        ):
            candidates = {
                _normalize_label(value): index
                for index, value in enumerate(row)
                if value is not None
            }
            if {"id", "actividad", "presupuesto sub total usd"}.issubset(candidates):
                header_row = row_number
                gantt_columns = candidates
                break
        if not header_row:
            raise ReportDataError(
                "No se encontraron ID, ACTIVIDAD y PRESUPUESTO SUB TOTAL (USD) "
                "en la hoja Gantt_Drenaje integral E01."
            )

        logistics_by_normalized_partida = {
            _normalize_label(partida): (partida, amount)
            for partida, amount in partida_totals.items()
        }
        aliases = {
            "movimiento de tierras": "movimiento de tierra",
            "obras electromecanicas y bombeo": "obras electricas magneticas y bombeo",
        }
        desired_ids = {1, 2, 3, 4, 5}
        seen_ids: set[int] = set()
        activity_costs: list[ActivityCost] = []
        total_budget = 0.0
        total_executed = 0.0

        for row in gantt.iter_rows(min_row=header_row + 1, values_only=True):
            raw_id = row[gantt_columns["id"]]
            if isinstance(raw_id, bool) or not isinstance(raw_id, (int, float)):
                continue
            if not float(raw_id).is_integer() or int(raw_id) not in desired_ids:
                continue

            activity_id = int(raw_id)
            if activity_id in seen_ids:
                raise ReportDataError(
                    f"El ID {activity_id} aparece más de una vez como actividad principal."
                )
            seen_ids.add(activity_id)

            activity = str(row[gantt_columns["actividad"]] or "").strip()
            raw_budget = row[gantt_columns["presupuesto sub total usd"]]
            if isinstance(raw_budget, bool) or not isinstance(raw_budget, (int, float)):
                raise ReportDataError(
                    f"El presupuesto de la actividad {activity_id} no es numérico."
                )
            if not math.isfinite(raw_budget):
                raise ReportDataError(
                    f"El presupuesto de la actividad {activity_id} no es finito."
                )

            normalized_activity = _normalize_label(activity)
            matched = logistics_by_normalized_partida.get(normalized_activity)
            if matched is None:
                legacy_alias = aliases.get(normalized_activity)
                if legacy_alias is not None:
                    matched = logistics_by_normalized_partida.get(legacy_alias)
            if matched is None:
                partida = "Sin partida equivalente"
                executed = 0.0
                executed_display = "—"
            else:
                partida, executed = matched
                executed_display = f"US$ {_format_number(executed)}"

            budget = float(raw_budget)
            total_budget += budget
            total_executed += executed
            activity_costs.append(
                ActivityCost(
                    activity_id=activity_id,
                    activity=activity,
                    partida=partida,
                    budget_usd=f"US$ {_format_number(budget)}",
                    executed_usd=executed_display,
                )
            )

        if seen_ids != desired_ids:
            missing_ids = ", ".join(map(str, sorted(desired_ids - seen_ids)))
            raise ReportDataError(
                f"No se encontraron todas las actividades principales del Gantt. "
                f"Faltan los ID: {missing_ids}."
            )

        return LogisticsSummary(
            project=", ".join(sorted(project_labels)),
            records=records,
            priced_records=priced_records,
            total_budget_usd=f"US$ {_format_number(total_budget)}",
            total_executed_usd=f"US$ {_format_number(total_executed)}",
            activity_costs=activity_costs,
            detail_records=detail_records,
            detail_total_usd=f"US$ {_format_number(sum(partida_totals.values()))}",
        )
    finally:
        machinery_workbook.close()


def _stop_category(reason: str) -> str:
    normalized = reason.casefold()
    if "operador" in normalized:
        return "Falta de operador"
    if "manguera" in normalized:
        return "Falla / manguera"
    if "malograd" in normalized or "averiad" in normalized:
        return "Excavadora averiada"
    if "retir" in normalized:
        return "Retiro del fundo"
    return "Otra causa"


def _vendor_share(workbook: Any, total_cost: float) -> tuple[str, str]:
    worksheet = workbook["Costo Histórico"]
    vendor_cost = 0.0
    has_vendor_rows = False
    for row in range(17, worksheet.max_row + 1):
        if _text(worksheet, f"C{row}").casefold() != "c&v":
            continue
        machine = _text(worksheet, f"B{row}").casefold()
        if not machine or machine.startswith("total"):
            continue
        vendor_cost += _number(worksheet, f"G{row}")
        has_vendor_rows = True
    if not has_vendor_rows or total_cost <= 0:
        raise ReportDataError("No se pudo calcular la concentración de costo de C&V.")
    share = vendor_cost / total_cost * 100
    return "C&V", _format_number(share, 0)


def load_report_data(source_file: Path = SOURCE_FILE) -> ReportData:
    if not source_file.is_file():
        raise ReportDataError(f"No se encontró el archivo fuente: {source_file}")

    try:
        workbook = load_workbook(source_file, read_only=True, data_only=True)
    except (BadZipFile, InvalidFileException, OSError) as error:
        raise ReportDataError(
            f"No se pudo abrir el libro de maquinaria: {source_file}"
        ) from error

    required_sheets = {
        "Detalle Reporte",
        "Reporte Ejecutivo",
        "Reporte Diario",
        "Costo Histórico",
    }
    missing_sheets = required_sheets.difference(workbook.sheetnames)
    if missing_sheets:
        workbook.close()
        names = ", ".join(sorted(missing_sheets))
        raise ReportDataError(f"Faltan hojas requeridas en el Excel: {names}.")

    try:
        detail = workbook["Detalle Reporte"]
        executive = workbook["Reporte Ejecutivo"]
        daily = workbook["Reporte Diario"]

        total_cost = _number(detail, "B9")
        hours = _number(detail, "B8")
        parts = int(_number(detail, "B7"))
        paid = _number(detail, "B11")
        pending = _number(detail, "B12")
        paid_percent = _number(detail, "B13") * 100
        meters = _number(detail, "B17")
        scheduled = int(_number(daily, "C8"))
        operational = int(_number(daily, "D8"))
        nonoperational = int(_number(daily, "E8"))

        if scheduled <= 0:
            raise ReportDataError("El parte diario no tiene equipos programados.")
        if hours <= 0 or parts <= 0:
            raise ReportDataError(
                "El corte financiero no contiene horas o partes para reportar."
            )
        if operational + nonoperational != scheduled:
            raise ReportDataError(
                "Los equipos operativos y no operativos no coinciden con el "
                "total programado del parte diario."
            )

        stop_records: list[StopRecord] = []
        reason_counts: dict[str, int] = {}
        mosqueta_total = 0
        mosqueta_operational = 0
        for row in range(11, daily.max_row + 1):
            machine = _text(daily, f"D{row}")
            if not machine:
                continue
            fund = _text(daily, f"C{row}")
            provider = _text(daily, f"E{row}")
            reason = _text(daily, f"G{row}")
            state = _text(daily, f"I{row}").casefold()
            if fund.casefold() == "mosqueta":
                mosqueta_total += 1
                if state == "operativa":
                    mosqueta_operational += 1
            if state == "no operativa":
                stop_records.append(
                    StopRecord(fund=fund, machine=machine, provider=provider, reason=reason)
                )
                category = _stop_category(reason)
                reason_counts[category] = reason_counts.get(category, 0) + 1

        max_reason_count = max(reason_counts.values(), default=1)
        stop_reasons = [
            StopReason(
                label=label,
                count=count,
                width=max(8, round(count / max_reason_count * 100)),
            )
            for label, count in sorted(
                reason_counts.items(), key=lambda item: (-item[1], item[0])
            )
        ]

        quality_data = [
            ("Sin estado", int(_number(detail, "E9")), "registros"),
            ("Sin marca de pago", int(_number(detail, "E12")), "registros"),
            ("Sin número de factura", int(_number(detail, "E11")), "registros"),
            ("Partes sin tarifa", int(_number(detail, "E7")), "partes"),
        ]
        quality_metrics = [
            QualityMetric(
                label=label,
                count=count,
                percent=count / parts * 100 if parts else 0,
                display=(
                    f"{_format_number(count / parts * 100, 1)}%"
                    if unit == "registros"
                    else f"{_format_number(_number(detail, 'E8'), 1)} h sin valorizar"
                ),
                width=max(1, round(count / parts * 100)) if parts else 0,
            )
            for label, count, unit in quality_data
        ]

        vendor, vendor_share = _vendor_share(workbook, total_cost)
        executive_text = _text(executive, "D17")
        purchase_orders = ", ".join(
            re.findall(r"\b\d{3,}\b", executive_text)
        )
        if not purchase_orders:
            purchase_orders = "Sin números de OC identificados"

        error_cells = ("B12", "B13", "H13", "D19")
        has_formula_errors = any(
            _text(executive, cell).startswith("#") for cell in error_cells
        )

        return ReportData(
            start_date=_format_date(detail["B5"].value, workbook.epoch),
            end_date=_format_date(detail["B6"].value, workbook.epoch),
            daily_date=_format_date(daily["D5"].value, workbook.epoch),
            cost_total=f"S/ {_format_number(total_cost)}",
            cost_usd=f"US$ {_format_number(_number(detail, 'B10'))}",
            hours=f"{_format_number(hours, 1)} h",
            parts=parts,
            hours_per_part=f"{_format_number(_number(detail, 'B14'))} h",
            meters=f"{_format_number(meters, 0)} m",
            cost_per_meter=f"S/ {_format_number(_number(detail, 'B18'))}/m",
            paid=f"S/ {_format_number(paid)}",
            paid_percent=f"{_format_number(paid_percent, 1)}%",
            paid_width=round(paid_percent),
            pending=f"S/ {_format_number(pending)}",
            pending_percent=f"{_format_number(pending / total_cost * 100, 1)}%",
            hours_with_tariff_rate=f"S/ {_format_number(_number(detail, 'B15'))}/h",
            blended_rate=f"S/ {_format_number(total_cost / hours)}/h",
            unpriced_hours=f"{_format_number(_number(detail, 'E8'), 1)} h",
            zero_hour_parts=int(_number(detail, "B16")),
            vendor=vendor,
            vendor_share=f"{vendor_share}%",
            scheduled=scheduled,
            operational=operational,
            nonoperational=nonoperational,
            availability=f"{_format_number(operational / scheduled * 100, 1)}%",
            availability_width=operational / scheduled * 100,
            mosqueta_total=mosqueta_total,
            mosqueta_operational=mosqueta_operational,
            stop_reasons=stop_reasons,
            stop_records=stop_records,
            quality_metrics=quality_metrics,
            missing_fund=int(_number(detail, "E10")),
            zero_hour_percent=f"{_format_number(_number(detail, 'B16') / parts * 100, 1)}%",
            purchase_orders=purchase_orders,
            purchase_order_status=_text(executive, "K8"),
            has_formula_errors=has_formula_errors,
        )
    finally:
        workbook.close()


@app.get("/")
def machinery_report() -> str | tuple[str, int]:
    try:
        logistics = load_logistics_summary()
        report = load_report_data()
    except ReportDataError as error:
        app.logger.error("No se pudo generar el reporte: %s", error)
        return render_template("error.html", message=str(error)), 503
    return render_template("report.html", report=report, logistics=logistics)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
