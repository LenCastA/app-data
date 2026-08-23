#!/usr/bin/env python3
"""Convert the FIIS hourly-load workbook into an API seed CSV."""

import argparse
import csv
from pathlib import Path

from openpyxl import load_workbook


SOURCE_HEADERS = {
    "CÓDIGO": "course",
    "NOMBRE DEL CURSO": "course_name",
    "SECCIÓN": "section",
    "SISTEMA DE EVALUACIÓN": "evaluation",
    "APELLIDOS Y NOMBRES DEL DOCENTE": "teacher",
    "TIPO CLASE": "session_type",
    "AULA": "classroom",
    "DÍA": "day",
    "HORA INICIO": "start",
    "HORA FINAL": "end",
    "DNI": "teacher_dni",
    "VACANTES": "vacancies",
}

OUTPUT_HEADERS = [
    "codigo_facultad",
    "codigo_curso",
    "nombre_curso",
    "seccion",
    "evaluacion",
    "vacantes",
    "inicio",
    "fin",
    "aula",
    "dni_docente",
    "nombre_docente",
    "tipo",
    "dia",
]


def clean(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_header(sheet) -> tuple[int, dict[str, int]]:
    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        positions = {clean(value).upper(): index for index, value in enumerate(row)}
        if SOURCE_HEADERS.keys() <= positions.keys():
            return row_number, {
                target: positions[source] for source, target in SOURCE_HEADERS.items()
            }
    raise ValueError("No se encontró la cabecera esperada de la carga horaria")


def previous_fallbacks(path: Path | None) -> dict[tuple[str, ...], tuple[str, str]]:
    if path is None:
        return {}
    with path.open(encoding="utf-8", newline="") as source:
        rows = csv.DictReader(source)
        return {
            (
                row["codigo_curso"].strip(),
                row["seccion"].strip(),
                row["dia"].strip(),
                row["inicio"].strip(),
                row["fin"].strip(),
            ): (row["nombre_docente"].strip(), row["tipo"].strip())
            for row in rows
        }


def convert(args: argparse.Namespace) -> int:
    workbook = load_workbook(args.workbook, data_only=True, read_only=True)
    sheet = workbook[args.sheet] if args.sheet else workbook.worksheets[0]
    header_row, columns = find_header(sheet)
    fallbacks = previous_fallbacks(args.previous)
    output: list[dict[str, str]] = []

    required = ("course", "course_name", "section", "evaluation", "day", "start", "end")
    for source_row_number, values in enumerate(
        sheet.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1
    ):
        row = {
            name: clean(values[index] if index < len(values) else None)
            for name, index in columns.items()
        }
        if not any(row.values()):
            continue
        missing = [name for name in required if not row[name]]
        if missing:
            raise ValueError(
                f"Fila {source_row_number}: faltan campos obligatorios: {', '.join(missing)}"
            )

        key = (row["course"], row["section"], row["day"], row["start"], row["end"])
        previous_teacher, previous_type = fallbacks.get(key, ("", ""))
        teacher = row["teacher"] or previous_teacher
        session_type = row["session_type"] or previous_type
        if not teacher:
            teacher = "AYUDANTE CATEDRA" if row["teacher_dni"].upper() == "SOLO PARA PC" else "NN"
        if not session_type and row["teacher_dni"].upper() == "SOLO PARA PC":
            session_type = "PC"

        output.append(
            {
                "codigo_facultad": args.faculty,
                "codigo_curso": row["course"].replace("-", ""),
                "nombre_curso": row["course_name"],
                "seccion": row["section"],
                "evaluacion": row["evaluation"],
                "vacantes": row["vacancies"],
                "inicio": row["start"],
                "fin": row["end"],
                "aula": row["classroom"],
                "dni_docente": "" if row["teacher_dni"].upper() == "SOLO PARA PC" else row["teacher_dni"],
                "nombre_docente": teacher,
                "tipo": session_type,
                "dia": row["day"],
            }
        )

    output.sort(key=lambda row: (row["codigo_curso"], row["seccion"], row["dia"], row["inicio"], row["tipo"]))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8", newline="") as target:
        writer = csv.DictWriter(target, fieldnames=OUTPUT_HEADERS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(output)
    print(f"generated: {args.output} ({len(output)} rows)")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--sheet", default="CARGA HORARIA 2026-2")
    parser.add_argument("--faculty", default="I")
    parser.add_argument("--previous", type=Path)
    return convert(parser.parse_args())


if __name__ == "__main__":
    raise SystemExit(main())
