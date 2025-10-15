# Instalar: 
# py -m pip install --upgrade pip
# py -m pip install pdf2docx

#  Ejecución:
# python convert-pdf-to-docx_v-fileEH.py "C:\Users\user\Downloads\test\sssssssssssssss.pdf"
# Esta versión convierte un PDF a word.
# 
#
# Recomendado usar con el script de lectura de CSV "conver-docx-to-csv_v2.py":
#   python convert-pdf-to-docx_v-fileEH.py "C:/Users/user/Downloads/ssssssssssss.pdf"
#   python conver-docx-to-csv_v2.py "C:/Users/user/Downloads/ssssssssssss.docx" --csv "C:/Users/user/Downloads/ssssssssssss.csv"
#


import argparse
import os
import sys
from typing import List, Tuple, Optional

# -----------------------------
# Capa de utilidades (helpers)
# -----------------------------
def parse_page_ranges(spec: Optional[str]) -> Optional[List[Tuple[int, int]]]:
    """
    Convierte una cadena tipo '1-3,5,7-9' a una lista de tuplas [(1,3), (5,5), (7,9)] (1-based).
    Devuelve None si no se especifica.
    """
    if not spec:
        return None
    ranges: List[Tuple[int, int]] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start = int(a)
            end = int(b)
        else:
            start = int(part)
            end = start
        if start <= 0 or end <= 0:
            raise ValueError("Las páginas deben ser >= 1.")
        if end < start:
            raise ValueError(f"Rango inválido: {part}")
        ranges.append((start, end))
    return ranges


def default_output_path(input_pdf: str) -> str:
    base, _ = os.path.splitext(input_pdf)
    return base + ".docx"


def ensure_readable_file(path: str) -> None:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"No se encontró el archivo: {path}")
    if not os.access(path, os.R_OK):
        raise PermissionError(f"Sin permisos de lectura: {path}")


def ensure_writable_path(path: str, overwrite: bool) -> None:
    out_dir = os.path.dirname(path) or "."
    if not os.path.isdir(out_dir):
        raise FileNotFoundError(f"No existe el directorio de salida: {out_dir}")
    if os.path.exists(path) and not overwrite:
        raise FileExistsError(
            f"El archivo de salida ya existe: {path}. Usa --overwrite para sobrescribir."
        )


# --------------------------------
# Capa de conversión (motor)
# --------------------------------
def convert_with_pdf2docx(
    input_pdf: str,
    output_docx: str,
    page_ranges: Optional[List[Tuple[int, int]]] = None,
) -> None:
    """
    Convierte con pdf2docx. Si page_ranges es None, convierte todo el documento.
    """
    try:
        from pdf2docx import Converter
    except ImportError:
        raise RuntimeError(
            "No se encontró 'pdf2docx'. Instálalo con:  pip install pdf2docx"
        )

    print("📄 Iniciando conversión con pdf2docx...")
    cv = Converter(input_pdf)
    try:
        if not page_ranges:
            # Documento completo
            cv.convert(output_docx)
        else:
            # Convertir por rangos (1-based inclusive)
            # pdf2docx usa 0-based y end exclusive, así que ajustamos
            for i, (start, end) in enumerate(page_ranges, start=1):
                print(f"  → Rango {i}: páginas {start}-{end}")
                cv.convert(
                    output_docx,
                    start=start - 1,
                    end=end,  # end ya funciona como exclusivo al usar entero 0-based
                    # layout_kwargs podría afinar el reconocimiento, pero lo dejamos simple y estable
                )
    finally:
        cv.close()
    print(f"✅ Listo: {output_docx}")


# --------------------------------
# Capa de orquestación (CLI)
# --------------------------------
def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pdf_a_word.py",
        description="Convierte un PDF a Word (DOCX) preservando formato (similar a iLovePDF).",
    )
    p.add_argument("input_pdf", help="Ruta al archivo PDF de entrada")
    p.add_argument(
        "-o",
        "--output",
        dest="output_docx",
        help="Ruta al DOCX de salida (opcional). Por defecto, mismo nombre que el PDF.",
    )
    p.add_argument(
        "--pages",
        dest="pages",
        help='Rango(s) de páginas, ej. "1-3,5" (1-based). Si no se indica, convierte todo.',
    )
    p.add_argument(
        "--overwrite",
        action="store_true",
        help="Sobrescribe el archivo de salida si ya existe.",
    )
    return p


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    input_pdf = os.path.abspath(args.input_pdf)
    output_docx = (
        os.path.abspath(args.output_docx) if args.output_docx else default_output_path(input_pdf)
    )

    try:
        ensure_readable_file(input_pdf)
        ensure_writable_path(output_docx, overwrite=args.overwrite)
        page_ranges = parse_page_ranges(args.pages)
        convert_with_pdf2docx(input_pdf, output_docx, page_ranges)
    except Exception as ex:
        print(f"❌ Error: {ex}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
