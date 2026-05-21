import sys

import fitz


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: extract_pdf_text.py <pdf_path>", file=sys.stderr)
        return 2

    sys.stdout.reconfigure(encoding="utf-8")
    doc = fitz.open(sys.argv[1])
    try:
        pages = [page.get_text("text") for page in doc]
    finally:
        doc.close()

    print("\n".join(pages))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
