"""Sample 02 — Batch generation (local .NET worker).

Generate one DOCX per record by calling the worker in a loop. Each call
spawns a short-lived `dotnet DocumentBridge.dll` subprocess, so the overhead is
linear in the number of records.
"""
from pathlib import Path

import document_sdk  # type: ignore[import-not-found]


RECORDS = [
    {"id": "INV-1001", "customer": "Acme Corp", "amount": "$1,250.00"},
    {"id": "INV-1002", "customer": "Globex",    "amount": "$890.50"},
    {"id": "INV-1003", "customer": "Initech",   "amount": "$2,400.00"},
]


def main() -> None:
    service = document_sdk.load_service()
    out_dir = (Path(__file__).resolve().parent.parent / "output" / "invoices").resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    for record in RECORDS:
        text = f"Invoice {record['id']} for {record['customer']} — {record['amount']}"
        out_path = out_dir / f"{record['id']}.docx"
        service.create_docx(text, out_path)
        print(f"  {record['id']} -> {out_path.name}")

    print(f"Wrote {len(RECORDS)} invoices into {out_dir}")


if __name__ == "__main__":
    main()
