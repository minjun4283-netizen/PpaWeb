#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from excel_writer import upsert_row_to_table


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsm", required=True)
    parser.add_argument("--table", required=True)
    parser.add_argument("--pk", required=True)
    parser.add_argument("--payload-file", required=True)
    args = parser.parse_args()

    try:
        payload = json.loads(Path(args.payload_file).read_text(encoding="utf-8"))
        result = upsert_row_to_table(args.xlsm, args.table, args.pk, payload)
        print(json.dumps({"ok": True, "result": result}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()