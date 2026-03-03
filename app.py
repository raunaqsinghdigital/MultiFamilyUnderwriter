#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import hashlib
import hmac
import json
import math
import mimetypes
import os
import re
import threading
import time
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import urlopen
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import encode_dss_signature
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric.ec import EllipticCurvePublicNumbers, SECP256R1
from cryptography.hazmat.backends import default_backend


ROOT_DIR = Path(__file__).resolve().parent
DATA_DIR = ROOT_DIR / "data"
STATIC_DIR = ROOT_DIR / "static"
MODEL_PATH = DATA_DIR / "workbook_model.json"
FORMULAS_PATH = DATA_DIR / "formulas_expanded.json"
OVERRIDES_PATH = DATA_DIR / "admin_overrides.json"
MORTGAGE_SHEET_NAME = "Mortgage"
ONE_POINT_ONE_SHEET_NAME = "1.1"
HIDDEN_ANALYST_SHEETS = {MORTGAGE_SHEET_NAME, ONE_POINT_ONE_SHEET_NAME}
MANUAL_INPUT_CELL_OVERRIDES: dict[str, dict[str, str]] = {
    # Keep Property Taxes $/year user-editable; dependent ratios remain formula-driven.
    "Valuation!E18": {"value_type": "number"},
}

CELL_REF_RE = re.compile(r"^\$?([A-Z]{1,3})\$?([0-9]+)$")

# ── Auth / CORS ────────────────────────────────────────────────────────────────
SUPABASE_JWT_SECRET: str = os.environ.get("SUPABASE_JWT_SECRET", "")
_jwks_cache: dict | None = None  # { kid: jwk_dict } — populated on first ES256 verification
_ALLOWED_ORIGINS_RAW: str = os.environ.get("ALLOWED_ORIGINS", "")
_DEFAULT_ALLOWED_ORIGINS: set[str] = (
    {o.strip() for o in _ALLOWED_ORIGINS_RAW.split(",") if o.strip()}
    if _ALLOWED_ORIGINS_RAW.strip()
    else set()
)


def _b64url_decode(s: str) -> bytes:
    padding = 4 - len(s) % 4
    return base64.urlsafe_b64decode(s + "=" * padding)


def _get_jwks(issuer: str) -> dict:
    """Fetch Supabase JWKS and cache by kid. issuer = JWT 'iss' claim."""
    global _jwks_cache
    if _jwks_cache is not None:
        return _jwks_cache
    jwks_url = issuer.rstrip("/") + "/.well-known/jwks.json"
    with urlopen(jwks_url, timeout=5) as resp:
        keys = json.loads(resp.read())["keys"]
    _jwks_cache = {k["kid"]: k for k in keys}
    return _jwks_cache


def _verify_es256_jwt(token: str) -> dict | None:
    """Verify an ES256 JWT using Supabase JWKS. Returns payload or None."""
    try:
        parts = token.split(".")
        if len(parts) != 3:
            return None
        header_b64, payload_b64, sig_b64 = parts
        header = json.loads(_b64url_decode(header_b64))
        kid = header.get("kid")
        payload = json.loads(_b64url_decode(payload_b64))
        issuer = payload.get("iss", "")
        if not issuer or not kid:
            return None
        jwks = _get_jwks(issuer)
        jwk = jwks.get(kid)
        if not jwk:
            return None
        x = int.from_bytes(_b64url_decode(jwk["x"]), "big")
        y = int.from_bytes(_b64url_decode(jwk["y"]), "big")
        pub_key = EllipticCurvePublicNumbers(x, y, SECP256R1()).public_key(default_backend())
        sig_bytes = _b64url_decode(sig_b64)
        r = int.from_bytes(sig_bytes[:32], "big")
        s = int.from_bytes(sig_bytes[32:], "big")
        der_sig = encode_dss_signature(r, s)
        pub_key.verify(der_sig, f"{header_b64}.{payload_b64}".encode(), ec.ECDSA(hashes.SHA256()))
        if payload.get("exp", 0) < time.time():
            return None
        return payload
    except Exception:
        return None


def _verify_supabase_jwt(token: str) -> dict | None:
    """Verify a Supabase JWT (ES256 or HS256). Returns payload or None."""
    try:
        parts = token.split(".")
        if len(parts) != 3:
            return None
        header = json.loads(_b64url_decode(parts[0]))
        alg = header.get("alg", "HS256")
    except Exception:
        return None

    if alg == "ES256":
        return _verify_es256_jwt(token)

    # HS256 path (legacy Supabase projects)
    if not SUPABASE_JWT_SECRET:
        return None
    try:
        header_b64, payload_b64, sig_b64 = parts
        expected = hmac.new(
            SUPABASE_JWT_SECRET.encode(),
            f"{header_b64}.{payload_b64}".encode(),
            hashlib.sha256,
        ).digest()
        if not hmac.compare_digest(_b64url_decode(sig_b64), expected):
            return None
        payload = json.loads(_b64url_decode(payload_b64))
        if payload.get("exp", 0) < time.time():
            return None
        return payload
    except Exception:
        return None


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def parse_number(raw: str) -> int | float:
    value = float(raw)
    if value.is_integer():
        return int(value)
    return value


def col_to_index(col: str) -> int:
    value = 0
    for ch in col:
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value


def index_to_col(index: int) -> str:
    letters: list[str] = []
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def parse_cell_address(address: str) -> tuple[int, int]:
    normalized = address.replace("$", "").upper()
    match = CELL_REF_RE.match(normalized)
    if not match:
        raise ValueError(f"Invalid cell reference: {address}")
    col_letters, row_str = match.groups()
    return col_to_index(col_letters), int(row_str)


def normalize_cell_address(address: str) -> str:
    col, row = parse_cell_address(address)
    return f"{index_to_col(col)}{row}"


def iter_range_addresses(start: str, end: str) -> list[str]:
    start_col, start_row = parse_cell_address(start)
    end_col, end_row = parse_cell_address(end)

    min_col = min(start_col, end_col)
    max_col = max(start_col, end_col)
    min_row = min(start_row, end_row)
    max_row = max(start_row, end_row)

    out: list[str] = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            out.append(f"{index_to_col(col)}{row}")
    return out


def load_model(model_path: Path, formulas_path: Path) -> dict:
    if not model_path.exists():
        raise FileNotFoundError(
            f"Missing model file: {model_path}. "
            "Expected workbook metadata to exist in data/workbook_model.json"
        )
    if not formulas_path.exists():
        raise FileNotFoundError(
            f"Missing formulas file: {formulas_path}. "
            "Expected expanded formula map in data/formulas_expanded.json"
        )

    model = json.loads(model_path.read_text(encoding="utf-8"))
    formula_map = json.loads(formulas_path.read_text(encoding="utf-8"))

    missing: list[str] = []
    for sheet in model.get("sheets", []):
        for cell in sheet.get("cells", []):
            if not cell.get("has_formula"):
                continue
            key = cell["key"]
            formula = formula_map.get(key, cell.get("formula") or "")
            if not formula:
                missing.append(key)
                continue
            cell["formula"] = formula

    for item in model.get("formula_cells", []):
        key = item.get("key")
        if not key:
            continue
        item["formula"] = formula_map.get(key, item.get("formula") or "")

    if missing:
        sample = ", ".join(missing[:8])
        raise RuntimeError(f"Missing formulas for {len(missing)} cells. Sample: {sample}")

    model["loaded_at"] = now_iso()
    return model


def deep_clone(value):
    return json.loads(json.dumps(value))


def apply_formula_overrides(model: dict, overrides: dict[str, str]) -> None:
    if not overrides:
        return
    for sheet in model.get("sheets", []):
        for cell in sheet.get("cells", []):
            if not cell.get("has_formula"):
                continue
            key = cell.get("key")
            if key in overrides:
                cell["formula"] = overrides[key]

    for item in model.get("formula_cells", []):
        key = item.get("key")
        if key in overrides:
            item["formula"] = overrides[key]


def apply_manual_input_cell_overrides(model: dict) -> None:
    if not MANUAL_INPUT_CELL_OVERRIDES:
        return

    cell_index: dict[str, tuple[str, dict]] = {}
    for sheet in model.get("sheets", []):
        sheet_name = sheet.get("name", "")
        for cell in sheet.get("cells", []):
            key = cell.get("key")
            if key:
                cell_index[key] = (sheet_name, cell)

    input_cells = model.setdefault("input_cells", [])
    formula_cells = model.get("formula_cells", [])

    for key, config in MANUAL_INPUT_CELL_OVERRIDES.items():
        record = cell_index.get(key)
        if record is None:
            continue

        sheet_name, cell = record
        cell["is_input"] = True
        cell["has_formula"] = False
        cell["formula"] = ""

        existing_input = next((entry for entry in input_cells if entry.get("key") == key), None)
        if existing_input is None:
            input_cells.append(
                {
                    "key": key,
                    "sheet": sheet_name,
                    "address": cell.get("address", ""),
                    "value_type": config.get("value_type", cell.get("value_type", "text")),
                    "default_value": cell.get("value"),
                }
            )
        else:
            existing_input["sheet"] = sheet_name
            existing_input["address"] = cell.get("address", "")
            existing_input["value_type"] = config.get(
                "value_type", existing_input.get("value_type", cell.get("value_type", "text"))
            )
            if "default_value" not in existing_input:
                existing_input["default_value"] = cell.get("value")

    if formula_cells:
        model["formula_cells"] = [
            entry for entry in formula_cells if entry.get("key") not in MANUAL_INPUT_CELL_OVERRIDES
        ]


def filter_public_model(model: dict) -> dict:
    public_model = deep_clone(model)
    public_model["sheets"] = [
        sheet
        for sheet in public_model.get("sheets", [])
        if sheet.get("name") not in HIDDEN_ANALYST_SHEETS
    ]
    public_model["input_cells"] = [
        cell
        for cell in public_model.get("input_cells", [])
        if cell.get("sheet") not in HIDDEN_ANALYST_SHEETS
    ]
    public_model["formula_cells"] = [
        cell
        for cell in public_model.get("formula_cells", [])
        if cell.get("sheet") not in HIDDEN_ANALYST_SHEETS
    ]
    public_model["hidden_sheets"] = sorted(HIDDEN_ANALYST_SHEETS)
    return public_model


def load_formula_overrides(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        return {}
    overrides = payload.get("formula_overrides", {})
    if not isinstance(overrides, dict):
        return {}
    normalized: dict[str, str] = {}
    for key, value in overrides.items():
        if not isinstance(key, str):
            continue
        if not isinstance(value, str):
            continue
        text = value.strip()
        if text:
            normalized[key] = text
    return normalized


def save_formula_overrides(path: Path, overrides: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "updated_at": now_iso(),
        "formula_overrides": dict(sorted(overrides.items())),
    }
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=True), encoding="utf-8")


class FormulaParseError(RuntimeError):
    pass


Token = tuple[str, str]


class FormulaTokenizer:
    number_re = re.compile(r"\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?")
    ident_re = re.compile(r"[A-Za-z0-9_.$]+")

    def tokenize(self, formula: str) -> list[Token]:
        text = formula.strip()
        if text.startswith("="):
            text = text[1:]

        tokens: list[Token] = []
        i = 0
        n = len(text)
        while i < n:
            ch = text[i]
            if ch.isspace():
                i += 1
                continue

            if ch in "(),:!":
                token_type = {
                    "(": "LPAREN",
                    ")": "RPAREN",
                    ",": "COMMA",
                    ":": "COLON",
                    "!": "BANG",
                }[ch]
                tokens.append((token_type, ch))
                i += 1
                continue

            if text.startswith("<=", i) or text.startswith(">=", i) or text.startswith("<>", i):
                tokens.append(("OP", text[i : i + 2]))
                i += 2
                continue

            if ch in "+-*/^&=<>" and ch != ":":
                tokens.append(("OP", ch))
                i += 1
                continue

            if ch == '"':
                i += 1
                chars: list[str] = []
                while i < n:
                    if text[i] == '"' and i + 1 < n and text[i + 1] == '"':
                        chars.append('"')
                        i += 2
                        continue
                    if text[i] == '"':
                        i += 1
                        break
                    chars.append(text[i])
                    i += 1
                tokens.append(("STRING", "".join(chars)))
                continue

            if ch == "'":
                i += 1
                chars = []
                while i < n:
                    if text[i] == "'" and i + 1 < n and text[i + 1] == "'":
                        chars.append("'")
                        i += 2
                        continue
                    if text[i] == "'":
                        i += 1
                        break
                    chars.append(text[i])
                    i += 1
                tokens.append(("IDENT", "".join(chars)))
                continue

            number_match = self.number_re.match(text, i)
            if number_match:
                tokens.append(("NUMBER", number_match.group(0)))
                i = number_match.end()
                continue

            ident_match = self.ident_re.match(text, i)
            if ident_match:
                tokens.append(("IDENT", ident_match.group(0)))
                i = ident_match.end()
                continue

            raise FormulaParseError(f"Unexpected token near: {text[i:i+20]!r}")

        tokens.append(("EOF", ""))
        return tokens


class FormulaParser:
    def __init__(self, formula: str) -> None:
        self.tokens = FormulaTokenizer().tokenize(formula)
        self.pos = 0

    def parse(self):
        node = self._parse_comparison()
        self._expect("EOF")
        return node

    def _peek(self) -> Token:
        return self.tokens[self.pos]

    def _consume(self) -> Token:
        token = self.tokens[self.pos]
        self.pos += 1
        return token

    def _accept(self, token_type: str, value: str | None = None) -> bool:
        t_type, t_value = self._peek()
        if t_type != token_type:
            return False
        if value is not None and t_value != value:
            return False
        self.pos += 1
        return True

    def _expect(self, token_type: str, value: str | None = None) -> Token:
        t_type, t_value = self._peek()
        if t_type != token_type or (value is not None and t_value != value):
            wanted = f"{token_type}:{value}" if value is not None else token_type
            got = f"{t_type}:{t_value}"
            raise FormulaParseError(f"Expected {wanted} but got {got}")
        return self._consume()

    def _parse_comparison(self):
        node = self._parse_concat()
        while True:
            t_type, t_value = self._peek()
            if t_type == "OP" and t_value in {"=", "<>", "<", ">", "<=", ">="}:
                self._consume()
                right = self._parse_concat()
                node = ("bin", t_value, node, right)
                continue
            break
        return node

    def _parse_concat(self):
        node = self._parse_additive()
        while True:
            t_type, t_value = self._peek()
            if t_type == "OP" and t_value == "&":
                self._consume()
                right = self._parse_additive()
                node = ("bin", "&", node, right)
                continue
            break
        return node

    def _parse_additive(self):
        node = self._parse_multiplicative()
        while True:
            t_type, t_value = self._peek()
            if t_type == "OP" and t_value in {"+", "-"}:
                self._consume()
                right = self._parse_multiplicative()
                node = ("bin", t_value, node, right)
                continue
            break
        return node

    def _parse_multiplicative(self):
        node = self._parse_power()
        while True:
            t_type, t_value = self._peek()
            if t_type == "OP" and t_value in {"*", "/"}:
                self._consume()
                right = self._parse_power()
                node = ("bin", t_value, node, right)
                continue
            break
        return node

    def _parse_power(self):
        node = self._parse_unary()
        t_type, t_value = self._peek()
        if t_type == "OP" and t_value == "^":
            self._consume()
            right = self._parse_power()
            node = ("bin", "^", node, right)
        return node

    def _parse_unary(self):
        t_type, t_value = self._peek()
        if t_type == "OP" and t_value in {"+", "-"}:
            self._consume()
            node = self._parse_unary()
            return ("unary", t_value, node)
        return self._parse_primary()

    def _parse_primary(self):
        t_type, t_value = self._peek()
        if t_type == "NUMBER":
            self._consume()
            return ("num", float(t_value))

        if t_type == "STRING":
            self._consume()
            return ("str", t_value)

        if t_type == "LPAREN":
            self._consume()
            node = self._parse_comparison()
            self._expect("RPAREN")
            return node

        if t_type == "IDENT":
            ident = self._consume()[1]
            if self._accept("LPAREN"):
                args: list = []
                if not self._accept("RPAREN"):
                    while True:
                        args.append(self._parse_comparison())
                        if self._accept("COMMA"):
                            continue
                        self._expect("RPAREN")
                        break
                return ("func", ident.upper(), args)

            if self._accept("BANG"):
                cell = self._parse_cell_ref(sheet=ident)
                return self._parse_range_suffix(cell)

            upper_ident = ident.upper()
            if upper_ident == "TRUE":
                return ("bool", True)
            if upper_ident == "FALSE":
                return ("bool", False)

            if CELL_REF_RE.match(ident.replace("$", "").upper()):
                cell = ("cell", None, normalize_cell_address(ident))
                return self._parse_range_suffix(cell)

            raise FormulaParseError(f"Unexpected identifier: {ident}")

        raise FormulaParseError(f"Unexpected token: {t_type}:{t_value}")

    def _parse_cell_ref(self, sheet: str | None):
        token_type, token_value = self._expect("IDENT")
        _ = token_type
        if not CELL_REF_RE.match(token_value.replace("$", "").upper()):
            raise FormulaParseError(f"Expected cell ref after sheet name but got {token_value}")
        return ("cell", sheet, normalize_cell_address(token_value))

    def _parse_ref_endpoint(self, default_sheet: str | None):
        token_type, token_value = self._expect("IDENT")
        _ = token_type
        if self._accept("BANG"):
            sheet = token_value
            _, ref = self._expect("IDENT")
            if not CELL_REF_RE.match(ref.replace("$", "").upper()):
                raise FormulaParseError(f"Invalid range endpoint ref: {ref}")
            return ("cell", sheet, normalize_cell_address(ref))
        if not CELL_REF_RE.match(token_value.replace("$", "").upper()):
            raise FormulaParseError(f"Invalid range endpoint ref: {token_value}")
        return ("cell", default_sheet, normalize_cell_address(token_value))

    def _parse_range_suffix(self, start_cell):
        if not self._accept("COLON"):
            return start_cell
        start_sheet = start_cell[1]
        end_cell = self._parse_ref_endpoint(default_sheet=start_sheet)
        return ("range", start_cell, end_cell)


class PurePythonUnderwriterEngine:
    def __init__(self, model: dict) -> None:
        self.model = model
        self._lock = threading.Lock()

        self.input_meta: dict[str, dict] = {
            entry["key"]: entry for entry in model.get("input_cells", [])
        }
        self.formula_keys: list[str] = [entry["key"] for entry in model.get("formula_cells", []) if entry.get("key")]

        self.base_values: dict[str, object] = {}
        self.formulas: dict[str, str] = {}
        for sheet in model.get("sheets", []):
            for cell in sheet.get("cells", []):
                key = cell["key"]
                self.base_values[key] = cell.get("value", "")
                if cell.get("has_formula"):
                    self.formulas[key] = cell.get("formula", "")

        self.formula_ast: dict[str, object] = {}
        for key, formula in self.formulas.items():
            if not formula:
                raise RuntimeError(f"Missing formula text for {key}")
            self.formula_ast[key] = FormulaParser(formula).parse()

    def _to_positive_int(self, value: object, fallback: int) -> int:
        if isinstance(value, bool):
            return fallback
        if isinstance(value, int):
            return max(1, value)
        if isinstance(value, float):
            if math.isnan(value) or math.isinf(value):
                return fallback
            return max(1, int(round(value)))
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return fallback
            try:
                return max(1, int(float(text)))
            except ValueError:
                return fallback
        return fallback

    def _to_clean_text(self, value: object) -> str:
        if value is None:
            return ""
        return str(value).strip()

    def _to_clean_number(self, value: object, fallback: float = 0.0) -> float:
        number = self._try_number(value)
        if number is None:
            return fallback
        if math.isnan(number) or math.isinf(number):
            return fallback
        return float(number)

    def _normalize_rent_roll_payload(self, payload: object) -> dict | None:
        if not isinstance(payload, dict):
            return None

        property_name = self._to_clean_text(payload.get("property_name"))
        property_address = self._to_clean_text(payload.get("property_address"))

        incoming_units = payload.get("units")
        unit_rows: list[dict] = []
        if isinstance(incoming_units, list):
            for raw_row in incoming_units:
                if not isinstance(raw_row, dict):
                    continue
                regular_rent = self._to_clean_number(raw_row.get("regular_rent"), 0.0)
                utilities = self._to_clean_number(raw_row.get("utilities"), 0.0)
                parking = self._to_clean_number(raw_row.get("parking"), 0.0)
                pet_fee = self._to_clean_number(raw_row.get("pet_fee"), 0.0)
                projected_rent = self._to_clean_number(raw_row.get("projected_rent"), 0.0)
                security_dep = self._to_clean_number(raw_row.get("security_dep"), 0.0)

                row = {
                    "tenant_name": self._to_clean_text(raw_row.get("tenant_name")),
                    "unit": self._to_clean_text(raw_row.get("unit")),
                    "regular_rent": regular_rent,
                    "utilities": utilities,
                    "parking": parking,
                    "pet_fee": pet_fee,
                    "projected_rent": projected_rent,
                    "lease_date": self._to_clean_text(raw_row.get("lease_date")),
                    "lease_end": self._to_clean_text(raw_row.get("lease_end")),
                    "security_dep": security_dep,
                }
                row["total_rent"] = regular_rent + utilities + parking + pet_fee
                unit_rows.append(row)

        requested_count = self._to_positive_int(payload.get("unit_count"), max(1, len(unit_rows) or 1))
        if len(unit_rows) < requested_count:
            for _ in range(requested_count - len(unit_rows)):
                unit_rows.append(
                    {
                        "tenant_name": "",
                        "unit": "",
                        "regular_rent": 0.0,
                        "utilities": 0.0,
                        "parking": 0.0,
                        "pet_fee": 0.0,
                        "projected_rent": 0.0,
                        "lease_date": "",
                        "lease_end": "",
                        "security_dep": 0.0,
                        "total_rent": 0.0,
                    }
                )
        elif len(unit_rows) > requested_count:
            unit_rows = unit_rows[:requested_count]

        totals = {
            "utilities": sum(row["utilities"] for row in unit_rows),
            "parking": sum(row["parking"] for row in unit_rows),
            "pet_fee": sum(row["pet_fee"] for row in unit_rows),
            "total_rent": sum(row["total_rent"] for row in unit_rows),
            "projected_rent": sum(row["projected_rent"] for row in unit_rows),
            "security_dep": sum(row["security_dep"] for row in unit_rows),
        }

        property_parts = [part for part in [property_name, property_address] if part]
        combined_property = " - ".join(property_parts)

        return {
            "property_name": property_name,
            "property_address": property_address,
            "combined_property": combined_property,
            "unit_count": requested_count,
            "units": unit_rows,
            "totals": totals,
        }

    def _apply_dynamic_rent_roll(
        self,
        state: dict[str, object],
        normalized_inputs: dict[str, object],
        formula_overrides: dict[str, object],
        rent_roll: dict,
    ) -> None:
        if rent_roll["combined_property"]:
            state["Valuation!D4"] = rent_roll["combined_property"]
            normalized_inputs["Valuation!D4"] = rent_roll["combined_property"]

        state["Valuation!D5"] = rent_roll["unit_count"]
        normalized_inputs["Valuation!D5"] = rent_roll["unit_count"]

        # Workbook has canonical unit rows 6..13. We hydrate these for compatibility,
        # while aggregate totals use dynamic row count via formula overrides below.
        base_row = 6
        workbook_row_count = 8
        for idx in range(workbook_row_count):
            row_index = base_row + idx
            if idx < len(rent_roll["units"]):
                row = rent_roll["units"][idx]
                state[f"Rent Roll!A{row_index}"] = row["tenant_name"]
                state[f"Rent Roll!B{row_index}"] = row["unit"]
                # Rent Roll Rev 5.3.2026 layout:
                # C=#units, D=regular, E=utilities, F=parking, G=pet fee,
                # H=total rent(formula), I=total all units(formula), J=projected,
                # K=lease date, L=lease end, M=security dep(formula in template).
                state[f"Rent Roll!C{row_index}"] = 1
                state[f"Rent Roll!D{row_index}"] = row["regular_rent"]
                state[f"Rent Roll!E{row_index}"] = row["utilities"]
                state[f"Rent Roll!F{row_index}"] = row["parking"]
                state[f"Rent Roll!G{row_index}"] = row["pet_fee"]
                state[f"Rent Roll!J{row_index}"] = row["projected_rent"]
                state[f"Rent Roll!K{row_index}"] = row["lease_date"]
                state[f"Rent Roll!L{row_index}"] = row["lease_end"]
            else:
                state[f"Rent Roll!A{row_index}"] = ""
                state[f"Rent Roll!B{row_index}"] = ""
                state[f"Rent Roll!C{row_index}"] = 0
                state[f"Rent Roll!D{row_index}"] = 0
                state[f"Rent Roll!E{row_index}"] = 0
                state[f"Rent Roll!F{row_index}"] = 0
                state[f"Rent Roll!G{row_index}"] = 0
                state[f"Rent Roll!J{row_index}"] = 0
                state[f"Rent Roll!K{row_index}"] = ""
                state[f"Rent Roll!L{row_index}"] = ""

        totals = rent_roll["totals"]
        state["Rent Roll!E15"] = totals["utilities"]
        state["Rent Roll!I15"] = totals["total_rent"]
        state["Rent Roll!J15"] = totals["projected_rent"]

        formula_overrides["Rent Roll!E15"] = totals["utilities"]
        formula_overrides["Rent Roll!I15"] = totals["total_rent"]
        formula_overrides["Rent Roll!J15"] = totals["projected_rent"]

    def _to_number(self, value: object) -> float:
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if value is None:
            return 0.0
        if isinstance(value, str):
            if value.startswith("#"):
                return 0.0
            text = value.strip().replace(",", "")
            if text in {"", "-"}:
                return 0.0
            try:
                return float(text)
            except ValueError:
                return 0.0
        return 0.0

    def _try_number(self, value: object) -> float | None:
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            if value.startswith("#"):
                return None
            text = value.strip().replace(",", "")
            if text == "":
                return 0.0
            try:
                return float(text)
            except ValueError:
                return None
        if value is None:
            return 0.0
        return None

    def _is_error(self, value: object) -> bool:
        return isinstance(value, str) and value.startswith("#")

    def _truthy(self, value: object) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return value != 0
        if isinstance(value, str):
            return value.strip() != ""
        return bool(value)

    def _normalize_number(self, value: float) -> int | float:
        if math.isfinite(value) and abs(value - round(value)) < 1e-12:
            return int(round(value))
        return value

    def _flatten(self, value: object) -> list[object]:
        if isinstance(value, list):
            out: list[object] = []
            for item in value:
                out.extend(self._flatten(item))
            return out
        return [value]

    def _resolve_cell(
        self,
        sheet_name: str,
        address: str,
        state: dict[str, object],
        cache: dict[str, object],
        visiting: set[str],
    ) -> object:
        key = f"{sheet_name}!{address}"
        if key in self.formula_ast:
            return self._evaluate_formula_cell(key, state, cache, visiting)
        return state.get(key, 0)

    def _resolve_range(
        self,
        start_cell,
        end_cell,
        current_sheet: str,
        state: dict[str, object],
        cache: dict[str, object],
        visiting: set[str],
    ) -> list[object]:
        start_sheet = start_cell[1] or current_sheet
        end_sheet = end_cell[1] or start_sheet
        if start_sheet != end_sheet:
            raise RuntimeError(
                f"Cross-sheet ranges are not supported: {start_sheet}:{end_sheet}"
            )
        values: list[object] = []
        for address in iter_range_addresses(start_cell[2], end_cell[2]):
            values.append(self._resolve_cell(start_sheet, address, state, cache, visiting))
        return values

    def _apply_binary_op(self, op: str, left: object, right: object) -> object:
        if self._is_error(left):
            return left
        if self._is_error(right):
            return right

        if op == "&":
            return f"{left}{right}"

        if op in {"=", "<>", "<", ">", "<=", ">="}:
            left_num = self._try_number(left)
            right_num = self._try_number(right)
            if left_num is not None and right_num is not None:
                a = left_num
                b = right_num
            else:
                a = str(left)
                b = str(right)
            if op == "=":
                return a == b
            if op == "<>":
                return a != b
            if op == "<":
                return a < b
            if op == ">":
                return a > b
            if op == "<=":
                return a <= b
            return a >= b

        a = self._try_number(left)
        b = self._try_number(right)
        if a is None or b is None:
            return "#VALUE!"
        if op == "+":
            return self._normalize_number(a + b)
        if op == "-":
            return self._normalize_number(a - b)
        if op == "*":
            return self._normalize_number(a * b)
        if op == "/":
            if abs(b) < 1e-15:
                return "#DIV/0!"
            return self._normalize_number(a / b)
        if op == "^":
            return self._normalize_number(a**b)
        raise RuntimeError(f"Unsupported operator: {op}")

    def _apply_function(self, name: str, args: list[object]) -> object:
        if name == "IF":
            if len(args) < 2:
                raise RuntimeError("IF requires at least 2 args")
            cond = args[0]
            if self._is_error(cond):
                return cond
            true_value = args[1]
            false_value = args[2] if len(args) >= 3 else 0
            return true_value if self._truthy(cond) else false_value

        if name == "SUM":
            total = 0.0
            for arg in args:
                for item in self._flatten(arg):
                    if self._is_error(item):
                        return item
                    number = self._try_number(item)
                    if number is not None:
                        total += number
            return self._normalize_number(total)

        if name == "MIN":
            numbers: list[float] = []
            for arg in args:
                for item in self._flatten(arg):
                    if self._is_error(item):
                        return item
                    number = self._try_number(item)
                    if number is not None:
                        numbers.append(number)
            if not numbers:
                return 0
            return self._normalize_number(min(numbers))

        if name == "EXP":
            if not args:
                return 0
            if self._is_error(args[0]):
                return args[0]
            number = self._try_number(args[0])
            if number is None:
                return "#VALUE!"
            return self._normalize_number(math.exp(number))

        if name == "LN":
            if not args:
                return 0
            if self._is_error(args[0]):
                return args[0]
            number = self._try_number(args[0])
            if number is None:
                return "#VALUE!"
            value = number
            if value <= 0:
                return "#NUM!"
            return self._normalize_number(math.log(value))

        if name == "PV":
            if len(args) < 3:
                raise RuntimeError("PV requires at least 3 args")
            for arg in args:
                if self._is_error(arg):
                    return arg
            rate = self._try_number(args[0])
            nper = self._try_number(args[1])
            pmt = self._try_number(args[2])
            fv = self._try_number(args[3]) if len(args) >= 4 else 0.0
            when = self._try_number(args[4]) if len(args) >= 5 else 0.0
            if None in {rate, nper, pmt, fv, when}:
                return "#VALUE!"
            if abs(rate) < 1e-15:
                return self._normalize_number(-(fv + pmt * nper))
            factor = (1 + rate) ** nper
            pv = -(fv + pmt * (1 + rate * when) * (factor - 1) / rate) / factor
            return self._normalize_number(pv)

        raise RuntimeError(f"Unsupported function: {name}")

    def _eval_ast(
        self,
        node,
        current_sheet: str,
        state: dict[str, object],
        cache: dict[str, object],
        visiting: set[str],
    ) -> object:
        node_type = node[0]
        if node_type == "num":
            return self._normalize_number(node[1])
        if node_type == "str":
            return node[1]
        if node_type == "bool":
            return node[1]
        if node_type == "cell":
            sheet = node[1] or current_sheet
            address = node[2]
            return self._resolve_cell(sheet, address, state, cache, visiting)
        if node_type == "range":
            return self._resolve_range(node[1], node[2], current_sheet, state, cache, visiting)
        if node_type == "unary":
            op = node[1]
            value = self._eval_ast(node[2], current_sheet, state, cache, visiting)
            number = self._to_number(value)
            if op == "+":
                return self._normalize_number(number)
            return self._normalize_number(-number)
        if node_type == "bin":
            left = self._eval_ast(node[2], current_sheet, state, cache, visiting)
            right = self._eval_ast(node[3], current_sheet, state, cache, visiting)
            return self._apply_binary_op(node[1], left, right)
        if node_type == "func":
            args = [
                self._eval_ast(arg, current_sheet, state, cache, visiting)
                for arg in node[2]
            ]
            return self._apply_function(node[1], args)
        raise RuntimeError(f"Unsupported AST node: {node_type}")

    def _evaluate_formula_cell(
        self,
        key: str,
        state: dict[str, object],
        cache: dict[str, object],
        visiting: set[str],
    ) -> object:
        if key in cache:
            return cache[key]
        if key in visiting:
            raise RuntimeError(f"Cyclic formula dependency detected: {key}")

        visiting.add(key)
        try:
            sheet_name, _ = key.split("!", 1)
            value = self._eval_ast(self.formula_ast[key], sheet_name, state, cache, visiting)
            cache[key] = value
            state[key] = value
            return value
        finally:
            visiting.remove(key)

    def _coerce_input_value(self, raw: object, value_type: str, default_value: object) -> object:
        if raw is None:
            raw = default_value
        if value_type == "number":
            if isinstance(raw, (int, float)):
                return raw
            text = str(raw).strip()
            if text == "":
                return 0
            try:
                return parse_number(text)
            except ValueError:
                if isinstance(default_value, (int, float)):
                    return default_value
                return 0
        if value_type == "bool":
            if isinstance(raw, bool):
                return raw
            return str(raw).strip().lower() in {"1", "true", "yes", "y"}
        return str(raw)

    def calculate(
        self,
        incoming_inputs: dict[str, object],
        rent_roll_payload: object | None = None,
    ) -> dict[str, object]:
        with self._lock:
            state: dict[str, object] = dict(self.base_values)
            normalized_inputs: dict[str, object] = {}
            for key, meta in self.input_meta.items():
                normalized = self._coerce_input_value(
                    incoming_inputs.get(key),
                    meta.get("value_type", "text"),
                    meta.get("default_value"),
                )
                normalized_inputs[key] = normalized
                state[key] = normalized

            formula_overrides: dict[str, object] = {}
            normalized_rent_roll = self._normalize_rent_roll_payload(rent_roll_payload)
            if normalized_rent_roll is not None:
                self._apply_dynamic_rent_roll(
                    state=state,
                    normalized_inputs=normalized_inputs,
                    formula_overrides=formula_overrides,
                    rent_roll=normalized_rent_roll,
                )

            cache: dict[str, object] = dict(formula_overrides)
            for key, value in formula_overrides.items():
                state[key] = value
            visiting: set[str] = set()
            formula_values: dict[str, object] = {}
            for key in self.formula_keys:
                if key not in self.formula_ast:
                    continue
                if key in cache:
                    formula_values[key] = cache[key]
                else:
                    formula_values[key] = self._evaluate_formula_cell(key, state, cache, visiting)

            return {
                "calculated_at": now_iso(),
                "inputs": normalized_inputs,
                "formula_values": formula_values,
                "rent_roll": normalized_rent_roll,
            }


class AppState:
    def __init__(self, base_model: dict, overrides_path: Path) -> None:
        self.base_model = deep_clone(base_model)
        self.overrides_path = overrides_path
        self._lock = threading.RLock()

        self.mortgage_default_formulas = self._collect_mortgage_defaults(self.base_model)
        loaded_overrides = load_formula_overrides(self.overrides_path)
        self.formula_overrides: dict[str, str] = {}
        for key, value in loaded_overrides.items():
            if key not in self.mortgage_default_formulas:
                continue
            if value != self.mortgage_default_formulas[key]:
                self.formula_overrides[key] = value

        self.model_full: dict = {}
        self.model_public: dict = {}
        self.engine: PurePythonUnderwriterEngine | None = None
        self.mortgage_admin_records: list[dict] = []
        self._rebuild_locked()

    def _collect_mortgage_defaults(self, model: dict) -> dict[str, str]:
        defaults: dict[str, str] = {}
        for sheet in model.get("sheets", []):
            if sheet.get("name") != MORTGAGE_SHEET_NAME:
                continue
            for cell in sheet.get("cells", []):
                if cell.get("has_formula"):
                    key = cell.get("key")
                    if key:
                        defaults[key] = cell.get("formula", "")
        return defaults

    def _build_mortgage_records(self, model_full: dict) -> list[dict]:
        records: list[dict] = []
        for sheet in model_full.get("sheets", []):
            if sheet.get("name") != MORTGAGE_SHEET_NAME:
                continue
            for cell in sheet.get("cells", []):
                if not cell.get("has_formula"):
                    continue
                key = cell.get("key")
                if not key:
                    continue
                default_formula = self.mortgage_default_formulas.get(key, cell.get("formula", ""))
                current_formula = cell.get("formula", "")
                records.append(
                    {
                        "key": key,
                        "address": cell.get("address", ""),
                        "row": cell.get("row", 0),
                        "col": cell.get("col", 0),
                        "default_formula": default_formula,
                        "current_formula": current_formula,
                        "is_overridden": key in self.formula_overrides,
                    }
                )
        records.sort(key=lambda item: (item.get("row", 0), item.get("col", 0)))
        return records

    def _rebuild_locked(self) -> None:
        model_full = deep_clone(self.base_model)
        apply_formula_overrides(model_full, self.formula_overrides)
        apply_manual_input_cell_overrides(model_full)

        self.model_full = model_full
        self.model_public = filter_public_model(model_full)
        self.engine = PurePythonUnderwriterEngine(model_full)
        self.mortgage_admin_records = self._build_mortgage_records(model_full)

    def get_public_model(self) -> dict:
        with self._lock:
            return deep_clone(self.model_public)

    def calculate(self, inputs: dict[str, object], rent_roll_payload: object | None = None) -> dict:
        with self._lock:
            engine = self.engine
        if engine is None:
            raise RuntimeError("Engine not initialized")
        return engine.calculate(inputs, rent_roll_payload=rent_roll_payload)

    def get_mortgage_admin_payload(self) -> dict:
        with self._lock:
            return {
                "sheet": MORTGAGE_SHEET_NAME,
                "total_records": len(self.mortgage_admin_records),
                "overrides_count": len(self.formula_overrides),
                "overrides_path": str(self.overrides_path.resolve()),
                "records": deep_clone(self.mortgage_admin_records),
            }

    def update_mortgage_overrides(self, incoming_overrides: dict[str, object]) -> dict:
        normalized: dict[str, str] = {}
        for key, raw_formula in incoming_overrides.items():
            if not isinstance(key, str):
                raise ValueError("Override keys must be strings.")
            if key not in self.mortgage_default_formulas:
                raise ValueError(f"Unsupported override key: {key}")
            if not isinstance(raw_formula, str):
                raise ValueError(f"Formula override for {key} must be a string.")

            formula = raw_formula.strip()
            if formula.startswith("="):
                formula = formula[1:].strip()
            if not formula:
                continue

            try:
                FormulaParser(formula).parse()
            except FormulaParseError as exc:
                raise ValueError(f"Invalid formula for {key}: {exc}") from exc

            if formula != self.mortgage_default_formulas[key]:
                normalized[key] = formula

        with self._lock:
            self.formula_overrides = normalized
            save_formula_overrides(self.overrides_path, self.formula_overrides)
            self._rebuild_locked()
            return {
                "updated_at": now_iso(),
                "overrides_count": len(self.formula_overrides),
                "total_records": len(self.mortgage_admin_records),
            }


class UnderwriterHandler(BaseHTTPRequestHandler):
    state: AppState | None = None

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        # Block cross-origin requests from non-whitelisted origins on API routes
        if parsed.path.startswith("/api/") and self._cors_origin() == "__blocked__":
            self._send_json({"error": "Origin not allowed"}, status=403)
            return
        if parsed.path == "/api/model":
            user, blocked = self._require_auth()
            if blocked:
                return
            if self.state is None:
                self._send_json({"error": "Server not initialized"}, status=500)
                return
            self._send_json(self.state.get_public_model())
            return
        if parsed.path == "/api/admin/mortgage":
            user, blocked = self._require_auth(required_role="admin")
            if blocked:
                return
            if self.state is None:
                self._send_json({"error": "Server not initialized"}, status=500)
                return
            self._send_json(self.state.get_mortgage_admin_payload())
            return
        self._serve_static(parsed.path)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        # Block cross-origin requests from non-whitelisted origins
        if self._cors_origin() == "__blocked__":
            self._send_json({"error": "Origin not allowed"}, status=403)
            return
        if parsed.path not in {"/api/calculate", "/api/admin/mortgage-overrides"}:
            self._send_json({"error": "Not found"}, status=404)
            return
        # Auth gate
        required_role = "admin" if parsed.path == "/api/admin/mortgage-overrides" else None
        user, blocked = self._require_auth(required_role=required_role)
        if blocked:
            return
        if self.state is None:
            self._send_json({"error": "Server not initialized"}, status=500)
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(content_length) if content_length else b"{}"
        try:
            payload = json.loads(raw_body.decode("utf-8"))
        except json.JSONDecodeError:
            self._send_json({"error": "Invalid JSON body"}, status=400)
            return

        try:
            if parsed.path == "/api/calculate":
                if not isinstance(payload, dict) or not isinstance(payload.get("inputs"), dict):
                    self._send_json({"error": "Body must include an inputs object"}, status=400)
                    return
                result = self.state.calculate(
                    inputs=payload["inputs"],
                    rent_roll_payload=payload.get("rent_roll"),
                )
            else:
                if not isinstance(payload, dict) or not isinstance(payload.get("overrides"), dict):
                    self._send_json({"error": "Body must include an overrides object"}, status=400)
                    return
                result = self.state.update_mortgage_overrides(payload["overrides"])
        except ValueError as exc:
            self._send_json({"error": str(exc)}, status=400)
            return
        except Exception as exc:
            self._send_json({"error": str(exc)}, status=500)
            return
        self._send_json(result)

    def _send_json(self, obj: dict, status: int = 200) -> None:
        payload = json.dumps(obj, ensure_ascii=True).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self._write_cors_headers()
        self.end_headers()
        self.wfile.write(payload)

    def _serve_static(self, raw_path: str) -> None:
        path = raw_path or "/"
        if path == "/":
            path = "/index.html"
        candidate = (STATIC_DIR / path.lstrip("/")).resolve()
        static_root = STATIC_DIR.resolve()
        if static_root not in candidate.parents and candidate != static_root:
            self._send_json({"error": "Invalid path"}, status=400)
            return
        if not candidate.exists() or not candidate.is_file():
            self._send_json({"error": "Not found"}, status=404)
            return

        mime_type, _ = mimetypes.guess_type(str(candidate))
        data = candidate.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime_type or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self._write_cors_headers()
        self.end_headers()
        self.wfile.write(data)


    # ── CORS helpers ──────────────────────────────────────────────────────────

    def _cors_origin(self) -> str:
        """Return the CORS origin to echo, '*' for open, '' if blocked.

        - No Origin header  → same-origin / direct request → always ok (return '')
          (we treat '' as "no CORS header needed", not blocked in this case — callers
          handle the '' sentinel only when an Origin header *was* provided)
        - Origin present + whitelist empty → public mode → return '*'
        - Origin present + in whitelist    → return origin
        - Origin present + NOT in whitelist → blocked → return sentinel '__blocked__'
        """
        origin = self.headers.get("Origin", "")
        if not origin:
            return ""  # same-origin; no CORS header needed
        allowed: set[str] = getattr(self.server, "allowed_origins", set())
        if not allowed:
            return "*"
        if origin in allowed:
            return origin
        return "__blocked__"

    def _write_cors_headers(self) -> None:
        """Write CORS headers if applicable. Call before end_headers()."""
        cors = self._cors_origin()
        if not cors or cors == "__blocked__":
            return
        self.send_header("Access-Control-Allow-Origin", cors)
        if cors != "*":
            self.send_header("Vary", "Origin")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")

    def do_OPTIONS(self) -> None:
        """Handle CORS preflight requests."""
        cors = self._cors_origin()
        if cors == "__blocked__":
            self.send_response(403)
            self.end_headers()
            return
        self.send_response(204)
        if cors:
            self.send_header("Access-Control-Allow-Origin", cors)
            if cors != "*":
                self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
            self.send_header("Access-Control-Max-Age", "86400")
        self.end_headers()

    # ── Auth helper ───────────────────────────────────────────────────────────

    def _require_auth(
        self, required_role: str | None = None
    ) -> tuple[dict | None, bool]:
        """Validate the Bearer token and optionally check role.

        Returns (payload, False) on success, (None, True) after sending the
        error response on failure.
        """
        auth_header = self.headers.get("Authorization", "")
        if not auth_header.startswith("Bearer "):
            self._send_json({"error": "Authentication required"}, status=401)
            return None, True
        token = auth_header[len("Bearer "):]
        payload = _verify_supabase_jwt(token)
        if payload is None:
            self._send_json({"error": "Invalid or expired token"}, status=401)
            return None, True
        if required_role:
            role = (payload.get("app_metadata") or payload.get("raw_app_meta_data") or {}).get("role", "analyst")
            if role != required_role:
                self._send_json({"error": "Insufficient permissions"}, status=403)
                return None, True
        return payload, False

    def log_message(self, fmt: str, *args: object) -> None:  # type: ignore[override]
        pass  # Suppress default request logging noise


class UnderwriterServer(ThreadingHTTPServer):
    """ThreadingHTTPServer that carries CORS origin whitelist on the server object."""

    def __init__(
        self,
        server_address: tuple[str, int],
        RequestHandlerClass: type,
        allowed_origins: set[str] | None = None,
    ) -> None:
        super().__init__(server_address, RequestHandlerClass)
        self.allowed_origins: set[str] = allowed_origins or set()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Pure Python Multifamily Underwriter (no Excel dependency)"
    )
    parser.add_argument("--host", default=os.environ.get("HOST", "0.0.0.0"))
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8000")))
    parser.add_argument(
        "--allowed-origins",
        default=os.environ.get("ALLOWED_ORIGINS", ""),
        help="Comma-separated list of allowed CORS origins (empty = allow all)",
    )
    parser.add_argument("--model", default=str(MODEL_PATH), help="Path to workbook model JSON")
    parser.add_argument(
        "--formulas",
        default=str(FORMULAS_PATH),
        help="Path to expanded formulas JSON",
    )
    parser.add_argument(
        "--overrides",
        default=str(OVERRIDES_PATH),
        help="Path to admin formula override JSON",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    STATIC_DIR.mkdir(parents=True, exist_ok=True)

    model = load_model(Path(args.model).expanduser(), Path(args.formulas).expanduser())
    overrides_path = Path(args.overrides).expanduser()
    UnderwriterHandler.state = AppState(base_model=model, overrides_path=overrides_path)

    allowed_origins: set[str] = (
        {o.strip() for o in args.allowed_origins.split(",") if o.strip()}
        if args.allowed_origins.strip()
        else set()
    )

    if not SUPABASE_JWT_SECRET:
        print("WARNING: SUPABASE_JWT_SECRET is not set. HS256 JWT verification will fail. "
              "(ES256/JWKS projects issued after Aug 2024 do not need this.)")

    server = UnderwriterServer((args.host, args.port), UnderwriterHandler, allowed_origins)
    print(f"Underwriter running at http://{args.host}:{args.port}")
    print("Calculation engine: pure Python (no Excel runtime dependency)")
    if allowed_origins:
        print(f"CORS: restricted to {', '.join(sorted(allowed_origins))}")
    else:
        print("CORS: open (no origin restrictions)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
