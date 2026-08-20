from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None

from ebay_api import EbayConfig, EbayOAuthError, exchange_authorization_code, mask_token, refresh_access_token, update_env_values


def load_env(env_path: Path) -> None:
    if load_dotenv:
        load_dotenv(env_path, override=True)
        return
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


def main() -> int:
    parser = argparse.ArgumentParser(description="Check or exchange LUCAS eBay OAuth credentials without printing secrets.")
    parser.add_argument("--env-file", default=str(ROOT / ".env"), help="Path to the LUCAS .env file.")
    parser.add_argument("--exchange-code", action="store_true", help="Exchange a one-time eBay authorization code for a refresh token.")
    parser.add_argument("--code", default="", help="One-time authorization code. Prefer EBAY_AUTH_CODE in .env when possible.")
    parser.add_argument(
        "--code-from-refresh-token",
        action="store_true",
        help="Treat the current EBAY_REFRESH_TOKEN value as the one-time authorization code. Useful if the callback code was pasted there by mistake.",
    )
    parser.add_argument("--write-env", action="store_true", help="Write returned EBAY_REFRESH_TOKEN and EBAY_ACCESS_TOKEN back to .env.")
    args = parser.parse_args()

    env_path = Path(args.env_file).expanduser()
    load_env(env_path)
    config = EbayConfig.from_env()

    try:
        if args.exchange_code:
            code = args.code.strip() or os.environ.get("EBAY_AUTH_CODE", "").strip() or os.environ.get("EBAY_AUTHORIZATION_CODE", "").strip()
            if not code and args.code_from_refresh_token:
                code = os.environ.get("EBAY_REFRESH_TOKEN", "").strip()
            result = exchange_authorization_code(config, code)
            refresh_token = str(result.get("refresh_token") or "").strip()
            access_token = str(result.get("access_token") or "").strip()
            if not refresh_token:
                raise EbayOAuthError("eBay accepted the code but did not return a refresh_token.")
            print("eBay authorization code exchange: OK")
            print(f"Access token: {mask_token(access_token)}")
            print(f"Refresh token: {mask_token(refresh_token)}")
            print(f"Access expires in: {result.get('expires_in', 'unknown')} seconds")
            print(f"Refresh expires in: {result.get('refresh_token_expires_in', 'unknown')} seconds")
            if args.write_env:
                update_env_values(
                    env_path,
                    {
                        "EBAY_ACCESS_TOKEN": access_token,
                        "EBAY_REFRESH_TOKEN": refresh_token,
                    },
                )
                print(f"Updated {env_path}")
            return 0

        result = refresh_access_token(config, os.environ.get("EBAY_REFRESH_TOKEN", ""))
        access_token = str(result.get("access_token") or "").strip()
        if not access_token:
            raise EbayOAuthError("eBay did not return an access token from EBAY_REFRESH_TOKEN.")
        print("eBay refresh token check: OK")
        print(f"New access token: {mask_token(access_token)}")
        print(f"Access expires in: {result.get('expires_in', 'unknown')} seconds")
        if args.write_env:
            update_env_values(env_path, {"EBAY_ACCESS_TOKEN": access_token})
            print(f"Updated {env_path}")
        return 0
    except EbayOAuthError as error:
        print(f"eBay OAuth check failed: {error}", file=sys.stderr)
        print("", file=sys.stderr)
        print("If you pasted the callback page's giant code into EBAY_REFRESH_TOKEN, run:", file=sys.stderr)
        print("  python scripts/check_ebay_oauth.py --exchange-code --code-from-refresh-token --write-env", file=sys.stderr)
        print("If that says the code expired, sign in to Production again and put the new callback code in EBAY_AUTH_CODE.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
