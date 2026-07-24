"""Network + TLS: LAN IP discovery, Tailscale detection/serve, self-signed cert.

Same behaviour as v1: the app listens on HTTPS with a self-signed cert for
LAN/localhost, and if Tailscale is up with HTTPS certs enabled it maps a
tailnet HTTPS port (TS_HTTPS_PORT) to us, giving Sheila's phone a proper
no-warning URL that works from anywhere.
"""
from __future__ import annotations

import datetime
import ipaddress
import json
import socket
import subprocess
from pathlib import Path
from typing import List, Optional

from . import config

CERT_KEY = config.CERT_DIR / "key.pem"
CERT_PEM = config.CERT_DIR / "cert.pem"


def get_local_ips() -> List[str]:
    ips = set()
    try:
        hostname = socket.gethostname()
        for info in socket.getaddrinfo(hostname, None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127."):
                ips.add(ip)
    except OSError:
        pass
    # UDP trick — finds the outbound interface even when getaddrinfo is sparse.
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ips.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    return sorted(ips)


def get_tailscale_ip() -> str:
    for ip in get_local_ips():
        parts = ip.split(".")
        if parts[0] == "100" and 64 <= int(parts[1]) <= 127:
            return ip
    return ""


def _tailscale_bin() -> str:
    win = r"C:\Program Files\Tailscale\tailscale.exe"
    return win if Path(win).exists() else "tailscale"


def _run_tailscale(args: List[str]) -> Optional[str]:
    try:
        out = subprocess.run(
            [_tailscale_bin()] + args, capture_output=True, text=True, timeout=7,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return out.stdout if out.returncode == 0 else None
    except (OSError, subprocess.TimeoutExpired):
        return None


def tailscale_info() -> Optional[dict]:
    out = _run_tailscale(["status", "--json"])
    if not out:
        return None
    try:
        j = json.loads(out)
        dns = (j.get("Self", {}).get("DNSName") or "").rstrip(".")
        https_enabled = bool(j.get("CertDomains"))
        return {"dns": dns, "httpsEnabled": https_enabled} if dns else None
    except ValueError:
        return None


def ensure_tailscale_serve(local_port: int) -> str:
    """Map https://<tailnet-host>:TS_HTTPS_PORT -> local HTTPS port. Returns URL or ''."""
    info = tailscale_info()
    if not info:
        return ""
    if not info["httpsEnabled"]:
        print(f"  Tailscale:   {info['dns']} — enable HTTPS at "
              "https://login.tailscale.com/admin/dns for a no-warning cert")
        return ""
    _run_tailscale(["serve", "--bg", f"--https={config.TS_HTTPS_PORT}",
                    f"https+insecure://127.0.0.1:{local_port}"])
    suffix = "" if config.TS_HTTPS_PORT == 443 else f":{config.TS_HTTPS_PORT}"
    return f"https://{info['dns']}{suffix}"


def load_or_generate_cert(ips: List[str], hostnames: List[str]) -> bool:
    """Ensure key.pem/cert.pem exist and cover all names. True if regenerated."""
    if CERT_KEY.exists() and CERT_PEM.exists():
        try:
            from cryptography import x509
            from cryptography.hazmat.primitives import serialization
            cert = x509.load_pem_x509_certificate(CERT_PEM.read_bytes())
            key = serialization.load_pem_private_key(CERT_KEY.read_bytes(), password=None)
            now = datetime.datetime.now(datetime.timezone.utc)
            expiry = (cert.not_valid_after_utc if hasattr(cert, "not_valid_after_utc")
                      else cert.not_valid_after.replace(tzinfo=datetime.timezone.utc))
            names = {"localhost", *hostnames, *ips, "127.0.0.1"}
            san = cert.extensions.get_extension_for_class(
                x509.SubjectAlternativeName
            ).value
            cert_names = {str(n.value) for n in san}
            key_matches = cert.public_key().public_bytes(
                serialization.Encoding.DER,
                serialization.PublicFormat.SubjectPublicKeyInfo,
            ) == key.public_key().public_bytes(
                serialization.Encoding.DER,
                serialization.PublicFormat.SubjectPublicKeyInfo,
            )
            if now < expiry and expiry - now > datetime.timedelta(days=30) \
                    and names <= cert_names and key_matches:
                return False
            print("  TLS certificate names, expiry, or key changed - regenerating...")
        except OSError:
            print("  Existing cert unreadable, regenerating...")
        except Exception as e:  # noqa: BLE001 - invalid certs must rotate safely
            print(f"  Existing cert invalid ({e}), regenerating...")

    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, config.HOSTNAME_LOCAL)])
    alt_names: list = [x509.DNSName("localhost")]
    alt_names += [x509.DNSName(h) for h in hostnames]
    alt_names += [x509.IPAddress(ipaddress.ip_address("127.0.0.1"))]
    for ip in ips:
        try:
            alt_names.append(x509.IPAddress(ipaddress.ip_address(ip)))
        except ValueError:
            pass
    now = datetime.datetime.utcnow()
    cert = (
        x509.CertificateBuilder()
        .subject_name(name).issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=825))
        .add_extension(x509.SubjectAlternativeName(alt_names), critical=False)
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(digital_signature=True, key_encipherment=True,
                          content_commitment=False, data_encipherment=False,
                          key_agreement=False, key_cert_sign=False, crl_sign=False,
                          encipher_only=False, decipher_only=False),
            critical=False)
        .add_extension(
            x509.ExtendedKeyUsage([x509.oid.ExtendedKeyUsageOID.SERVER_AUTH]),
            critical=False)
        .sign(key, hashes.SHA256())
    )
    CERT_KEY.write_bytes(key.private_bytes(
        serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL,
        serialization.NoEncryption()))
    CERT_PEM.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    return True


def resolve_public_url() -> str:
    return config.PUBLIC_URL or ensure_tailscale_serve(config.PORT)
