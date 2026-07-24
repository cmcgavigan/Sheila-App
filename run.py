#!/usr/bin/env python3
"""Entry point: prepare the TLS cert, then serve the app over HTTPS.

Run directly (python run.py) or as the Windows service installed by
install/INSTALL.cmd — same thing either way.
"""
import uvicorn

from app import config, netinfo


def main() -> None:
    ips = netinfo.get_local_ips()
    hostnames = [config.HOSTNAME_LOCAL]
    ts_info = netinfo.tailscale_info()
    if ts_info:
        hostnames.append(ts_info["dns"])
    netinfo.load_or_generate_cert(ips, hostnames)

    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=config.PORT,
        ssl_keyfile=str(netinfo.CERT_KEY),
        ssl_certfile=str(netinfo.CERT_PEM),
        log_level="info",
    )


if __name__ == "__main__":
    main()
