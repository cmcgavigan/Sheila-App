#!/usr/bin/env python3
"""
watchdog.py - keep the receipt servers alive and reachable, alert if not.

Runs every 4 hours (Windows Task Scheduler). Each run it:
  1. Checks each server is reachable over its TAILSCALE TUNNEL (not just localhost).
  2. Restarts any server that's unreachable, via its launch.vbs.
  3. Tracks how long each server has been continuously failing WHILE THE LAPTOP
     WAS ON - gaps longer than the schedule (laptop off) do NOT count.
  4. Alerts via ntfy.sh after ALERT_AFTER_HOURS of real on-time failure.
  5. On the first run after a gap, checks the Windows event log: clean shutdown =
     silent; unexpected power loss = a notification (and whether servers recovered).

Config (.env): NTFY_TOPIC, NTFY_SERVER, WATCHDOG_ALERT_HOURS, WATCHDOG_GAP_MINUTES.
Test:  python watchdog.py            (one real cycle)
       python watchdog.py --dry-run  (checks only, no restart/alert)
"""
import os, sys, json, time, ssl, subprocess, datetime, urllib.request, urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = os.path.join(HERE, "watchdog-state.json")
LOG_FILE = os.path.join(HERE, "watchdog.log")
DRY_RUN = "--dry-run" in sys.argv


def env(key, default=None):
    v = os.environ.get(key)
    if v is not None:
        return v
    p = os.path.join(HERE, ".env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if line.startswith(key + "="):
                return line.split("=", 1)[1].strip()
    return default


NTFY_TOPIC = env("NTFY_TOPIC", "")
NTFY_SERVER = env("NTFY_SERVER", "https://ntfy.sh").rstrip("/")
ALERT_AFTER_HOURS = float(env("WATCHDOG_ALERT_HOURS", "24"))
GAP_MINUTES = float(env("WATCHDOG_GAP_MINUTES", "290"))

SERVERS = [
    {"name": "Sheila",
     "health_url": "https://desktop-rh302uu.tail999c81.ts.net:8443/api/health",
     "launch_vbs": r"C:\Users\CMcGavigan\Documents\Sheilas app\launch.vbs"},
    {"name": "Receipts",
     "health_url": "https://desktop-rh302uu.tail999c81.ts.net/api/health",
     "launch_vbs": r"C:\Users\CMcGavigan\Documents\Receipts\launch.vbs"},
]


def log(msg):
    line = f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def load_state():
    try:
        return json.load(open(STATE_FILE, encoding="utf-8"))
    except Exception:
        return {}


def save_state(state):
    try:
        json.dump(state, open(STATE_FILE, "w", encoding="utf-8"), indent=2)
    except OSError as e:
        log(f"could not write state: {e}")


def check_health(url, timeout=12):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    try:
        req = urllib.request.Request(url, headers={"Cache-Control": "no-store"})
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
            if r.status != 200:
                return False
            body = r.read(2000).decode("utf-8", "ignore")
            return '"ok":true' in body.replace(" ", "")
    except Exception:
        return False


def restart(server):
    if DRY_RUN:
        log(f"[dry-run] would restart {server['name']}")
        return
    vbs = server["launch_vbs"]
    if not os.path.exists(vbs):
        log(f"cannot restart {server['name']}: launcher not found at {vbs}")
        return
    try:
        subprocess.Popen(["wscript", vbs], shell=False)
        log(f"restart triggered for {server['name']}")
    except Exception as e:
        log(f"restart FAILED for {server['name']}: {e}")


def ntfy(title, message, priority="default", tags=""):
    if not NTFY_TOPIC:
        log("NTFY_TOPIC not set - cannot send alert (set it in .env)")
        return
    if DRY_RUN:
        log(f"[dry-run] would ntfy: {title} - {message}")
        return
    url = f"{NTFY_SERVER}/{urllib.parse.quote(NTFY_TOPIC)}"
    headers = {"Title": title, "Priority": priority}
    if tags:
        headers["Tags"] = tags
    try:
        req = urllib.request.Request(url, data=message.encode("utf-8"), headers=headers, method="POST")
        urllib.request.urlopen(req, timeout=12)
        log(f"alert sent: {title}")
    except Exception as e:
        log(f"ntfy send failed: {e}")


def last_shutdown_event():
    """Most recent shutdown from the Windows System log, classified.
    Returns {kind: clean|powerloss|unknown, when: epoch, id}."""
    if os.name != "nt":
        return {"kind": "unknown", "when": 0, "id": 0}
    ps = ("$e = Get-WinEvent -FilterHashtable @{LogName='System'; Id=6008,41,1074,6006} "
          "-MaxEvents 1 -ErrorAction SilentlyContinue; "
          "if ($e) { '{0}|{1}' -f $e.Id, ([int][double]::Parse((Get-Date $e.TimeCreated -UFormat %s))) }")
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                             capture_output=True, text=True, timeout=20).stdout.strip()
        if not out or "|" not in out:
            return {"kind": "unknown", "when": 0, "id": 0}
        eid_s, when_s = out.split("|", 1)
        eid, when = int(eid_s), int(when_s)
        kind = "powerloss" if eid in (6008, 41) else "clean"
        return {"kind": kind, "when": when, "id": eid}
    except Exception as e:
        log(f"event-log read failed: {e}")
        return {"kind": "unknown", "when": 0, "id": 0}


def main():
    now = time.time()
    state = load_state()
    last_run = state.get("last_run", 0)
    gap_min = (now - last_run) / 60 if last_run else 0
    laptop_was_off = bool(last_run) and gap_min > GAP_MINUTES
    if laptop_was_off:
        log(f"gap of {int(gap_min)} min > {int(GAP_MINUTES)} - treating as laptop-off (not counted as failure time)")

    servers_state = state.get("servers", {})

    for srv in SERVERS:
        name = srv["name"]
        ss = servers_state.get(name, {"failing_since": None, "alerted": False})
        ok = check_health(srv["health_url"])

        if ok:
            if ss.get("failing_since"):
                if ss.get("alerted"):
                    ntfy(f"{name} server back up", "It's reachable again. All good.", tags="white_check_mark")
                log(f"{name}: OK (recovered)")
            else:
                log(f"{name}: OK")
            ss = {"failing_since": None, "alerted": False}
        else:
            log(f"{name}: UNREACHABLE over tunnel")
            restart(srv)
            if not ss.get("failing_since") or laptop_was_off:
                ss["failing_since"] = now
                ss["alerted"] = False
            failing_hours = (now - ss["failing_since"]) / 3600
            log(f"{name}: failing for ~{failing_hours:.1f}h of on-time")
            if failing_hours >= ALERT_AFTER_HOURS and not ss.get("alerted"):
                ntfy(f"WARNING: {name} server down >{int(ALERT_AFTER_HOURS)}h",
                     f"{name} has been unreachable over Tailscale for over {int(ALERT_AFTER_HOURS)} hours "
                     f"and isn't recovering on restart. Sheila can't reach it. Worth a look.",
                     priority="high", tags="warning")
                ss["alerted"] = True

        servers_state[name] = ss

    # First run after a gap: clean shutdown -> silent; power loss -> notify once.
    if laptop_was_off:
        ev = last_shutdown_event()
        already = state.get("last_shutdown_reported", 0)
        if ev["kind"] == "powerloss" and ev["when"] and ev["when"] != already:
            all_ok = all(servers_state.get(s["name"], {}).get("failing_since") is None for s in SERVERS)
            if all_ok:
                ntfy("Power loss recovered",
                     "The laptop lost power unexpectedly, but it's back and both servers are running fine.",
                     tags="electric_plug")
            else:
                down = [s["name"] for s in SERVERS if servers_state.get(s["name"], {}).get("failing_since")]
                ntfy("WARNING: Power loss - server didn't come back",
                     f"The laptop lost power unexpectedly and these aren't reachable yet: {', '.join(down)}. "
                     f"They may still be starting; the watchdog will keep retrying.",
                     priority="high", tags="warning")
            state["last_shutdown_reported"] = ev["when"]
            log(f"power-loss event reported (event id {ev['id']})")
        elif ev["kind"] == "clean":
            log("last shutdown was clean - no notification")

    state["servers"] = servers_state
    state["last_run"] = now
    save_state(state)
    log("watchdog run complete.")


def test_alert():
    # Fire one of each real alert type through ntfy(), so you can confirm the
    # whole chain (watchdog -> ntfy -> your phone) works. Triggered by --test-alert.
    if not NTFY_TOPIC:
        log("NTFY_TOPIC not set in .env - nothing to test. Set it first.")
        return
    log(f"sending test alerts to ntfy topic '{NTFY_TOPIC}' ...")
    ntfy("Sheila server down >24h (TEST)",
         "This is a TEST of the 'server down' alert. If you got this, real alerts will reach you.",
         priority="high", tags="warning")
    ntfy("Power loss recovered (TEST)",
         "This is a TEST of the 'power loss recovered' alert.",
         tags="electric_plug")
    ntfy("Sheila server back up (TEST)",
         "This is a TEST of the 'recovered' alert. All three test alerts sent.",
         tags="white_check_mark")
    log("test alerts sent - check your phone.")


if __name__ == "__main__":
    if "--test-alert" in sys.argv:
        test_alert()
    else:
        main()
