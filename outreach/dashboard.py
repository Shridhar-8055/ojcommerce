"""
OJCommerce Outreach Dashboard — HTML Generator
Reads prospects.csv and sent_log.csv to generate a live-status dashboard.
Called automatically after every send from outreach_emails.py.
"""

import csv
import subprocess
from datetime import datetime
from pathlib import Path

import config


def get_monitor_status():
    """Check if email_monitor.py is running. Returns (is_running, last_check)."""
    try:
        # Try multiple detection methods
        is_running = False

        # Method 1: pgrep
        result = subprocess.run(
            ["pgrep", "-f", "email_monitor"],
            capture_output=True, text=True
        )
        if result.returncode == 0 and result.stdout.strip():
            is_running = True

        # Method 2: ps aux fallback
        if not is_running:
            result = subprocess.run(
                ["ps", "aux"],
                capture_output=True, text=True
            )
            if "email_monitor.py" in result.stdout and "--watch" in result.stdout:
                is_running = True

        # Method 3: check PID file if exists
        pid_file = config.BASE_DIR / ".monitor_pid"
        if not is_running and pid_file.exists():
            try:
                pid = int(pid_file.read_text().strip())
                import os
                os.kill(pid, 0)  # signal 0 = check if process exists
                is_running = True
            except (ProcessLookupError, ValueError, PermissionError):
                pass

        # Get last check time from monitor state file
        last_check = "Never"
        state_file = config.MONITOR_STATE_FILE
        if state_file.exists():
            import os
            mtime = os.path.getmtime(state_file)
            last_check = datetime.fromtimestamp(mtime).strftime("%I:%M %p")

        return is_running, last_check
    except Exception:
        return False, "Unknown"


def load_prospects():
    rows = []
    with open(config.PROSPECTS_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            rows.append(row)
    return rows


def load_sent_log():
    if not config.SENT_LOG_CSV.exists():
        return []
    with open(config.SENT_LOG_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def load_replies_log():
    if not config.REPLIES_LOG_CSV.exists():
        return []
    with open(config.REPLIES_LOG_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def generate_dashboard():
    prospects = load_prospects()
    sent_log = load_sent_log()
    replies_log = load_replies_log()

    # Build lookup: site_name -> list of log entries
    log_by_site = {}
    for entry in sent_log:
        name = entry.get("site_name", "")
        log_by_site.setdefault(name, []).append(entry)

    # Stats
    total = len(prospects)
    tier1 = [p for p in prospects if p.get("tier") == "1"]
    tier2 = [p for p in prospects if p.get("tier") == "2"]
    pitched = sum(1 for p in prospects if p.get("status") not in ("Not Started", ""))
    sent_count = sum(1 for e in sent_log if e.get("status") == "sent")
    failed_count = sum(1 for e in sent_log if e.get("status") == "failed")
    replied = sum(1 for p in prospects if p.get("status") == "Replied")
    accepted = sum(1 for p in prospects if p.get("status") == "Accepted")
    published = sum(1 for p in prospects if p.get("status") == "Published")

    # Status color map
    status_colors = {
        "Not Started": ("#94a3b8", "#f1f5f9"),
        "Pitched": ("#0066FF", "#e0f0ff"),
        "Replied": ("#f59e0b", "#fef3c7"),
        "Accepted": ("#10b981", "#d1fae5"),
        "Rejected": ("#ef4444", "#fef2f2"),
        "Published": ("#8b5cf6", "#f5f3ff"),
    }

    def status_badge(status):
        color, bg = status_colors.get(status, ("#94a3b8", "#f1f5f9"))
        return f'<span style="color:{color};background:{bg};padding:4px 12px;border-radius:12px;font-size:12px;font-weight:700;">{status}</span>'

    def dr_badge(dr, tier):
        if tier == "1":
            return f'<span style="background:#FFF8E1;color:#B8860B;padding:4px 12px;border-radius:12px;font-size:12px;font-weight:700;">{dr}</span>'
        return f'<span style="background:#E8F4FD;color:#0066FF;padding:4px 12px;border-radius:12px;font-size:12px;font-weight:700;">{dr}</span>'

    def log_entries_html(site_name):
        entries = log_by_site.get(site_name, [])
        if not entries:
            return '<span style="color:#94a3b8;font-size:12px;">—</span>'
        html = ""
        for e in entries:
            ts = e.get("timestamp", "")[:16].replace("T", " ")
            tmpl = e.get("template", "")
            st = e.get("status", "")
            err = e.get("error", "")
            icon = "&#9989;" if st == "sent" else "&#10060;" if st == "failed" else "&#9898;"
            line = f'{icon} {ts} — {tmpl}'
            if err:
                line += f' <span style="color:#ef4444;">({err})</span>'
            html += f'<div style="font-size:11px;color:#555;margin:2px 0;">{line}</div>'
        return html

    # Build prospect rows by tier
    tier1_rows = ""
    for i, p in enumerate(tier1, 1):
        tier1_rows += f"""<tr>
            <td style="text-align:center;color:#94a3b8;font-weight:600;">{i}</td>
            <td><strong style="color:#0a1628;">{p['site_name']}</strong><br><a href="https://{p['site_url']}" target="_blank" rel="noopener noreferrer" style="color:#0066FF;font-size:12px;text-decoration:none;cursor:pointer;" onclick="window.open(this.href);return false;">{p['site_url']}</a></td>
            <td>{dr_badge(p['dr'], '1')}</td>
            <td style="font-size:12px;">{p['niche']}</td>
            <td style="font-size:12px;">{p['opportunity_type']}</td>
            <td style="font-size:12px;">{p.get('contact_email') or '<span style="color:#ccc;">needs research</span>'}</td>
            <td>{status_badge(p['status'])}</td>
            <td>{log_entries_html(p['site_name'])}</td>
        </tr>"""

    tier2_rows = ""
    for i, p in enumerate(tier2, 1):
        tier2_rows += f"""<tr>
            <td style="text-align:center;color:#94a3b8;font-weight:600;">{i}</td>
            <td><strong style="color:#0a1628;">{p['site_name']}</strong><br><a href="https://{p['site_url']}" target="_blank" rel="noopener noreferrer" style="color:#0066FF;font-size:12px;text-decoration:none;cursor:pointer;" onclick="window.open(this.href);return false;">{p['site_url']}</a></td>
            <td>{dr_badge(p['dr'], '2')}</td>
            <td style="font-size:12px;">{p['niche']}</td>
            <td style="font-size:12px;">{p['opportunity_type']}</td>
            <td style="font-size:12px;">{p.get('contact_email') or '<span style="color:#ccc;">needs research</span>'}</td>
            <td>{status_badge(p['status'])}</td>
            <td>{log_entries_html(p['site_name'])}</td>
        </tr>"""

    # Recent activity log
    recent_log_html = ""
    for entry in reversed(sent_log[-20:]):
        ts = entry.get("timestamp", "")[:16].replace("T", " ")
        st = entry.get("status", "")
        icon = "&#9989;" if st == "sent" else "&#10060;" if st == "failed" else "&#9898;"
        color = "#10b981" if st == "sent" else "#ef4444" if st == "failed" else "#94a3b8"
        recent_log_html += f"""<tr>
            <td style="font-size:12px;color:#555;">{ts}</td>
            <td style="font-size:12px;"><strong>{entry.get('site_name','')}</strong></td>
            <td style="font-size:12px;">{entry.get('template','')}</td>
            <td style="font-size:12px;">{entry.get('email','')}</td>
            <td style="font-size:12px;color:{color};font-weight:700;">{icon} {st}</td>
        </tr>"""
    if not recent_log_html:
        recent_log_html = '<tr><td colspan="5" style="text-align:center;color:#94a3b8;padding:30px;">No emails sent yet. Run a send command to see activity here.</td></tr>'

    # Build reply cards (newest first)
    import html as html_module
    replies_html = ""
    sorted_replies = sorted(replies_log, key=lambda r: r.get("timestamp", ""), reverse=True)
    for reply in sorted_replies:
        ts = reply.get("timestamp", "")[:16].replace("T", " ")
        site = html_module.escape(reply.get("site_name", "Unknown"))
        sender = html_module.escape(reply.get("from_email", ""))
        subj = html_module.escape(reply.get("subject", "(no subject)"))
        body = html_module.escape(reply.get("body", "")).replace("\n", "<br>")
        if not body:
            body = '<span style="color:#94a3b8;font-style:italic;">No body content available</span>'
        replies_html += f"""
        <div style="background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.06);margin-bottom:16px;overflow:hidden;">
          <div style="background:linear-gradient(135deg,#0a1628,#1a365d);color:#fff;padding:14px 20px;display:flex;justify-content:space-between;align-items:center;">
            <div>
              <strong style="font-size:16px;">{site}</strong>
              <span style="color:#a0c4ff;font-size:12px;margin-left:10px;">{sender}</span>
            </div>
            <span style="color:#a0c4ff;font-size:12px;">{ts}</span>
          </div>
          <div style="padding:16px 20px;border-bottom:1px solid #f0f0f0;">
            <div style="font-size:12px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">Subject</div>
            <div style="font-size:14px;font-weight:600;color:#0a1628;">{subj}</div>
          </div>
          <div style="padding:16px 20px;">
            <div style="font-size:12px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Message</div>
            <div style="font-size:14px;color:#333;line-height:1.6;white-space:pre-wrap;word-wrap:break-word;">{body}</div>
          </div>
        </div>"""
    if not replies_html:
        replies_html = '<div style="text-align:center;color:#94a3b8;padding:40px;">No replies received yet.</div>'

    now = datetime.now().strftime("%b %d, %Y at %I:%M %p")
    monitor_running, last_check = get_monitor_status()
    monitor_dot = "#00C9A7" if monitor_running else "#ef4444"
    monitor_label = "Monitor: Connected" if monitor_running else "Monitor: Offline"
    monitor_sub = f"Last check: {last_check}" if monitor_running else "Run: python3 email_monitor.py --watch"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OJCommerce Outreach Dashboard (Month 1)</title>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f4f8; color: #333; }}
  .header {{
    background: linear-gradient(135deg, #0a1628 0%, #1a365d 50%, #0066FF 100%);
    padding: 40px 40px 30px;
    color: #fff;
  }}
  .header h1 {{ font-size: 28px; font-weight: 800; margin-bottom: 4px; }}
  .header .sub {{ font-size: 14px; color: #a0c4ff; margin-bottom: 20px; }}
  .stats {{ display: flex; gap: 16px; flex-wrap: wrap; }}
  .stat {{
    background: rgba(255,255,255,0.1);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255,255,255,0.15);
    border-radius: 12px;
    padding: 14px 24px;
    text-align: center;
    min-width: 130px;
  }}
  .stat .num {{ font-size: 28px; font-weight: 800; }}
  .stat .num.green {{ color: #00C9A7; }}
  .stat .num.blue {{ color: #60a5fa; }}
  .stat .num.yellow {{ color: #fbbf24; }}
  .stat .num.red {{ color: #f87171; }}
  .stat .num.purple {{ color: #a78bfa; }}
  .stat .lbl {{ font-size: 11px; color: #a0c4ff; text-transform: uppercase; letter-spacing: 1px; margin-top: 2px; }}

  .container {{ max-width: 1500px; margin: 0 auto; padding: 24px 30px 60px; }}

  .tabs {{ display: flex; gap: 4px; margin-bottom: 20px; }}
  .tab-btn {{
    padding: 10px 24px; border: none; background: #fff; color: #555;
    font-size: 14px; font-weight: 600; cursor: pointer; border-radius: 8px 8px 0 0;
    border-bottom: 3px solid transparent; transition: all 0.2s;
  }}
  .tab-btn:hover {{ color: #0066FF; }}
  .tab-btn.active {{ color: #0066FF; border-bottom-color: #0066FF; background: #fff; }}
  .tab-btn .badge {{
    background: #0066FF; color: #fff; font-size: 11px; padding: 2px 8px;
    border-radius: 10px; margin-left: 6px;
  }}

  .panel {{ display: none; }}
  .panel.active {{ display: block; }}

  .card {{
    background: #fff; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    overflow: hidden; margin-bottom: 24px;
  }}
  .card-title {{
    background: #0a1628; color: #fff; padding: 14px 20px; font-size: 14px;
    font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;
  }}

  table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
  thead th {{
    background: #0a1628; color: #fff; padding: 12px 14px; text-align: left;
    font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600;
    position: sticky; top: 0;
  }}
  tbody tr {{ border-bottom: 1px solid #f0f0f0; transition: background 0.15s; }}
  tbody tr:hover {{ background: #f0f7ff; }}
  tbody td {{ padding: 12px 14px; vertical-align: middle; }}

  .updated {{ color: #94a3b8; font-size: 12px; text-align: right; margin-top: 20px; }}

  @media (max-width: 768px) {{
    .header {{ padding: 24px 16px; }}
    .container {{ padding: 16px; }}
    .stats {{ gap: 8px; }}
    .stat {{ min-width: 100px; padding: 10px 14px; }}
    .stat .num {{ font-size: 22px; }}
    .card {{ overflow-x: auto; }}
    table {{ min-width: 900px; }}
  }}
</style>
</head>
<body>

<div class="header">
  <h1>OJCommerce Outreach Dashboard (Month 1)</h1>
  <p class="sub">Email campaign tracker — auto-updated after every send &nbsp;|&nbsp; <span style="color:{monitor_dot};font-weight:700;">&#9679; {monitor_label}</span> <span style="color:#a0c4ff;font-size:12px;">({monitor_sub})</span></p>
  <div class="stats">
    <div class="stat"><div class="num green">{total}</div><div class="lbl">Total Prospects</div></div>
    <div class="stat"><div class="num blue">{pitched}</div><div class="lbl">Pitched</div></div>
    <div class="stat"><div class="num green">{sent_count}</div><div class="lbl">Emails Sent</div></div>
    <div class="stat"><div class="num red">{failed_count}</div><div class="lbl">Failed</div></div>
    <div class="stat"><div class="num yellow">{replied}</div><div class="lbl">Replied</div></div>
    <div class="stat"><div class="num green">{accepted}</div><div class="lbl">Accepted</div></div>
    <div class="stat"><div class="num purple">{published}</div><div class="lbl">Published</div></div>
  </div>
</div>

<div class="container">
  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab('tier1')">Tier 1 — DR 80+ <span class="badge">{len(tier1)}</span></button>
    <button class="tab-btn" onclick="switchTab('tier2')">Tier 2 — DR 60-80 <span class="badge">{len(tier2)}</span></button>
    <button class="tab-btn" onclick="switchTab('log')">Activity Log <span class="badge">{len(sent_log)}</span></button>
    <button class="tab-btn" onclick="switchTab('replies')">Replies <span class="badge" style="background:#f59e0b;">{len(replies_log)}</span></button>
  </div>

  <div class="panel active" id="panel-tier1">
    <div class="card">
      <table>
        <thead><tr>
          <th>#</th><th>Website</th><th>DR</th><th>Niche</th><th>Opportunity</th><th>Email</th><th>Status</th><th>Send History</th>
        </tr></thead>
        <tbody>{tier1_rows}</tbody>
      </table>
    </div>
  </div>

  <div class="panel" id="panel-tier2">
    <div class="card">
      <table>
        <thead><tr>
          <th>#</th><th>Website</th><th>DR</th><th>Niche</th><th>Opportunity</th><th>Email</th><th>Status</th><th>Send History</th>
        </tr></thead>
        <tbody>{tier2_rows}</tbody>
      </table>
    </div>
  </div>

  <div class="panel" id="panel-log">
    <div class="card">
      <div class="card-title">Recent Activity (Last 20)</div>
      <table>
        <thead><tr>
          <th>Time</th><th>Prospect</th><th>Template</th><th>Email</th><th>Result</th>
        </tr></thead>
        <tbody>{recent_log_html}</tbody>
      </table>
    </div>
  </div>

  <div class="panel" id="panel-replies">
    {replies_html}
  </div>

  <p class="updated">Last updated: {now}</p>
</div>

<script>
function switchTab(id) {{
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-' + id).classList.add('active');
  event.target.closest('.tab-btn').classList.add('active');
}}
</script>
</body>
</html>"""

    output_path = config.BASE_DIR / "dashboard.html"
    output_path.write_text(html, encoding="utf-8")
    return output_path
