"""Generate ~300 fictitious documents with the same structure as the Documents tab."""
import random, datetime as dt, json
random.seed(7)

TODAY = dt.date(2026, 11, 10)
START = dt.date(2026, 7, 6)
END = dt.date(2027, 4, 30)

WPS = [("WP00 - Project Setup", "WP00", 30), ("WP01 - Concept Design", "WP01", 55),
       ("WP02 - System Design", "WP02", 50), ("WP03 - Utilities", "WP03", 40),
       ("WP04 - Auxiliary Systems", "WP04", 35), ("WP05 - Site & Buildings", "WP05", 30),
       ("WP06 - Commissioning Preparation", "WP06", 35), ("WP07 - Operations Readiness", "WP07", 25)]
OWNERS = ["Owner A", "Owner B", "Owner C", "Owner D", "Owner E", "Owner F", "Owner G",
          "Owner H", "Owner I", "Owner J", "Owner K", "Owner L"]
TYPES = [("TEC", "Technical note"), ("DRW", "Drawing"), ("SPC", "Specification"),
         ("ANA", "Analysis"), ("PLN", "Plan"), ("RPT", "Report")]
MILESTONES = ["Outline approved", "Draft in review",
              "Sent for approval", "Sent for sign-off",
              "Approved and filed"]
# cumulative fraction of a document's duration at which each milestone is reached
MS_FRAC = [0.15, 0.5, 0.7, 0.85, 1.0]


def rnd_date(a, b):
    return a + dt.timedelta(days=random.randint(0, (b - a).days))


rows = []
for wp, ref, n in WPS:
    for i in range(1, n + 1):
        code, title = random.choice(TYPES)
        start = rnd_date(START, dt.date(2027, 1, 15))
        dur = random.randint(35, 110)
        baseline = min(start + dt.timedelta(days=dur), END)
        # realism: most documents slip a little
        slip = random.choice([0, 0, 7, 14, 14, 21, 28, 42, -7])
        forecast = baseline + dt.timedelta(days=slip)
        actual_dur = (forecast - start).days
        ms_dates = []
        for f in MS_FRAC:
            d = start + dt.timedelta(days=round(actual_dur * f))
            ms_dates.append(d if d <= TODAY else None)
        # some documents get stuck: forecast date passed but not yet approved -> overdue
        if ms_dates[-1] and random.random() < 0.22:
            keep = random.randint(1, 4)
            ms_dates = ms_dates[:keep] + [None] * (5 - keep)
        started = start <= TODAY
        reached = sum(1 for d in ms_dates if d)
        status = (MILESTONES[reached - 1] if reached else ("Started" if started else "Not started"))
        rows.append({
            "wp": wp, "ref": f"{ref}.{random.randint(1, 4):02d}",
            "id": f"DEMO-{code}-{len(rows) + 1:04d}",
            "title": f"{title} {ref}-{i:03d}", "owner": random.choice(OWNERS),
            "status": status,
            "baseline": baseline, "forecast": forecast,
            "start": start if started else None,
            "ms": ms_dates,
            "key": "Yes" if random.random() < 0.3 else "No",
            "dr": "Yes" if random.random() < 0.2 else "No",
        })

print(len(rows))
import collections
print(collections.Counter(r["status"] for r in rows))
