import json, urllib.request, collections, statistics as st

LOT = (33.9410, -82.5646)

DEST = {
    "Lincolnton, GA (county seat)":        (33.7940, -82.4793),
    "Washington, GA (Wills Memorial Hosp)": (33.7365, -82.7393),
    "Thomson, GA (Walmart Supercenter)":   (33.4697, -82.5046),
    "Elijah Clark State Park":             (33.8549, -82.3932),
    "Piedmont Augusta Hospital":           (33.4735, -81.9876),
    "Augusta Regional Airport (AGS)":      (33.3699, -81.9645),
    "Athens, GA":                          (33.9519, -83.3576),
    "Columbia Metro Airport (CAE)":        (33.9388, -81.1195),
    "Hartsfield-Jackson Atlanta (ATL)":    (33.6407, -84.4277),
}

print("--- Driving distance / time from the lot (OSRM road network) ---")
print(f"{'Destination':38} {'km':>7} {'miles':>7} {'drive':>9}")
for name, (la, lo) in DEST.items():
    url = (f"https://router.project-osrm.org/route/v1/driving/"
           f"{LOT[1]},{LOT[0]};{lo},{la}?overview=false")
    try:
        d = json.load(urllib.request.urlopen(url, timeout=60))
        r = d["routes"][0]
        km = r["distance"] / 1000
        mi = km * 0.621371
        mins = r["duration"] / 60
        hh = f"{int(mins//60)}h {int(mins%60)}m" if mins >= 60 else f"{int(mins)} min"
        print(f"{name:38} {km:7.1f} {mi:7.1f} {hh:>9}")
    except Exception as e:
        print(f"{name:38}   ERROR {e}")

print("\n--- Growing season / frost dates (ERA5 reanalysis 1995-2024, 2 m temp) ---")
url = ("https://archive-api.open-meteo.com/v1/archive?latitude=33.9410&longitude=-82.5646"
       "&start_date=1995-01-01&end_date=2024-12-31"
       "&daily=temperature_2m_min,temperature_2m_max&timezone=America%2FNew_York")
d = json.load(urllib.request.urlopen(url, timeout=180))
dd = d["daily"]
byyear = collections.defaultdict(list)
for day, mn in zip(dd["time"], dd["temperature_2m_min"]):
    if mn is not None:
        byyear[day[:4]].append((day, mn))

last_spring, first_fall, lengths = [], [], []
for y, days in byyear.items():
    ls = None
    ff = None
    for day, mn in days:
        doy = (int(day[5:7]), int(day[8:10]))
        md = int(day[5:7]) * 100 + int(day[8:10])
        if mn <= 0 and md < 700:
            ls = md
        if mn <= 0 and md > 700 and ff is None:
            ff = md
    if ls and ff:
        last_spring.append(ls); first_fall.append(ff)
        # approximate day-of-year difference
        import datetime
        a = datetime.date(int(y), ls // 100, ls % 100)
        b = datetime.date(int(y), ff // 100, ff % 100)
        lengths.append((b - a).days)


def fmt(md):
    mm, dd_ = int(md) // 100, int(md) % 100
    return f"{['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][mm]} {dd_}"


if last_spring:
    print(f"  mean last spring frost (<=0 C): {fmt(st.mean(last_spring))}")
    print(f"  mean first autumn frost (<=0 C): {fmt(st.mean(first_fall))}")
    print(f"  mean frost-free season: {st.mean(lengths):.0f} days")
    print(f"  latest spring frost on record: {fmt(max(last_spring))}")
    print(f"  earliest autumn frost on record: {fmt(min(first_fall))}")

mins = [m for _, m in sum(byyear.values(), [])]
annmin = []
for y, days in byyear.items():
    annmin.append(min(m for _, m in days))
print(f"  mean annual extreme min: {st.mean(annmin):.1f} C ({st.mean(annmin)*9/5+32:.1f} F)"
      "  -> USDA hardiness zone check")
print(f"  coldest single day 1995-2024: {min(mins):.1f} C ({min(mins)*9/5+32:.1f} F)")

# heat / cooling demand
mx = [m for m in dd["temperature_2m_max"] if m is not None]
print(f"  days/yr with max >= 32 C (90 F): {sum(1 for m in mx if m >= 32)/30:.0f}")
print(f"  days/yr with max >= 35 C (95 F): {sum(1 for m in mx if m >= 35)/30:.0f}")
