import csv, math, json

M2FT = 3.280839895
FULL_POOL = 330.0          # ft-msl, USACE Thurmond Lake full pool
WINTER_POOL = 326.0        # typical seasonal drawdown target
TOP_OF_GATES = 335.0

rows = list(csv.DictReader(open("/projects/sandbox/work/terrain_grid.csv")))
pts = {(float(r["lat"]), float(r["lon"])): float(r["elev_ft"]) for r in rows}
lats = sorted({la for la, lo in pts})
lons = sorted({lo for la, lo in pts})

LAT0, LON0 = 33.9410, -82.5646
MPD_LAT = 111132.0
MPD_LON = 111320.0 * math.cos(math.radians(LAT0))
FT = 3.280839895


def to_xy(la, lo):
    """local metres east/north from reference"""
    return ((lo - LON0) * MPD_LON, (la - LAT0) * MPD_LAT)


def bilinear(la, lo):
    """interpolate elevation (ft) at arbitrary point"""
    if la < lats[0] or la > lats[-1] or lo < lons[0] or lo > lons[-1]:
        return None
    i = min(range(len(lats) - 1), key=lambda k: abs(lats[k] - la) if lats[k] <= la else 9e9)
    while i + 1 < len(lats) - 1 and lats[i + 1] <= la:
        i += 1
    j = 0
    while j + 1 < len(lons) - 1 and lons[j + 1] <= lo:
        j += 1
    la0, la1 = lats[i], lats[i + 1]
    lo0, lo1 = lons[j], lons[j + 1]
    try:
        z00, z01 = pts[(la0, lo0)], pts[(la0, lo1)]
        z10, z11 = pts[(la1, lo0)], pts[(la1, lo1)]
    except KeyError:
        return None
    tu = (lo - lo0) / (lo1 - lo0) if lo1 != lo0 else 0
    tv = (la - la0) / (la1 - la0) if la1 != la0 else 0
    return (z00 * (1 - tu) * (1 - tv) + z01 * tu * (1 - tv)
            + z10 * (1 - tu) * tv + z11 * tu * tv)


# ---------------------------------------------------------------- shoreline
shore = [(la, lo, e) for (la, lo), e in pts.items() if FULL_POOL - 2.5 <= e <= FULL_POOL + 2.5]
below = [(la, lo, e) for (la, lo), e in pts.items() if e < FULL_POOL]
print("grid elev range: %.1f – %.1f ft" % (min(pts.values()), max(pts.values())))
print("samples at/below full pool (330 ft):", len(below))
print("samples within +/-2.5 ft of full pool:", len(shore))
if below:
    print("  lat range of sub-330 ground: %.5f .. %.5f" % (min(p[0] for p in below), max(p[0] for p in below)))
    print("  lon range of sub-330 ground: %.5f .. %.5f" % (min(p[1] for p in below), max(p[1] for p in below)))

# road centreline (OSM way 409863360) - Watersedge Cove
road = [(33.9407835, -82.5657162), (33.9409225, -82.5650618), (33.9414833, -82.5624936),
        (33.9414955, -82.5622414), (33.9414588, -82.5620336), (33.9413030, -82.5615937)]
cul = [(33.9405682, -82.5656653), (33.9404514, -82.5656586), (33.9403217, -82.5657196),
       (33.9402477, -82.5658329), (33.9402294, -82.5659985), (33.9402728, -82.5661514),
       (33.9403729, -82.5662366), (33.9404853, -82.5662507), (33.9406121, -82.5661937),
       (33.9407134, -82.5660488)]


def dist_m(p, q):
    (x1, y1), (x2, y2) = to_xy(*p), to_xy(*q)
    return math.hypot(x2 - x1, y2 - y1)


def nearest_shore(la, lo):
    best = None
    for sla, slo, se in shore:
        d = dist_m((la, lo), (sla, slo))
        if best is None or d < best[0]:
            best = (d, sla, slo, se)
    return best


print("\n--- Distance from road centreline to full-pool shoreline ---")
stations = road + cul
for la, lo in stations:
    e = bilinear(la, lo)
    ns = nearest_shore(la, lo)
    if ns and e:
        print("  road %.6f,%.6f  elev %6.1f ft   shore %5.0f m (%4.0f ft) away, bearing-to N%s"
              % (la, lo, e, ns[0], ns[0] * FT,
                 "NE" if ns[2] > lo else "NW"))

# ---------------------------------------------------------- transect builder
def transect(start, end, n=160):
    (la0, lo0), (la1, lo1) = start, end
    out = []
    total = dist_m(start, end)
    for k in range(n + 1):
        t = k / n
        la = la0 + (la1 - la0) * t
        lo = lo0 + (lo1 - lo0) * t
        z = bilinear(la, lo)
        if z is None:
            continue
        out.append((t * total, z, la, lo))     # (metres along, ft elev, lat, lon)
    return out


def slope_stats(prof, window_ft=50.0):
    if len(prof) < 2:
        return {}
    d0, z0 = prof[0][0], prof[0][1]
    d1, z1 = prof[-1][0], prof[-1][1]
    run_ft = (d1 - d0) * FT
    rise_ft = z0 - z1
    avg_pct = (rise_ft / run_ft * 100) if run_ft else 0
    # steepest slope over any `window_ft` horizontal window
    worst = (0, None, None)
    for i in range(len(prof)):
        for j in range(i + 1, len(prof)):
            run = (prof[j][0] - prof[i][0]) * FT
            if run < window_ft * 0.9:
                continue
            if run > window_ft * 1.1:
                break
            drop = abs(prof[i][1] - prof[j][1])
            pct = drop / run * 100
            if pct > worst[0]:
                worst = (pct, prof[i], prof[j])
    zs = [p[1] for p in prof]
    return {
        "length_m": round(d1 - d0, 1),
        "length_ft": round(run_ft, 1),
        "z_start_ft": round(z0, 1),
        "z_end_ft": round(z1, 1),
        "z_max_ft": round(max(zs), 1),
        "z_min_ft": round(min(zs), 1),
        "relief_ft": round(max(zs) - min(zs), 1),
        "avg_slope_pct": round(avg_pct, 1),
        "avg_slope_deg": round(math.degrees(math.atan(avg_pct / 100)), 1),
        "max50_slope_pct": round(worst[0], 1),
        "max50_slope_deg": round(math.degrees(math.atan(worst[0] / 100)), 1),
    }


# Build transects from several road stations straight to the nearest shoreline
print("\n--- Road-to-lake transects (1 m lidar, NAVD88) ---")
results = {}
for name, (la, lo) in [
        ("W cul-de-sac head", (33.9403729, -82.5662366)),
        ("W cul-de-sac N side", (33.9407134, -82.5660488)),
        ("Road sta. A", (33.9407835, -82.5657162)),
        ("Road sta. B", (33.9409225, -82.5650618)),
        ("Road sta. C", (33.9412000, -82.5638000)),
        ("Road sta. D", (33.9414833, -82.5624936)),
]:
    ns = nearest_shore(la, lo)
    if not ns:
        continue
    prof = transect((la, lo), (ns[1], ns[2]), 200)
    st = slope_stats(prof)
    results[name] = {"from": [la, lo], "to": [ns[1], ns[2]], "stats": st,
                     "profile": [[round(p[0], 2), round(p[1], 2)] for p in prof]}
    print(f"\n  {name}: road {la:.6f},{lo:.6f} -> shore {ns[1]:.6f},{ns[2]:.6f}")
    for k, v in st.items():
        print(f"      {k:18} {v}")

json.dump(results, open("/projects/sandbox/work/transects.json", "w"), indent=1)
print("\nwrote transects.json")
