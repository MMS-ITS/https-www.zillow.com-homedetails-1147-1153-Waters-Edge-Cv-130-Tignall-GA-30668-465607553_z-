"""Build-envelope terrain analysis: first 350 ft back from the road centreline,
which is where a 1.07-acre lot's buildable area must sit."""
import csv, math, json

rows = list(csv.DictReader(open("/projects/sandbox/work/terrain_grid.csv")))
pts = {(float(r["lat"]), float(r["lon"])): float(r["elev_ft"]) for r in rows}
lats = sorted({la for la, lo in pts})
lons = sorted({lo for la, lo in pts})
LAT0, LON0 = 33.9410, -82.5646
MPD_LAT, MPD_LON = 111132.0, 111320.0 * math.cos(math.radians(LAT0))
FT = 3.280839895


def bilinear(la, lo):
    if la < lats[0] or la > lats[-1] or lo < lons[0] or lo > lons[-1]:
        return None
    i = 0
    while i + 1 < len(lats) - 1 and lats[i + 1] <= la:
        i += 1
    j = 0
    while j + 1 < len(lons) - 1 and lons[j + 1] <= lo:
        j += 1
    la0, la1, lo0, lo1 = lats[i], lats[i + 1], lons[j], lons[j + 1]
    try:
        z00, z01, z10, z11 = pts[(la0, lo0)], pts[(la0, lo1)], pts[(la1, lo0)], pts[(la1, lo1)]
    except KeyError:
        return None
    tu = (lo - lo0) / (lo1 - lo0)
    tv = (la - la0) / (la1 - la0)
    return z00 * (1 - tu) * (1 - tv) + z01 * tu * (1 - tv) + z10 * (1 - tu) * tv + z11 * tu * tv


def offset(la, lo, north_m, east_m):
    return la + north_m / MPD_LAT, lo + east_m / MPD_LON


# Lot geometry assumption: 1.07 ac = 46,609 sq ft. Listing/plat width unknown;
# model a 150 ft x 310 ft lot (46,500 sq ft) with frontage on Watersedge Cove,
# rear boundary toward the lake (north).
LOT_W_FT, LOT_D_FT = 150.0, 310.0
print("Modelled single lot: %.0f ft frontage x %.0f ft depth = %.0f sq ft = %.2f ac"
      % (LOT_W_FT, LOT_D_FT, LOT_W_FT * LOT_D_FT, LOT_W_FT * LOT_D_FT / 43560))
print("Combined two lots (side by side): %.0f ft frontage x %.0f ft depth = %.2f ac\n"
      % (LOT_W_FT * 2, LOT_D_FT, 2 * LOT_W_FT * LOT_D_FT / 43560))

# Depth profile straight back (north) from road, 0-350 ft, at 3 frontage positions
print("--- Elevation going back (north) from road centreline ---")
print("  depth_ft |  west sta. |  mid sta.  |  east sta.")
stations = [(33.9407835, -82.5657162), (33.9409225, -82.5650618), (33.9412000, -82.5638000)]
rowsout = []
for d_ft in range(0, 360, 20):
    vals = []
    for la, lo in stations:
        nla, nlo = offset(la, lo, d_ft / FT, 0)
        z = bilinear(nla, nlo)
        vals.append(z)
    rowsout.append((d_ft, vals))
    print("   %6d   | %9s | %9s | %9s" % (
        d_ft, *["%.1f" % v if v else "  n/a" for v in vals]))

print("\n--- Slope within the first 310 ft (buildable depth) ---")
for idx, (la, lo) in enumerate(stations):
    prof = []
    for k in range(0, 311, 5):
        nla, nlo = offset(la, lo, k / FT, 0)
        z = bilinear(nla, nlo)
        if z is not None:
            prof.append((k, z))
    if len(prof) < 2:
        continue
    rise = prof[0][1] - prof[-1][1]
    run = prof[-1][0] - prof[0][0]
    avg = rise / run * 100
    worst = 0
    wi = None
    for i in range(len(prof)):
        for j in range(i + 1, len(prof)):
            r = prof[j][0] - prof[i][0]
            if r < 45:
                continue
            if r > 55:
                break
            p = abs(prof[i][1] - prof[j][1]) / r * 100
            if p > worst:
                worst, wi = p, (prof[i], prof[j])
    name = ["west", "mid", "east"][idx]
    print("  %-5s  start %.1f ft  ->  at 310 ft back %.1f ft | drop %.1f ft"
          % (name, prof[0][1], prof[-1][1], rise))
    print("         avg slope %.1f%% (%.1f deg) | steepest 50 ft window %.1f%% (%.1f deg)"
          % (avg, math.degrees(math.atan(avg / 100)), worst,
             math.degrees(math.atan(worst / 100))))

# Where does the ground reach key elevations going back from the road?
print("\n--- Depth back from road at which ground falls to given elevation ---")
for target in (380, 370, 360, 350, 340, 335, 330):
    out = []
    for la, lo in stations:
        hit = None
        for k in range(0, 900, 5):
            nla, nlo = offset(la, lo, k / FT, 0)
            z = bilinear(nla, nlo)
            if z is not None and z <= target:
                hit = k
                break
        out.append(hit)
    print("   %4d ft-msl : west %s | mid %s | east %s" % (
        target, *["%4s ft" % (h if h is not None else "  >") for h in out]))
