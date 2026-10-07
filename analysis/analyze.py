import csv, math

rows = list(csv.DictReader(open("/projects/sandbox/work/terrain_grid.csv")))
pts = {}
for r in rows:
    pts[(float(r["lat"]), float(r["lon"]))] = float(r["elev_ft"])

lats = sorted({la for la, lo in pts})
lons = sorted({lo for la, lo in pts})
print("rows(lat)", len(lats), "cols(lon)", len(lons))
vals = list(pts.values())
print("elev ft: min %.1f max %.1f" % (min(vals), max(vals)))

FULL_POOL = 330.0

# ASCII map, north at top
def band(e):
    if e < 320: return "~"      # below winter pool / in water
    if e < 330: return "="      # between winter and full pool
    if e < 340: return "."
    if e < 350: return ":"
    if e < 360: return "-"
    if e < 370: return "+"
    if e < 380: return "o"
    if e < 390: return "O"
    if e < 400: return "#"
    return "@"

print("\nTerrain map (N up, W left) ~15m cells")
print("legend: ~<320  =320-330  .330-340  :340-350  -350-360  +360-370  o370-380  O380-390  #390-400  @>400")
hdr = "      " + "".join("|" if i % 10 == 0 else " " for i in range(len(lons)))
print(hdr)
for la in reversed(lats):
    line = "".join(band(pts.get((la, lo), 9999)) if (la, lo) in pts else " " for lo in lons)
    print(f"{la:.5f} {line}")
print(hdr)
print("lon range %.5f .. %.5f" % (lons[0], lons[-1]))

# Road centerline nodes (OSM way 409863360)
road = [(33.9407835,-82.5657162),(33.9405682,-82.5656653),(33.9404514,-82.5656586),
(33.9403217,-82.5657196),(33.9402477,-82.5658329),(33.9402294,-82.5659985),
(33.9402728,-82.5661514),(33.9403729,-82.5662366),(33.9404853,-82.5662507),
(33.9406121,-82.5661937),(33.9407134,-82.5660488),(33.9407835,-82.5657162),
(33.9409225,-82.5650618),(33.9414833,-82.5624936),(33.9414955,-82.5622414),
(33.9414588,-82.5620336),(33.9413030,-82.5615937)]

def nearest(la, lo):
    best = None
    for (a, o), e in pts.items():
        d = (a-la)**2 + ((o-lo)*math.cos(math.radians(la)))**2
        if best is None or d < best[0]:
            best = (d, e)
    return best[1]

print("\nRoad centerline elevations (ft, nearest grid sample):")
for la, lo in road:
    print(f"  {la:.6f} {lo:.6f}  {nearest(la,lo):7.1f}")
