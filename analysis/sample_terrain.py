import json, urllib.parse, urllib.request, time, csv

M2FT = 3.280839895

def sample(points, retries=4):
    geom = {"points": points, "spatialReference": {"wkid": 4326}}
    params = {
        "geometry": json.dumps(geom),
        "geometryType": "esriGeometryMultipoint",
        "returnFirstValueOnly": "true",
        "f": "json",
        "mosaicRule": json.dumps({"mosaicMethod": "esriMosaicAttribute",
                                  "sortField": "Best", "ascending": True}),
    }
    url = "https://elevation.nationalmap.gov/arcgis/rest/services/3DEPElevation/ImageServer/getSamples"
    data = urllib.parse.urlencode(params).encode()
    for a in range(retries):
        try:
            r = json.load(urllib.request.urlopen(urllib.request.Request(url, data=data), timeout=120))
            out = {}
            for s in r.get("samples", []):
                try:
                    out[s["locationId"]] = (float(s["value"]), s.get("resolution"))
                except (TypeError, ValueError):
                    pass
            return out
        except Exception as e:
            print("retry", a, e)
            time.sleep(3)
    return {}

lat0, lat1, step_lat = 33.9372, 33.9432, 0.00015
lon0, lon1, step_lon = -82.5695, -82.5598, 0.00015

lats = []
v = lat0
while v <= lat1 + 1e-9:
    lats.append(round(v, 6)); v += step_lat
lons = []
v = lon0
while v <= lon1 + 1e-9:
    lons.append(round(v, 6)); v += step_lon

grid = [(la, lo) for la in lats for lo in lons]
print("grid points:", len(grid), "rows", len(lats), "cols", len(lons))

results = {}
CH = 250
for i in range(0, len(grid), CH):
    chunk = grid[i:i+CH]
    pts = [[lo, la] for (la, lo) in chunk]
    res = sample(pts)
    for j, (la, lo) in enumerate(chunk):
        if j in res:
            results[(la, lo)] = res[j]
    print(f"{i+len(chunk)}/{len(grid)} ok={len(results)}", flush=True)
    time.sleep(0.3)

with open("/projects/sandbox/work/terrain_grid.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["lat", "lon", "elev_m", "elev_ft", "res_m"])
    for (la, lo), (m, r) in sorted(results.items()):
        w.writerow([la, lo, round(m, 3), round(m * M2FT, 2), r])
print("wrote", len(results), "rows")
