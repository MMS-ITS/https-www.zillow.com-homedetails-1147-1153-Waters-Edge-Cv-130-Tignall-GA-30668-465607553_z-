# Analysis artifacts

Everything here supports **`../Property LOT Due Diligence - Report.docx`**, the answer to the
61-point checklist in `../Property LOT Due diligence.docx`.

Subject property: 1147 (Lot 130) and 1153 (Lot 131) Watersedge Cove, Tignall,
**Lincoln County**, GA 30668 — 2 × 1.07 ac in Stillwater Coves, on Clarks Hill /
J. Strom Thurmond Lake. Reference point `33.94100 N, -82.56460 W`.

## Scripts (run in this order)

| Script | What it does |
|---|---|
| `sample_terrain.py` | Samples 2,665 USGS 3DEP **1-metre lidar** elevations (NAVD 88) over a 1,070 × 670 m grid → `data/terrain_grid.csv` |
| `analyze.py` | ASCII terrain map, elevation range, road-centreline elevations |
| `terrain_analysis.py` | Locates the 330 ft full-pool shoreline; builds road→lake transects; computes average and steepest-50 ft slopes → `data/transects.json` |
| `envelope.py` | Depth-from-road elevation profiles; answers "how far back does the ground reach 330 ft?" |
| `location.py` | OSRM road-network distances/drive times; ERA5 frost dates and growing season |
| `figures.py` | Generates all eight figures → `figures/` |
| `build_docx.py` | Assembles the A4 portrait Word report with page numbering |

## Data

- `data/terrain_grid.csv` — the lidar sample grid (lat, lon, elevation in m and ft, source resolution)
- `data/transects.json` — road→lake profiles with computed slope statistics

`fema.json` is **not** committed (7.7 MB of raw shoreline geometry). Regenerate with:

```bash
curl -s "https://hazards.fema.gov/arcgis/rest/services/public/NFHL/MapServer/28/query?geometry=%7B%22xmin%22%3A-82.5700%2C%22ymin%22%3A33.9380%2C%22xmax%22%3A-82.5580%2C%22ymax%22%3A33.9435%2C%22spatialReference%22%3A%7B%22wkid%22%3A4326%7D%7D&geometryType=esriGeometryEnvelope&inSR=4326&outSR=4326&spatialRel=esriSpatialRelIntersects&outFields=FLD_ZONE,SFHA_TF&returnGeometry=true&f=json" -o fema.json
```

## Figures

| File | Figure |
|---|---|
| `fig1_topo.png` | True-scale topographic map, 5 ft contours, 330 ft shoreline, drainage arrows |
| `fig2_section.png` | Road→lake section, **vertically exaggerated and true 1:1 side by side** (checklist requirement) |
| `fig3_3d.png` | 3-D terrain model with the full-pool plane |
| `fig4_slope.png` | Slope classification map |
| `fig5_fema.png` | FEMA NFHL Zone A / Zone X, DFIRM 13181C |
| `fig6_envelope.png` | Illustrative building envelope, house, driveway, septic primary + reserve |
| `fig7_location.png` | Regional location and access |
| `fig8_climate.png` | Monthly temperature and precipitation, ERA5 1995–2024 |

## Dependencies

```bash
pip install python-docx matplotlib numpy scipy pypdf
```

## Reading the numbers

Three things are worth knowing before you trust any figure:

1. **Lot geometry is modelled, not surveyed.** Lincoln County's parcel GIS returns
   `499 Token Required`, so no parcel polygon exists in this analysis. Lots are drawn as
   150 × 310 ft rectangles (1.07 ac). Boundary-dependent numbers — frontage, depth, setbacks,
   buildable area — are illustrative. Terrain, flood and soil findings do **not** depend on
   that assumption.
2. **Elevations are planning-grade.** 1 m lidar to NAVD 88 is excellent for slope and relief,
   but flood and permit decisions need a surveyed Elevation Certificate.
3. **Soil is survey-scale SSURGO.** NRCS states plainly that survey-scale maps are insufficient
   to approve a septic system; a stamped Level 3/4 on-site report is required.

## Data sources

USGS 3DEP (1 m lidar, NAVD 88) · FEMA National Flood Hazard Layer (DFIRM 13181C) ·
USDA NRCS SSURGO via Soil Data Access · OpenStreetMap (© contributors, ODbL 1.0) ·
OSRM · ERA5 via Open-Meteo · USACE Savannah District · Georgia DPH · US Census Bureau.
Full source list is in the report's Sources section.
