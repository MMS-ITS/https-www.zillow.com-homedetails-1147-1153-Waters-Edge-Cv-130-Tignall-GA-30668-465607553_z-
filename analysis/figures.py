import csv, json, math, os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MplPoly, Rectangle, FancyArrow
from matplotlib.lines import Line2D
from scipy.ndimage import gaussian_filter

OUT = "/projects/sandbox/work/fig"
os.makedirs(OUT, exist_ok=True)
FT = 3.280839895
FULL_POOL = 330.0
LAT0, LON0 = 33.9410, -82.5646
MPD_LAT = 111132.0
MPD_LON = 111320.0 * math.cos(math.radians(LAT0))

# ---------------------------------------------------------------- load grid
rows = list(csv.DictReader(open("/projects/sandbox/work/terrain_grid.csv")))
lats = sorted({float(r["lat"]) for r in rows})
lons = sorted({float(r["lon"]) for r in rows})
Z = np.full((len(lats), len(lons)), np.nan)
li = {v: i for i, v in enumerate(lats)}
oi = {v: i for i, v in enumerate(lons)}
for r in rows:
    Z[li[float(r["lat"])], oi[float(r["lon"])]] = float(r["elev_ft"])

# local metric coords (feet, east/north from SW corner of grid)
X_ft = (np.array(lons) - lons[0]) * MPD_LON * FT
Y_ft = (np.array(lats) - lats[0]) * MPD_LAT * FT
XX, YY = np.meshgrid(X_ft, Y_ft)

Zs = gaussian_filter(np.nan_to_num(Z, nan=np.nanmean(Z)), 0.8)

# road polyline (OSM way 409863360, Watersedge Cove) -> ft coords
ROAD = [(33.9407835, -82.5657162), (33.9405682, -82.5656653), (33.9404514, -82.5656586),
        (33.9403217, -82.5657196), (33.9402477, -82.5658329), (33.9402294, -82.5659985),
        (33.9402728, -82.5661514), (33.9403729, -82.5662366), (33.9404853, -82.5662507),
        (33.9406121, -82.5661937), (33.9407134, -82.5660488), (33.9407835, -82.5657162),
        (33.9409225, -82.5650618), (33.9414833, -82.5624936), (33.9414955, -82.5622414),
        (33.9414588, -82.5620336), (33.9413030, -82.5615937)]


def ll2ft(la, lo):
    return ((lo - lons[0]) * MPD_LON * FT, (la - lats[0]) * MPD_LAT * FT)


road_xy = np.array([ll2ft(la, lo) for la, lo in ROAD])

# Modelled study lots: 2 x 150 ft frontage x 310 ft deep, north side of road,
# centred on the reach of Watersedge Cove with lake exposure.
ANCHOR = (33.9408200, -82.5655000)
ax0, ay0 = ll2ft(*ANCHOR)
LOTW, LOTD = 150.0, 310.0
lot_a = np.array([[ax0, ay0], [ax0 + LOTW, ay0], [ax0 + LOTW, ay0 + LOTD], [ax0, ay0 + LOTD]])
lot_b = np.array([[ax0 - LOTW, ay0], [ax0, ay0], [ax0, ay0 + LOTD], [ax0 - LOTW, ay0 + LOTD]])


def scalebar(ax, length_ft=200, label=None, loc=(0.06, 0.06)):
    x0 = ax.get_xlim()[0] + loc[0] * (ax.get_xlim()[1] - ax.get_xlim()[0])
    y0 = ax.get_ylim()[0] + loc[1] * (ax.get_ylim()[1] - ax.get_ylim()[0])
    ax.plot([x0, x0 + length_ft], [y0, y0], "k-", lw=3, solid_capstyle="butt", zorder=20)
    ax.plot([x0, x0 + length_ft / 2], [y0, y0], "w-", lw=3, solid_capstyle="butt", zorder=21)
    ax.plot([x0, x0 + length_ft], [y0, y0], "k-", lw=0.8, zorder=22)
    ax.text(x0 + length_ft / 2, y0 + 0.012 * (ax.get_ylim()[1] - ax.get_ylim()[0]),
            label or f"{length_ft:.0f} ft  ({length_ft/FT:.0f} m)",
            ha="center", va="bottom", fontsize=8, zorder=22)


def northarrow(ax, loc=(0.94, 0.86)):
    xl, yl = ax.get_xlim(), ax.get_ylim()
    x = xl[0] + loc[0] * (xl[1] - xl[0])
    y = yl[0] + loc[1] * (yl[1] - yl[0])
    dy = 0.07 * (yl[1] - yl[0])
    ax.annotate("", xy=(x, y + dy), xytext=(x, y),
                arrowprops=dict(arrowstyle="-|>", color="k", lw=1.6), zorder=22)
    ax.text(x, y + dy * 1.05, "N", ha="center", va="bottom", fontweight="bold", fontsize=10, zorder=22)


# ============================================== FIG 1 : topographic map
fig, ax = plt.subplots(figsize=(9.2, 7.4))
levels = np.arange(325, 455, 5)
cf = ax.contourf(XX, YY, Zs, levels=levels, cmap="terrain", alpha=0.85)
cs = ax.contour(XX, YY, Zs, levels=np.arange(325, 455, 10), colors="k", linewidths=0.5, alpha=0.7)
ax.clabel(cs, fmt="%d", fontsize=6, inline=True)
# full-pool shoreline
sh = ax.contour(XX, YY, Zs, levels=[FULL_POOL], colors="#0050b3", linewidths=2.4)
ax.contour(XX, YY, Zs, levels=[335.0], colors="#0050b3", linewidths=1.0, linestyles=":")
ax.plot(road_xy[:, 0], road_xy[:, 1], color="#5c3317", lw=2.6, zorder=8)
ax.plot(road_xy[:, 0], road_xy[:, 1], color="w", lw=0.8, ls=(0, (6, 6)), zorder=9)
for lot, lab in ((lot_b, "Lot 130\n(1147)"), (lot_a, "Lot 131\n(1153)")):
    ax.add_patch(MplPoly(lot, closed=True, fill=False, ec="#c1121f", lw=2.0, zorder=10))
    _dx = -95 if lab.startswith("Lot 130") else 95
    ax.text(lot[:, 0].mean() + _dx, lot[:, 1].mean(), lab, ha="center", va="center",
            fontsize=7.5, color="#c1121f", fontweight="bold", zorder=11,
            bbox=dict(fc="white", alpha=.8, ec="#c1121f", lw=.5, pad=1.5))
# high / low point
hi = np.unravel_index(np.nanargmax(Z), Z.shape)
lo_ = np.unravel_index(np.nanargmin(Z), Z.shape)
ax.plot(XX[hi], YY[hi], "k^", ms=8, zorder=15)
ax.text(XX[hi], YY[hi], f"  high {Z[hi]:.0f} ft", fontsize=7, va="center", zorder=15)
ax.plot(XX[lo_], YY[lo_], "bv", ms=8, zorder=15)
ax.text(XX[lo_], YY[lo_], f"  low {Z[lo_]:.0f} ft", fontsize=7, va="center", zorder=15)
# drainage arrows (down-gradient)
gy, gx = np.gradient(Zs)
for iy in range(3, len(lats) - 3, 7):
    for ix in range(3, len(lons) - 3, 9):
        u, v = -gx[iy, ix], -gy[iy, ix]
        n = math.hypot(u, v)
        if n < 1e-6:
            continue
        ax.arrow(XX[iy, ix], YY[iy, ix], u / n * 42, v / n * 42, head_width=14,
                 head_length=14, fc="#1f4e79", ec="#1f4e79", alpha=.55, lw=0.5, zorder=7)
ax.set_aspect("equal")
ax.set_xlabel("feet east of grid SW corner")
ax.set_ylabel("feet north of grid SW corner")
ax.set_title("Fig. 1  True-scale topographic map — Watersedge Cove study area\n"
             "USGS 3DEP 1 m lidar (NAVD 88) | contours 5 ft | blue = 330 ft full-pool shoreline",
             fontsize=10)
cb = fig.colorbar(cf, ax=ax, shrink=.85, pad=.02)
cb.set_label("Elevation (ft, NAVD 88)")
scalebar(ax, 400)
northarrow(ax)
ax.text(0.015, 0.985, f"SW corner {lats[0]:.5f} N, {lons[0]:.5f} W\n"
                      f"NE corner {lats[-1]:.5f} N, {lons[-1]:.5f} W\n"
                      "Lot rectangles are MODELLED (150x310 ft) — not surveyed",
        transform=ax.transAxes, va="top", fontsize=7,
        bbox=dict(fc="#fffbe6", ec="#999", lw=.5))
ax.legend(handles=[
    Line2D([], [], color="#0050b3", lw=2.4, label="330 ft full pool (lake edge)"),
    Line2D([], [], color="#0050b3", lw=1.0, ls=":", label="335 ft top-of-gates"),
    Line2D([], [], color="#5c3317", lw=2.6, label="Watersedge Cove (road)"),
    Line2D([], [], color="#c1121f", lw=2.0, label="Modelled lot outline"),
    Line2D([], [], color="#1f4e79", lw=1, marker=">", label="Drainage direction"),
], loc="lower right", fontsize=7, framealpha=.9)
fig.tight_layout()
fig.savefig(f"{OUT}/fig1_topo.png", dpi=170)
plt.close(fig)
print("fig1 ok")

# ============================================== FIG 2 : cross-sections
tr = json.load(open("/projects/sandbox/work/transects.json"))
key = "Road sta. A" if "Road sta. A" in tr else list(tr)[0]
prof = np.array(tr[key]["profile"])          # metres along, ft elev
d_ft = prof[:, 0] * FT
z_ft = prof[:, 1]
st_ = tr[key]["stats"]

fig = plt.figure(figsize=(11.6, 7.6))
gs = fig.add_gridspec(2, 1, height_ratios=[1, 1], hspace=.38)

# (a) vertically exaggerated
axa = fig.add_subplot(gs[0])
axa.fill_between(d_ft, 300, z_ft, color="#c9a227", alpha=.45)
axa.plot(d_ft, z_ft, "k-", lw=1.6)
axa.axhline(FULL_POOL, color="#0050b3", lw=1.6)
axa.axhline(335, color="#0050b3", lw=0.9, ls=":")
axa.fill_between(d_ft, 300, FULL_POOL, color="#9ecae1", alpha=.75)
axa.text(6, FULL_POOL - 2.5, "330 ft full pool", color="#0050b3", fontsize=8, va="top", ha="left")
axa.text(6, 336.5, "335 ft top of spillway gates", color="#0050b3", fontsize=7, va="bottom", ha="left")
axa.axvline(310, color="#c1121f", lw=1.3, ls="--")
axa.text(310, z_ft.max() - 3, " modelled rear lot line\n (310 ft deep)", color="#c1121f", fontsize=8, va="top")
axa.set_xlim(0, d_ft[-1])
axa.set_ylim(300, z_ft.max() + 12)
ve = ((d_ft[-1]) / (z_ft.max() + 12 - 300))
axa.set_title(f"(a) Road-to-lake section, VERTICALLY EXAGGERATED (~{ve:.0f}:1) — slopes look far steeper than reality",
              fontsize=10)
axa.set_xlabel("Horizontal distance from Watersedge Cove centreline (ft)")
axa.set_ylabel("Elevation (ft NAVD 88)")
axa.grid(alpha=.3)

# (b) true 1:1
axb = fig.add_subplot(gs[1])
axb.fill_between(d_ft, 300, z_ft, color="#c9a227", alpha=.45)
axb.plot(d_ft, z_ft, "k-", lw=1.6)
axb.fill_between(d_ft, 300, FULL_POOL, color="#9ecae1", alpha=.75)
axb.axhline(FULL_POOL, color="#0050b3", lw=1.4)
axb.axvline(310, color="#c1121f", lw=1.3, ls="--")
axb.set_aspect("equal")
axb.set_xlim(0, d_ft[-1])
axb.set_ylim(300, 300 + d_ft[-1] * 0.34)
axb.set_title("(b) The SAME section at TRUE 1:1 scale — actual perceived steepness", fontsize=10)
axb.set_xlabel("Horizontal distance (ft)   [1 ft horizontal = 1 ft vertical]")
axb.set_ylabel("Elevation (ft)")
axb.grid(alpha=.3)
txt = (f"Road crown  {st_['z_start_ft']} ft   ->   full pool {FULL_POOL:.0f} ft\n"
       f"Total fall  {st_['z_start_ft']-FULL_POOL:.1f} ft over {d_ft[-1]:.0f} ft of run\n"
       f"Average slope  {st_['avg_slope_pct']}%  ({st_['avg_slope_deg']} deg)\n"
       f"Steepest 50 ft window  {st_['max50_slope_pct']}%  ({st_['max50_slope_deg']} deg)")
axb.text(.985, .94, txt, transform=axb.transAxes, ha="right", va="top", fontsize=8.5,
         family="monospace", bbox=dict(fc="#fffbe6", ec="#999", lw=.6))
fig.suptitle("Fig. 2  Longitudinal terrain section, road -> lake (1 m lidar). "
             "Exaggerated and true-scale shown together, as required.", fontsize=11)
fig.savefig(f"{OUT}/fig2_section.png", dpi=170, bbox_inches="tight")
plt.close(fig)
print("fig2 ok")

# ============================================== FIG 3 : 3-D terrain
from mpl_toolkits.mplot3d import Axes3D  # noqa
fig = plt.figure(figsize=(11, 7.6))
ax3 = fig.add_subplot(111, projection="3d")
step = 1
s = ax3.plot_surface(XX[::step, ::step], YY[::step, ::step], Zs[::step, ::step],
                     cmap="terrain", linewidth=0, antialiased=True, alpha=.97,
                     rstride=1, cstride=1, vmin=325, vmax=450)
# lake plane
msk = Zs <= FULL_POOL
if msk.any():
    ax3.plot_surface(XX, YY, np.where(msk, FULL_POOL, np.nan),
                     color="#1b6ca8", alpha=.75, linewidth=0, shade=False)
rz = []
for x, y in road_xy:
    ix = int(np.clip(np.searchsorted(X_ft, x), 0, len(X_ft) - 1))
    iy = int(np.clip(np.searchsorted(Y_ft, y), 0, len(Y_ft) - 1))
    rz.append(Zs[iy, ix] + 4)
ax3.plot(road_xy[:, 0], road_xy[:, 1], rz, color="#8b0000", lw=3, label="Watersedge Cove")
ax3.set_xlabel("ft east")
ax3.set_ylabel("ft north")
ax3.set_zlabel("ft NAVD 88")
ax3.set_zlim(320, 460)
ax3.view_init(elev=36, azim=-126)
ax3.set_title("Fig. 3  3-D terrain model — Watersedge Cove / Fishing Creek arm of Thurmond Lake\n"
              "blue plane = 330 ft full pool;  dark red = road;  vertical scale exaggerated",
              fontsize=10)
fig.colorbar(s, ax=ax3, shrink=.6, pad=.08, label="Elevation (ft)")
ax3.legend(loc="upper left", fontsize=8)
fig.savefig(f"{OUT}/fig3_3d.png", dpi=165, bbox_inches="tight")
plt.close(fig)
print("fig3 ok")

# ============================================== FIG 4 : slope map
slope_pct = np.hypot(np.gradient(Zs, axis=1) / np.gradient(X_ft).mean(),
                     np.gradient(Zs, axis=0) / np.gradient(Y_ft).mean()) * 100
fig, ax = plt.subplots(figsize=(9.2, 7.2))
bounds = [0, 2, 5, 8, 10, 15, 20, 25, 35, 60]
cmap = plt.get_cmap("RdYlGn_r", len(bounds) - 1)
norm = matplotlib.colors.BoundaryNorm(bounds, cmap.N)
im = ax.pcolormesh(XX, YY, slope_pct, cmap=cmap, norm=norm, shading="auto")
ax.contour(XX, YY, Zs, levels=[FULL_POOL], colors="#0050b3", linewidths=2)
ax.plot(road_xy[:, 0], road_xy[:, 1], color="k", lw=2.4)
for lot in (lot_a, lot_b):
    ax.add_patch(MplPoly(lot, closed=True, fill=False, ec="#111", lw=2.2, ls="--"))
ax.set_aspect("equal")
cb = fig.colorbar(im, ax=ax, shrink=.86, pad=.02, ticks=bounds)
cb.set_label("Ground slope (%)")
ax.set_title("Fig. 4  Slope classification map (derived from 1 m lidar)\n"
             "green <8% buildable with minimal grading | orange 10-20% | red >20% needs engineering",
             fontsize=10)
ax.set_xlabel("ft east");  ax.set_ylabel("ft north")
scalebar(ax, 400); northarrow(ax)
fig.tight_layout(); fig.savefig(f"{OUT}/fig4_slope.png", dpi=170); plt.close(fig)
print("fig4 ok")

# ============================================== FIG 5 : FEMA flood map
fema = json.load(open("/projects/sandbox/work/fema.json"))
fig, ax = plt.subplots(figsize=(9.2, 7.2))
ax.contourf(XX, YY, Zs, levels=np.arange(325, 455, 10), cmap="Greys", alpha=.25)
drew = {}
for f in fema["features"]:
    zone = f["attributes"]["FLD_ZONE"]
    col = "#2171b5" if zone == "A" else "#d9f0d3"
    for ring in f["geometry"]["rings"]:
        arr = np.array([ll2ft(p[1], p[0]) for p in ring])
        if arr[:, 0].max() < -800 or arr[:, 0].min() > X_ft[-1] + 800:
            continue
        if arr[:, 1].max() < -800 or arr[:, 1].min() > Y_ft[-1] + 800:
            continue
        ax.add_patch(MplPoly(arr, closed=True, fc=col, ec="#08306b" if zone == "A" else "#4daf4a",
                             lw=.6, alpha=.55 if zone == "A" else .25, zorder=3))
        drew[zone] = True
ax.contour(XX, YY, Zs, levels=[FULL_POOL], colors="#0050b3", linewidths=1.8, zorder=5)
ax.plot(road_xy[:, 0], road_xy[:, 1], color="k", lw=2.4, zorder=6)
for lot, lab in ((lot_b, "1147 / Lot 130"), (lot_a, "1153 / Lot 131")):
    ax.add_patch(MplPoly(lot, closed=True, fill=False, ec="#c1121f", lw=2.2, zorder=7))
ax.set_xlim(X_ft[0], X_ft[-1]); ax.set_ylim(Y_ft[0], Y_ft[-1])
ax.set_aspect("equal")
ax.set_title("Fig. 5  FEMA National Flood Hazard Layer — DFIRM 13181C (Lincoln County, GA)\n"
             "Zone A = Special Flood Hazard Area, NO base flood elevation published; "
             "balance of area Zone X (minimal hazard)", fontsize=9.5)
ax.legend(handles=[
    MplPoly([[0, 0]], fc="#2171b5", ec="#08306b", alpha=.55, label="Zone A (SFHA, no BFE)"),
    MplPoly([[0, 0]], fc="#d9f0d3", ec="#4daf4a", alpha=.5, label="Zone X (minimal)"),
    Line2D([], [], color="#0050b3", lw=1.8, label="330 ft full pool"),
    Line2D([], [], color="#c1121f", lw=2.2, label="Modelled lots"),
], loc="upper right", fontsize=8)
ax.set_xlabel("ft east"); ax.set_ylabel("ft north")
scalebar(ax, 400); northarrow(ax, loc=(.06, .8))
fig.tight_layout(); fig.savefig(f"{OUT}/fig5_fema.png", dpi=170); plt.close(fig)
print("fig5 ok")

# ============================================== FIG 6 : building envelope
fig, ax = plt.subplots(figsize=(9.6, 7.6))
sub_x = (X_ft > ax0 - 330) & (X_ft < ax0 + 330)
sub_y = (Y_ft > ay0 - 90) & (Y_ft < ay0 + 420)
cs = ax.contour(XX, YY, Zs, levels=np.arange(330, 450, 5), colors="#999", linewidths=.6)
ax.clabel(cs, fmt="%d", fontsize=6)
for lot, lab in ((lot_b, "Lot 130 (1147)\n1.07 ac"), (lot_a, "Lot 131 (1153)\n1.07 ac")):
    ax.add_patch(MplPoly(lot, closed=True, fc="#fff3cd", ec="#c1121f", lw=2.2, alpha=.5, zorder=4))
    ax.text(lot[:, 0].mean(), lot[:, 1].min() - 22, lab.replace("\n", "  "), ha="center",
            fontsize=7.5, color="#c1121f", fontweight="bold", zorder=9,
            bbox=dict(fc="white", alpha=.85, ec="#c1121f", lw=.5, pad=1.4))
# setbacks (ILLUSTRATIVE - county confirmation required)
FS, SS, RS = 35.0, 15.0, 35.0
for lot in (lot_a, lot_b):
    x0, y0 = lot[0]
    env = np.array([[x0 + SS, y0 + FS], [x0 + LOTW - SS, y0 + FS],
                    [x0 + LOTW - SS, y0 + LOTD - RS], [x0 + SS, y0 + LOTD - RS]])
    ax.add_patch(MplPoly(env, closed=True, fc="#b7e4c7", ec="#2d6a4f", lw=1.4, ls="--", alpha=.6, zorder=5))
# house footprint 2,800 sq ft 56x50 on the flatter upslope third
hx, hy = ax0 - LOTW / 2 - 28, ay0 + 55
ax.add_patch(Rectangle((hx, hy), 56, 50, fc="#6c757d", ec="k", lw=1.3, alpha=.9, zorder=8))
ax.text(hx + 28, hy + 25, "house\n56x50 ft\n(2,800 sf)", ha="center", va="center",
        color="w", fontsize=7, fontweight="bold", zorder=9)
# driveway
ax.plot([hx + 28, hx + 28], [ay0, hy], color="#495057", lw=7, solid_capstyle="butt", zorder=6)
ax.text(hx + 34, ay0 + 22, "drive", fontsize=7, rotation=90, va="bottom", zorder=9)
# septic primary + reserve (3,000 sf/bedroom x 4 = 12,000 sf each)
ax.add_patch(Rectangle((ax0 + 20, ay0 + 70), 110, 110, fc="#a5d8ff", ec="#1864ab",
                       lw=1.2, ls="-.", alpha=.75, zorder=7))
ax.text(ax0 + 75, ay0 + 125, "PRIMARY\nabsorption\nfield\n~12,000 sf", ha="center", va="center",
        fontsize=6.5, zorder=9)
ax.add_patch(Rectangle((ax0 + 20, ay0 + 190), 110, 110, fc="#dee2e6", ec="#495057",
                       lw=1.2, ls="-.", alpha=.75, zorder=7))
ax.text(ax0 + 75, ay0 + 245, "RESERVE\nfield\n~12,000 sf", ha="center", va="center",
        fontsize=6.5, zorder=9)
ax.plot(road_xy[:, 0], road_xy[:, 1], color="#5c3317", lw=4, zorder=3)
ax.text(ax0, ay0 - 30, "WATERSEDGE COVE", fontsize=8, ha="center", fontweight="bold")
ax.annotate("", xy=(ax0 + 230, ay0 + 330), xytext=(ax0 + 230, ay0 + 250),
            arrowprops=dict(arrowstyle="-|>", color="#0050b3", lw=2))
ax.text(ax0 + 222, ay0 + 345, "to USACE boundary\n& lake (several hundred ft)",
        fontsize=7, color="#0050b3", va="center", ha="right")
ax.set_xlim(ax0 - 360, ax0 + 300); ax.set_ylim(ay0 - 90, ay0 + 395)
ax.set_aspect("equal")
ax.set_title("Fig. 6  Illustrative building-envelope / site-layout diagram (2 combined lots)\n"
             "SETBACKS AND LOT SHAPE ARE ASSUMED — confirm with recorded plat, HOA and Lincoln County",
             fontsize=9.5)
ax.set_xlabel("ft east"); ax.set_ylabel("ft north")
ax.legend(handles=[
    MplPoly([[0, 0]], fc="#fff3cd", ec="#c1121f", label="Lot boundary (modelled)"),
    MplPoly([[0, 0]], fc="#b7e4c7", ec="#2d6a4f", ls="--", label="Buildable envelope (35/15/35 ft assumed)"),
    MplPoly([[0, 0]], fc="#6c757d", ec="k", label="4-bed house footprint"),
    MplPoly([[0, 0]], fc="#a5d8ff", ec="#1864ab", label="Septic primary field"),
    MplPoly([[0, 0]], fc="#dee2e6", ec="#495057", label="Septic reserve field"),
], loc="upper left", fontsize=7, framealpha=.92)
scalebar(ax, 100, loc=(.72, .04)); northarrow(ax, loc=(.95, .55))
fig.tight_layout(); fig.savefig(f"{OUT}/fig6_envelope.png", dpi=170); plt.close(fig)
print("fig6 ok")

# ============================================== FIG 7 : location / access
fig, ax = plt.subplots(figsize=(9.6, 6.6))
places = {
    "LOT\n1147/1153 Watersedge Cv": (33.9410, -82.5646, "#c1121f", 95),
    "Lincolnton\n(county seat, 17 mi)": (33.7940, -82.4793, "#333", 45),
    "Washington GA\nWills Mem. Hosp (20 mi)": (33.7365, -82.7393, "#1864ab", 45),
    "Thomson\nWalmart (44 mi)": (33.4697, -82.5046, "#333", 40),
    "Elijah Clark SP\n(21 mi)": (33.8549, -82.3932, "#2d6a4f", 40),
    "Augusta\nPiedmont Hosp (55 mi)": (33.4735, -81.9876, "#1864ab", 55),
    "Augusta Rgnl\nAGS (63 mi)": (33.3699, -81.9645, "#6a040f", 45),
    "Athens GA\n(61 mi)": (33.9519, -83.3576, "#333", 40),
}
for n, (la, lo, c, s) in places.items():
    ax.scatter(lo, la, s=s, c=c, zorder=5, edgecolor="w", linewidth=.8)
    ax.annotate(n, (lo, la), textcoords="offset points", xytext=(7, 5), fontsize=7.4, color=c)
    if "LOT" not in n:
        ax.plot([-82.5646, lo], [33.9410, la], color="#aaa", lw=.8, ls=":", zorder=2)
ax.set_xlabel("Longitude (deg W)"); ax.set_ylabel("Latitude (deg N)")
ax.set_title("Fig. 7  Regional location and access — Lincoln County, Georgia\n"
             "straight lines are reference only; mileages are OSRM road-network driving distances",
             fontsize=10)
ax.grid(alpha=.3)
ax.set_aspect(1 / math.cos(math.radians(33.9)))
fig.tight_layout(); fig.savefig(f"{OUT}/fig7_location.png", dpi=170); plt.close(fig)
print("fig7 ok")

# ============================================== FIG 8 : climate
months = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
tmax = [12.3, 14.5, 18.5, 23.2, 27.6, 30.9, 32.3, 31.7, 28.6, 23.6, 17.7, 13.5]
tmin = [3.5, 4.9, 8.2, 12.2, 17.1, 21.2, 23.1, 22.7, 19.7, 14.2, 8.6, 5.1]
prcp = [98.5, 100.5, 107.0, 85.7, 70.9, 88.8, 95.5, 84.3, 87.7, 64.7, 77.0, 109.0]
fig, ax = plt.subplots(figsize=(9.4, 5.0))
ax2 = ax.twinx()
ax2.bar(months, prcp, color="#9ecae1", alpha=.75, label="Precipitation")
ax.plot(months, tmax, "o-", color="#c1121f", label="Mean daily max")
ax.plot(months, tmin, "o-", color="#1864ab", label="Mean daily min")
ax.fill_between(months, tmin, tmax, color="#ffd6a5", alpha=.4)
ax.axhline(0, color="k", lw=.8, ls=":")
ax.set_ylabel("Temperature (deg C)"); ax2.set_ylabel("Precipitation (mm / month)")
ax.set_title("Fig. 8  Climate at the site (33.9410 N, 82.5646 W) — ERA5 reanalysis 1995-2024\n"
             "annual mean 18.2 C | annual precipitation 1,070 mm (42.1 in) | ~283 frost-free days",
             fontsize=10)
ax.legend(loc="upper left", fontsize=8); ax2.legend(loc="upper right", fontsize=8)
ax.grid(alpha=.3)
fig.tight_layout(); fig.savefig(f"{OUT}/fig8_climate.png", dpi=170); plt.close(fig)
print("fig8 ok")

print("\nall figures written to", OUT)
for f in sorted(os.listdir(OUT)):
    print("  ", f, os.path.getsize(os.path.join(OUT, f)) // 1024, "KB")
