"""
Auto-publish every (non-basemap) feature layer in one or more maps of an
ArcGIS Pro project as individual hosted feature layers on a portal.

  * Service name           = source feature class name
  * Item title (after pub) = layer name as shown in ArcGIS Pro
  * Existing service       -> overwrite ; otherwise -> publish new

Settings come from a JSON config file (default: config.json next to this
script). Any setting can be overridden on the command line:

  python publish_maps.py
  python publish_maps.py --maps "Map" "Utilities"
  python publish_maps.py --config C:\secrets\prod.json --folder Scratch
"""
import argparse
import getpass
import json
import os
import re
import sys
import tempfile
import traceback

import arcpy
from arcgis.gis import GIS
from arcgis.features import FeatureLayerCollection

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG = os.path.join(SCRIPT_DIR, "config.json")

# Built-in defaults; config file and CLI layer on top of these
DEFAULTS = {
    "aprx": "..\\WKCDA_MAP.aprx",
    "portal": "https://www-uat.wkcdagis.com/portal",
    "user": "appadmin",
    "password": "",
    "password_env_var": "PORTAL_PWD",
    "folder": "",
    "maps": [],
    "workdir": "",
    "tags": "auto-publish, arcpy",
}
REQUIRED = ("aprx", "portal", "user", "maps")


# --------------------------------------------------------------------------- #
# Config loading
# --------------------------------------------------------------------------- #
def load_config(path):
    if not os.path.isfile(path):
        sys.exit(
            f"Config file not found: {path}\n"
            f"Copy config.example.json to config.json and fill in your settings."
        )
    with open(path, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
        except json.JSONDecodeError as e:
            sys.exit(f"Config file is not valid JSON ({path}): {e}")

    unknown = set(data) - set(DEFAULTS)
    if unknown:
        print(f"WARNING: unknown keys in config ignored: {', '.join(sorted(unknown))}")

    cfg = dict(DEFAULTS)
    cfg.update({k: v for k, v in data.items() if k in DEFAULTS})
    return cfg


def parse_args():
    ap = argparse.ArgumentParser(
        description="Auto-publish map layers as hosted feature layers. "
                    "Settings are read from a JSON config; CLI flags override it.")
    ap.add_argument("--config", default=DEFAULT_CONFIG,
                    help=f"Path to config JSON (default: {DEFAULT_CONFIG})")
    ap.add_argument("--maps", nargs="+", help="Map name(s) to publish")
    ap.add_argument("--user", help="Portal username")
    ap.add_argument("--password", help="Portal password (else config / env var / prompt)")
    ap.add_argument("--aprx", help="Project (.aprx) path")
    ap.add_argument("--portal", help="Portal URL")
    ap.add_argument("--folder", help="Portal folder")
    ap.add_argument("--workdir", help="Working folder for .sddraft/.sd")
    a = ap.parse_args()

    cfg = load_config(a.config)

    # CLI overrides
    for key in ("maps", "user", "password", "aprx", "portal", "folder", "workdir"):
        val = getattr(a, key)
        if val is not None:
            cfg[key] = val

    # Validate
    missing = [k for k in REQUIRED if not cfg[k]]
    if missing:
        ap.error(f"Missing required setting(s): {', '.join(missing)} "
                 f"(set in {a.config} or pass on the command line)")
    if isinstance(cfg["maps"], str):
        cfg["maps"] = [cfg["maps"]]
    if not os.path.isfile(cfg["aprx"]):
        ap.error(f"Project file not found: {cfg['aprx']}")
    if not cfg["workdir"]:
        cfg["workdir"] = os.path.join(tempfile.gettempdir(), "auto_publish")

    # Password: CLI > config > env var > prompt
    if not cfg["password"]:
        cfg["password"] = os.environ.get(cfg["password_env_var"] or "", "")
    if not cfg["password"]:
        cfg["password"] = getpass.getpass(f"Password for {cfg['user']} @ {cfg['portal']}: ")

    cfg["_config_path"] = a.config
    return cfg


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def sanitize_service_name(name):
    name = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if re.match(r"^\d", name):
        name = "_" + name
    return name


def get_source_fc_name(lyr):
    ds = None
    try:
        cp = lyr.connectionProperties
        if cp:
            ds = cp.get("dataset")
            if not ds and "source" in cp:
                ds = cp["source"].get("dataset")
    except Exception:
        pass
    if not ds:
        ds = os.path.basename(lyr.dataSource)
    ds = ds.split(".")[-1]
    ds = os.path.splitext(ds)[0]
    return ds


def find_hosted_service(gis, username, service_name):
    q = f'owner:{username} AND type:"Feature Service"'
    for item in gis.content.search(query=q, max_items=2000):
        url = (item.url or "").rstrip("/")
        if url.lower().endswith(f"/{service_name.lower()}/featureserver"):
            if "Hosted Service" in item.typeKeywords and "View Service" not in item.typeKeywords:
                return item
    return None


def _folder_name(f):
    """Return a folder's name whether it's a Folder object (API >= 2.3) or a dict (older)."""
    if isinstance(f, dict):
        return f.get("title") or f.get("name")
    return getattr(f, "name", None) or getattr(f, "title", None) or f.properties.get("title")


def ensure_folder(gis, folder_name):
    if not folder_name:
        return
    existing = [_folder_name(f) for f in gis.users.me.folders]
    if folder_name in existing:
        return

    folders_mgr = getattr(gis.content, "folders", None)
    if folders_mgr is not None:                       # API >= 2.3
        folders_mgr.create(folder_name)
    else:                                             # older API
        gis.content.create_folder(folder_name)
    print(f"  Created portal folder '{folder_name}'")

def set_item_title(item, title):
    """Rename a portal item if its title differs. Works on arcgis 2.x and older."""
    if not title or item.title == title:
        return
    try:
        item.update(item_properties={"title": title})
        print(f"  Renamed feature layer item -> '{title}'")
    except Exception as e:
        print(f"  WARNING: could not rename item '{item.title}' -> '{title}': {e}")


def rename_items(gis, fs_item, new_title):
    if fs_item.title != new_title:
        fs_item.update(item_properties={"title": new_title})
        print(f"  Renamed feature layer item -> '{new_title}'")
    try:
        for sd_item in fs_item.related_items("Service2Data", "forward"):
            if sd_item.type == "Service Definition" and sd_item.title != new_title:
                sd_item.update(item_properties={"title": new_title})
                print(f"  Renamed service definition item -> '{new_title}'")
    except Exception as e:
        print(f"  WARNING: could not rename SD item: {e}")


# --------------------------------------------------------------------------- #
# Publish ONE layer
# --------------------------------------------------------------------------- #
def publish_layer(m, lyr, gis, username, cfg):
    layer_title = lyr.name
    fc_name = get_source_fc_name(lyr)
    service_name = sanitize_service_name(fc_name)

    print(f"\n--- Layer '{layer_title}'  (source FC: {fc_name} -> service: {service_name})")

    existing = find_hosted_service(gis, username, service_name)
    overwrite = existing is not None
    print("  Existing service found -> OVERWRITE" if overwrite else "  No existing service -> PUBLISH NEW")

    sddraft = os.path.join(cfg["workdir"], service_name + ".sddraft")
    sd = os.path.join(cfg["workdir"], service_name + ".sd")
    for f in (sddraft, sd):
        if os.path.exists(f):
            os.remove(f)

    draft = m.getWebLayerSharingDraft("HOSTING_SERVER", "FEATURE", service_name, [lyr])
    draft.summary = f"{layer_title} (auto-published from map '{m.name}')"
    draft.tags = cfg["tags"]
    draft.description = f"Source feature class: {fc_name}"
    draft.credits = ""
    draft.useLimitations = ""
    if cfg["folder"]:
        draft.portalFolder = cfg["folder"]
    draft.overwriteExistingService = overwrite
    draft.copyDataToServer = True
    draft.exportToSDDraft(sddraft)

    arcpy.server.StageService(sddraft, sd)

    if overwrite:
        try:
            arcpy.server.UploadServiceDefinition(sd, "My Hosted Services")
        except arcpy.ExecuteError:
            print("  arcpy overwrite failed, falling back to FeatureLayerCollection.manager.overwrite()")
            print("  " + arcpy.GetMessages(2).replace("\n", "\n  "))
            flc = FeatureLayerCollection.fromitem(existing)
            result = flc.manager.overwrite(sd)
            if not result.get("success", False):
                raise RuntimeError(f"Overwrite failed: {result}")
        # Re-fetch: the overwrite resets the item title, and `existing` holds stale cached properties
        fs_item = gis.content.get(existing.id)
    else:
        arcpy.server.UploadServiceDefinition(sd, "My Hosted Services")
        fs_item = find_hosted_service(gis, username, service_name)
        if fs_item is None:
            raise RuntimeError("Published, but could not locate the new item on the portal.")

    rename_items(gis, fs_item, layer_title)
    print(f"  DONE  ->  {fs_item.homepage}")
    return fs_item


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    cfg = parse_args()
    os.makedirs(cfg["workdir"], exist_ok=True)

    print(f"Config  : {cfg['_config_path']}")
    print(f"Project : {cfg['aprx']}")
    print(f"Portal  : {cfg['portal']}")
    print(f"User    : {cfg['user']}")
    print(f"Folder  : {cfg['folder'] or '(root)'}")
    print(f"Maps    : {', '.join(cfg['maps'])}")

    print(f"\nSigning in to {cfg['portal']} ...")
    arcpy.SignInToPortal(cfg["portal"], cfg["user"], cfg["password"])
    gis = GIS(cfg["portal"], cfg["user"], cfg["password"])
    username = gis.users.me.username
    ensure_folder(gis, cfg["folder"])

    aprx = arcpy.mp.ArcGISProject(cfg["aprx"])
    results = {"published": [], "failed": [], "skipped": []}

    for map_name in cfg["maps"]:
        maps = aprx.listMaps(map_name)
        if not maps:
            print(f"\n!!! Map '{map_name}' not found in project - skipping")
            results["skipped"].append(f"{map_name} (map not found)")
            continue
        m = maps[0]
        print(f"\n=============== MAP: {m.name} ===============")

        for lyr in m.listLayers():
            tag = f"{m.name} / {lyr.longName}"
            if lyr.isBasemapLayer:
                results["skipped"].append(f"{tag} (basemap)"); continue
            if lyr.isGroupLayer:
                continue
            if not lyr.isFeatureLayer:
                results["skipped"].append(f"{tag} (not a feature layer)"); continue
            if lyr.isBroken:
                results["skipped"].append(f"{tag} (broken data source)"); continue
            if lyr.isWebLayer:
                results["skipped"].append(f"{tag} (already a web layer)"); continue

            try:
                item = publish_layer(m, lyr, gis, username, cfg)
                results["published"].append(f"{tag} -> {item.title} ({item.id})")
            except Exception:
                print(f"  !!! FAILED: {tag}")
                print("  " + arcpy.GetMessages(2).replace("\n", "\n  "))
                traceback.print_exc()
                results["failed"].append(tag)

    print("\n================ SUMMARY ================")
    for k in ("published", "failed", "skipped"):
        print(f"{k.upper()} ({len(results[k])}):")
        for r in results[k]:
            print("   ", r)

    del aprx
    sys.exit(1 if results["failed"] else 0)


if __name__ == "__main__":
    main()
