#!/usr/bin/env python3

"""
ArcGIS Enterprise Content Migration

Migrates:
- Hosted Feature Layers
- Web Maps
- Web Scenes

while preserving Portal Item IDs when supported.

Author: GIS Administrator
"""

import json
import logging
from datetime import datetime
from arcgis.gis import GIS

# =====================================================
# CONFIGURATION
# =====================================================

SOURCE_PORTAL = "https://www-uat.wkcdagis.com/portal"
SOURCE_USERNAME = "appadmin"
SOURCE_PASSWORD = ""

TARGET_PORTAL = "https://www.wkcdagis.com/portal"
TARGET_USERNAME = "appAdmin"
TARGET_PASSWORD = ""

# Item IDs to migrate
ITEM_IDS = [
    "3719892a188a4f8181280fb0c7a1b048",
    "ac6aafe1c1ba488ba0615566a2b298eb",
    "357075b18e1b474aa1f8581d6e6f41b3",
    "57a30fd4875e4ee49d128ff59417e0b7"
]

# =====================================================
# LOGGING
# =====================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler("migration.log"),
        logging.StreamHandler()
    ]
)

# =====================================================
# CONNECT
# =====================================================

def connect_portals():
    logging.info("Connecting to source portal...")
    source = GIS(
        SOURCE_PORTAL,
        SOURCE_USERNAME,
        SOURCE_PASSWORD
    )

    logging.info("Connecting to target portal...")
    target = GIS(
        TARGET_PORTAL,
        TARGET_USERNAME,
        TARGET_PASSWORD
    )

    return source, target

# =====================================================
# VALIDATION
# =====================================================

SUPPORTED_TYPES = [
    "Feature Layer",
    "Feature Service",
    "Web Map",
    "Web Scene"
]

def validate_item(item):
    if item is None:
        return False

    if item.type not in SUPPORTED_TYPES:
        logging.warning(
            f"Unsupported item type: {item.type} ({item.title})"
        )
        return False

    return True

# =====================================================
# CLONE
# =====================================================

def clone_item(source_item, target):
    """
    Clone item preserving ID.
    """

    logging.info(
        f"Cloning: {source_item.title} "
        f"[{source_item.type}] "
        f"ID={source_item.itemid}"
    )

    try:

        cloned_items = target.content.clone_items(
            items=[source_item],

            preserve_item_id=True,

            copy_data=True,

            search_existing_items=True
        )

        return cloned_items

    except Exception as ex:
        logging.exception(
            f"Failed cloning {source_item.itemid}: {ex}"
        )
        return None

# =====================================================
# REPORT
# =====================================================

def write_report(results):

    report_file = (
        f"migration_report_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    )

    with open(report_file, "w", encoding="utf-8") as f:
        json.dump(
            results,
            f,
            indent=4
        )

    logging.info(f"Report written to {report_file}")

# =====================================================
# MAIN
# =====================================================

def main():

    source, target = connect_portals()

    results = []

    for item_id in ITEM_IDS:

        logging.info(
            f"Processing source item: {item_id}"
        )

        item = source.content.get(item_id)

        if not validate_item(item):
            continue

        cloned = clone_item(item, target)

        if not cloned:
            continue

        for new_item in cloned:

            result = {
                "title": new_item.title,
                "type": new_item.type,
                "source_item_id": item.itemid,
                "target_item_id": new_item.itemid,
                "preserved":
                    item.itemid.lower()
                    ==
                    new_item.itemid.lower(),
                "url": getattr(new_item, "url", None)
            }

            logging.info(
                f"Source ID : {item.itemid}"
            )

            logging.info(
                f"Target ID : {new_item.itemid}"
            )

            logging.info(
                f"Preserved : "
                f"{result['preserved']}"
            )

            results.append(result)

##    write_report(results)

    logging.info("Migration completed")


if __name__ == "__main__":
    main()
