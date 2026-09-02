import arcpy
import os

arcpy.env.overwriteOutput = True

# ------------------------------------------------------------------
# Inputs
# ------------------------------------------------------------------
dgn_file = r"D:\Project-WKCDA-DCS-082026\03_DATA\Data_20260812\Batch 1\P18 4192-EA12\4192-EA12.dwg"
output_gdb = r"D:\Project-WKCDA-DCS-082026\04_MAP\WKCDA_Plans.gdb"
dataset_name = "P18_4192_EA12"

# ------------------------------------------------------------------
# Create GDB if it does not exist
# ------------------------------------------------------------------
gdb_folder = os.path.dirname(output_gdb)
gdb_name = os.path.basename(output_gdb)

if not arcpy.Exists(output_gdb):
    arcpy.management.CreateFileGDB(gdb_folder, gdb_name)

# ------------------------------------------------------------------
# Convert DGN to Geodatabase
# ------------------------------------------------------------------
print("Converting DGN to Geodatabase...")

arcpy.conversion.CADToGeodatabase(
    input_cad_datasets=dgn_file,
    out_gdb_path=output_gdb,
    out_dataset_name=dataset_name,
    reference_scale=1000
)

print("Conversion completed.")

# ------------------------------------------------------------------
# Feature classes expected from CAD conversion
# ------------------------------------------------------------------
feature_dataset = os.path.join(output_gdb, dataset_name)

arcpy.env.workspace = feature_dataset

feature_classes = arcpy.ListFeatureClasses()

# Output report
report_file = os.path.join(gdb_folder, "DGN_Layers_Report.txt")

with open(report_file, "w") as rpt:

    rpt.write(f"DGN File: {dgn_file}\n")
    rpt.write(f"Feature Dataset: {feature_dataset}\n")
    rpt.write("=" * 80 + "\n\n")

    for fc in feature_classes:

        fc_path = os.path.join(feature_dataset, fc)

        print(f"\nFeature Class: {fc}")
        rpt.write(f"Feature Class: {fc}\n")

        fields = [f.name for f in arcpy.ListFields(fc_path)]

        # CAD layer names are usually stored in the Layer field
        if "Layer" in fields:

            layer_names = set()

            with arcpy.da.SearchCursor(fc_path, ["Layer"]) as cursor:
                for row in cursor:
                    if row[0] not in layer_names:
                        layer_names.add(row[0])

            layer_names = sorted(layer_names)

            print(f"  Layers found: {len(layer_names)}")
            rpt.write(f"Layers found: {len(layer_names)}\n")

            for layer in layer_names:
                print(f"    {layer}")
                rpt.write(f"    {layer}\n")

        else:
            print("  No Layer field found.")
            rpt
