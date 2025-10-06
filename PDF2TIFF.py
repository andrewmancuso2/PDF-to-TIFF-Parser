# -*- coding: utf-8 -*-
import os
import arcpy
import uuid

class Toolbox:
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the .pyt file)."""
        self.label = "Toolbox"
        self.alias = "toolbox"
        self.tools = [Tool]

class Tool:
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "PDF2TIFF"
        self.description = "PDF to TIFF conversion tool"

    def getParameterInfo(self):
        """Define the tool parameters."""
        params = []

        # Input PDF file parameter
        input_file = arcpy.Parameter(
            displayName="Input PDF File",
            name="input_pdf",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        input_file.filter.list = ["pdf"]

        # Output folder parameter
        output_folder = arcpy.Parameter(
            displayName="Output Folder Path",
            name="output_folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input"
        )

        # Page start parameter
        page_start = arcpy.Parameter(
            displayName="Page Start",
            name="page_start",
            datatype="GPLong",
            parameterType="Required",
            direction="Input"
        )

        # Page end parameter
        page_end = arcpy.Parameter(
            displayName="Page End",
            name="page_end",
            datatype="GPLong",
            parameterType="Required",
            direction="Input"
        )

        # Project ID string parameter
        project_id = arcpy.Parameter(
            displayName="Project ID",
            name="project_id",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        # County string parameter
        county = arcpy.Parameter(
            displayName="County",
            name="county",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        params.append(county)
        params.append(project_id)
        params.append(page_start)
        params.append(page_end)
        params.append(input_file)
        params.append(output_folder)
        # The order above must match the order used in the execute method.
        return params

    def isLicensed(self):
        """Set whether the tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Called whenever a parameter has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify messages created by internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        # Get inputs in the same positional order as getParameterInfo
        county = parameters[0].valueAsText
        project_id = parameters[1].valueAsText
        start_page = int(parameters[2].value) - 1  # zero-based internal
        end_page = int(parameters[3].value) - 1
        input_pdf_path = parameters[4].valueAsText
        output_folder = parameters[5].valueAsText

        aprx = arcpy.mp.ArcGISProject("CURRENT")
        map_obj = aprx.activeMap or aprx.listMaps()[0]
        target_sr = map_obj.spatialReference

        # Ensure output folder exists
        if not os.path.isdir(output_folder):
            os.makedirs(output_folder, exist_ok=True)

        # --- helpers (scoped to execute) ---
        def pick_transform(in_sr, out_sr):
            try:
                cands = arcpy.ListTransformations(in_sr, out_sr)
                return cands[0] if cands else ""
            except Exception:
                return ""

        def project_raster_in_place(in_path, out_sr):
            """
            Reproject 'in_path' to 'out_sr' IN PLACE:
            - If input has Unknown CRS -> no-op.
            - If CRS matches out_sr -> no-op.
            - Else ProjectRaster to a temporary path in same folder,
              then replace original file with projected version.
            Always returns the final path to add to the map (same as in_path).
            """
            try:
                desc = arcpy.Describe(in_path)
                in_sr = getattr(desc, "spatialReference", None)
            except Exception as ex:
                arcpy.AddWarning(f"[WARN] Could not Describe() {in_path}: {ex}")
                return in_path

            if (in_sr is None) or (in_sr.name == "Unknown"):
                arcpy.AddMessage(f"[INFO] {os.path.basename(in_path)}: no defined CRS. Leaving as-is.")
                return in_path

            # Same CRS? nothing to do
            if (in_sr.factoryCode == out_sr.factoryCode) and (in_sr.name == out_sr.name):
                arcpy.AddMessage(f"[INFO] {os.path.basename(in_path)}: already in target CRS.")
                return in_path

            folder = os.path.dirname(in_path)
            base, ext = os.path.splitext(os.path.basename(in_path))
            # temp name alongside original (to avoid cross-device moves)
            tmp_name = f"{base}.__tmp_{uuid.uuid4().hex}.tif"
            tmp_path = os.path.join(folder, tmp_name)

            gtrans = pick_transform(in_sr, out_sr)
            arcpy.AddMessage(
                f"[INFO] Reprojecting {os.path.basename(in_path)} "
                f"from '{in_sr.name}' ({in_sr.factoryCode}) to '{out_sr.name}' ({out_sr.factoryCode}) "
                f"{'(transformation: ' + gtrans + ')' if gtrans else '(no transformation)'}"
            )

            # Use NEAREST for scanned/linework TIFFs (keeps labels sharp)
            arcpy.management.ProjectRaster(
                in_raster=in_path,
                out_raster=tmp_path,
                out_coor_system=out_sr,
                resampling_type="NEAREST",
                cell_size=None,
                geographic_transform=gtrans if gtrans else ""
            )

            # Replace original safely
            try:
                # Some environments need explicit delete to release locks
                arcpy.management.Delete(in_path)
            except Exception:
                # If lock persists, try renaming original aside and then delete
                try:
                    aside = os.path.join(folder, f"{base}.__old_{uuid.uuid4().hex}.tif")
                    os.replace(in_path, aside)
                    arcpy.management.Delete(aside)
                except Exception as ex:
                    arcpy.AddWarning(f"[WARN] Could not remove original before replace: {ex}")

            # Move tmp into original name
            os.replace(tmp_path, in_path)
            arcpy.AddMessage(f"[INFO] Updated in place: {os.path.basename(in_path)}")
            return in_path
        # --- end helpers ---

        for page_num in range(start_page, end_page + 1):
            pagestring = str(page_num + 1).zfill(2)  # 01, 02, ...
            county_suffix = f"{county}_County_Project"
            output_filename = f"{project_id}_p{pagestring}_{county_suffix}.tif"
            output_path = os.path.join(output_folder, output_filename)

            # Export the specific PDF page (1-based page index)
            arcpy.conversion.PDFToTIFF(
                input_pdf_path,
                output_path,
                None,            # pdf_password (not used)
                page_num + 1     # pdf_page_number (1-based)
            )

            # Reproject IN PLACE if/when needed
            final_path = project_raster_in_place(output_path, target_sr)

            # Add the one-and-only file to the map
            map_obj.addDataFromPath(final_path)

        del aprx
        return

    def postExecute(self, parameters):
        """Runs after outputs are processed and added to the display."""
        return
