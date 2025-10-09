# -*- coding: utf-8 -*-
import os
import arcpy
import uuid

class Toolbox:
    def __init__(self):
        self.label = "Toolbox"
        self.alias = "toolbox"
        self.tools = [Tool]

class Tool:
    def __init__(self):
        self.label = "PDF2TIFF"
        self.description = "PDF to TIFF conversion tool (old-behavior export + QoL UI)"

    def getParameterInfo(self):
        params = []

        # 1) Input PDF at top
        input_file = arcpy.Parameter(
            displayName="Input PDF File",
            name="input_pdf",
            datatype="DEFile",
            parameterType="Required",
            direction="Input",
        )
        input_file.filter.list = ["pdf"]

        # 2) Output folder (left blank; no autofill)
        output_folder = arcpy.Parameter(
            displayName="Output Folder Path",
            name="output_folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input",
        )

        # 3) Page Start dropdown (GPString for reliable ValueList UI)
        page_start = arcpy.Parameter(
            displayName="Page Start",
            name="page_start",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
        )
        page_start.filter.type = "ValueList"
        page_start.filter.list = []  # populated after PDF is chosen

        # 4) Page End dropdown
        page_end = arcpy.Parameter(
            displayName="Page End",
            name="page_end",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
        )
        page_end.filter.type = "ValueList"
        page_end.filter.list = []  # populated after PDF is chosen

        # 5) Project ID
        project_id = arcpy.Parameter(
            displayName="Project ID",
            name="project_id",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
        )

        # 6) County
        county = arcpy.Parameter(
            displayName="County",
            name="county",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
        )

        # UI order
        params.append(input_file)
        params.append(output_folder)
        params.append(page_start)
        params.append(page_end)
        params.append(project_id)
        params.append(county)

        # Refresh page lists when PDF changes
        page_start.parameterDependencies = [input_file.name]
        page_end.parameterDependencies   = [input_file.name]

        return params

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        """Populate page dropdowns from selected PDF; keep End >= Start; do not autofill output folder."""
        input_pdf, output_folder, page_start, page_end, project_id, county = parameters

        # Populate dropdowns from the chosen PDF
        if input_pdf.altered and input_pdf.valueAsText and os.path.isfile(input_pdf.valueAsText):
            try:
                pdf_doc = arcpy.mp.PDFDocumentOpen(input_pdf.valueAsText)
                page_count = pdf_doc.pageCount  # 1-based
                del pdf_doc

                pages = [str(i) for i in range(1, page_count + 1)]
                page_start.filter.list = pages
                page_end.filter.list   = pages

                # Set defaults only if user hasn't set them yet
                if not page_start.altered:
                    page_start.value = "1"
                if not page_end.altered:
                    page_end.value = str(page_count)

            except Exception as ex:
                arcpy.AddWarning(f"Could not read PDF page count: {ex}")
                page_start.filter.list = []
                page_end.filter.list   = []

        # If user pasted a file path into the folder field, coerce to its directory (no autofill otherwise)
        if output_folder.valueAsText:
            path = output_folder.valueAsText
            root, ext = os.path.splitext(path)
            if ext and (os.path.isfile(path) or ext.lower() in [".pdf", ".tif", ".tiff", ".jpg", ".png"]):
                dir_only = os.path.dirname(path)
                if dir_only and os.path.isdir(dir_only):
                    output_folder.value = dir_only

        # Ensure End >= Start if both are set
        try:
            if page_start.value and page_end.value:
                if int(page_end.value) < int(page_start.value):
                    page_end.value = page_start.value
        except Exception:
            pass

        return

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        arcpy.env.overwriteOutput = True

        input_pdf_path = parameters[0].valueAsText
        output_folder  = parameters[1].valueAsText
        start_page = int(parameters[2].value) - 1   # 0-based internal
        end_page   = int(parameters[3].value) - 1
        project_id = parameters[4].valueAsText
        county     = parameters[5].valueAsText

        # Final safety: if output_folder points to a file, use its directory
        if output_folder:
            root, ext = os.path.splitext(output_folder)
            if ext and not os.path.isdir(output_folder):
                maybe_dir = os.path.dirname(output_folder)
                if maybe_dir and os.path.isdir(maybe_dir):
                    output_folder = maybe_dir

        aprx = arcpy.mp.ArcGISProject("CURRENT")
        map_obj = aprx.activeMap or aprx.listMaps()[0]
        target_sr = map_obj.spatialReference

        # Ensure output folder exists (user chose it; we don't autofill)
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
              then atomically replace original file with projected version.
            Returns the final path (same as input).
            """
            try:
                desc = arcpy.Describe(in_path)
                in_sr = getattr(desc, "spatialReference", None)
            except Exception as ex:
                arcpy.AddWarning(f"[WARN] Describe failed on {in_path}: {ex}")
                return in_path

            if (in_sr is None) or (in_sr.name == "Unknown"):
                arcpy.AddMessage(f"[INFO] {os.path.basename(in_path)}: no defined CRS. Leaving as-is.")
                return in_path

            if (in_sr.factoryCode == out_sr.factoryCode) and (in_sr.name == out_sr.name):
                arcpy.AddMessage(f"[INFO] {os.path.basename(in_path)}: already in target CRS.")
                return in_path

            folder = os.path.dirname(in_path)
            base, _ = os.path.splitext(os.path.basename(in_path))
            tmp_path = os.path.join(folder, f"{base}.__tmp_{uuid.uuid4().hex}.tif")

            gtrans = pick_transform(in_sr, out_sr)
            arcpy.AddMessage(
                f"[INFO] Reprojecting {os.path.basename(in_path)} "
                f"from '{in_sr.name}' ({in_sr.factoryCode}) to '{out_sr.name}' ({out_sr.factoryCode}) "
                f"{'(transformation: ' + gtrans + ')' if gtrans else '(no transformation)'}"
            )

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
                arcpy.management.Delete(in_path)
            except Exception:
                try:
                    aside = os.path.join(folder, f"{base}.__old_{uuid.uuid4().hex}.tif")
                    os.replace(in_path, aside)
                    arcpy.management.Delete(aside)
                except Exception as ex:
                    arcpy.AddWarning(f"[WARN] Could not remove original before replace: {ex}")

            os.replace(tmp_path, in_path)
            arcpy.AddMessage(f"[INFO] Updated in place: {os.path.basename(in_path)}")
            return in_path
        # --- end helpers ---

        # Preflight: page count check (avoid 003464 on out-of-range)
        try:
            pdf_doc = arcpy.mp.PDFDocumentOpen(input_pdf_path)
            page_count = pdf_doc.pageCount
            del pdf_doc
        except Exception as ex:
            arcpy.AddWarning(f"[WARN] Could not read page count: {ex}")
            page_count = None

        if page_count is not None:
            req_start_1b = start_page + 1
            req_end_1b   = end_page + 1
            if req_start_1b < 1 or req_end_1b > page_count:
                arcpy.AddError(
                    f"Requested page range {req_start_1b}-{req_end_1b} exceeds this PDF's pageCount={page_count}."
                )
                raise arcpy.ExecuteError

        # Export loop — EXACTLY the old working behavior:
        # arcpy.conversion.PDFToTIFF(input_pdf, output_tif, None, page_num+1)
        for page_num in range(start_page, end_page + 1):
            page_1b = page_num + 1
            pagestring = str(page_1b).zfill(2)  # 01, 02, ...
            county_suffix = f"{county}_County_Project"
            output_filename = f"{project_id}_p{pagestring}_{county_suffix}.tif"
            output_path = os.path.join(output_folder, output_filename)

            # Clean any previous export to avoid locks/warnings
            if os.path.exists(output_path):
                try:
                    arcpy.management.Delete(output_path)
                except Exception:
                    pass

            # --- The core export call (same as your old code) ---
            try:
                arcpy.conversion.PDFToTIFF(
                    input_pdf_path,
                    output_path,
                    None,            # pdf_password (not used)
                    page_1b          # pdf_page_number (1-based)
                )
                arcpy.AddMessage(f"[INFO] Exported page {page_1b} via 4-arg PDFToTIFF.")
            except Exception as ex:
                arcpy.AddError(f"Failed exporting page {page_1b}: {ex}")
                raise

            # Reproject IN PLACE if/when needed
            final_path = project_raster_in_place(output_path, target_sr)

            # Add the one-and-only file to the map
            map_obj.addDataFromPath(final_path)

        del aprx
        return

    def postExecute(self, parameters):
        return

