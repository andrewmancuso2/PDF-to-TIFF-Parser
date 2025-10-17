# -*- coding: utf-8 -*-
import os
import re
import uuid
import arcpy

class Toolbox:
    def __init__(self):
        self.label = "Toolbox"
        self.alias = "toolbox"
        self.tools = [Tool]

class Tool:
    def __init__(self):
        self.label = "PDF2TIFF"
        self.description = "PDF to TIFF (base behavior) + skip failures + failures table named by Project ID in project GDB"

    def getParameterInfo(self):
        params = []

        # 1) Input PDF
        input_file = arcpy.Parameter(
            displayName="Input PDF File",
            name="input_pdf",
            datatype="DEFile",
            parameterType="Required",
            direction="Input",
        )
        input_file.filter.list = ["pdf"]

        # 2) Output folder (no autofill)
        output_folder = arcpy.Parameter(
            displayName="Output Folder Path",
            name="output_folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input",
        )

        # 3/4) Page range dropdowns
        page_start = arcpy.Parameter(
            displayName="Page Start",
            name="page_start",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
        )
        page_start.filter.type = "ValueList"
        page_start.filter.list = []

        page_end = arcpy.Parameter(
            displayName="Page End",
            name="page_end",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
        )
        page_end.filter.type = "ValueList"
        page_end.filter.list = []

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

        params.extend([input_file, output_folder, page_start, page_end, project_id, county])

        # Populate page lists when PDF changes
        page_start.parameterDependencies = [input_file.name]
        page_end.parameterDependencies   = [input_file.name]
        return params

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        input_pdf, output_folder, page_start, page_end, project_id, county = parameters

        # Populate page dropdowns from the chosen PDF
        if input_pdf.altered and input_pdf.valueAsText and os.path.isfile(input_pdf.valueAsText):
            try:
                pdf_doc = arcpy.mp.PDFDocumentOpen(input_pdf.valueAsText)
                page_count = pdf_doc.pageCount  # 1-based
                del pdf_doc

                pages = [str(i) for i in range(1, page_count + 1)]
                page_start.filter.list = pages
                page_end.filter.list   = pages

                # Defaults (only if user hasn't set them yet)
                if not page_start.altered:
                    page_start.value = "1"
                if not page_end.altered:
                    page_end.value = str(page_count)

                # Clamp stale selections if user changed PDFs
                try:
                    if page_start.value and int(page_start.value) > page_count:
                        page_start.value = str(page_count)
                    if page_end.value and int(page_end.value) > page_count:
                        page_end.value = str(page_count)
                except Exception:
                    pass

            except Exception as ex:
                arcpy.AddWarning(f"Could not read PDF page count: {ex}")
                page_start.filter.list = []
                page_end.filter.list   = []

        # If a file path was pasted into folder field, coerce to its directory
        if output_folder.valueAsText:
            path = output_folder.valueAsText
            root, ext = os.path.splitext(path)
            if ext and (os.path.isfile(path) or ext.lower() in [".pdf", ".tif", ".tiff", ".jpg", ".png"]):
                dir_only = os.path.dirname(path)
                if dir_only and os.path.isdir(dir_only):
                    output_folder.value = dir_only

        # Keep End >= Start
        try:
            if page_start.value and page_end.value:
                if int(page_end.value) < int(page_start.value):
                    page_end.value = page_start.value
        except Exception:
            pass

        return

    def updateMessages(self, parameters):
        return

    # ---------- helpers ----------
    @staticmethod
    def _sanitize_for_fs(s, maxlen=120):
        return re.sub(r'[<>:"/\\|?*\s]+', '_', s.strip())[:maxlen]

    @staticmethod
    def _sr_equal(a, b):
        try:
            if a.factoryCode and b.factoryCode:
                return a.factoryCode == b.factoryCode
            return a.name == b.name
        except Exception:
            return False

    @staticmethod
    def _pick_transform(in_sr, out_sr):
        try:
            cands = arcpy.ListTransformations(in_sr, out_sr)
            if not cands:
                return ""
            for key in ("NAD_1983", "HARN", "NSRS2007", "2011", "ETRS", "NAVD"):
                for c in cands:
                    if key in c:
                        return c
            return cands[0]
        except Exception:
            return ""
    # -----------------------------

    def execute(self, parameters, messages):
        arcpy.env.overwriteOutput = True

        input_pdf_path = parameters[0].valueAsText
        output_folder  = parameters[1].valueAsText
        start_page     = int(parameters[2].value) - 1   # 0-based internal
        end_page       = int(parameters[3].value) - 1
        project_id     = parameters[4].valueAsText
        county         = parameters[5].valueAsText

        # Ensure output folder exists (user chooses it; we don't autofill)
        if output_folder:
            root, ext = os.path.splitext(output_folder)
            if ext and not os.path.isdir(output_folder):
                maybe_dir = os.path.dirname(output_folder)
                if maybe_dir and os.path.isdir(maybe_dir):
                    output_folder = maybe_dir
        if not os.path.isdir(output_folder):
            os.makedirs(output_folder, exist_ok=True)

        # Get map + target SR
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        map_obj = aprx.activeMap or (aprx.listMaps()[0] if aprx.listMaps() else None)
        target_sr = map_obj.spatialReference if map_obj else arcpy.SpatialReference(4326)

        # Preflight: page count & requested range
        try:
            pdf_doc = arcpy.mp.PDFDocumentOpen(input_pdf_path)
            page_count = pdf_doc.pageCount
            del pdf_doc
        except Exception as ex:
            arcpy.AddError(f"Could not read page count: {ex}")
            raise arcpy.ExecuteError

        req_start_1b = start_page + 1
        req_end_1b   = end_page + 1
        if req_start_1b < 1 or req_end_1b > page_count:
            arcpy.AddError(
                f"Requested page range {req_start_1b}-{req_end_1b} exceeds this PDF's pageCount={page_count}."
            )
            raise arcpy.ExecuteError

        # Failure collector — recorded only when an export fails
        failures = []

        # --- helpers (scoped to execute) ---
        def pick_transform(in_sr, out_sr):
            return self._pick_transform(in_sr, out_sr)

        def project_raster_in_place(in_path, out_sr):
            """
            Reproject 'in_path' to 'out_sr' IN PLACE (base behavior):
            - If input has Unknown CRS -> no-op.
            - If CRS matches out_sr -> no-op.
            - Else ProjectRaster to a temporary path in same folder, then replace original.
            """
            try:
                desc = arcpy.Describe(in_path)
                in_sr = getattr(desc, "spatialReference", None)
            except Exception as ex:
                arcpy.AddWarning(f"[WARN] Describe failed on {in_path}: {ex}")
                return in_path

            if (in_sr is None) or (in_sr.name == "Unknown"):
                # Export succeeded; just no defined CRS — treat as success
                return in_path

            if self._sr_equal(in_sr, out_sr):
                return in_path

            folder = os.path.dirname(in_path)
            base, _ = os.path.splitext(os.path.basename(in_path))
            tmp_path = os.path.join(folder, f"{base}.__tmp_{uuid.uuid4().hex}.tif")

            gtrans = pick_transform(in_sr, out_sr)
            try:
                arcpy.management.ProjectRaster(
                    in_raster=in_path,
                    out_raster=tmp_path,
                    out_coor_system=out_sr,
                    resampling_type="NEAREST",
                    cell_size=None,
                    geographic_transform=gtrans if gtrans else ""
                )
                try:
                    arcpy.management.Delete(in_path)
                except Exception:
                    pass
                os.replace(tmp_path, in_path)
            except Exception as ex_proj:
                # Reprojection failed — keep original (don’t record as a page failure)
                arcpy.AddWarning(f"[WARN] Reprojection failed for {os.path.basename(in_path)}: {ex_proj}")
            return in_path
        # --- end helpers ---

        # Export loop — EXACT base working call, but skip pages on failure
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

            # --- Core export (legacy 4-arg) ---
            try:
                # Base behavior that worked for you: (in_pdf, out_tif, pdf_password, pdf_page_number)
                arcpy.conversion.PDFToTIFF(
                    input_pdf_path,
                    output_path,
                    None,
                    page_1b
                )
                arcpy.AddMessage(f"[OK] Exported page {page_1b} → {output_filename}")

            except Exception as ex_export:
                # SKIP this page and record failure (do not raise)
                failures.append({
                    "pdf_path": input_pdf_path,
                    "page": page_1b,
                    "output": output_path,
                    "error": str(ex_export)
                })
                arcpy.AddWarning(f"[SKIP] Page {page_1b} failed: {ex_export}")
                continue

            # Reproject IN PLACE if/when needed (doesn't affect failure list)
            try:
                final_path = project_raster_in_place(output_path, target_sr)
            except Exception as ex_post:
                arcpy.AddWarning(f"[WARN] Post-process issue on page {page_1b}: {ex_post}")
                final_path = output_path

            # Add the one-and-only file to the map (same as base)
            try:
                if map_obj:
                    map_obj.addDataFromPath(final_path)
            except Exception as ex_add:
                arcpy.AddWarning(f"[WARN] Added TIFF but couldn't add to map (page {page_1b}): {ex_add}")

        # ---------- Write failures (if any) into the Project's Default GDB and add table to map ----------
        if failures:
            try:
                gdb = aprx.defaultGeodatabase
                if not gdb or not os.path.isdir(gdb):
                    # Fallback to project home if default GDB missing (rare)
                    gdb = aprx.homeFolder

                # Table name includes Project ID; overwrite if already exists
                safe_pid = self._sanitize_for_fs(project_id)
                tbl_name = f"PDF2TIFF_Failures_{safe_pid}"
                tbl_path = os.path.join(gdb, tbl_name)
                if arcpy.Exists(tbl_path):
                    arcpy.management.Delete(tbl_path)

                # Create table & fields
                arcpy.management.CreateTable(os.path.dirname(tbl_path), os.path.basename(tbl_path))
                arcpy.management.AddField(tbl_path, "Page", "LONG")
                arcpy.management.AddField(tbl_path, "PDF_Path", "TEXT", field_length=512)
                arcpy.management.AddField(tbl_path, "Output_Path", "TEXT", field_length=512)
                arcpy.management.AddField(tbl_path, "Error", "TEXT", field_length=1024)

                # Insert rows
                with arcpy.da.InsertCursor(tbl_path, ["Page", "PDF_Path", "Output_Path", "Error"]) as cur:
                    for f in failures:
                        cur.insertRow([f["page"], f["pdf_path"], f["output"], f["error"]])

                # Add the table to the current map so it's visible in Contents
                if map_obj:
                    try:
                        map_obj.addDataFromPath(tbl_path)
                    except Exception:
                        pass

                # Print a compact summary in the GP Messages pane
                pages_str = ", ".join(str(f["page"]) for f in failures)
                arcpy.AddWarning(f"[SUMMARY] Skipped pages ({len(failures)}): {pages_str}")
                arcpy.AddMessage(f"[INFO] Failure details table added to project GDB: {tbl_name}")

            except Exception as ex_tbl:
                arcpy.AddWarning(f"[WARN] Could not write failures table to project GDB: {ex_tbl}")
                pages_str = ", ".join(str(f["page"]) for f in failures)
                arcpy.AddWarning(f"[SUMMARY] Skipped pages ({len(failures)}): {pages_str}")
        else:
            arcpy.AddMessage("[SUMMARY] No pages were skipped.")

        del aprx
        return

    def postExecute(self, parameters):
        return
