# -*- coding: utf-8 -*-

import arcpy

class Toolbox:
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
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
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter. This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        # Get the input PDF file path from the parameters
        input_pdf_path = parameters[4].valueAsText
        # Get output folder, page range, project ID, and county
        output_folder = parameters[5].valueAsText
        start_page = int(parameters[2].value) - 1  # zero-based index
        end_page = int(parameters[3].value) - 1
        project_id = parameters[1].valueAsText
        county = parameters[0].valueAsText

        aprx = arcpy.mp.ArcGISProject("CURRENT")
        map_obj = aprx.listMaps()[0] 

        for page_num in range(start_page, end_page + 1):
            pagestring = str(page_num + 1)
            if page_num <= 10:
                pagestring = "0" + str(page_num + 1)
            output_filename = f"{project_id}_p{pagestring}_{county}.tif"
            output_path = arcpy.os.path.join(output_folder, output_filename)
            arcpy.conversion.PDFToTIFF(input_pdf_path, output_path)
            map_obj.addDataFromPath(output_path)

        del aprx
        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return