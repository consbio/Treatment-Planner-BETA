########################################################################################################################
# File name: Treatment Planner.pyt (Python Toolbox Tool)
# Author: Mike Gough
# Date created: 05/25/2019
# Date last modified: 10/19/2019
# Python Version: 2.7
# Script Version: 4.0.1
# Updates: "Allows for multiple management actions. Ported to Python Toolbox"
# Description: A tool for landscape/forest management and scenario planning.
# This tool takes user input (stand layer, fvs database, and a set of treatments), runs FVS and optionally sends 
# the FVS output to an EEMS Fuzzy Logic model.
########################################################################################################################

import arcpy
import os
import subprocess

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Treatment Planner"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [Run_Simulation]


class Run_Simulation(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Run Simulation"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        param0 = arcpy.Parameter(
            displayName='Forest Stands Layer',
            name='forest_stands_layer',
            datatype='GPFeatureLayer',
        )
        
        param1 = arcpy.Parameter(
            displayName='FVS Database',
            name='fvs_database',
            datatype='DEFile',
        )
        #param1.value = r'F:\Projects2\USFS-052_PSW_Forest_Treatment_Planner_2019_mike_gough\Tasks\FVS_Arc_4_Box_model\Tools\TreatmentPlanner_dev\Data\Inputs\FVS_Databases\LakeGillette.accdb'
        param1.value = r'F:\Projects2\USFS-052_PSW_Forest_Treatment_Planner_2019_mike_gough\Tasks\FVS_Arc_4_Box_model\Tools\Treatment_Planner_BETA\Data\Inputs\FVS_Databases\ShaverLake.accdb'

        param2 = arcpy.Parameter(
            displayName='Number of Years to Simulate',
            name='num_years',
            datatype='Long'
        )
        
        param3 = arcpy.Parameter(
            displayName='Output Fields to Join',
            name='output_fields',
            datatype='String',
            parameterType='Optional',
            multiValue=True,
        )
        param3.filter.list = ['Stand Age', 'Trees per acre (Tpa)', 'Basal area per acre (BA)', 'Stand density index (Sdi)', 'Crown competition factor (CCF)', 'Average dominant height (TopHt)', 'Quadratic mean DBH (QMD)', 'Forest cover type (ForTyp)', 'Stand size class (SizeCls)', 'Merchantable cuft volume (MCuFt)']
        
        param4 = arcpy.Parameter(
            displayName='Management Action(s)',
            name='management_actions',
            datatype='GPValueTable',
            parameterType='Optional',
            direction='Input',
            category="Management Actions",
        )
        param4.columns = [['GPString', 'Action'], ['GPString', 'Year'], ['GPString', 'Advanced Settings (See Help)']]
        param4.filters[0].type = 'ValueList'
        param4.filters[0].list = ['Thin from Above', 'Thin from Below', 'Clearcut', 'Prescribed Burn']
        
        param5 = arcpy.Parameter(
            displayName='Run simulation on non-selected Stands',
            name='no_selected',
            datatype='GPBoolean',
            parameterType='Optional',
            category="Management Actions",
        )
        param5.value = "True"
        
        param6 = arcpy.Parameter(
            displayName='EEMS Model',
            name='eems_model',
            datatype='String',
            parameterType='Optional',
            category="Post-Processing Options"
        )
        param6.filter.list = ["Severe Fire Risk", "Thinning Assessment"]

        params = [param0, param1, param2, param3, param4, param5, param6]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        
        # Check the fields needed to run the selected EEMS model.
        if parameters[6].value == "Thinning Assessment":
            required_inputs_list = ["Stand Age", "Trees per acre (Tpa)", "Basal area per acre (BA)", "Stand density index (Sdi)",
                                    "Merchantable cuft volume (MCuFt)"]
        elif parameters[6].value == "Severe Fire Risk":
            required_inputs_list = ["Stand Age", "Trees per acre (Tpa)", "Basal area per acre (BA)", "Stand density index (Sdi)",
                                    "Crown competition factor (CCF)"]
        if parameters[6].value:
            if parameters[3].value:
                selected_value_list = str(parameters[3].value).strip().replace("'", "").split(";")
                unique_list = list(set().union(selected_value_list, required_inputs_list))
                parameters[3].value = unique_list
            else:
                parameters[3].value = required_inputs_list

        # Set the Year, and the Advanced Settings based on the selected action.
        if not parameters[4].hasBeenValidated and parameters[4].value:
            default_advanced_settings = {
                "Clearcut": "0,0,999",
                "Thin from Above" :  "60,1,0,999,0,999",
                "Thin from Below" :  "60,1,0,999,0,999",
                "Prescribed Burn" :  "8,2,70,1,70,1",
            }
            
            param_list = parameters[4].values
            action = param_list[-1][0]
            current_year = param_list[-1][1]
            current_advanced_settings = param_list[-1][2]

            if current_year != '':
                year = current_year
            else: 
                year = 0
            
            if current_advanced_settings != '':
                advanced_settings = current_advanced_settings
            else:
                advanced_settings = default_advanced_settings[action]
                
            update = [action,year,advanced_settings]
            
            del param_list[-1]
            param_list.append(update)
            parameters[4].values = param_list 

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        
        arcpy.env.overwriteOutput = True
        fields_to_join = []
        
        # Maps the user selected variables to FVS database output fieldnames.
        variable_code_lookup = {
            "Stand Age": "Age",
            "Trees per acre (Tpa)": "Tpa",
            "Basal area per acre (BA)": "BA",
            "Stand density index (Sdi)": "Sdi",
            "Crown competition factor (CCF)": "CCF",
            "Average dominant height (TopHt)": "TopHt",
            "Quadratic mean DBH (QMD)": "QMD",
            "Forest cover type (ForTyp)": "ForTyp",
            "Stand size class (SizeCls)": "SizeCls",
            "Merchantable cuft volume (MCuFt)": "MCuFt"
        }

        # Gather input parameters
        input_stands = parameters[0].valueAsText
        fvs_database = parameters[1].valueAsText
        num_years = int(parameters[2].valueAsText)
        if parameters[3].valueAsText:
            # Build a list of user selected variables.
            variables = parameters[3].valueAsText.split(";")
            if variables[0] != "":
                for variable in variables:
                    variable_code = variable_code_lookup[variable.replace("'", "")]
                    fields_to_join.append(variable_code)
        management_actions = parameters[4].values 
        simulate_all_stands = parameters[5].valueAsText
        eems_model = parameters[6].valueAsText

        # Ensures EEMS command file can be overwritten.
        arcpy.env.overwriteOutput = True

        # Path to FVS
        fvs_executable = "C:/FVSbin/FVSie.exe"

        # Set the workspace to the ArcFVS directory. Based on the location of this python script.
        workspace = os.path.dirname(os.path.realpath(__file__))

        # Output directory for FVS files.
        output_dir = workspace + "\\Data\\Outputs"
        os.chdir(output_dir)

        arcpy.AddMessage("\nInitializing...")

        arcpy.AddMessage("Deleting Fields...")
        current_fields = arcpy.ListFields(input_stands)
        for field in current_fields:
            if field.name.split("_")[0] in variable_code_lookup.values():
                arcpy.DeleteField_management(input_stands, field.name)


        # FVS Files
        rsp_file = "arc_fvs.rsp"
        keyword_file = "arc_fvs.key"
        tre_file = "arc_fvs.tre"
        out_file = "arc_fvs.out"
        trl_file = "arc_fvs.trl"
        sum_file = "arc_fvs.sum"
        chp_file = "arc_fvs.chp"

        arcpy.AddMessage("Creating Output Database...")
        output_database_name = "arc_fvs_out.mdb"
        output_database_path = output_dir + os.sep + output_database_name
        output_database_table = output_database_path + os.sep + "FVS_Summary"
        if arcpy.Exists(output_database_path):
            arcpy.Delete_management(output_database_path)
        arcpy.CreatePersonalGDB_management(output_dir, output_database_name)

        # Remove files from the previous run.
        if os.path.isfile(keyword_file):
            os.remove(keyword_file)

        if os.path.isfile(rsp_file):
            os.remove(rsp_file)

        desc = arcpy.Describe(input_stands)
        input_stands_path = desc.CatalogPath

        # Get the stand ids for the selected features.
        selected_stand_ids = []
        sc = arcpy.SearchCursor(input_stands)
        try:
            for row in sc:
                selected_stand_ids.append(row.getValue("STAND_ID"))
            del sc, row
        except arcpy.ExecuteError:
            arcpy.AddError("\nERROR: The Forest Stands Layer needs a 'STAND_ID' field.")

        all_stand_ids = []
        if simulate_all_stands == 'true':
            # Get all the stand ids in the feature class.
            sc = arcpy.SearchCursor(input_stands_path)
            for row in sc:
                all_stand_ids.append(row.getValue("STAND_ID"))
            del sc, row
        else:
            all_stand_ids = selected_stand_ids

        def create_rsp_file():
            ''' Creates the rsp file which just has the names of the other files in it (most importantly the keyword file). '''

            with open(rsp_file, "a") as f:
                f.write(keyword_file + "\n")
                f.write(tre_file + "\n")
                f.write(out_file + "\n")
                f.write(trl_file + "\n")
                f.write(sum_file + "\n")
                f.write(chp_file + "\n")

            return rsp_file

        def create_keyword_file(fvs_database, input_stand_id, num_years, management_action_code=" ",
                                output_database=""):
            ''' Creates the keyword file for the FVS run. '''
            keyword_file_content = "Comment\n\
FVS Keyword file generated by ArcFVS\n\
End\n\
StdIdent\n\
{0} Stand {0}\n\
Screen\n\
StandCN\n\
{0}\n\
InvYear         0\n\
TimeInt                      1\n\
NumCycle          {1}\n\
DataBase \n\
DSNOut\n\
{4}\n\
Summary\n\
End\n\
Database\n\
DSNIn\n\
{2}\n\
StandSQL\n\
SELECT * \n\
FROM FVS_StandInit \n\
WHERE Stand_ID = '%StandID%' \n\
EndSQL\n\
TreeSQL\n\
SELECT * \n\
FROM FVS_TreeInit \n\
WHERE Stand_ID = '%StandID%' \n\
EndSQL\n\
END\n\
{3}\n\
SPLabel\n\
   All, &\n\
   Lake_Gillette, &\n\
   Wildland-Urban-Interface\n\
Process\n\
STOP".format(input_stand_id, num_years, fvs_database, management_action_code, output_database)

            with open(keyword_file, "w") as f:
                f.write(keyword_file_content)

            return keyword_file_content

        def create_management_action_code(management_actions):

            management_action_code = ""
            for management_list in management_actions:
                try: 
                    int(management_list[1])
                except:
                    arcpy.AddError("\nERROR: Invalid Management Action Year")
                management_action = management_list[0]
                management_year = str(int(management_list[1]) + 1)
                advanced_settings = management_list[2].split(",")
                ''' Creates the management action code to be injected into the keywords file.'''

                if management_action == "Thin from Above":
                    management_action_code += "thinABA         {}      {}        {}.        {}.      {}.       {}.      {}.\n".format(management_year, *advanced_settings)

                elif management_action == "Thin from Below":
                    management_action_code += "thinBBA         {}      {}        {}.        {}.      {}.       {}.      {}.\n".format(management_year, *advanced_settings)

                elif management_action == "Prescribed Burn":
                    management_action_code += "FMIn\nSimFire           {}        {}.         {}       {}.         {}       {}.         {}\nEnd\n".format(management_year, *advanced_settings)

                elif management_action == "Clearcut":
                    management_action_code += "ThinDBH         {}      {}        {}.       1.0      0.0       0.0      0.0\nThinDBH         {}      {}      999.0      1.0      0.0       {}       0.0".format(management_year, advanced_settings[0], advanced_settings[2], management_year, advanced_settings[2], advanced_settings[1])

            return management_action_code

        # For each stand_id, generate a .key file and run FVS.
        arcpy.AddMessage("Starting Simulation...")
        for stand_id in all_stand_ids:

            arcpy.AddMessage("\nSTAND ID :" + str(stand_id))

            # Create the keyword file.
            arcpy.AddMessage("Creating FVS Keyword File...")
            # If a management action is defined and the stand is selected, add the management action code to the keyword file.
            if management_actions and stand_id in selected_stand_ids:
                management_action_code = create_management_action_code(management_actions)
                create_keyword_file(fvs_database, stand_id, num_years, management_action_code, output_database_path)

            # Otherwise, create the keyword file without the management action code.
            else:
                create_keyword_file(fvs_database, stand_id, num_years, " ", output_database_path)

            # Create the rsp file.
            rsp_file = create_rsp_file()

            # Run FVS by feeding it the rsp file which just contains a list of filenames to process.
            arcpy.AddMessage("Running FVS...")
            fvs_command = fvs_executable + " < " + rsp_file
            fvs_run = subprocess.Popen(fvs_command, shell=True, stdout=subprocess.PIPE)

            # Get the output summary statistics that FVS sends to standard out.
            output, err = fvs_run.communicate()
            print output

            # Parse the summary statistics (stdout from FVS), and add to the ArcGIS Script Tool window.
            results_dict = {}
            for line in output.splitlines():
                if "ENTER" not in line:
                    arcpy.AddMessage(line)

            # Option to write all output from FVS in the .out file.
            def write_all_lines():
                with open(out_file) as f:
                    data = file.readlines(f)
                    for line in data:
                        arcpy.AddMessage(line)

        arcpy.RegisterWithGeodatabase_management(output_database_table)
        sc = arcpy.SearchCursor(output_database_table)
        years = []
        for row in sc:
            years.append(row.getValue('Year'))
        del sc, row

        first_year = min(years)
        last_year = max(years)
        
        if fields_to_join:

            arcpy.AddMessage("\nJoining Results...")
            arcpy.JoinField_management(input_stands_path, "Stand_ID", output_database_table, "StandID", fields_to_join)

            for field in fields_to_join:
                arcpy.AlterField_management(input_stands_path, field, field + "_" + str(first_year))

            query = "[Year] = " + str(last_year)

            arcpy.MakeTableView_management(output_database_table, 'last_year_table', query)
            arcpy.JoinField_management(input_stands_path, "Stand_ID", 'last_year_table', "StandID", fields_to_join)

            arcpy.Delete_management('last_year_table')

            for field in fields_to_join:
                arcpy.AlterField_management(input_stands_path, field, field + "_" + str(last_year))

        # If error, check to make sure parameters in the model are in the right order.
        # Make sure alias on the Model matches the name in tbx. call below.

        ################################################# EEMS MODELS ##################################################

        # Note: Problems with toolbox imports may be caused by other toolboxes with the same name added to the ArcToolbox window.
        # Toolbox import form is tbx.<toolname>_<toolboxalias>

        if eems_model:
            mxd = arcpy.mapping.MapDocument("CURRENT")
            df = arcpy.mapping.ListDataFrames(mxd)[0]

            # AddToolbox only works on system toolboxes or toolboxes that have already been added to ArcToolbox in the mxd.
            #tbx = arcpy.AddToolbox(workspace + "//EEMS/EEMS_Models.tbx")
            tbx = arcpy.ImportToolbox(workspace + "//EEMS/EEMS_Models.tbx")
            eems_command_file_dir = workspace + "//Data//Intermediate//EEMS_Command_Files//"
            eems_lyr_dir = workspace + "//Data//Outputs//lyr//"

            arcpy.AddMessage("\nRunning EEMS " + eems_model + " Model...")
            eems_command_file = eems_command_file_dir + eems_model.replace(" ", "_") + ".mpt"

            if eems_model == "Thinning Assessment":
                arcpy.ThinningAssessment_EEMSModels(last_year, input_stands_path, eems_command_file)
                eems_lyr_file = eems_lyr_dir + "Thinning_Assessment.lyr"

            if eems_model == "Severe Fire Risk":
                arcpy.SevereFireRisk_EEMSModels(last_year, input_stands_path, eems_command_file)
                eems_lyr_file = eems_lyr_dir + "Severe_Fire_Risk.lyr"

            if arcpy.Exists(eems_lyr_file):
                layer = arcpy.mapping.Layer(eems_lyr_file)
                arcpy.mapping.AddLayer(df, layer, "TOP")

        def add_veg_type_fields():
            ''' works, but appends to end of table. Not pretty. '''
            arcpy.AddField_management(input_stands_path, "ForType_" + str(first_year), "STRING")
            arcpy.AddField_management(input_stands_path, "ForType_" + str(last_year), "STRING")

            uc = arcpy.UpdateCursor(input_stands_path)
            for row in uc:
                forest_type_code = row.getValue("ForTyp_" + str(first_year))
                forest_type = forest_type_codes[forest_type_code]
                row.setValue("ForType_" + str(first_year), forest_type)
                forest_type_code = row.getValue("ForTyp_" + str(last_year))
                forest_type = forest_type_codes[forest_type_code]
                row.setValue("ForType_" + str(last_year), forest_type)
                uc.updateRow(row)

            del uc, row

        # add_veg_type_fields()
        return

