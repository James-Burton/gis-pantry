################################################################################
##########################   NEW FN Reporting TOOL #############################
################################################################################
## Second Release
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# Name:    Automated First Nation Historical Stumpage Calculator
#
# Keywords:                      Union - Calculate - Report
#                                   version 2.4
#
# Author:      James Burton|Government of British Columbia
#              Geospatial Analyst: Skeena-Stikine District
#
# Created:     November-December 2020
#
# Introduction: A second iteration of the popular VnV tool, version 1.
# This tool provides a better report in less time.
# The script has been adapted to perform the averaging calculations previously
# employed while incorperating more raw data. This raw data can now include
# specific species removed from the Area of Interest. The raw data can also
# exceed 20 years of histroical statistics. When dealing with such old data,
# some items are often missing. A report of Missing Timber Mark Spatials is also
# created by iterating through the unique timber marks in the raw data and
# comparing them to the uniwue timber marks found within your input data. The
# input data is detailed further below. The script itself uses minimal ArcPy
# lines in an effort to reduce run time, which was successful. This tool has
# been successfully utilized a handful of times, and it is anticipated that it
# will be used yearly here-on-in.
#
# Outputs:     1 spatial of the unioned timber marks with total timber mark area
# and total sub-unit timber mark area. This is found in the tempGDB created by
# the script.
#              A spreadsheet (XLSX) for every year in the analysis that contains
# timber marks ONLY found in the FN boundary, categorized by attributes included
# in the FN boundary shape.
#              1 Spreadsheet (XLSX) of unmatched timber marks, represented by
# full records obtained from the raw tabular table.
#
# Dependancies: Python 2.7, Arc GDB 10.x software (ArcMap, ArcCatalog, ArcPro,
# QGIS to view output spatial.
#               Must have noted spatials and xlsx inputs.
#
# Copyright:   (c) bc.gov.ca, James Burton 2020
#-------------------------------------------------------------------------------
# Potential edits:
# Edit as necessary for your analysis. Try to keep shared edits consistent with
# needs of your colleagues (don't edit it to work with constrained data).
#-------------------------------------------------------------------------------
# Edit History:
#
# Date:December 2020
# Author: James Burton
# Modification Notes:
# 1. Adapted script to use pandas for math instead of arcpy calc fields.
# 2. Removed redundancies in spatial data creation
# 3. Pretty much a made a new script based on previous principles.
#
# Date:December 11/12, 2020
# Author:James Burton
# Modification Notes: 
# 1. Changed lambdas row function to handle 0 volume or value 
# 2. Changed year dataframe compilation to skip years with no harvest in FN bdy
# 3. Enabled script to join all harvest years in FN to on xlsx doc
# 3. Edited script to python3 in prep for R implementation
#-------------------------------------------------------------------------------
#
# Date:
# Author:
# Modification Notes:
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!BIG NOTE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# Spatial Layer creations! : a number of steps are taken in your GIS program of
# choice. There are three input data types for this program to operate. The tabular
# data can come in only as an XLSX in this version. To have diferent data types
# like CSV, please modify the necessary components. At time of writing there are
# two seperate functions that access the tabular data. Firstly, there is an
# openpyxl function that obtains the headers, which takes xlsx. Second there is
# some panda dataframe creations that are not set to handle csvs. You'll figure it
# out I am sure.

# Two spatial data sets are required to make this program run. The first data set is
# easy to obtain. Simply input a shapefile of the Nation you are working with. The
# program will take all the rows from the input spatial so you do not need to tell
# the program what column you want to harness. This is to ensure multiple attributes
# are retained in the event one may need House and Watershed names in the output, for
# example. This FN boundary layer will now be know as the AOI for area of interest.

# The script's input shapefile is a manual creation currently.
# Through a combined dataset, the areas associated with the Timber Mark and FN
# shape are calculated.

# Up to now, the process of spatial data (and tabular) compilation is as follows:

##1. WHSE_ADMIN_BOUNDARIES.FADM_TSA
##   Dissolve by TSA_NUMBER or TSA_NUMBER_DESCRIPTION
##   Using Select by Location, select the TSAs that intersect the FN shape.
##   Save this layer if you want as Selected_TSAs
##
##   This is done because a sub-report is generated. The sub-report contains the Timber Marks from the
##   initial report that do not have a spatial relation. This information is important to acknowledge as
##   as footnote identifying the Volume and Value that may or may not have been removed from the FN Nation.
##
##2. WHSE_FOREST_TENURE.FTEN_HARVEST_AUTH_POLY_SVW
##   From this layer, using Select by Location, select all records that intersect with the Selected_TSAs
##   shape.
##   Do not clear the selection.
##   Dissolve the layer based on TIMBER_MARK_PRIME
##   Open the attribute table, add a field called TM
##   Calculate Field, TM = !TIMBER_MARK_PRIME! (or alter to your choice of language)
##
##   This is done because the total Timber Mark area is needed for the average calculation. Clipping the data
##   to the selection would reduce the total value if the Timber Mark/ Licence straddled a TSA inside and
##   outside the area of interest.
##
##3. WHSE_FOREST_TENURE.FTEN_ROAD_SEGMENT_LINES_SVW
##   From this layer, using Select by Location, select all the records that intersect with the Selected_TSAs
##   shape.
##   Do not clear the selection.
##   Buffer the selection by 17.5 meters
##   Open the attribute table, add field called TM
##   Calculate Field, TM = !TIMBER_MARK!
##   Select by Attributes, TM is Null
##   Calculate Field again on selected Null attributes, TM = FOREST_FILE_ID
##
##   This is done because the Timber Marks associated with roads is represented in a linear format. For
##   area calculation, the buffer is obviously necessary. The Select by Location instead of a Clip ensures
##   we do not omit sections of the Timber Mark/ Licence.
##
##4. WHSE_FOREST_VEGETATION.RSLT_OPENING_SVW
##   From this layer, using Select by Location, select all the records that intersect with the Selected_TSAs.
##   I would export this selection as a stand-alone feature class. Examine how many Null Timber Marks there are.
##   If you feel confident you cannot match any Timber Marks from the Null values..
##   Create new field TM, calculate field TM = !TIMBER_MARK!
##   Dissolve on TM
##
##   This is done to satisfy historical stumpage reports. Some shapes may have been retired however their
##   silviculture obligations may remain and are spatially represented.
##
##5. WHSE_FOREST_TENURE.FTEN_CUT_BLOCK_POLY_SVW
##   From this layer, Select by Location all records that intersect the Selected_TSAs.
##   Do not clear your selection.
##   Dissolve the Timber Mark field
##   add new field TM, calculate field TM=!CUT_BLOCK_TIMBER_MARK!
##
##   This layer is used because I have discovered inconsistencies with the location of data storage.
##   This layer contains items that are not contained in layers above.
##
##6. Merge final output from step 2,3,4 and 5. Dissolve again, this time by TM.
##
##7. This final dissolved shape should be saved as a shape file and inside the code, set the HarvAuth file path
##   to this newly created shapefile.
##
##8. While inside the script, change the Housez variable to become the path to the First Nation shape you are
##   analyzing. This shape can as simple as the PIP boundary or as specific as a Watershed/Wilp/House/etc.
################################################################################
################################################################################
__file__ = sys.argv[0]
import os, sys, arcpy, shutil, xlsxwriter
from arcpy import env
import pandas as pd
from datetime import datetime
import time

#------------------------------------------------------------------------------#
##################### Set Date, Time; start timer ##############################

## starts program timer
t0= time.clock()
## Sets date for temporary folder creation
rightnow= datetime.now()
Today= rightnow.strftime('%Y_%m_%d')

## the container is set to the scripts file for now.
Container = (os.path.dirname(__file__))
TempContainer=os.path.join(Container+ '\\0800004Temp%s' % Today)

Test=1
## This layer will become the AOI
Housez = r'W:\FOR\RNI\DSS\General_User_Data\users\jamburto\!scripts\Gitanyow\Spatials\Gitanyow_wilp_sharepoint_July_2013.shp'
AOI='AOI'

## this is the location of the incoming scaled
ScaleData = r'W:\FOR\RNI\DSS\General_User_Data\users\jamburto\!scripts\Gitanyow\Reports\ScaleData2.xlsx'
##OurData=os.path.join(TempContainer,'GitScale.xlsx' )

HarvAuth = r'W:\FOR\RNI\DSS\General_User_Data\users\jamburto\!scripts\Gitanyow\Spatials\TSA56_TM_All_Dissolve.shp'
AOI_Harv = 'AOI_Harv'
TimberMarkList=[]
YrList =[]
YrListSum = []
sortedyr = []
rprthead = []
Git_HATM_List =[]
ALL_HATM_List=[]
MatchTMs=[]
NotMatchTMs=[]
House_TM=pd.DataFrame()
AOI_H_D = 'AOI_Harv_Diss' ## timber marks dissolved by timber marks selected to intersect the FN BNDY
data = pd.read_excel(ScaleData)
Onion = 'AOI_TM_D_Housez'
print(data.head())
#------------------------------------------------------------------------------#
##################### Set up workspace #################################

def MakeTempFolder():
    global TempGDB
    if os.path.exists(TempContainer):
        shutil.rmtree(TempContainer)
        time.sleep (1)
        os.makedirs(TempContainer)
    else:
        os.makedirs(TempContainer)
    TempGDB = arcpy.CreateFileGDB_management(TempContainer, 'TempGDB.gdb', '10.0')
    ##shutil.copy(ScaleData, OurData)

#------------------------------------------------------------------------------#
def HarvestTimberMarks():

    arcpy.env.workspace=str(TempGDB)
    arcpy.CopyFeatures_management(Housez, AOI)
    arcpy.MakeFeatureLayer_management(HarvAuth, 'Harv_lyr')
    arcpy.SelectLayerByLocation_management('Harv_lyr', 'intersect', AOI)
    arcpy.CopyFeatures_management('Harv_lyr', AOI_Harv)
    arcpy.Dissolve_management(AOI_Harv, AOI_H_D, "TM")

    with arcpy.da.SearchCursor(AOI_H_D, 'TM') as cursor:
            for row in cursor:
                field_value=str(cursor)
                field_val_temp = field_value.strip("(u'")
                field_value= field_val_temp.strip("'),")
                Git_HATM_List.append(field_value)

    print(len(Git_HATM_List))
    with arcpy.da.SearchCursor(HarvAuth, 'TM') as cursor:
            for row in cursor:
                field_value=str(cursor)
                field_val_temp = field_value.strip("(u'")
                field_value= field_val_temp.strip("'),")
                ALL_HATM_List.append(field_value)
    tm13=data

    tm14=tm13.loc[~tm13["Timber_Mark"].isin(ALL_HATM_List)]
    filename = str('ALLMissingMarks.xlsx')
    filefile = os.path.join(TempContainer,filename)
    tm14.to_excel(filefile, engine='xlsxwriter')


def calculate_vol(row):
    try:
        return row['Total_Volume']/row['TM_TOT_AREA']
    except:
        return 0

def calculate_val(row):
    try:
        return row['Total_Value']/row['TM_TOT_AREA']
    except:
        return 0

def calculate_vol2(row):
    try:
        return row['TM_Vol_Av_HA']/row['WILP_TM_AREA']
    except:
        return 0

def calculate_val2(row):
    try:
        return row['TM_Val_Av_HA']/row['WILP_TM_AREA']
    except:
        return 0

#------------------------------------------------------------------------------#
## this opens the original data and exports the year's timber marks that exist spatially to it's own xlsx
def YearXLSX():
    newdf2=pd.DataFrame()
    YearList = data.Scaled_Year.unique()

    for x in YearList:
        scayea = pd.DataFrame()
        HA=pd.DataFrame()
        ha13=pd.DataFrame()
        scayea = data[data.Scaled_Year == x]

        for y in Git_HATM_List:

            HA=scayea[scayea.Timber_Mark == y]

            if not HA.empty:
                if not ha13.empty:
                    ha13 = ha13.append(HA,ignore_index=True)
                else:
                    ha13=HA

        if not ha13.empty:
            thisyeardata=ha13
            print(thisyeardata.head())
            print('hi rex you were here recently')
            newdf = thisyeardata.join(House_TM.set_index('WILP_TM'), on='Timber_Mark')

##            print newdf
##            try:
            newdf['TM_Vol_Av_HA']=newdf.apply(lambda row: 0 if row.Total_Volume==0 or row.TM_TOT_AREA ==0 else row.Total_Volume/row.TM_TOT_AREA, axis=1)
##            except:
##                newdf['TM_Vol_Av_HA']=0
##            try:
            newdf['TM_Val_Av_HA']=newdf.apply(lambda row: 0 if row.Total_Value==0 or row.TM_TOT_AREA ==0 else row.Total_Value/row.TM_TOT_AREA, axis=1)
##            except:
##                newdf['TM_Val_Av_HA']=0
            newdf['TM_Hz_Vol']=newdf.apply(lambda row: 0 if row.TM_Vol_Av_HA==0 else row.TM_Vol_Av_HA*row.WILP_TM_AREA, axis=1)
            newdf['TM_Hz_Val']=newdf.apply(lambda row: 0 if row.TM_Val_Av_HA==0 else row.TM_Val_Av_HA*row.WILP_TM_AREA, axis=1)

            newdf2=newdf2.append(newdf,ignore_index=True)


    print(newdf2.head())
    filename ='rexington13.xlsx'## str('scaleyear_%s' %x+'.xlsx')
    filefile = os.path.join(TempContainer,filename)
    newdf2.to_excel(filefile, engine='xlsxwriter')

    arcpy.env.workspace=str(TempGDB)
    arcpy.ExcelToTable_conversion(filefile,filename[:-5])
#------------------------------------------------------------------------------#



def trial():
    global House_TM
    Onion = 'AOI_TM_D_Housez'
    arcpy.env.workspace=str(TempGDB)
    arcpy.AddField_management(AOI_H_D,"TM_TOT_HA","FLOAT","","","","","NULLABLE")
    arcpy.CalculateField_management(AOI_H_D, "TM_TOT_HA","!SHAPE.AREA@HECTARES!","PYTHON_9.3")
    arcpy.Union_analysis([AOI_H_D,Housez],Onion)################################################
    field_names = [f.name for f in arcpy.ListFields(Onion)]
    print(field_names)
    for fname in field_names:
        if 'FID' in fname:
            field_value=fname
            fieldname_val_temp = field_value.strip("u'")
            fieldname_value= fieldname_val_temp.strip("',")

            with arcpy.da.UpdateCursor(Onion, fieldname_value) as cursor:
                for row in cursor:
                    if row[0] == -1:
                        cursor.deleteRow()

    arcpy.Dissolve_management(Onion,'Onion',['TM','WILPNAMES',"PDEEK",'TM_TOT_HA'])
    arcpy.AddField_management('Onion',"TM_Hz_TOT_HA","FLOAT","","","","","NULLABLE")
    arcpy.CalculateField_management('Onion', "TM_Hz_TOT_HA","!SHAPE.AREA@HECTARES!","PYTHON_9.3")

############THIS Section will need to be edited for your report.
    col1_col=[]
    col2_col=[]
    col3_col=[]
    col4_col=[]
    col1 = "WILP_TM"
    col2 ="WILP_NAME"
    col3= "TM_TOT_AREA"
    col4="WILP_TM_AREA"

    l=0

    with arcpy.da.SearchCursor('Onion', ["TM","WILPNAMES","TM_TOT_HA","TM_Hz_TOT_HA"]) as cursor:
        for row in cursor:
            col1_col.append(cursor[0])
            col2_col.append(cursor[1])
            col3_col.append(cursor[2])
            col4_col.append(cursor[3])

    House_TM = pd.DataFrame(list(zip(col1_col,col2_col, col3_col,col4_col)),columns=[col1,col2,col3,col4])

MakeTempFolder()
HarvestTimberMarks()

trial()
YearXLSX()
print('DONE :)')
t1 = time.clock() - t0
print('Time elapsed: ', (t1 - t0)) # CPU seconds elapsed (floating point)