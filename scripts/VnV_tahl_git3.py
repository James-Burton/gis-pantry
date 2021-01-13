
import sys, arcpy, shutil, xlsxwriter
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