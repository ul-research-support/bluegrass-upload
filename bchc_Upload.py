# -*- coding: utf-8 -*-
import pandas as pd
#import pycurl, cStringIO
#import certifi         # lines 3 & 4 are necessary for utilizing REDCap API
pd.set_option('display.height', 1500)
pd.set_option('display.max_rows', 1500)
pd.set_option('display.max_columns', 1500)
pd.set_option('display.width', 150)

"""
 :synopsis: Program to parse through excel-recorded refugee data for upload to REDCap. 
            Combines Bluegrass ARHCA worksheets into a single dataframe consisting of
            REDcap-relevant data, reformatted acc. to REDCap standards. 
            **Upload via REDCap API or manual .csv import**
            
  :notes:   Check the excel dataset for duplicate A#s being assigned to different patients. 
            This occurs from time to time when, for example, two patients are meant to be
            represented by A#s within one digit of each other (such as Axxx-xxx-056 and -057), 
            yet are erroneously assigned the same exact number. In this instance, you should 
            check if they have an existing record on REDCap to potentially verify their 
            correct A#. Otherwise, contact the Refugee Health Program Coordinator. 
"""
#%% - Read excel file into the program 
dr = 'mmddyyyy'  # change this to reflect your upload directory, my standard is date of upload
input_file_date = 'm-yyyy' # change this to reflect your input file's month/year
input_file = 'ARHCA_'+input_file_date
XL_FILE = pd.ExcelFile('C:\\Users\\japese01\\My Documents\\RefugeeHealth\\uploads\\uploads\\'+dr+'\\'+input_file+'.xls')

#%% - Load User-Defined Fields tab to access Alien Number
DF = XL_FILE.parse(sheetname=2)

#%% - Reindex the User Defined Fields tab (eliminate rows where Patient # is null)
DF = DF.set_index('Patient #', drop=False) 
DF = DF.loc[DF.index.to_series().dropna()]

#%% - Reformat the alien_no so it matches our storage format, drop unecessary fields
DF['Value'] = DF['Value'].astype(str)
DF['Value'] = DF['Value'].str.zfill(9)                                                  #fill with leading zeros if < 9 chars
d = [str(row)[:3] + '-' + str(row)[3:6] + '-' + str(row)[6:] for row in DF['Value']]    #insert dashes into alien_no fields
DF['alien_no'] = ['A'+str(row) for row in d]                                            #prefix each alien_no with 'A' 
DF.drop(['Value'], inplace=True, axis=1) 
DF.drop(['Field Name'], inplace=True, axis=1)

#%% - Load the Patient Demo tab to access patient data, drop unecessary fields, rename columns to match our format
DF_A = XL_FILE.parse(sheetname=0)
DF_A.rename(columns={'Patient Name': 'name', 'Date of Birth': 'date_of_birth', \
'Gender': 'gender', 'Marriage Status': 'marriage_status', 'Insurance': 'health_insurance', \
'Resettlement Agency': 'resettlement_agency', 'Zip Code': 'zip_code'}, inplace=True)

DF_A.clinic = 'bchc'
DF_A.zip_code = DF_A.zip_code.astype(str)
DF_A.zip_code = DF_A.zip_code.str[:5]
DF_A.drop(['Age'], inplace=True, axis=1)     # autocalculated field; unnecessary
DF_A['Patient #'].drop_duplicates()
DF_A = DF_A.set_index('Patient #', drop=True)

#%% - Join the dataframes on the Patient Number, remove null/duplicate entries
RESULT = DF.join(DF_A, on=DF['Patient #'], how='outer')
RESULT.drop('Patient #', axis=1, inplace=True)
RESULT = RESULT[pd.notnull(RESULT['alien_no'])]
RESULT.drop_duplicates('alien_no', inplace=True)

#%% - Convert demographic values into REDcap format
# - Tip: View the data dictionary 'Codebook' in REDcap to see correct variable names and field attributes.
RESULT.health_insurance.loc[RESULT.health_insurance == 'WELLCARE OF KENTUCKY'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'WELLCARE MEDICAID OF KENTUCKY'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'PASSPORT HEALTH PLAN'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'HUMANA CARESOURCE KY MEDICAID'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'HUMANA CARESOURCE KY MEDICAID '] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'COVENTRY CARES OF KY'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'AETNA BETTER HEALTH OF KENTUCKY'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'ANTHEM KENTUCKY MEDICAID'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'ANTHEM BCBS MEDICAID'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'BEACON HEALTH'] = 2
RESULT.health_insurance.loc[RESULT.health_insurance == 'HUMANA CARESOURCE MEDICARE ADVANTAGE'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'MEDICAID'] = '2'
RESULT.health_insurance.loc[RESULT.health_insurance == 'BLUE CROSS BLUE SHIELD'] = '2'

RESULT.gender.loc[RESULT.gender == 'M'] = '1'
RESULT.gender.loc[RESULT.gender == 'F'] = '2'

RESULT.marriage_status.loc[RESULT.marriage_status == 'MARRIED'] = '1'
RESULT.marriage_status.loc[RESULT.marriage_status == 'DIVORCED'] = '2'
RESULT.marriage_status.loc[RESULT.marriage_status == 'WIDOWED'] = '3'
RESULT.marriage_status.loc[RESULT.marriage_status == 'SEPARATED'] = '4'
RESULT.marriage_status.loc[RESULT.marriage_status == 'SINGLE'] = '5'
RESULT.marriage_status.loc[RESULT.marriage_status == 'SINGLE LIVING WITH PARTNER'] = '6'
RESULT.marriage_status.loc[RESULT.marriage_status == 'UNKNOWN'] = ''

if 'resettlement_agency' in RESULT:
    RESULT.resettlement_agency.loc[RESULT.resettlement_agency == 'KRM'] = '3'
    
#%% - Load the Vitals tab, remove duplicate indices, rename columns, split blood pressure column into systolic & diastolic 
DF_E = XL_FILE.parse(sheetname=6)
DF_E.drop(['Date'], inplace=True,axis=1)
DF_VITALS = DF_E.groupby(DF_E['Patient #']).first()

DF_VITALS.rename(columns={'Height': 'vsd1_height', 'Weight': 'vsd1_weight'}, inplace=True)
DF_VITALS.vsd1_sys_bp, DF_VITALS.vsd1_dia_bp = zip(*DF_VITALS.BP.map(lambda x: x.split('/')))
DF_VITALS.drop(['BP'], inplace=True, axis=1)
DF_VITALS.drop(['BMI'], inplace=True, axis=1)     # autocalculated field; unnecessary

#%% - height and weight conversion
DF_VITALS.vsd1_height = DF_VITALS.vsd1_height * 2.54
DF_VITALS.vsd1_weight = DF_VITALS.vsd1_weight * 0.45

#%% - Load the Medcin tab, drop unecessary fields
DF_C = XL_FILE.parse(sheetname=3)
DF_C.drop(['Enc Date'], inplace=True, axis=1)
DF_C.drop(['Medcin Id'], inplace=True, axis=1)
DF_C.drop(['Value'], inplace=True, axis=1)
DF_C.drop(['Onset Date'], inplace=True, axis=1)

# - This section and others like it will search for cells in the designated ['column'] that contain certain ('strings')
# - It will effectively filter out irrelevant data that is not to be stored in REDcap
DF_C = DF_C[DF_C['Medcin Description'].str.contains('STRONGYLOIDIASIS') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Strongyloidiasis') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('summary') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Anthelmintics') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Ivermectin') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Praziquantel') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('screened') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Schistosomiasis') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Immigration') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('method') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Presumptive') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('under') == False]
DF_C = DF_C.drop_duplicates()

#%% - Create secondary medcin dataframes to isolate & handle some misplaced demographic data (departure/origin, arrival date, language)
DF_MED1 = DF_C[DF_C['Medcin Description'].str.contains('Country') == True]
DF_MED2 = DF_C[DF_C['Medcin Description'].str.contains('Arrival') == True]
DF_MED1 = DF_MED1.pivot_table(index='Patient #', columns='Medcin Description', values='Note', aggfunc='first')
DF_MED2 = DF_MED2.pivot_table(index='Patient #', columns='Medcin Description', values='Note', aggfunc='first')
DF_MED1.rename(columns={'Country of Departure': 'cntry_dept', 'Country of Origin': 'cntry_origin'}, inplace=True)
DF_MED2.rename(columns={'Date of U.S. Arrival': 'us_arrival_date'}, inplace=True)
RESULT = RESULT.join(DF_MED1, how='outer')
RESULT = RESULT.join(DF_MED2, how='outer')


if (DF_C['Medcin Description'] == 'preferred language').any(): #conditional due to "preferred language" appearing inconsistently in excel forms
    DF_MED3 = DF_C[DF_C['Medcin Description'].str.contains('language') == True]
    DF_MED4 = DF_C[DF_C['Medcin Description'].str.contains('language') == True]
    DF_MED3 = DF_MED3.pivot('Patient #', 'Medcin Description', 'Note')
    DF_MED4 = DF_MED4.pivot('Patient #', 'Medcin Description', 'Note')
    DF_MED3.rename(columns={'preferred language': 'preferred_language'}, inplace=True)
    DF_MED4.rename(columns={'preferred language': 'prefered_language_other'}, inplace=True)
    DF_MED3['preferred_language'].loc[DF_MED3['preferred_language'] != 'Spanish'] = 0
    DF_MED3['preferred_language'].loc[DF_MED3['preferred_language'] == 'Spanish'] = 1
    DF_MED3.rename(columns={'preferred language': 'preferred_language'}, inplace=True)
    DF_MED4.rename(columns={'preferred language': 'prefered_language_other'}, inplace=True)
    RESULT = RESULT.join(DF_MED3, how='outer')
    RESULT = RESULT.join(DF_MED4, how='outer')

#%% - Drop demographic fields from parent dataframe after they have been isolated
DF_C.drop(['Note'], inplace=True,axis=1)
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Country') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('Arrival') == False]
DF_C = DF_C[DF_C['Medcin Description'].str.contains('language') == False]

#%% - Create Medcin dataframe
DF_MEDCIN = DF_C.pivot_table(index='Patient #', columns='Medcin Description', values='Result', aggfunc=sum)

#%% - Rename Medcin Columns to correct REDcap format
DF_MEDCIN.rename(columns={'Betel Nut use': 'nw_betel_nut', 'Blood Transfusion (___ ml)': 'nw_bld_tx', \
'Class B TB Status': 'ovs_class_btb', 'Currently breastfeeding': 'nw_breastfeeding', \
'Overseas medical records indicate a diagnosis of mental illness': 'ovs_mntl_hlth', \
'Patient has experienced imprisonment - torture or violence.  Effect on patient:': 'seh_exprcd_torture_yn', \
'Patient has witnessed someone experiencing torture or violence.': 'seh_wtnss_torture_yn', \
'Pre-departure treatment for Malaria': 'ovs_malaria', 'Pre-departure treatment given for intestinal parasites': 'ovs_intstnl_parasites', \
'Regular medications - vitamins - herbs or traditional medications used.': 'seh_rglr_meds_vit_yn', \
'Secondary Migrant': 'secondary_migrant', 'Street drugs': 'nw_strt_drugs_yn', 'Tattoo': 'nw_tattoo', \
'There is a faith tradition / religion that the patient practices.': 'seh_faith_yn', 'Use of injection drugs ever': 'nw_injctd_drgs', \
'alcohol use': 'nw_ethoh_yn', 'chewing nicotine-containing substances': 'nw_chews_yn', 'current smoker': 'nw_smokes_yn', \
'patient thinks she may be pregnant': 'nw_pregnant', 'sexually active': 'nw_sexually_act'}, inplace=True)

DF_MEDCIN = DF_MEDCIN.drop_duplicates() 

#%% - Convert Medcin values into REDcap's format  
if 'nw_betel_nut' in DF_MEDCIN:
    DF_MEDCIN.nw_betel_nut.loc[DF_MEDCIN.nw_betel_nut == 'N'] = 0
    DF_MEDCIN.nw_betel_nut.loc[DF_MEDCIN.nw_betel_nut == 'NN'] = 0
    DF_MEDCIN.nw_betel_nut.loc[DF_MEDCIN.nw_betel_nut == 'Y'] = 1
    DF_MEDCIN.nw_betel_nut.loc[DF_MEDCIN.nw_betel_nut == 'YY'] = 1

if 'nw_bld_tx' in DF_MEDCIN:
    DF_MEDCIN.nw_bld_tx.loc[DF_MEDCIN.nw_bld_tx == 'N'] = 0
    DF_MEDCIN.nw_bld_tx.loc[DF_MEDCIN.nw_bld_tx == 'NN'] = 0
    DF_MEDCIN.nw_bld_tx.loc[DF_MEDCIN.nw_bld_tx == 'Y'] = 1
    DF_MEDCIN.nw_bld_tx.loc[DF_MEDCIN.nw_bld_tx == 'YY'] = 1
    DF_MEDCIN.nw_bld_tx.loc[DF_MEDCIN.nw_bld_tx == 'None'] = 2

if 'ovs_class_btb' in DF_MEDCIN:
    DF_MEDCIN.ovs_class_btb.loc[DF_MEDCIN.ovs_class_btb == 'N'] = 0
    DF_MEDCIN.ovs_class_btb.loc[DF_MEDCIN.ovs_class_btb == 'NN'] = 0
    DF_MEDCIN.ovs_class_btb.loc[DF_MEDCIN.ovs_class_btb == 'Y'] = 1
    DF_MEDCIN.ovs_class_btb.loc[DF_MEDCIN.ovs_class_btb == 'YY'] = 1
    DF_MEDCIN.ovs_class_btb.loc[DF_MEDCIN.ovs_class_btb == 'None'] = 2

if 'nw_breastfeeding' in DF_MEDCIN:
    DF_MEDCIN.nw_breastfeeding.loc[DF_MEDCIN.nw_breastfeeding == 'N'] = 0
    DF_MEDCIN.nw_breastfeeding.loc[DF_MEDCIN.nw_breastfeeding == 'NN'] = 0
    DF_MEDCIN.nw_breastfeeding.loc[DF_MEDCIN.nw_breastfeeding == 'Y'] = 1
    DF_MEDCIN.nw_breastfeeding.loc[DF_MEDCIN.nw_breastfeeding == 'YY'] = 1

if 'ovs_mntl_hlth' in DF_MEDCIN:
    DF_MEDCIN.ovs_mntl_hlth.loc[DF_MEDCIN.ovs_mntl_hlth == 'N'] = 0
    DF_MEDCIN.ovs_mntl_hlth.loc[DF_MEDCIN.ovs_mntl_hlth == 'NN'] = 0
    DF_MEDCIN.ovs_mntl_hlth.loc[DF_MEDCIN.ovs_mntl_hlth == 'Y'] = 1
    DF_MEDCIN.ovs_mntl_hlth.loc[DF_MEDCIN.ovs_mntl_hlth == 'YY'] = 1
    DF_MEDCIN.ovs_mntl_hlth.loc[DF_MEDCIN.ovs_mntl_hlth == 'None'] = 2

if 'seh_exprcd_torture_yn' in DF_MEDCIN:
    DF_MEDCIN.seh_exprcd_torture_yn.loc[DF_MEDCIN.seh_exprcd_torture_yn == 'N'] = 0
    DF_MEDCIN.seh_exprcd_torture_yn.loc[DF_MEDCIN.seh_exprcd_torture_yn == 'NN'] = 0
    DF_MEDCIN.seh_exprcd_torture_yn.loc[DF_MEDCIN.seh_exprcd_torture_yn == 'Y'] = 1
    DF_MEDCIN.seh_exprcd_torture_yn.loc[DF_MEDCIN.seh_exprcd_torture_yn == 'YY'] = 1

if 'seh_wtnss_torture_yn' in DF_MEDCIN:
    DF_MEDCIN.seh_wtnss_torture_yn.loc[DF_MEDCIN.seh_wtnss_torture_yn == 'N'] = 0
    DF_MEDCIN.seh_wtnss_torture_yn.loc[DF_MEDCIN.seh_wtnss_torture_yn == 'NN'] = 0
    DF_MEDCIN.seh_wtnss_torture_yn.loc[DF_MEDCIN.seh_wtnss_torture_yn == 'Y'] = 1
    DF_MEDCIN.seh_wtnss_torture_yn.loc[DF_MEDCIN.seh_wtnss_torture_yn == 'YY'] = 1

if 'ovs_malaria' in DF_MEDCIN:
    DF_MEDCIN.ovs_malaria.loc[DF_MEDCIN.ovs_malaria == 'N'] = 0
    DF_MEDCIN.ovs_malaria.loc[DF_MEDCIN.ovs_malaria == 'NN'] = 0
    DF_MEDCIN.ovs_malaria.loc[DF_MEDCIN.ovs_malaria == 'Y'] = 1
    DF_MEDCIN.ovs_malaria.loc[DF_MEDCIN.ovs_malaria == 'YY'] = 1
    DF_MEDCIN.ovs_malaria.loc[DF_MEDCIN.ovs_malaria == 'None'] = 2
    DF_MEDCIN.ovs_malaria.loc[DF_MEDCIN.ovs_malaria == 'Not Indicated'] = 3
    DF_MEDCIN.ovs_malaria.loc[DF_MEDCIN.ovs_malaria == 'No documentation but patient reports treatment'] = 4

if 'ovs_intstnl_parasites' in DF_MEDCIN:
    DF_MEDCIN.ovs_intstnl_parasites.loc[DF_MEDCIN.ovs_intstnl_parasites == 'N'] = 0
    DF_MEDCIN.ovs_intstnl_parasites.loc[DF_MEDCIN.ovs_intstnl_parasites == 'NN'] = 0
    DF_MEDCIN.ovs_intstnl_parasites.loc[DF_MEDCIN.ovs_intstnl_parasites == 'Y'] = 1
    DF_MEDCIN.ovs_intstnl_parasites.loc[DF_MEDCIN.ovs_intstnl_parasites == 'YY'] = 1
    DF_MEDCIN.ovs_intstnl_parasites.loc[DF_MEDCIN.ovs_intstnl_parasites == 'None'] = 3
    DF_MEDCIN.ovs_intstnl_parasites.loc[DF_MEDCIN.ovs_intstnl_parasites \
    == 'No documentation but patient reports treatment'] = 4

if 'seh_rglr_meds_vit_yn' in DF_MEDCIN:
    DF_MEDCIN.seh_rglr_meds_vit_yn.loc[DF_MEDCIN.seh_rglr_meds_vit_yn == 'N'] = 0
    DF_MEDCIN.seh_rglr_meds_vit_yn.loc[DF_MEDCIN.seh_rglr_meds_vit_yn == 'NN'] = 0
    DF_MEDCIN.seh_rglr_meds_vit_yn.loc[DF_MEDCIN.seh_rglr_meds_vit_yn == 'Y'] = 1
    DF_MEDCIN.seh_rglr_meds_vit_yn.loc[DF_MEDCIN.seh_rglr_meds_vit_yn == 'YY'] = 1

if 'secondary_migrant' in DF_MEDCIN:
    DF_MEDCIN.secondary_migrant.loc[DF_MEDCIN.secondary_migrant == 'N'] = 0
    DF_MEDCIN.secondary_migrant.loc[DF_MEDCIN.secondary_migrant == 'NN'] = 0
    DF_MEDCIN.secondary_migrant.loc[DF_MEDCIN.secondary_migrant == 'Y'] = 1
    DF_MEDCIN.secondary_migrant.loc[DF_MEDCIN.secondary_migrant == 'YY'] = 1

if 'nw_strt_drugs_yn' in DF_MEDCIN:
    DF_MEDCIN.nw_strt_drugs_yn.loc[DF_MEDCIN.nw_strt_drugs_yn == 'N'] = 0
    DF_MEDCIN.nw_strt_drugs_yn.loc[DF_MEDCIN.nw_strt_drugs_yn == 'NN'] = 0
    DF_MEDCIN.nw_strt_drugs_yn.loc[DF_MEDCIN.nw_strt_drugs_yn == 'Y'] = 1
    DF_MEDCIN.nw_strt_drugs_yn.loc[DF_MEDCIN.nw_strt_drugs_yn == 'YY'] = 1

if 'nw_tattoo' in DF_MEDCIN:
    DF_MEDCIN.nw_tattoo.loc[DF_MEDCIN.nw_tattoo == 'N'] = 0
    DF_MEDCIN.nw_tattoo.loc[DF_MEDCIN.nw_tattoo == 'NN'] = 0
    DF_MEDCIN.nw_tattoo.loc[DF_MEDCIN.nw_tattoo == 'Y'] = 1
    DF_MEDCIN.nw_tattoo.loc[DF_MEDCIN.nw_tattoo == 'YY'] = 1

if 'seh_faith_yn' in DF_MEDCIN:
    DF_MEDCIN.seh_faith_yn.loc[DF_MEDCIN.seh_faith_yn == 'N'] = 0
    DF_MEDCIN.seh_faith_yn.loc[DF_MEDCIN.seh_faith_yn == 'NN'] = 0
    DF_MEDCIN.seh_faith_yn.loc[DF_MEDCIN.seh_faith_yn == 'Y'] = 1
    DF_MEDCIN.seh_faith_yn.loc[DF_MEDCIN.seh_faith_yn == 'YY'] = 1

if 'nw_injctd_drgs' in DF_MEDCIN:
    DF_MEDCIN.nw_injctd_drgs.loc[DF_MEDCIN.nw_injctd_drgs == 'N'] = 0
    DF_MEDCIN.nw_injctd_drgs.loc[DF_MEDCIN.nw_injctd_drgs == 'NN'] = 0
    DF_MEDCIN.nw_injctd_drgs.loc[DF_MEDCIN.nw_injctd_drgs == 'Y'] = 1
    DF_MEDCIN.nw_injctd_drgs.loc[DF_MEDCIN.nw_injctd_drgs == 'YY'] = 1

if 'nw_ethoh_yn' in DF_MEDCIN:
    DF_MEDCIN.nw_ethoh_yn.loc[DF_MEDCIN.nw_ethoh_yn == 'N'] = 0
    DF_MEDCIN.nw_ethoh_yn.loc[DF_MEDCIN.nw_ethoh_yn == 'NN'] = 0
    DF_MEDCIN.nw_ethoh_yn.loc[DF_MEDCIN.nw_ethoh_yn == 'Y'] = 1
    DF_MEDCIN.nw_ethoh_yn.loc[DF_MEDCIN.nw_ethoh_yn == 'YY'] = 1

if 'nw_chews_yn' in DF_MEDCIN:
    DF_MEDCIN.nw_chews_yn.loc[DF_MEDCIN.nw_chews_yn == 'N'] = 0
    DF_MEDCIN.nw_chews_yn.loc[DF_MEDCIN.nw_chews_yn == 'NN'] = 0
    DF_MEDCIN.nw_chews_yn.loc[DF_MEDCIN.nw_chews_yn == 'Y'] = 1
    DF_MEDCIN.nw_chews_yn.loc[DF_MEDCIN.nw_chews_yn == 'YY'] = 1

if 'nw_smokes_yn' in DF_MEDCIN:
    DF_MEDCIN.nw_smokes_yn.loc[DF_MEDCIN.nw_smokes_yn == 'N'] = 0
    DF_MEDCIN.nw_smokes_yn.loc[DF_MEDCIN.nw_smokes_yn == 'NN'] = 0
    DF_MEDCIN.nw_smokes_yn.loc[DF_MEDCIN.nw_smokes_yn == 'Y'] = 1
    DF_MEDCIN.nw_smokes_yn.loc[DF_MEDCIN.nw_smokes_yn == 'YY'] = 1

if 'nw_pregnant' in DF_MEDCIN:
    DF_MEDCIN.nw_pregnant.loc[DF_MEDCIN.nw_pregnant == 'N'] = 0
    DF_MEDCIN.nw_pregnant.loc[DF_MEDCIN.nw_pregnant == 'NN'] = 0
    DF_MEDCIN.nw_pregnant.loc[DF_MEDCIN.nw_pregnant == 'Y'] = 1
    DF_MEDCIN.nw_pregnant.loc[DF_MEDCIN.nw_pregnant == 'YY'] = 1

if 'nw_sexually_act' in DF_MEDCIN:
    DF_MEDCIN.nw_sexually_act.loc[DF_MEDCIN.nw_sexually_act == 'N'] = 0
    DF_MEDCIN.nw_sexually_act.loc[DF_MEDCIN.nw_sexually_act == 'NN'] = 0
    DF_MEDCIN.nw_sexually_act.loc[DF_MEDCIN.nw_sexually_act == 'Y'] = 1
    DF_MEDCIN.nw_sexually_act.loc[DF_MEDCIN.nw_sexually_act == 'YY'] = 1
                                 
if 'nw_ethoh_yn' in DF_MEDCIN:
    DF_MEDCIN.nw_ethoh_yn.loc[DF_MEDCIN.nw_ethoh_yn == 'NY'] = ''
    
#%% - Load the Orders tab, drop unecessary fields, filter out irrelevant data, remove duplicates
DF_D = XL_FILE.parse(sheetname=4)
DF_D.drop(['Order Code'], inplace=True, axis=1)

DF_D = DF_D[DF_D['Order Description'].str.contains('ASSAY') == False]
DF_D = DF_D[DF_D['Order Description'].str.contains('PURE') == False]
DF_D = DF_D[DF_D['Order Description'].str.contains('RPR') == False]
DF_D = DF_D[DF_D['Order Description'].str.contains('VARICELLA') == False]
DF_D = DF_D[DF_D['Order Description'].str.contains('RHS-15') == False]

DF_D = DF_D.drop_duplicates()

#%%  - Create a boolean column for patient screening check
DF_D['Screened?'] = DF_D['Order Description'].isnull() == False
DF_D['Screened?'] = DF_D['Screened?'].astype(int)

#%% - Create Orders dataframe, pivot and index data
""" 
            Order Description reflects whether or not patient was screened (y/n bool)
            Has patient been tested?   0 = No | 1 = Yes 
"""
DF_ORD1 = DF_D.pivot_table(index='Patient #', columns='Order Description', values='Screened?')
DF_ORD1.fillna(value=0, inplace=True)
DF_ORD1 = DF_ORD1.astype(int)

#%% - Concat Hep B types into one column and drop original segmented fields
DF_ORD1.lab_hepb_scrn = DF_ORD1['HEPATITIS B SURFACE ANTIGEN (HBsAG)'].map(int) + \
DF_ORD1['HEP B SURFACE ANTIBODY'] + DF_ORD1['HEPATITIS B CORE AB TOTAL']

# - Sometimes excel files contain multiple field names for single fields and they must be joined
if 'HEPATITIS B SURFACE ANTIGEN (HBSAG)' in DF_ORD1: 
    DF_ORD1.lab_hepb_scrn = DF_ORD1['HEPATITIS B SURFACE ANTIGEN (HBSAG)'].map(int) + DF_ORD1.lab_hepb_scrn
    DF_ORD1.drop(['HEPATITIS B SURFACE ANTIGEN (HBSAG)'], inplace=True, axis=1)
    
DF_ORD1.lab_hepb_scrn.replace(2, 1, inplace=True) 
DF_ORD1.lab_hepb_scrn.replace(3, 1, inplace=True) 

DF_ORD1.drop(['HEP B SURFACE ANTIBODY'], inplace=True, axis=1)
DF_ORD1.drop(['HEPATITIS B CORE AB TOTAL'], inplace=True, axis=1)
DF_ORD1.drop(['HEPATITIS B SURFACE ANTIGEN (HBsAG)'], inplace=True, axis=1)

DF_ORD1.rename(columns={'CBC': 'lab_cbc_scrnd', 'CMP': 'alr_cmp_scrnd', \
'Ova and Parasites, Stool Conc/Perm Smear, 2 spec': 'ips_scrnd', 'TB AG RESPONSE T-CELL SUSP': 'lab_tb_test_type', \
'URINALYSIS, AUTO, W/O SCOPE': 'lab_ua_scrnd', \
'VISUAL ACUITY SCREEN': 'vsd1_vsn_scrnd'}, inplace=True)

# - To account for unnecessary CAPS LOCKing of the ips screening field
if 'OVA AND PARASITES, STOOL CONC/PERM SMEAR, 2 SPEC' in DF_ORD1:
    DF_ORD1.rename(columns={'OVA AND PARASITES, STOOL CONC/PERM SMEAR, 2 SPEC': 'ips_scrnd'}, inplace=True)

# - Replace TB column values with integers corresponging to test types
if 'lab_tb_test_type' in DF_ORD1:
    DF_ORD1.lab_tb_test_type.loc[DF_ORD1.lab_tb_test_type == 'TSPOT'] = 1
    DF_ORD1.lab_tb_test_type.loc[DF_ORD1.lab_tb_test_type == 'TST'] = 2
    DF_ORD1.lab_tb_test_type.loc[DF_ORD1.lab_tb_test_type == 'Not Done'] = 3
                               
if 'COMPREHENSIVE METABOLIC PANEL W/O EGFR' in DF_ORD1:
    DF_ORD1.drop(['COMPREHENSIVE METABOLIC PANEL W/O EGFR'], inplace=True, axis=1)

if 'LIPID PANEL' in DF_ORD1:
    DF_ORD1.drop(['LIPID PANEL'], inplace=True, axis=1)
    
if 'SYPHILIS TEST, NON-TREP, QUALITATIVE' in DF_ORD1:
    DF_ORD1.drop(['SYPHILIS TEST, NON-TREP, QUALITATIVE'], inplace=True, axis=1)
    
if 'URINE PREGNANCY TEST' in DF_ORD1:
    DF_ORD1.drop(['URINE PREGNANCY TEST'], inplace=True, axis=1)

#%% - Filter data
DF_D = DF_D[DF_D['Result Component'].str.contains('ABSOLUTE') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('MCH') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('MCHC') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('MPV') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('BASOPHILS') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('COMMENT') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('LYMPHOCYTES') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('MONOCYTES') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('MYELOCYTES') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('PROMYELOCYTES') == False]

DF_D = DF_D[DF_D['Result Component'].str.contains('eGFR') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('Specific Gravity') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('ESTIMATION') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('CONFIRMATION') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('BUN') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('Nitrite') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('NUCLEATED') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('GLOBULIN') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('ALKALINE PHOSPHATASE') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('NEUTROPHILS') == False]

DF_D = DF_D[DF_D['Result Component'].str.contains('CARBON DIOXIDE') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('MORPHOLOGY') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('CHOL/HDLC') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('NON-HDL CHOLESTEROL') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('CONCENTRATE') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('TRICHROME') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('TSPOT') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('Bilirubin') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('BLASTS') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('EOSINOPHILS') == False]

DF_D = DF_D[DF_D['Result Component'].str.contains('Ketones') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('Leukocytes') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('Urobilinogen') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('pH') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('RED') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('QUESTION') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('CONTAINER') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('RESOLUTION') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('MESSAGE:') == False]
DF_D = DF_D[DF_D['Result Component'].str.contains('FECAL') == False]


#%% Create Result Component dataframe (Orders 2)
DF_ORD2 = DF_D.pivot_table(index='Patient #', columns='Result Component', values='Result', aggfunc='first')

#%%
DF_ORD2.rename(columns={'ALBUMIN': 'alr_cmp_albumin', 'ALT': 'alr_cmp_alt', 'AST': 'alr_cmp_ast', \
'Blood': 'lab_ua_blood', 'BILIRUBIN, TOTAL': 'alr_cmp_bilirubin', 'CALCIUM': 'alr_cmp_ca', \
'Both Eyes': 'vsd1_vision_both', 'CHLORIDE': 'alr_cmp_cl', 'CHOLESTEROL, TOTAL': 'lab_cholesterol_rslt', \
'CREATININE': 'alr_cmp_creatinine', 'GLUCOSE': 'alr_cmp_glucose', 'Glucose': 'lab_ua_glucose', \
'HDL CHOLESTEROL': 'lab_hdl_rslt', 'HEMATOCRIT': 'lab_hematocrit', 'HEMOGLOBIN': 'lab_hemoglobin', \
'HEPATITIS B CORE AB TOTAL': 'lab_hbcab', 'HEPATITIS B SURFACE ANTIBODY QL': 'lab_hbsab', \
'HEPATITIS B SURFACE$ANTIGEN': 'lab_hbsag', 'LDL-CHOLESTEROL': 'alr_cmp_ldl', 'Left Eye': 'vsd1_vision_left', \
'MCV': 'lab_mcv', 'RDW': 'lab_rdw', 'PLATELET COUNT': 'lab_platelet', 'POTASSIUM': 'alr_cmp_k', 'RHS 15 Score': 'rhs15_score', \
'PROTEIN, TOTAL': 'alr_cmp_ttlprotein', 'Protein': 'lab_ua_protein', 'Right Eye': 'vsd1_vision_right', \
'SODIUM': 'alr_cmp_na', 'TRIGLYCERIDES': 'alr_cmp_tryglycde', 'URINE PREGNANCY TEST': 'lab_pregnant_rslt', \
'WHITE BLOOD CELL COUNT': 'lab_wbc', 'RPR (DX) W/REFL TITER AND CONFIRMATORY TESTING': 'lab_syphilis_rslts'}, inplace=True) 
 
#%% - Assign REDCap dropdown values based on cell contents 
# - This section of the code can call for occasional updating due to the wide variance of columns values
# - seen in the Orders sheet. Bluegrass' fields often contains typos, some of which can be new to the script.

#%%

# DF_ORD2.loc[DF_ORD2.loc.contains(r'[Nn]egative') == True] = 0

if 'lab_ua_blood' in DF_ORD2:
    DF_ORD2.lab_ua_blood = DF_ORD2.lab_ua_blood.str.strip() #strip any leading and trailing whitespace
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'Not Performed'] = ''
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'None'] = 0
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'negative'] = 0
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'Negative'] = 0
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'NEGATIVE'] = 0
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'neg'] = 0
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'small'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'SMALL'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'trace'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'trace-lysed'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'trace-lysd'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'tace-lysed'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'Trace'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'Trace-intact'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'trace-intact'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'TRACE-INTACT'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'moderate'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'Moderate'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood == 'large'] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood.str.contains('RBC/uL', na=False)] = 1
    DF_ORD2.lab_ua_blood.loc[DF_ORD2.lab_ua_blood.str.contains('rbc/ul', na=False)] = 1

if 'lab_ua_glucose' in DF_ORD2:
    DF_ORD2.lab_ua_glucose = DF_ORD2.lab_ua_glucose.str.strip()
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'Not Performed'] = ''
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'None'] = ''
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'negative'] = 0
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'negatie'] = 0
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'negatve'] = 0
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'neg'] = 0
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'eneg'] = 0
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'Negative'] = 0
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'NEGATIVE'] = 0
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'Moderate'] = 1
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'Trace'] = 1
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'positive'] = 1
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == 'Positive'] = 1
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == '100'] = 1
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == '5.5'] = 1
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose == '250'] = 1
    DF_ORD2.lab_ua_glucose.loc[DF_ORD2.lab_ua_glucose.str.contains('mg/dl', na=False)] = 1

if 'lab_ua_protein' in DF_ORD2:
    DF_ORD2.lab_ua_protein = DF_ORD2.lab_ua_protein.str.strip()
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'Not Performed'] = ''
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'None'] = ''
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'negative'] = 0
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'negarive'] = 0
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'Negative'] = 0
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'NEGATIVE'] = 0
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'neg'] = 0
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'Neg'] = 0
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'positive'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'Positive'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'trace'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'traace'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'Trace'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == 'TRACE'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == '6.0'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == '100'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein == '30'] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein.str.contains('mg', na=False)] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein.str.contains('mg/dl', na=False)] = 1
    DF_ORD2.lab_ua_protein.loc[DF_ORD2.lab_ua_protein.str.contains('300', na=False)] = 1

if 'lab_pregnant_rslt' in DF_ORD2:
    DF_ORD2.lab_pregnant_rslt = DF_ORD2.lab_pregnant_rslt.str.strip()
    DF_ORD2.lab_pregnant_rslt.loc[DF_ORD2.lab_pregnant_rslt == 'Not Performed'] = ''
    DF_ORD2.lab_pregnant_rslt.loc[DF_ORD2.lab_pregnant_rslt == 'Negative'] = 0
    DF_ORD2.lab_pregnant_rslt.loc[DF_ORD2.lab_pregnant_rslt == 'Positive'] = 1

if 'lab_hbcab' in DF_ORD2:
    DF_ORD2.lab_hbcab = DF_ORD2.lab_hbcab.str.strip()
    DF_ORD2.lab_hbcab.loc[DF_ORD2.lab_hbcab == 'Not Performed'] = ''
    DF_ORD2.lab_hbcab.loc[DF_ORD2.lab_hbcab == 'NON-REACTIVE'] = 0
    DF_ORD2.lab_hbcab.loc[DF_ORD2.lab_hbcab == 'REACTIVE'] = 1
    DF_ORD2.lab_hbcab.loc[DF_ORD2.lab_hbcab == 'BORDERLINE'] = 2

if 'lab_hbsab' in DF_ORD2:
    DF_ORD2.lab_hbsab = DF_ORD2.lab_hbsab.str.strip()
    DF_ORD2.lab_hbsab.loc[DF_ORD2.lab_hbsab == 'Not Performed'] = ''
    DF_ORD2.lab_hbsab.loc[DF_ORD2.lab_hbsab == 'NON-REACTIVE'] = 0
    DF_ORD2.lab_hbsab.loc[DF_ORD2.lab_hbsab == 'REACTIVE'] = 1
    DF_ORD2.lab_hbsab.loc[DF_ORD2.lab_hbsab == 'BORDERLINE'] = 2

if 'lab_hbsag' in DF_ORD2:
    DF_ORD2.lab_hbsag = DF_ORD2.lab_hbsag.str.strip()
    DF_ORD2.lab_hbsag.loc[DF_ORD2.lab_hbsag == 'Not Performed'] = ''
    DF_ORD2.lab_hbsag.loc[DF_ORD2.lab_hbsag == 'NON-REACTIVE'] = 0
    DF_ORD2.lab_hbsag.loc[DF_ORD2.lab_hbsag == 'REACTIVE'] = 1
    DF_ORD2.lab_hbsag.loc[DF_ORD2.lab_hbsag == 'BORDERLINE'] = 2

if 'lab_syphilis_rslts' in DF_ORD2:             
    DF_ORD2.lab_syphilis_rslts = DF_ORD2.lab_syphilis_rslts.str.strip()
    DF_ORD2.lab_syphilis_rslts.loc[DF_ORD2.lab_syphilis_rslts == 'NON-REACTIVE'] = 0
    DF_ORD2.lab_syphilis_rslts.loc[DF_ORD2.lab_syphilis_rslts == 'REACTIVE'] = 1

if 'lab_hematocrit' in DF_ORD2:
    DF_ORD2.lab_hematocrit = DF_ORD2.lab_hematocrit.str.strip()
    DF_ORD2.lab_hemoglobin = DF_ORD2.lab_hemoglobin.str.strip()
    DF_ORD2.lab_mcv = DF_ORD2.lab_mcv.str.strip()
    DF_ORD2.lab_platelet = DF_ORD2.lab_platelet.str.strip()
    DF_ORD2.lab_rdw = DF_ORD2.lab_rdw.str.strip()
    DF_ORD2.lab_hematocrit.loc[DF_ORD2.lab_hematocrit == 'Not Performed'] = ''
    DF_ORD2.lab_hemoglobin.loc[DF_ORD2.lab_hemoglobin == 'Not Performed'] = ''
    DF_ORD2.lab_mcv.loc[DF_ORD2.lab_mcv == 'Not Performed'] = ''
    DF_ORD2.lab_platelet.loc[DF_ORD2.lab_platelet == 'Not Performed'] = ''
    DF_ORD2.lab_rdw.loc[DF_ORD2.lab_rdw == 'Not Performed'] = ''
    DF_ORD2.lab_wbc.loc[DF_ORD2.lab_wbc == 'TNP'] = ''

if 'vsd1_vision_both' in DF_ORD2:
    DF_ORD2.vsd1_vision_both = '20/' + DF_ORD2.vsd1_vision_right

#%% - Merge both frames into DF_ORDERS and sort alphabetically
DF_ORDERS = DF_ORD1.join(DF_ORD2, how='outer')
DF_ORDERS.sort_index(axis=1, ascending=True, inplace=True)

#%% - Load the Immunizations tab, drop unecessary fields
DF_F = XL_FILE.parse(sheetname=8)
DF_F.drop(['Code'], inplace=True, axis=1)
DF_F.drop(['Date Ordered'], inplace=True, axis=1)
DF_F = DF_F[DF_F.Description.str.contains('IMMUNIZATION') == False]

#%% - Create Immun dataframe, pivot and index Immunizations data
""" 
             Immunization descriptions are columnized
             Values reflect vaccines given:
                 0 = not given | 1 = given
"""
DF_F['Vac Given?'] = DF_F.Description.isnull() == False
DF_F['Vac Given?'] = DF_F['Vac Given?'].astype(int) 
DF_F['Description'] = DF_F['Description'].str.strip() #strip leading and trailing whitespace to prevent potential errors
DF_IMMUN = DF_F.pivot_table(index='Patient #', columns = 'Description', values = 'Vac Given?', aggfunc=sum)
DF_IMMUN.fillna(value='0', inplace=True)

#%% - Merge corresponding vaccine columns together, change to int, rename to REDcap's format. Drop the old ones and sort
# - Varicella
if 'CHICKEN POX VACCINE, SC (VARIVAX)' in DF_IMMUN:
    DF_IMMUN.imm_varicella = DF_IMMUN['CHICKEN POX VACCINE, SC (VARIVAX)'].astype(int)
    DF_IMMUN.drop(['CHICKEN POX VACCINE, SC (VARIVAX)'], inplace=True, axis=1)    
if 'CHICKEN POX (VFC) VACCINE, SC' in DF_IMMUN:
    try:
        DF_IMMUN.imm_varicella = DF_IMMUN['CHICKEN POX (VFC) VACCINE, SC'].map(int) + DF_IMMUN.imm_varicella
        DF_IMMUN.drop(['CHICKEN POX (VFC) VACCINE, SC'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_varicella = DF_IMMUN['CHICKEN POX (VFC) VACCINE, SC'].astype(int)
        DF_IMMUN.drop(['CHICKEN POX (VFC) VACCINE, SC'], inplace=True, axis=1)    
if 'MMRV (VFC) VACCINE, SC' in DF_IMMUN: 
    try:
        DF_IMMUN.imm_varicella = DF_IMMUN['MMRV (VFC) VACCINE, SC'].map(int) + DF_IMMUN.imm_varicella
    except:
        DF_IMMUN.imm_varicella = DF_IMMUN['MMRV (VFC) VACCINE, SC'].astype(int)

#%% - DTAP
if 'DTAP (VFC) VACCINE, < 7 YRS, IM' in DF_IMMUN:
    DF_IMMUN.imm_dtap_dtp_dose1 = DF_IMMUN['DTAP (VFC) VACCINE, < 7 YRS, IM'].astype(int)
    DF_IMMUN.drop(['DTAP (VFC) VACCINE, < 7 YRS, IM'], inplace=True, axis=1)    
if 'DTAP-HEP B-IPV (VFC) VACCINE, IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_dtap_dtp_dose1 = DF_IMMUN['DTAP-HEP B-IPV (VFC) VACCINE, IM'].map(int) + DF_IMMUN.imm_dtap_dtp_dose1 
    except:
        DF_IMMUN.imm_dtap_dtp_dose1 = DF_IMMUN['DTAP-HEP B-IPV (VFC) VACCINE, IM'].astype(int)    
if 'DTAP-HIB-IPV (VFC) VACCINE, IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_dtap_dtp_dose1 = DF_IMMUN['DTAP-HIB-IPV (VFC) VACCINE, IM'].map(int) + DF_IMMUN.imm_dtap_dtp_dose1
    except:
        DF_IMMUN.imm_dtap_dtp_dose1 = DF_IMMUN['DTAP-HIB-IPV (VFC) VACCINE, IM'].map(int) 
if 'DTAP-IPV (VFC) VACC 4-6 YR IM' in DF_IMMUN:
    DF_IMMUN.imm_dtap_dtp_dose1 = DF_IMMUN['DTAP-IPV (VFC) VACC 4-6 YR IM'].map(int) + DF_IMMUN.imm_dtap_dtp_dose1

#%% - Influenza 
if 'FLU VACC 4 VAL 3 YRS PLUS IM' in DF_IMMUN:
    DF_IMMUN['imm_flu'] = DF_IMMUN['FLU VACC 4 VAL 3 YRS PLUS IM'].astype(int)
    DF_IMMUN.drop(['FLU VACC 4 VAL 3 YRS PLUS IM'], inplace=True, axis=1)    
if'FLU VACCINE 6-35 MO (VFC), TRIVALENT, PRESERVATIVE FREE ' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLU VACCINE 6-35 MO (VFC), TRIVALENT, PRESERVATIVE FREE '].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLU VACCINE 6-35 MO (VFC), TRIVALENT, PRESERVATIVE FREE '], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLU VACCINE 6-35 MO (VFC), TRIVALENT, PRESERVATIVE FREE '].astype(int)
        DF_IMMUN.drop(['FLU VACCINE 6-35 MO (VFC), TRIVALENT, PRESERVATIVE FREE '], inplace=True, axis=1)
if 'FLUMIST (VFC) VACCINE, NASAL VFC ' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUMIST (VFC) VACCINE, NASAL VFC '].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLUMIST (VFC) VACCINE, NASAL VFC '], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUMIST (VFC) VACCINE, NASAL VFC '].astype(int)
        DF_IMMUN.drop(['FLUMIST (VFC) VACCINE, NASAL VFC '], inplace=True, axis=1)
if 'FLUARIX .5 ML SYRINGE 3 YRS+' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUARIX .5 ML SYRINGE 3 YRS+'].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLUARIX .5 ML SYRINGE 3 YRS+'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUARIX .5 ML SYRINGE 3 YRS+'].astype(int)
        DF_IMMUN.drop(['FLUARIX .5 ML SYRINGE 3 YRS+'], inplace=True, axis=1)
if 'FLUVIRIN .5 ML SYRINGE 4+YRS' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUVIRIN .5 ML SYRINGE 4+YRS'].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLUVIRIN .5 ML SYRINGE 4+YRS'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUVIRIN .5 ML SYRINGE 4+YRS']
        DF_IMMUN.drop(['FLUVIRIN .5 ML SYRINGE 4+YRS'], inplace=True, axis=1)
if 'FLUZONE .25 ML SYRINGE 6-35 MONTHS' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUZONE .25 ML SYRINGE 6-35 MONTHS'].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLUZONE .25 ML SYRINGE 6-35 MONTHS'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUZONE .25 ML SYRINGE 6-35 MONTHS']
        DF_IMMUN.drop(['FLUZONE .25 ML SYRINGE 6-35 MONTHS'], inplace=True, axis=1)
if 'FLUZONE (VFC) 4 VAL 3 YRS+' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUZONE (VFC) 4 VAL 3 YRS+'].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLUZONE (VFC) 4 VAL 3 YRS+'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUZONE (VFC) 4 VAL 3 YRS+']
        DF_IMMUN.drop(['FLUZONE (VFC) 4 VAL 3 YRS+'], inplace=True, axis=1)
if 'FLUZONE 3 YRS+ 0.5 ML SYRINGE' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUZONE 3 YRS+ 0.5 ML SYRINGE'].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLUZONE 3 YRS+ 0.5 ML SYRINGE'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLUZONE 3 YRS+ 0.5 ML SYRINGE']
        DF_IMMUN.drop(['FLUZONE 3 YRS+ 0.5 ML SYRINGE'], inplace=True, axis=1)
if 'FLULAVAL (IIV4) 0.5 ML SYRINGE' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLULAVAL (IIV4) 0.5 ML SYRINGE'].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['FLULAVAL (IIV4) 0.5 ML SYRINGE'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['FLULAVAL (IIV4) 0.5 ML SYRINGE']
        DF_IMMUN.drop(['FLULAVAL (IIV4) 0.5 ML SYRINGE'], inplace=True, axis=1)
if 'IIV4 VACCINE 3 YRS PLUS IM' in DF_IMMUN:
    try:
        DF_IMMUN['imm_flu'] = DF_IMMUN['IIV4 VACCINE 3 YRS PLUS IM'].map(int) + DF_IMMUN['imm_flu']
        DF_IMMUN.drop(['IIV4 VACCINE 3 YRS PLUS IM'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_flu'] = DF_IMMUN['IIV4 VACCINE 3 YRS PLUS IM']
        DF_IMMUN.drop(['IIV4 VACCINE 3 YRS PLUS IM'], inplace=True, axis=1)

#%% - Hep A
if 'HEP A (VFC) VACC, PED/ADOL, 2 DOSE' in DF_IMMUN:
    DF_IMMUN.oi_hepa = DF_IMMUN['HEP A (VFC) VACC, PED/ADOL, 2 DOSE'].astype(int)
    DF_IMMUN.drop(['HEP A (VFC) VACC, PED/ADOL, 2 DOSE'], inplace=True, axis=1)

#%% - Hep B
if 'HEP B VACCINE, ADULT, IM' in DF_IMMUN:
    DF_IMMUN.imm_hepb_d1 = DF_IMMUN['HEP B VACCINE, ADULT, IM'].astype(int)
    DF_IMMUN.drop(['HEP B VACCINE, ADULT, IM'], inplace=True, axis=1)
if 'HEPB (VFC) VACC PED/ADOL 3 DOSE IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_hepb_d1 = DF_IMMUN['HEPB (VFC) VACC PED/ADOL 3 DOSE IM'].map(int) + DF_IMMUN.imm_hepb_d1
        DF_IMMUN.drop(['HEPB (VFC) VACC PED/ADOL 3 DOSE IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_hepb_d1 = DF_IMMUN['HEPB (VFC) VACC PED/ADOL 3 DOSE IM'].astype(int)
        DF_IMMUN.drop(['HEPB (VFC) VACC PED/ADOL 3 DOSE IM'], inplace=True, axis=1)
if 'DTAP-HEP B-IPV (VFC) VACCINE, IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_hepb_d1 = DF_IMMUN['DTAP-HEP B-IPV (VFC) VACCINE, IM'].map(int) + DF_IMMUN.imm_hepb_d1 
        DF_IMMUN.drop(['DTAP-HEP B-IPV (VFC) VACCINE, IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_hepb_d1 = DF['DTAP-HEP B-IPV (VFC) VACCINE, IM'].astype(int)
        DF_IMMUN.drop(['DTAP-HEP B-IPV (VFC) VACCINE, IM'], inplace=True, axis=1)
if 'HEPB VACC PED/ADOL 3 DOSE IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_hepb_d1 = DF_IMMUN['HEPB VACC PED/ADOL 3 DOSE IM'].map(int) + DF_IMMUN.imm_hepb_d1
        DF_IMMUN.drop(['HEPB VACC PED/ADOL 3 DOSE IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_hepb_d1 = DF['HEPB VACC PED/ADOL 3 DOSE IM'].astype(int)
        DF_IMMUN.drop(['HEPB VACC PED/ADOL 3 DOSE IM'], inplace=True, axis=1)

#%% - Hib
if 'HIB (VFC) VACCINE, PRP-T, IM' in DF_IMMUN:
    DF_IMMUN.oi_hib = DF_IMMUN['HIB (VFC) VACCINE, PRP-T, IM'].astype(int)
    DF_IMMUN.drop(['HIB (VFC) VACCINE, PRP-T, IM'], inplace=True, axis=1)

#%% - HPV
if 'HPV VACCINE 9 VALENT IM' in DF_IMMUN:
    DF_IMMUN.oi_hpv = DF_IMMUN['HPV VACCINE 9 VALENT IM'].astype(int)
    DF_IMMUN.drop(['HPV VACCINE 9 VALENT IM'], inplace=True, axis=1)

#%% - Meningoccal
if 'MENINGOCOCCAL (VFC) VACCINE, IM' in DF_IMMUN:
    DF_IMMUN.imm_mening = DF_IMMUN['MENINGOCOCCAL (VFC) VACCINE, IM'].astype(int)
    DF_IMMUN.drop(['MENINGOCOCCAL (VFC) VACCINE, IM'], inplace=True, axis=1)    
if 'MENINGOCOCCAL GRP B VFC (10-25 YRS)' in DF_IMMUN:
    try:
        DF_IMMUN.imm_mening = DF_IMMUN['MENINGOCOCCAL GRP B VFC (10-25 YRS)'].map(int) + DF_IMMUN.imm_mening
        DF_IMMUN.drop(['MENINGOCOCCAL GRP B VFC (10-25 YRS)'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_mening = DF_IMMUN['MENINGOCOCCAL GRP B VFC (10-25 YRS)'].astype(int)
        DF_IMMUN.drop(['MENINGOCOCCAL GRP B VFC (10-25 YRS)'], inplace=True, axis=1)

#%% - MMR
if 'MMR (VFC) VACCINE, SC' in DF_IMMUN:
    DF_IMMUN['imm_mmr'] = DF_IMMUN['MMR (VFC) VACCINE, SC'].astype(int)
    DF_IMMUN.drop(['MMR (VFC) VACCINE, SC'], inplace=True, axis=1)
if 'MMR VACCINE, SC' in DF_IMMUN:
    try:
        DF_IMMUN['imm_mmr'] = DF_IMMUN['MMR VACCINE, SC'].map(int) + DF_IMMUN['imm_mmr']
        DF_IMMUN.drop(['MMR VACCINE, SC'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_mmr'] = DF_IMMUN['MMR VACCINE, SC'].astype(int)
        DF_IMMUN.drop(['MMR VACCINE, SC'], inplace=True, axis=1)    
if 'MMRV (VFC) VACCINE, SC' in DF_IMMUN:
    try:
        DF_IMMUN['imm_mmr'] = DF_IMMUN['MMRV (VFC) VACCINE, SC'].map(int) + DF_IMMUN['imm_mmr']
        DF_IMMUN.drop(['MMRV (VFC) VACCINE, SC'], inplace=True, axis=1)
    except:
        DF_IMMUN['imm_mmr'] = DF_IMMUN['MMRV (VFC) VACCINE, SC'].astype(int)
        DF_IMMUN.drop(['MMRV (VFC) VACCINE, SC'], inplace=True, axis=1)
        
#%% - Pneumococcal
if 'PNEUMOCOCCAL (VFC) VACC 13 VAL IM VFC' in DF_IMMUN:
    DF_IMMUN.imm_pneumo = DF_IMMUN['PNEUMOCOCCAL (VFC) VACC 13 VAL IM VFC'].astype(int)
    DF_IMMUN.drop(['PNEUMOCOCCAL (VFC) VACC 13 VAL IM VFC'], inplace=True, axis=1)
if 'PNEUMOCOCCAL VACC 13 VAL IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_pneumo = DF_IMMUN['PNEUMOCOCCAL VACC 13 VAL IM'].astype(int)
        DF_IMMUN.drop(['PNEUMOCOCCAL VACC 13 VAL IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_pneumo = DF_IMMUN['PNEUMOCOCCAL VACC 13 VAL IM'].map(int) + DF_IMMUN.imm_pneumo
        DF_IMMUN.drop(['PNEUMOCOCCAL VACC 13 VAL IM'], inplace=True, axis=1)
if 'PNEUMO VACC 23 VAL IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_pneumo = DF_IMMUN['PNEUMO VACC 23 VAL IM'].astype(int)
        DF_IMMUN.drop(['PNEUMO VACC 23 VAL IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_pneumo = DF_IMMUN['PNEUMO VACC 23 VAL IM'].map(int) + DF_IMMUN.imm_pneumo
        DF_IMMUN.drop(['PNEUMO VACC 23 VAL IM'], inplace=True, axis=1)
        
#%% - Polio
if 'POLIOVIRUS (VFC), IPV, SC/IM' in DF_IMMUN:
    DF_IMMUN.imm_polio_dose1 = DF_IMMUN['POLIOVIRUS (VFC), IPV, SC/IM'].astype(int)
    DF_IMMUN.drop(['POLIOVIRUS (VFC), IPV, SC/IM'], inplace=True, axis=1)   
if 'DTAP-HEP B-IPV (VFC) VACCINE, IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_polio_dose1 = DF_IMMUN['DTAP-HEP B-IPV (VFC) VACCINE, IM'].map(int) + DF_IMMUN.imm_polio_dose1
        DF_IMMUN.drop(['DTAP-HEP B-IPV (VFC) VACCINE, IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_polio_dose1 = DF_IMMUN['DTAP-HEP B-IPV (VFC) VACCINE, IM'].astype(int)
        DF_IMMUN.drop(['DTAP-HEP B-IPV (VFC) VACCINE, IM'], inplace=True, axis=1)
if 'DTAP-HIB-IPV (VFC) VACCINE, IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_polio_dose1 = DF_IMMUN['DTAP-HIB-IPV (VFC) VACCINE, IM'].map(int) + DF_IMMUN.imm_dtap_dtp_dose1
        DF_IMMUN.drop(['DTAP-HIB-IPV (VFC) VACCINE, IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_polio_dose1 = DF_IMMUN['DTAP-HIB-IPV (VFC) VACCINE, IM'].astype(int)
        DF_IMMUN.drop(['DTAP-HIB-IPV (VFC) VACCINE, IM'], inplace=True, axis=1)
if 'DTAP-IPV (VFC) VACC 4-6 YR IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_polio_dose1 = DF_IMMUN['DTAP-IPV (VFC) VACC 4-6 YR IM'].map(int) + DF_IMMUN.imm_dtap_dtp_dose1
        DF_IMMUN.drop(['DTAP-IPV (VFC) VACC 4-6 YR IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_polio_dose1 = DF_IMMUN['DTAP-IPV (VFC) VACC 4-6 YR IM'].astype(int)
        DF_IMMUN.drop(['DTAP-IPV (VFC) VACC 4-6 YR IM'], inplace=True, axis=1)
   
#%% - TDAP
if 'TDAP (VFC) VACCINE >7 IM' in DF_IMMUN:
    DF_IMMUN.imm_tdap = DF_IMMUN['TDAP (VFC) VACCINE >7 IM'].astype(int)
    DF_IMMUN.drop(['TDAP (VFC) VACCINE >7 IM'], inplace=True, axis=1)
if 'TDAP VACCINE >7 IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_tdap = DF_IMMUN['TDAP VACCINE >7 IM'].map(int) + DF_IMMUN.imm_tdap
        DF_IMMUN.drop(['TDAP VACCINE >7 IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_tdap = DF_IMMUN['TDAP VACCINE >7 IM'].astype(int)
        DF_IMMUN.drop(['TDAP VACCINE >7 IM'], inplace=True, axis=1)
if 'TD (VFC) VACCINE NO PRSRV >/= 7 IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_tdap = DF_IMMUN['TD (VFC) VACCINE NO PRSRV >/= 7 IM'].map(int) + DF_IMMUN.imm_tdap
        DF_IMMUN.drop(['TD (VFC) VACCINE NO PRSRV >/= 7 IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_tdap = DF_IMMUN['TD (VFC) VACCINE NO PRSRV >/= 7 IM'].astype(int)
        DF_IMMUN.drop(['TD (VFC) VACCINE NO PRSRV >/= 7 IM'], inplace=True, axis=1)   
if 'TD VACCINE NO PRSRV >/= 7 IM' in DF_IMMUN:
    try:
        DF_IMMUN.imm_tdap = DF_IMMUN['TD VACCINE NO PRSRV >/= 7 IM'].map(int) + DF_IMMUN.imm_tdap
        DF_IMMUN.drop(['TD VACCINE NO PRSRV >/= 7 IM'], inplace=True, axis=1)
    except:
        DF_IMMUN.imm_tdap = DF_IMMUN['TD VACCINE NO PRSRV >/= 7 IM'].astype(int)
        DF_IMMUN.drop(['TD VACCINE NO PRSRV >/= 7 IM'], inplace=True, axis=1)

#%% - Drop any remaining irrelevant columns and sort 
if 'Immunizations Delinquent' in DF_IMMUN:
    DF_IMMUN.drop(['Immunizations Delinquent'], inplace=True, axis=1)
if 'Immunizations Reviewed And Current' in DF_IMMUN:
    DF_IMMUN.drop(['Immunizations Reviewed And Current'], inplace=True, axis=1)
if 'Immunization Record Unavailable' in DF_IMMUN:
    DF_IMMUN.drop(['Immunization Record Unavailable'], inplace=True, axis=1)
DF_IMMUN.sort_index(axis=1, ascending=True, inplace=True)

#%% - Move everything into the RESULT dataframe
RESULT = RESULT.join(DF_VITALS, how='outer')
RESULT = RESULT.join(DF_ORDERS, how='outer')
RESULT = RESULT.join(DF_MEDCIN, how='outer')
RESULT = RESULT.join(DF_IMMUN, how='outer')

#%% - Lab fields -- min/max value range
if 'lab_platelet' in RESULT:
    RESULT.lab_platelet.loc[RESULT.lab_platelet >= 450.1] = ''
    RESULT.lab_platelet.loc[RESULT.lab_platelet <= 99.9] = ''
if 'lab_hematocrit' in RESULT:
    RESULT.lab_hematocrit.loc[RESULT.lab_hematocrit >= 54.1] = ''
    RESULT.lab_hematocrit.loc[RESULT.lab_hematocrit <= 24.9] = ''
if 'lab_hemoglobin' in RESULT:
    RESULT.lab_hemoglobin.loc[RESULT.lab_hemoglobin >= 18.1] = ''                   
    RESULT.lab_hemoglobin.loc[RESULT.lab_hemoglobin <= 9.9] = ''
if 'lab_cholesterol_rslt' in RESULT:
    RESULT.lab_cholesterol_rslt.loc[RESULT.lab_cholesterol_rslt >= 300.1] = ''
    RESULT.lab_cholesterol_rslt.loc[RESULT.lab_cholesterol_rslt <= 99.9] = ''
if 'lab_wbc' in RESULT:
    RESULT.lab_wbc.loc[RESULT.lab_wbc >= 14.1] = ''
    RESULT.lab_wbc.loc[RESULT.lab_wbc <= 2.9] = ''
if 'lab_mcv' in RESULT:
    RESULT.lab_mcv.loc[RESULT.lab_mcv >= 100.1] = ''
    RESULT.lab_mcv.loc[RESULT.lab_mcv <= 49.9] = ''
if 'lab_rdw' in RESULT:
    RESULT.lab_rdw.loc[RESULT.lab_rdw >= 20.1] = ''
    RESULT.lab_rdw.loc[RESULT.lab_rdw <= 10.9] = ''

#%% - Remove timestamps and NA values from Date fields
RESULT.us_arrival_date = pd.to_datetime(RESULT.us_arrival_date, errors='coerce')
RESULT.date_of_birth = pd.to_datetime(RESULT.date_of_birth, errors='coerce')
RESULT.date_of_birth = RESULT.date_of_birth.dt.date
RESULT.us_arrival_date = RESULT.us_arrival_date.dt.date

#%% - drop records where alien_no or Patient # = N/A
RESULT = RESULT[pd.notnull(RESULT['alien_no'])]
RESULT = RESULT.set_index('alien_no', drop=False)
RESULT = RESULT[RESULT.alien_no.str.contains('000') == False]

#%% - reformat demographics fields
RESULT.vsd1_height = RESULT.vsd1_height.round(2)
RESULT.vsd1_weight = RESULT.vsd1_weight.round(2)
 
#%% - By the end of the file, all NA fields should be empty in order for REDcap to accept the data
RESULT.fillna(value='', axis=1, inplace=True)

#%% - Perform the upload operation - Step one: create csv for use in REDCap import
path = ('C:\\Users\\japese01\Documents\\RefugeeHealth\\uploads\\uploads\\'+dr+'\\')   # insert the directory where you would like the output file to be created
RESULT.to_csv(path+'refHealthUpload_'+input_file_date+'.csv', index=False, date_format='%Y-%m-%d')

#%% - Alternatively, perform upload in one step using REDCap API (requires error-less dataset)
# REDCap import method will catch errors and allow you to make changes to the csv. 
"""
import pycurl, cStringIO
buf = cStringIO.StringIO()
data = {
    'token': '',         #insert API token here
    'content': 'project',
    'format': 'json',
    'returnFormat': 'json'
}
ch = pycurl.Curl()
ch.setopt(ch.URL, 'https://refugeehealth.louisville.edu/api/')
ch.setopt(ch.HTTPPOST, data.items())
ch.setopt(ch.WRITEFUNCTION, buf.write)
ch.perform()
ch.close()
print buf.getvalue()
buf.close()
"""
