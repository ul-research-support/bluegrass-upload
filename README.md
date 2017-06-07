# bluegrass-upload

BCHC Adult Refugee Health Clinical Assessment – Upload Script

***********************************************************************************************************************

 Project Overview

Program designed for automated entry of refugee health data received from Bluegrass Community Health Center. Data is processed and reformatted from excel files.

The script accomplishes this by:

1) reading the contents from the excel file into Python

2) creating modular dataframes from each sheet using PANDAS

3) filtering each excel sheet of unnecessary data

4) reformatting the relevant data in adherance with REDCap

5) merging the dataframes into a singular dataset (RESULT)

6) uploading the full dataset to REDCap

This upload is achieved using the REDCap API (a REDCap API token with import privileges is required to perform this operation) but it only works if there are zero errors in the BCHC dataset, which is rare. Alternatively, the easier upload method is to create a CSV and perform the upload operation as a manual data import. Performing a manual import allows you to view the data before uploading and check it against existing records to verify changes made. This process will let you clear up any errors within the CSV file itself. 

***********************************************************************************************************************

 Processing the Data
 
Read the excel file into Pandas by inserting its file path into the parentheses of pd.ExcelFile()

Execute the entire script or each block one-by-one as you go along, referring to comments in the code when needed. If a block fails to execute, then the excel file is likely contains a new data type.

Attribute errors can occur based on unpredictable variances from the Bluegrass file. If a code block is not executing, sometimes lines of code need to be added to adjust for previously unencountered fields. The console will display any lines with errors. Adjust the code accordingly and if necessary, comment out unneeded code for preservation.

Much of the data received from Bluegrass goes unrecorded into REDCap. Data that is filtered appears in the code as such:

                                [dataframe.str.contains(‘unused data’) == FALSE]

Anything present in the current excel tab that is not stored in REDCap is unneeded in the program. Refer to the data dictionary in REDCap for a better understanding of what data is to be collected.

If you find that you need to filter something else out, be careful to ensure that the same word isn’t being used in other (relevant) cells.

***********************************************************************************************************************

 Immunizations Troubleshooting 
 
Much of the variance you’ll find will likely be in the Immunizations component of the program. This has proven to be most often the case because different immunizations are administered each month.

Immunizations have their code isolated into codeblocks based immunization type (varicella, polio, TDAP, etc). This is to help simplify the process and increase readability.

Any previously unseen immunization will need to be incorporated into the script in order for the dataset to be accepted for upload. Incremental steps can be taken as time goes on in order to fully streamline this process, but it would require coming into contact with every possible vaccination so that nothing is left to surprise. For the time being, human judgment is necessary to account for any unpredictability.

***********************************************************************************************************************

 Prepping for Upload
 
At the end of the file, the modular dataframes are merged together and all NA values/timestamps are removed. Take time to look over and confirm that the RESULT dataframe is consistent with REDCap, and then perform the preferred upload operation.

If the RESULT dataframe contains values or columns that are not accepted in REDCap's codebook, you can handle this either in the code or in the resulting CSV file. Typically this means cleaning up typos that have gone unaccounted for in certain fields, for example, "negtve" in place of "negative," which might fail to convert to 0 unless a line of code was written in advance to handle that particular typo in that particular column. Unfortunately issues like this are quite common in BCHC files, which is why such a large section of the script has been written to deal with filtering typos and unecessary data.

Similarly, some columns may be present in the RESULT dataframe that do not belong in REDCap. Ensure that every column is REDCap-consistent by checking the data dictionary. Usually it will be easy to spot a column that does not belong in the final dataset because it will be written in all caps.

***********************************************************************************************************************

 Upload Operation 

    Using CSV (preferred method): Insert your filepath into designated line for where you want the file to be stored in your directory. 
    In the following line, you can name your upload file. After the file is created, use the Data Import Tool 
    on REDCap to upload the dataset. If there are any errors in the dataframe, the import tool will tell you where. 
    From here, you can fix these errors in the CSV directly and then perform the upload. 

    Using REDCap API: Insert your API token into the designated line and run the API. As long as you have 
    the required privileges for using the API and there are no errors in the RESULT dataframe, 
    the upload should be successful. 

