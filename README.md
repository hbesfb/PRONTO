# PRONTO
rePort geneRator fOr iNpred Tumor bOards    
This is a tool used to filter and analysis data from TSO500 results. And generate the report based on the results data and template file.        
This repository contains the configure, scripts, report template, file infrastructure and docker files to run this tool locally or in Docker/Singularity.

# Requirments
1. This tool needs to be run with python above version 3.
2. Module xlutils and pptx need to be insalled for python:                                                                                                       
To install module xlutils, run the command: [sudo pip install xlutils]                                                                      
To install module python-pptx, run the command: [sudp pip3 install python-pptx]                         
3. Module xlrd and xlwt are neede for MTF importing function:                          
Download the file from Internet. Please google search it and choose the correct version for your environment.                   
Under the downloaded folder and install the module with command: [sudo python setup.py install]

# File infrastructure
INPUT:                                      
Script/PRONTO.py                                      ---  The python script.                                                  
Config/configure_PRONTO.ini                           ---  The configure file. Local changes is needed to set up.                                        
In/Templates/MTB_template.pptx                        ---  The template file used for generating PP report.                                                      
In/InPreD_PRONTO_metadata.txt                         ---  The clinical data file. Reports will be generated for the Sample_id with "Create_report==Y" in this file. 
In/MTF/IPD-XXXX_Material Transit Form InPreD NGS.xlsx	---  The material file contains all patient personal information. (Used by OUS) This file will generate the inpred-samle-id following the nomenclature file.                          
                            
OUTPUT:                  
Out/$runID/IPDXXX					                          --- The folder contains all results for this sample.                       
Out/$runID/IPDXXX/extra_files				                --- The folder contains filter tables during the canculation process, and the patient material file from lab.
Out/$runID/IPDXXX/IPDXXX_MTB_report.pptx		        --- The PP report file.                       
Out/$runID/IPDXXX/IPDXXX_Remisse_draft.docx		      --- The remisse draft file for email. (Used by OUS)                         
Out/InPreD_PRONTO_metadata_tsoppi.txt			          --- The file contains clinical data and the SOPPI results for all sample reports.              
             
TESTING:                                                            
Testing_data/191206_NB501498_0174_AHWCNMBGXC_TSO_500_LocalApp_postprocessing_results.zip --- The testing data from AcroMetrix sample TSOPPI results which only contains the files PRONTO needs. Move this filder into your local TSOPPI result path for testing.                                                      
Testing_data/InPreD_PRONTO_metadata.txt                                                  --- The file contains clinical data of AcroMetrix samples for testing. Move this file into In/ for testing.                                                      
Testing_data/191206_NB501498_0174_AHWCNMBGXC.zip                                         --- The testing results from AcroMetrix sample TSOPPI results for your local comparisons.                                                                   

# Process of local running
1. LOCAL CONFIGURE:                       
In Config/configure_PRONTO.ini, please specify your InPreD node with "inpred_node = ". This will apprear in the header of the reports.               
In Config/configure_PRONTO.ini, please specify the local dataset file path of TSOPPI results with "data_path =".                         
2. Type clinical data into In/InPreD_PRONTO_metadata.txt                       
Manually write the clinical data into file In/InPreD_PRONTO_metadata.txt. Reports will be generated for the Sample_id with "Create_report==Y" in this file.     
3. Run Script/PRONTO.py                                                          
Please run the script with "-h" or "--help" to print the usage information.                                                                                                       
This script is a tool used to generate the paitent report based on the TSO500 analysis results and the personal intomation from the clinical data in In/InPreD_PRONTO_metadata.txt, and update the SOPPI results into the file Out/InPreD_PRONTO_metadata_tsoppi.txt when the reports are generated.          
This script could also fill the patient personal information into the clinical data file with the MTF files under the foder In/MTF/. (This fuction currently is only used by OUS)                                                                
To run this script tool in your system with python3, it will read the clinical data from In/InPreD_PRONTO_metadata.txt and generate reports for the Sample_id with "Create_report==Y".                             

# Example commands
[python3 Script/InPreD_PRONTO.py -h]
Print the usage information of this tool.             

[python3 Script/InPreD_PRONTO.py]                                           
Please remember to update the "Create_report" to "N" in file In/InPreD_PRONTO_metadata.txt manually after the report generation is finished!             

Special commands used by OUS:                                                
[python3 Script/InPreD_PRONTO.py -r <TSO500_runID> -D <DNA_sampleID> -c]                                                 
[python3 Script/InPreD_PRONTO.py -m]

# PRONTO Docker
PRONTO docker image is automatically pushed to dockerhub: https://hub.docker.com/r/inpred/pronto        
Please download the image with the latest tag: [docker pull inpred/pronto:latest]                 

Run PRONTO with docker image:                                                                        
[sudo docker run --rm -it -v $tsoppi_data:/pronto/tsoppi_data -v $InPreD_PRONTO_metadata_file:/pronto/In/InPreD_PRONTO_metadata.txt -v $pronto_output_dir:/pronto/Out inpred/pronto:latest python /pronto/Script/PRONTO.py]       
             
"$tsoppi_data" is the path of your local TSOPPI results, where contains all runs of TSOPPI data not the folder for individual runs.            
"$InPreD_PRONTO_metadata_file" is your local InPreD meta data file which contains clinical data for samples.                                      
"$pronto_output_dir" is the path in your local environment to store the reports generated by PRONTO.

NB:              
Running PRONTO with docker, you only need to modify the "inpred_node = " in Config/configure_PRONTO.ini file.            
Please keep the "data_path =/pronto/tsoppi_data" in Config/configure_PRONTO.ini.

# PRONTO Singularity
Download PRONTO docker image in the dockerhub with the latest tag: [docker pull inpred/pronto:latest]                                                         
Save the PRONTO docker image to a tar file: [docker save -o PRONTO_docker_image.tar inpred/pronto:latest]                                             
Generate PRONTO Singularity image by conversion from the docker image file. Example command:                                                   
[singularity build --disable-cache --tmpdir $SINGULARITY_TMP $dir/PRONTO_singularity_image.sif docker-archive://$dir/PRONTO_docker_image.tar]     

Run PRONTO image with Singularity:                                                                       
[singularity exec --no-home -B $tsoppi_data:/pronto/tsoppi_data -B $InPreD_PRONTO_metadata_file:/pronto/In/InPreD_PRONTO_metadata.txt -B $pronto_output_dir:/pronto/Out -W $SINGULARITY_TMP $dir/PRONTO_singularity_image.sif python /pronto/Script/PRONTO.py]                           

"$tsoppi_data" is the path of your local TSOPPI results, where contains all runs of TSOPPI data not the folder for individual runs.            
"$InPreD_PRONTO_metadata_file" is your local InPreD meta data file which contains clinical data for samples.                                      
"$pronto_output_dir" is the path in your local environment to store the reports generated by PRONTO.py.

NB:                   
Running PRONTO with singularity, you only need to modify the "inpred_node = " in Config/configure_PRONTO.ini file. Please keep the "data_path =/pronto/tsoppi_data" in Config/configure_PRONTO.ini.
