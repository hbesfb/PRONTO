PRONTO -  rePort geneRator fOr iNpred Tumor bOards

Script/PRONTO.py
This script is a tool used to generate the paitent report based on the TSO500 analysis results and the personal intomation from the clinical data in In/InPreD_PRONTO_metadata.txt, and update the SOPPI results into the file Out/InPreD_PRONTO_metadata_tsoppi.txt when the reports are generated.
This script could also fill the patient personal information into the clinical data file with the MTF files under the foder In/MTF/. (This fuction currently is only used by OUS)
To run this script tool in your computer with python3, it will read the clinical data from In/InPreD_PRONTO_metadata.txt and generate reports for the Sample_id with Create_report==Y.

INPUT:
Script/PRONTO.py		   			-  The python script.
Config/configure_PRONTO.ini				-  The configure file.
In/Templates/MTB_template.pptx		-  The template file used for generating PP report.
In/InPreD_PRONTO_metadata.txt				-  The clinical data file. Reports will be generated for the Sample_id with Create_report==Y in this file. 
In/MTF/IPD-XXXX_Material Transit Form InPreD NGS.xlsx	-  The material file contains all patient personal information. (only OUS)

OUTPUT:
Out/$runID/IPDXXX					- The folder contains all results for this sample.
Out/$runID/IPDXXX/extra_files				- The folder contains filter tables during the canculation process, and the patient material file from lab.
Out/$runID/IPDXXX/IPDXXX_MTB_report.pptx		- The PP report file.
Out/$runID/IPDXXX/IPDXXX_Remisse_draft.docx		- The remisse draft file for email. (only OUS)
Out/InPreD_PRONTO_metadata_tsoppi.txt			- The file contains clinical data and the SOPPI results for all sample reports.

LOCAL CONFIGURE:
In Config/configure_PRONTO.ini, please specify your InPreD node with "inpred_node = ". This will apprear in the header of the reports.
In Config/configure_PRONTO.ini, please specify the local dataset file path of TSOPPI results with "data_path =".

Example commands:
python3 Script/InPreD_PRONTO.py

(Please remember to update the "Create_report" to "N" in file In/InPreD_PRONTO_metadata.txt manually after the report generation is finished!)

Special commands used by OUS:
python3 Script/InPreD_PRONTO.py -r <TSO500_runID> -D <DNA_sampleID> -c
python3 Script/InPreD_PRONTO.py -m

Pre-work:
1)This script needs to be run with python above version 3.
2)Module xlutils and pptx need to be insalled for python:
To install module xlutils, run the command: sudo pip install xlutils
To install module python-pptx, run the command: sudp pip3 install python-pptx
3)Module xlrd and xlwt are neede for MTF importing function:
Download the file from Internet. Please google search it and choose the correct version for your environment.
Under the downloaded folder and install the module with command: sudo python setup.py install

PRONTO Docker:
Script/PRONTO_docker_build.sh
This script is used to build a PRONTO docker image "ipred/pronto:v1".

Script/PRONTO_docker_run.sh
This script is used to run a container once with command "python Script/InPreD_PRONTO.py" based on the docker image "inpred/pronto:v1". This script will ask you to put into your local "TSOPPI results" path, and the reports will be exported to "Out/$runID/IPDXXX" under your "PRONTO_report" folder.

PRONTO Singularity: Download PRONTO singularity image or build it based on the PRONTO docker image in your local system.
Script/PRONTO_singularity_run.sh
This script is used to execute with command "python Script/InPreD_PRONTO.py" based on the PRONTO Singularity image. This script will ask you to put into your local "TSOPPI results" path, and the reports will be exported to "Out/$runID/IPDXXX" under your "PRONTO_report" folder.
