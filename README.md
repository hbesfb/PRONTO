# PRONTO
**(rePort geneRator fOr iNpred Tumor bOards)**

This is a tool used to filter and analysis data from TSO500 results and generate the report based on the results data and template file.        
This repository contains the configure, scripts, report template, file infrastructure and docker files to run this tool locally or in Docker/Singularity.

## Contents

1. [Requirments for running PRONTO locally](#requirments-for-running-pronto-locally)
2. [Repository contents](#repository-contents)
3. [Run PRONTO locally](#run-pronto-locally)
4. [Example commands](#example-commands)
5. [PRONTO Docker](#pronto-docker)
6. [PRONTO Singularity](#pronto-singularity)
7. [ChangeLog](#changelog)

## Requirments for running PRONTO locally

1. This tool needs to be run with python above version 3.
2. Module xlutils and pptx need to be installed for python:                                                                                                              
	- To install module xlutils, run the command:
	 
		```
		sudo pip3 install xlutils
		```
	                                                                    
	- To install module python-pptx, run the command: 

		```
		sudo pip3 install python-pptx                        
		```
3. Module xlrd and xlwt are needed for MTF importing function:                          
	- Download the files from Internet. Please google search it and choose the correct version for your environment.                   
	- Under the downloaded folder and install the module with command: 
 
		```
		sudo python setup.py install
		```

## Repository contents

| INPUT file name: | Details: |
|:---|:---|
| `Script/PRONTO.py`   | The executable python script.|
| `Config/configure_PRONTO.ini` | The configure file. Needs to be modified prior to its first use.|
| `In/Templates/MTB_template.pptx` | The template file used for generating PP report. (TODO: explain PP) |
| `In/InPreD_PRONTO_metadata.txt` | The clinical data file. Reports will be generated for the `Sample_id` for which the `Create_report` value is set to `Y` in this file. |
| `In/MTF/IPD-XXXX_Material Transit Form InPreD NGS.xlsx` | The material file contains all patient personal information. (Used by OUS) This file will generate the inpred-samle-id following the nomenclature file. | 

</br>

| OUTPUT file/folder name: | Details: |
|:---|:---|
| `Out/$runID/IPZXXXX` | The folder contains all results for sample `IPZXXXX` from sequencing run $runID. |
| `Out/$runID/IPZXXXX/extra_files` | The folder contains filter tables during the calculation process, and the patient material file from lab. |
| `Out/$runID/IPZXXXX/IPZXXXX_MTB_report.pptx` | The PP report file. (TODO: explain PP) |
| `Out/$runID/IPZXXXX/IPZXXXX_Remisse_draft.docx` | The remise draft file for email. (Used by OUS) |
| `Out/InPreD_PRONTO_metadata_tsoppi.txt`	| The file contains clinical data and the TSOPPI results for all sample reports. |

</br>

| Files located in the `Testing_data` folder : | Details: |
|:---|:---|
| `$testRunID_TSO_500_LocalApp_postprocessing_results.zip` | The testing data from AcroMetrix sample TSOPPI results which only contains the files PRONTO needs. Move this folder into your local TSOPPI result path for testing. |
| `InPreD_PRONTO_metadata.txt` | The file contains clinical data of AcroMetrix samples for testing. Move this file into `In` folder of this repository for testing. |
| `$testRunID.zip` | The testing results from AcroMetrix sample TSOPPI results for your local comparisons. |
| `testRunID="191206_NB501498_0174_AHWCNMBGXC"` ||

## Run PRONTO locally

### Adapt the config file:

- In `Config/configure_PRONTO.ini`, please specify your InPreD node with `inpred_node = `. This will appear in the header of the reports.
- In `Config/configure_PRONTO.ini`, please specify the local dataset file path of TSOPPI results with `data_path = `.                         

### Type clinical data into `In/InPreD_PRONTO_metadata.txt`:

Manually write the clinical data into file `In/InPreD_PRONTO_metadata.txt`. Reports will be generated for the `Sample_id` for which the `Create_report` value is set to `Y` in this file.

### Run `Script/PRONTO.py:` 

Please run the script with `-h` or `--help` to print the usage information.                                                                                                       
This script is a tool used to generate the paitent report based on the TSO500 analysis results and the personal intomation from the clinical data in `In/InPreD_PRONTO_metadata.txt`, and update the TSOPPI results into the file `Out/InPreD_PRONTO_metadata_tsoppi.txt` when the reports are generated.          
This script could also fill the patient personal information into the clinical data file with the MTF files under the foder `In/MTF/`. (This fuction currently is only used by OUS)                                                                
To run this script tool in your system with python3, it will read the clinical data from `In/InPreD_PRONTO_metadata.txt` and generate reports for the `Sample_id` with the `Create_report` set to `Y`.

Afterwards, execute the command as follows: 

```
python3 Script/PRONTO.py 
```


## Example commands

### Print the usage information:

```
python3 Script/PRONTO.py -h
```

### Execute the report generating process:

```
python3 Script/PRONTO.py
```

After executing the script and generating all the required reports, update the `Create_report` value to `N` for every sample for which the report should not be re-generated in the future.

### Special commands used by OUS:

```
python3 Script/PRONTO.py -r <TSO500_runID> -D <DNA_sampleID> -c
python3 Script/PRONTO.py -m
```
- -c, --clinical_file: Fill the patient personal information into file InPreD_PRONTO_metadata.txt with the MTF files under the foder In/MTF/.
- -m, --mail_draft: Generate the Remisse_draft.docx file with report.



## PRONTO Docker

PRONTO docker image is automatically pushed to dockerhub: https://hub.docker.com/r/inpred/pronto.

### Download the image with the latest tag:

```
docker pull inpred/pronto:latest
```

### Modify content of the config file:

- Specify your InPreD node name as the value of the `inpred_node` parameter in the `Config/configure_PRONTO.ini` file.
- Keep the `data_path` parameter's value set to `/pronto/tsoppi_data/` in the same config file.


### Run PRONTO with docker image:

```   
sudo docker run \
	--rm -it \
	-v $tsoppi_data:/pronto/tsoppi_data \
	-v $InPreD_PRONTO_metadata_file:/pronto/In/InPreD_PRONTO_metadata.txt \
	-v $pronto_output_dir:/pronto/Out \
	inpred/pronto:latest \
	python /pronto/Script/PRONTO.py
``` 

- `$tsoppi_data` is the path of your local TSOPPI results, which contains all runs of TSOPPI data (not the folder for individual runs).
- `$InPreD_PRONTO_metadata_file` is your local InPreD meta data file which contains clinical data for the samples.
- `$pronto_output_dir` is the path in your local environment to store the reports generated by PRONTO.


## PRONTO Singularity

PRONTO docker image is automatically pushed to dockerhub: https://hub.docker.com/r/inpred/pronto and can be also pulled as a singularity image.


### Download PRONTO singularity image with the latest tag:

```
singularity pull PRONTO_singularity_image.sif docker://inpred/pronto:latest
```


### Modify content of the config file:

- Specify your InPreD node name as the value of the `inpred_node` parameter in the `Config/configure_PRONTO.ini` file.
- Keep the `data_path` parameter's value set to `/pronto/tsoppi_data/` in the same config file.


### Run PRONTO image with Singularity:

```
singularity exec \
	--no-home \
	-B $tsoppi_data:/pronto/tsoppi_data \
	-B $InPreD_PRONTO_metadata_file:/pronto/In/InPreD_PRONTO_metadata.txt \
	-B $pronto_output_dir:/pronto/Out \
	-W $SINGULARITY_TMP \
	$dir/PRONTO_singularity_image.sif \
	python /pronto/Script/PRONTO.py
```

- `$tsoppi_data` is the path of your local TSOPPI results, which contains all runs of TSOPPI data (not the folder for individual runs).
- `$InPreD_PRONTO_metadata_file` is your local InPreD meta data file which contains clinical data for the samples.                                      
- `$pronto_output_dir` is the path in your local environment to store the reports generated by PRONTO.py.


## ChangeLog

### v1.1

- New Features: (TODO: fill in)
- Resolved Issues:
	[ #5](https://github.com/InPreD/PRONTO/issues/5) 
	[ #7](https://github.com/InPreD/PRONTO/issues/7)
	[ #8](https://github.com/InPreD/PRONTO/issues/8) 
	[#14](https://github.com/InPreD/PRONTO/issues/14) 
	[#17](https://github.com/InPreD/PRONTO/issues/17)
- Other Changes: (TODO: fill in)

### v1.0
- First tracked version.