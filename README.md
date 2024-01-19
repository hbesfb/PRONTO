# PRONTO 
**(rePort geneRator fOr iNpred Tumor bOards)**

<br />

PRONTO is a tool used to filter and analyse data from TSO500 analysis (in the form of [TSOPPI](https://tsoppi.readthedocs.io/en/latest/) results). It generates powerpoint patient reports using predefined powerpoint template file. This repository contains the config file, metadata files, executable script, report template, testing data, and docker files to run this tool locally or in Docker/Singularity. The repository is accompanied with a docker image that is automatically pushed to [dockerhub](https://hub.docker.com/r/inpred/pronto) on each version release.

<br />

## Table of contents

1. [Requirments for running PRONTO locally](#requirments-for-running-pronto-locally)
2. [Repository contents](#repository-contents)
3. [Run PRONTO locally](#run-pronto-locally)
4. [Example commands](#example-commands)
5. [PRONTO Docker](#pronto-docker)
6. [PRONTO Singularity](#pronto-singularity)
7. [ChangeLog](#changelog)


<br />

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

<br />

## Repository contents

| INPUT file name: | Details: |
|:---|:---|
| `Script/PRONTO.py`   | The executable python script.|
| `Config/configure_PRONTO.ini` | The configure file. Needs to be modified prior to its first use.|
| `In/Templates/MTB_template.pptx` | The template file used for generating PP report. (TODO: explain PP) |
| `In/InPreD_PRONTO_metadata.txt` | The clinical data file. Reports will be generated for the `Sample_id` for which the `Create_report` value is set to `Y` in this file. |
| `In/MTF/IPD-XXXX_Material Transit Form InPreD NGS.xlsx` | The material file contains all patient personal information. (Used by OUS) This file will generate the inpred-samle-id following the nomenclature file. | 

<br />
<br />

| OUTPUT file/folder name: | Details: |
|:---|:---|
| `Out/$runID/IPZXXXX` | The folder contains all results for sample `IPZXXXX` from sequencing run $runID. |
| `Out/$runID/IPZXXXX/extra_files` | The folder contains filter tables during the calculation process, and the patient material file from lab. |
| `Out/$runID/IPZXXXX/IPZXXXX_MTB_report.pptx` | The PP report file. (TODO: explain PP) |
| `Out/$runID/IPZXXXX/IPZXXXX_Remisse_draft.docx` | The remise draft file for email. (Used by OUS) |
| `Out/InPreD_PRONTO_metadata_tsoppi.txt`	| The file contains clinical data and the TSOPPI results for all sample reports. |

<br />
<br />

| Files located in the `Testing_data` folder : | Details: |
|:---|:---|
| `testRunID="191206_NB501498_0174_AHWCNMBGXC"` | `testRunID` is used in the following lines of this table as a shortcut for the full ID of the sequencing run, which is `191206_NB501498_0174_AHWCNMBGXC`.|
| `$testRunID_TSO_500_LocalApp_postprocessing_results.zip` | The testing data from AcroMetrix sample TSOPPI results which only contains the files PRONTO needs. Move this folder into your local TSOPPI result path for testing. |
| `InPreD_PRONTO_metadata.txt` | The file contains clinical data of AcroMetrix samples for testing. Move this file into `In` folder of this repository for testing. |
| `$testRunID.zip` | The testing results from AcroMetrix sample TSOPPI results for your local comparisons. |



<br />

## Run PRONTO locally

### Adapt the config file:

- In `Config/configure_PRONTO.ini`, please specify your InPreD node by defining value of the `inpred_node` parameter. The node name will appear in the header of the reports.
- In `Config/configure_PRONTO.ini`, please specify the local dataset file path of TSOPPI results by defining value of the `data_path` parameter.

### Type clinical data into the input metadata file:

Manually write the clinical data into file `In/InPreD_PRONTO_metadata.txt`. Reports will be generated for the `Sample_id` for which the `Create_report` value is set to `Y` in this file.

### Run PRONTO: 

PRONTO takes in TSOPPI data, clinical info provided in `In/InPreD_PRONTO_metadata.txt`, and powerpoint template `In/Templates/MTB_template.pptx`. It generates a patient report for every `Sample_id` with the `Create_report` set to `Y` in the `Out` folder and updates the file `Out/InPreD_PRONTO_metadata_tsoppi.txt` with the TSOPPI results of the patients for which the reports were generated.


To print the usage information run the following command: 

```
python3 Script/PRONTO.py --help
```

To execute the command and generate the reports as defined in the input metadata file, run the command: 

```
python3 Script/PRONTO.py 
```


<br />

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


<br />

## PRONTO Docker

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


<br />

## PRONTO Singularity

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


<br />

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