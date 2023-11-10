# !/bin/bash
echo -n "´Please enter your local path for TSOPPI results:"
read tsoppi_data
pronto_dir=$(dirname "$pwd")
singularity exec --no-home -B $tsoppi_data:/pronto/tsoppi_data -B $pronto_dir/Config:/pronto/Config -B $pronto_dir/In:/pronto/In -B $pronto_dir/Out:/pronto/Out -B $pronto_dir/Script:/pronto/Script $pronto_dir/PRONTO_image/PRONTO_singularity/PRONTO_v1_singularity_image.sif python /pronto/Script/PRONTO.py
