# !/bin/bash
# Build the docker image first, or use command to load the PRONTO docker image: sudo docker load -i PRONTO_docker/PRONTO_v1_docker_image.tar
echo -n "´Please enter your local path for TSOPPI results:"
read tsoppi_data
pronto_dir=$(dirname "$PWD")
tag="docker_pronto:v1"
sudo docker run --rm -it -v $tsoppi_data:/pronto/tsoppi_data -v $pronto_dir/Config:/pronto/Config -v $pronto_dir/In:/pronto/In -v $pronto_dir/Out:/pronto/Out -v $pronto_dir/Script:/pronto/Script $tag
